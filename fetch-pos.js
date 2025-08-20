// fetch-pos.js (stealth + deep wait)
const { chromium } = require('playwright');
const fs = require('fs'), path = require('path');
const DEBUG_DIR = path.join(process.cwd(), 'debug'); fs.mkdirSync(DEBUG_DIR, { recursive: true });

function mask(s){ return s ? `(len:${String(s).length})` : '(missing)'; }

async function dump(page, name){
  await page.screenshot({ path: path.join(DEBUG_DIR, `${name}.png`), fullPage: true }).catch(()=>{});
  const html = await page.content().catch(()=> '');
  fs.writeFileSync(path.join(DEBUG_DIR, `${name}.html`), html);
  const inputs = await page.$$eval('input', els => els.map(e => ({
    type:e.type, name:e.name, id:e.id, ph:e.getAttribute('placeholder')||''
  })).slice(0,20)).catch(()=>[]);
  fs.writeFileSync(path.join(DEBUG_DIR, `${name}.inputs.json`), JSON.stringify(inputs,null,2));
  console.log(`[INFO] ${name} inputs:`, inputs);
}

(async ()=>{
  const {
    POS_EMAIL, POS_PASSWORD, GAS_WEBHOOK_URL, GAS_SHARED_SECRET,
    GENRE_ID='137', POS_BASE='https://pos.mycalinks.com', LOGIN_URL
  } = process.env;

  console.log('[ENV] POS_EMAIL=%s POS_PASSWORD=%s GAS_WEBHOOK_URL=%s GAS_SHARED_SECRET=%s GENRE_ID=%s POS_BASE=%s LOGIN_URL=%s',
    mask(POS_EMAIL), mask(POS_PASSWORD), mask(GAS_WEBHOOK_URL), mask(GAS_SHARED_SECRET), GENRE_ID, POS_BASE, LOGIN_URL||'(auto)');

  const browser = await chromium.launch({
    headless: true,
    args: ['--disable-blink-features=AutomationControlled']
  });

  const ctx = await browser.newContext({
    // “本物っぽい”環境
    userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36',
    locale: 'ja-JP',
    viewport: { width: 1366, height: 900 },
    javaScriptEnabled: true
  });

  // ステルス：webdriver等を隠す
  await ctx.addInitScript(() => {
    Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
    Object.defineProperty(Notification, 'permission', { get: () => 'denied' });
    // plugins と languages を偽装
    Object.defineProperty(navigator, 'languages', { get: () => ['ja-JP', 'ja'] });
    Object.defineProperty(navigator, 'plugins', { get: () => [1,2,3] });
  });

  const page = await ctx.newPage();
  try {
    // 1) ログインページ（候補を順に）
    const urls = [
      LOGIN_URL,
      `${POS_BASE}/auth/login`,
      `${POS_BASE}/login`,
      `${POS_BASE}/auth/signin`,
      `${POS_BASE}/auth`,
      POS_BASE
    ].filter(Boolean);

    const emailSels = [
      'input[type="email"]','input[name="email"]','input#email',
      'input[autocomplete="username"]','input[placeholder*="メール"]',
      'input[placeholder*="email" i]','input[name="username"]','input[name="loginId"]'
    ];
    const passSels  = [
      'input[type="password"]','input[name="password"]','input#password',
      'input[autocomplete="current-password"]','input[placeholder*="パスワード"]',
      'input[placeholder*="password" i]'
    ];
    const submitSels = [
      'button[type="submit"]','input[type="submit"]',
      'button:has-text("ログイン")','button:has-text("Sign in")','button:has-text("Sign In")'
    ];

    let loggedIn = false;

    for (const u of urls) {
      console.log('[STEP] goto', u);
      await page.goto(u, { waitUntil: 'domcontentloaded', timeout: 60000 });
      // JS 実行・hydrate 待ち
      await page.waitForLoadState('networkidle', { timeout: 20000 }).catch(()=>{});
      await page.waitForTimeout(1500);
      await dump(page, 'login-stage1');

      // “ログイン”ボタンがあれば押してモーダル表示
      const openBtn = await page.$('a:has-text("ログイン"), button:has-text("ログイン"), a:has-text("Sign in"), button:has-text("Sign in")');
      if (openBtn) { await openBtn.click().catch(()=>{}); await page.waitForTimeout(800); await dump(page, 'login-after-open'); }

      // Shadow DOM対策：全 input をJSで数える
      const count = await page.evaluate(() => document.querySelectorAll('input').length);
      console.log('[INFO] input count =', count);

      // 入力欄探索
      let emailBox=null, passBox=null;
      for (const s of emailSels){ emailBox = await page.$(s); if (emailBox) break; }
      for (const s of passSels){  passBox  = await page.$(s); if (passBox)  break; }

      if (emailBox && passBox) {
        await emailBox.fill(POS_EMAIL);
        await passBox.fill(POS_PASSWORD);

        let clicked=false;
        for (const s of submitSels) {
          const btn = await page.$(s);
          if (btn) {
            await Promise.all([
              page.waitForNavigation({ waitUntil: 'networkidle', timeout: 30000 }).catch(()=>{}),
              btn.click()
            ]);
            clicked = true; break;
          }
        }
        if (!clicked) {
          await passBox.press('Enter');
          await page.waitForLoadState('networkidle', { timeout: 15000 }).catch(()=>{});
        }

        const path = new URL(page.url()).pathname;
        if (!/\/auth\/?(login|signin)?$/i.test(path)) { loggedIn = true; break; }
      }
    }

    if (!loggedIn) {
      await dump(page, 'after-login-fail');
      console.error('Login failed (no inputs or blocked)');
      process.exit(1);
    }

    console.log('[INFO] logged in:', page.url());

    // 2) 対象ページへ
    const target = `${POS_BASE}/auth/item?genreId=${GENRE_ID}`;
    console.log('[STEP] goto target', target);
    await page.goto(target, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForLoadState('networkidle', { timeout: 20000 }).catch(()=>{});
    await page.waitForTimeout(1000);
    await dump(page, 'target');

    // 3) JSON API を捕捉
    console.log('[STEP] waitForResponse');
    const apiResp = await page.waitForResponse(r => {
      const ct = (r.headers()['content-type']||'').toLowerCase();
      const url = r.url();
      return r.status()===200 && ct.includes('application/json') && (url.includes('/api') || url.includes('genreId='));
    }, { timeout: 60000 });
    const raw = await apiResp.json();

    let items = Array.isArray(raw) ? raw : [];
    if (!items.length && raw && typeof raw==='object') {
      for (const k of ['data','items','records','result','payload']) {
        if (Array.isArray(raw[k])) { items = raw[k]; break; }
        if (raw[k] && typeof raw[k]==='object') {
          for (const k2 of ['data','items','records','result','payload']) {
            if (Array.isArray(raw[k][k2])) { items = raw[k][k2]; break; }
          }
          if (items.length) break;
        }
      }
    }
    if (!items.length) { console.error('No array in JSON. head=', JSON.stringify(raw).slice(0,300)); process.exit(2); }

    // 4) GASへPOST
    const res = await page.request.post(GAS_WEBHOOK_URL, {
      headers: { 'Content-Type':'application/json', 'Authorization': `Bearer ${GAS_SHARED_SECRET}` },
      data: { items }
    });
    console.log('GAS response:', await res.text());

  } finally {
    await ctx.close(); await browser.close();
  }
})();
