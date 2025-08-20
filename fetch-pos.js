// fetch-pos.debug.js
const { chromium, devices } = require('playwright');
const fs = require('fs');
const path = require('path');

const DEBUG_DIR = path.join(process.cwd(), 'debug');
fs.mkdirSync(DEBUG_DIR, { recursive: true });

function mask(s){ if(!s) return '(missing)'; return `(len:${String(s).length})`; }

async function main() {
  const env = {
    POS_EMAIL: process.env.POS_EMAIL,
    POS_PASSWORD: process.env.POS_PASSWORD,
    GAS_WEBHOOK_URL: process.env.GAS_WEBHOOK_URL,
    GAS_SHARED_SECRET: process.env.GAS_SHARED_SECRET,
    GENRE_ID: process.env.GENRE_ID || '137',
    POS_BASE: process.env.POS_BASE || 'https://pos.mycalinks.com',
  };

  console.log('[ENV] POS_EMAIL=%s POS_PASSWORD=%s GAS_WEBHOOK_URL=%s GAS_SHARED_SECRET=%s GENRE_ID=%s POS_BASE=%s',
    mask(env.POS_EMAIL), mask(env.POS_PASSWORD), mask(env.GAS_WEBHOOK_URL), mask(env.GAS_SHARED_SECRET),
    env.GENRE_ID, env.POS_BASE
  );

  let exitCode = 0;

  const browser = await chromium.launch({ headless: true });
  const ctx = await browser.newContext({
    userAgent: devices['Desktop Chrome'].userAgent,
    locale: 'ja-JP',
    recordHar: { path: path.join(DEBUG_DIR, 'network.har'), content: 'embed' }
  });
  const page = await ctx.newPage();

  page.on('console', m => console.log('[BROWSER]', m.type(), m.text()));
  page.on('requestfailed', r => console.warn('[REQ-FAILED]', r.method(), r.url(), r.failure()?.errorText));
  page.on('response', async r => { if (r.status() >= 400) console.warn('[RES-ERROR]', r.status(), r.url()); });

  try {
    // === 1) ログインページへ ===
    console.log('[STEP] goto login');
    await page.goto(`${env.POS_BASE}/auth/login`, { waitUntil: 'domcontentloaded', timeout: 60000 })
      .catch(async () => { await page.goto(`${env.POS_BASE}/auth`, { waitUntil: 'domcontentloaded', timeout: 60000 }); });

    // 3通りの方法でログイン試行
    let loggedIn = false;

    // A) 直接フォームに email/password がある場合
    try {
      const emailSel = 'input[type="email"], input[name="email"], input[autocomplete="username"]';
      const passSel  = 'input[type="password"], input[name="password"], input[autocomplete="current-password"]';
      await page.waitForSelector(emailSel, { timeout: 7000 });
      await page.fill(emailSel, env.POS_EMAIL);
      await page.fill(passSel, env.POS_PASSWORD);
      const loginBtn = 'button[type="submit"], button:has-text("ログイン"), button:has-text("Sign in"), button:has-text("Sign In")';
      await Promise.all([
        page.waitForNavigation({ waitUntil: 'networkidle', timeout: 30000 }).catch(()=>{}),
        page.click(loginBtn)
      ]);
      loggedIn = true;
    } catch {}

    // B) 画面に「ログイン」リンク/ボタンがあって、そこからフォームが出るタイプ
    if (!loggedIn) {
      try {
        const openBtn = page.getByRole ? page.getByRole('button', { name: /ログイン|sign ?in/i }) : null;
        if (openBtn) { await openBtn.click().catch(()=>{}); }
        const emailSel = 'input[type="email"], input[name="email"], input[autocomplete="username"]';
        const passSel  = 'input[type="password"], input[name="password"], input[autocomplete="current-password"]';
        await page.waitForSelector(emailSel, { timeout: 7000 });
        await page.fill(emailSel, env.POS_EMAIL);
        await page.fill(passSel, env.POS_PASSWORD);
        await Promise.all([
          page.waitForNavigation({ waitUntil: 'networkidle', timeout: 30000 }).catch(()=>{}),
          page.click('button[type="submit"], button:has-text("ログイン"), button:has-text("Sign in")')
        ]);
        loggedIn = true;
      } catch {}
    }

    // C) Next.js（NextAuth等）でフォームがモーダルにあるケース
    if (!loggedIn) {
      try {
        // ラベルから探す（アクセシビリティ対応サイト想定）
        const emailLabel = page.getByLabel ? page.getByLabel(/メール|email/i) : null;
        const passLabel  = page.getByLabel ? page.getByLabel(/パスワード|password/i) : null;
        if (emailLabel && passLabel) {
          await emailLabel.fill(env.POS_EMAIL);
          await passLabel.fill(env.POS_PASSWORD);
          await Promise.all([
            page.waitForNavigation({ waitUntil: 'networkidle', timeout: 30000 }).catch(()=>{}),
            page.getByRole('button', { name: /ログイン|sign ?in/i }).click()
          ]);
          loggedIn = true;
        }
      } catch {}
    }

    await page.waitForLoadState('networkidle', { timeout: 15000 }).catch(()=>{});
    await page.screenshot({ path: path.join(DEBUG_DIR, 'after-login.png'), fullPage: true });
    console.log('[INFO] after login url=', page.url());

    // === 2) 対象ページへ ===
    const targetUrl = `${env.POS_BASE}/auth/item?genreId=${env.GENRE_ID}`;
    console.log('[STEP] goto target', targetUrl);
    await page.goto(targetUrl, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForLoadState('networkidle', { timeout: 20000 }).catch(()=>{});
    await page.screenshot({ path: path.join(DEBUG_DIR, 'target-page.png'), fullPage: true });

    // === 3) JSON API レスポンス捕捉 ===
    console.log('[STEP] waitForResponse(JSON API)');
    const apiResp = await page.waitForResponse(res => {
      try {
        const url = res.url();
        const ct  = (res.headers()['content-type'] || '').toLowerCase();
        const ok  = res.status() === 200;
        const same = url.startsWith(env.POS_BASE);
        const json = ct.includes('application/json');
        return ok && same && json && (url.includes('/api') || url.includes('genreId='));
      } catch { return false; }
    }, { timeout: 60000 });

    const apiUrl = apiResp.url();
    const data = await apiResp.json();
    console.log('[INFO] captured API:', apiUrl);
    fs.writeFileSync(path.join(DEBUG_DIR, 'api-sample.json'), JSON.stringify(data, null, 2));

    // === 4) 配列抽出 ===
    let arr = [];
    if (Array.isArray(data)) arr = data;
    else if (data && typeof data === 'object') {
      for (const k of ['data','items','records','result','payload']) {
        if (Array.isArray(data[k])) { arr = data[k]; break; }
        if (data[k] && typeof data[k] === 'object') {
          for (const k2 of ['data','items','records','result','payload']) {
            if (Array.isArray(data[k][k2])) { arr = data[k][k2]; break; }
          }
          if (arr.length) break;
        }
      }
    }

    if (!arr.length) {
      console.error('[ERROR] No array found in API JSON. head=', JSON.stringify(data).slice(0,300));
      await page.screenshot({ path: path.join(DEBUG_DIR, 'no-array.png'), fullPage: true });
      exitCode = 2;
    } else {
      console.log('[INFO] items length =', arr.length, 'keys =', Object.keys(arr[0]||{}).slice(0,20));
      fs.writeFileSync(path.join(DEBUG_DIR, 'items-sample.json'), JSON.stringify(arr.slice(0,3), null, 2));

      // === 5) GASへPOST ===
      console.log('[STEP] POST to GAS');
      const res = await page.request.post(env.GAS_WEBHOOK_URL, {
        headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${env.GAS_SHARED_SECRET}` },
        data: { items: arr }
      });
      const txt = await res.text();
      console.log('GAS response:', txt);
      if (!txt.includes('"ok":true')) exitCode = 3;
    }

  } catch (err) {
    console.error('[FATAL]', err);
    exitCode = 1;
    try { await page.screenshot({ path: path.join(DEBUG_DIR, 'fatal.png'), fullPage: true }); } catch {}
  } finally {
    await ctx.close();
    await browser.close();
    process.exit(exitCode);
  }
}

main();
