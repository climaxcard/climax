// fetch-pos.js（堅牢ログイン版）
const { chromium, devices } = require('playwright');
const fs = require('fs');
const path = require('path');

const DEBUG_DIR = path.join(process.cwd(), 'debug');
fs.mkdirSync(DEBUG_DIR, { recursive: true });

function mask(s){ if(!s) return '(missing)'; return `(len:${String(s).length})`; }

async function tryLogin(page, email, password, base) {
  // 候補URL（順にトライ）
  const urls = [
    `${base}/auth/login`,
    `${base}/auth/signin`,
    `${base}/auth`,
    `${base}/login`,
    `${base}/signin`
  ];

  // 候補セレクタ
  const emailSels = [
    'input[type="email"]',
    'input[name="email"]',
    'input[autocomplete="username"]',
    'input[placeholder*="メール"]',
    'input[placeholder*="email" i]'
  ];
  const passSels = [
    'input[type="password"]',
    'input[name="password"]',
    'input[autocomplete="current-password"]',
    'input[placeholder*="パスワード"]',
    'input[placeholder*="password" i]'
  ];
  const submitSels = [
    'button[type="submit"]',
    'button:has-text("ログイン")',
    'button:has-text("Sign in")',
    'button:has-text("Sign In")',
    'input[type="submit"]'
  ];

  for (const url of urls) {
    try {
      await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
      await page.waitForLoadState('networkidle', { timeout: 15000 }).catch(()=>{});

      // 直接ページ内で探す
      for (const eSel of emailSels) {
        const emailBox = await page.$(eSel);
        if (!emailBox) continue;

        // パス欄も探す
        let passBox = null;
        for (const pSel of passSels) {
          passBox = await page.$(pSel);
          if (passBox) break;
        }
        if (!passBox) continue;

        await emailBox.fill(email);
        await passBox.fill(password);

        // 送信
        let clicked = false;
        for (const sSel of submitSels) {
          const btn = await page.$(sSel);
          if (btn) {
            await Promise.all([
              page.waitForNavigation({ waitUntil: 'networkidle', timeout: 30000 }).catch(()=>{}),
              btn.click()
            ]);
            clicked = true;
            break;
          }
        }
        if (!clicked) {
          // Enter送信
          await passBox.press('Enter');
          await page.waitForLoadState('networkidle', { timeout: 15000 }).catch(()=>{});
        }

        // ログイン判定：/auth/ 直下ダッシュボード等に遷移していればOKにする
        const cur = page.url();
        if (!/\/auth\/?(login|signin)?$/i.test(new URL(cur).pathname)) {
          return true;
        }
      }

      // フォームがモーダル/別DOMの場合に備えてボタン押してみる
      const openBtn = await page.$('button:has-text("ログイン")');
      if (openBtn) {
        await openBtn.click().catch(()=>{});
        await page.waitForTimeout(500);
        // 再度探す
        for (const eSel of emailSels) {
          const emailBox = await page.$(eSel);
          if (!emailBox) continue;
          let passBox = null;
          for (const pSel of passSels) {
            passBox = await page.$(pSel);
            if (passBox) break;
          }
          if (!passBox) continue;

          await emailBox.fill(email);
          await passBox.fill(password);
          const btn = await page.$('button[type="submit"], button:has-text("ログイン"), button:has-text("Sign in")');
          if (btn) {
            await Promise.all([
              page.waitForNavigation({ waitUntil: 'networkidle', timeout: 30000 }).catch(()=>{}),
              btn.click()
            ]);
          } else {
            await passBox.press('Enter');
            await page.waitForLoadState('networkidle', { timeout: 15000 }).catch(()=>{});
          }
          const cur = page.url();
          if (!/\/auth\/?(login|signin)?$/i.test(new URL(cur).pathname)) {
            return true;
          }
        }
      }

    } catch (e) {
      // 次のURL候補へ
    }
  }

  // 失敗時スクショ
  await page.screenshot({ path: path.join(DEBUG_DIR, 'after-login.png'), fullPage: true });
  return false;
}

async function main() {
  const {
    POS_EMAIL,
    POS_PASSWORD,
    GAS_WEBHOOK_URL,
    GAS_SHARED_SECRET,
    GENRE_ID = '137',
    POS_BASE = 'https://pos.mycalinks.com'
  } = process.env;

  console.log('[ENV] POS_EMAIL=%s POS_PASSWORD=%s GAS_WEBHOOK_URL=%s GAS_SHARED_SECRET=%s GENRE_ID=%s POS_BASE=%s',
    mask(POS_EMAIL), mask(POS_PASSWORD), mask(GAS_WEBHOOK_URL), mask(GAS_SHARED_SECRET), GENRE_ID, POS_BASE);

  const browser = await chromium.launch({ headless: true });
  const ctx = await browser.newContext({ userAgent: devices['Desktop Chrome'].userAgent, locale: 'ja-JP' });
  const page = await ctx.newPage();

  try {
    // 1) ログイン
    console.log('[STEP] login…');
    const ok = await tryLogin(page, POS_EMAIL, POS_PASSWORD, POS_BASE);
    if (!ok) {
      console.error('Login failed: unable to locate email/password fields or submit.');
      process.exit(1);
    }
    console.log('[INFO] logged in. url=', page.url());

    // 2) 対象ページへ
    const targetUrl = `${POS_BASE}/auth/item?genreId=${GENRE_ID}`;
    console.log('[STEP] goto target', targetUrl);
    await page.goto(targetUrl, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForLoadState('networkidle', { timeout: 20000 }).catch(()=>{});
    await page.screenshot({ path: path.join(DEBUG_DIR, 'target-page.png'), fullPage: true });

    // 3) JSON API レスポンス待ち（/api & JSON & 200 & genreId）
    console.log('[STEP] waitForResponse(JSON API)');
    const apiResp = await page.waitForResponse(res => {
      try {
        const url = res.url();
        const ct  = (res.headers()['content-type'] || '').toLowerCase();
        return res.status() === 200
          && ct.includes('application/json')
          && url.includes('/api')
          && (url.includes('genreId=') || url.includes(`/genreId=${GENRE_ID}`));
      } catch { return false; }
    }, { timeout: 60000 });

    const data = await apiResp.json();
    const apiUrl = apiResp.url();
    console.log('[INFO] captured API:', apiUrl);

    // 4) 配列抽出
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
      console.error('No array found in API JSON. head=', JSON.stringify(data).slice(0,300));
      process.exit(2);
    }

    // 5) GASへPOST
    const res = await page.request.post(GAS_WEBHOOK_URL, {
      headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${GAS_SHARED_SECRET}` },
      data: { items: arr }
    });
    const txt = await res.text();
    console.log('GAS response:', txt);
    if (!txt.includes('"ok":true')) process.exit(3);

  } finally {
    await ctx.close();
    await browser.close();
  }
}

main().catch(e => { console.error(e); process.exit(1); });
