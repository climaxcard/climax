// fetch-pos.js
const { chromium } = require('playwright');

async function main() {
  const {
    POS_EMAIL,
    POS_PASSWORD,
    GAS_WEBHOOK_URL,
    GAS_SHARED_SECRET,
    GENRE_ID = '137',
    POS_BASE = 'https://pos.mycalinks.com'
  } = process.env;

  if (!POS_EMAIL || !POS_PASSWORD || !GAS_WEBHOOK_URL || !GAS_SHARED_SECRET) {
    console.error('Missing env: POS_EMAIL, POS_PASSWORD, GAS_WEBHOOK_URL, GAS_SHARED_SECRET');
    process.exit(1);
  }

  const browser = await chromium.launch({ headless: true });
  const ctx = await browser.newContext();
  const page = await ctx.newPage();

  try {
    // 1) ログイン
    await page.goto(`${POS_BASE}/auth/login`, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.fill('input[type="email"], input[name="email"]', POS_EMAIL);
    await page.fill('input[type="password"], input[name="password"]', POS_PASSWORD);
    await Promise.all([
      page.waitForNavigation({ waitUntil: 'networkidle', timeout: 60000 }),
      page.click('button[type="submit"], button:has-text("ログイン"), button:has-text("Sign in")')
    ]);

    // 2) 対象ページへ
    const targetUrl = `${POS_BASE}/auth/item?genreId=${GENRE_ID}`;
    await page.goto(targetUrl, { waitUntil: 'domcontentloaded', timeout: 60000 });

    // 3) JSON API レスポンスを待つ（必要に応じてURL条件を狭める）
    const apiResp = await page.waitForResponse(resp => {
      try {
        const url = resp.url();
        const ok = resp.status() === 200;
        const ct = (resp.headers()['content-type'] || '').toLowerCase();
        const looksJson = ct.includes('application/json') || url.includes('api');
        const sameOrigin = url.startsWith(POS_BASE);
        // 例：特定エンドポイントに絞るなら ↓ を有効化
        // const endpoint = url.includes('/api/items') && url.includes(`genreId=${GENRE_ID}`);
        // return ok && looksJson && sameOrigin && endpoint;
        return ok && looksJson && sameOrigin;
      } catch (_) { return false; }
    }, { timeout: 60000 });

    const data = await apiResp.json();

    // 4) 配列を抽出
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
      console.error('No array found in API JSON. Dump head:', JSON.stringify(data).slice(0,300));
      process.exit(2);
    }

    // 5) GASへPOST
    const res = await page.request.post(GAS_WEBHOOK_URL, {
      headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${GAS_SHARED_SECRET}` },
      data: { items: arr }
    });
    const txt = await res.text();
    console.log('GAS response:', txt);
  } finally {
    await browser.close();
  }
}

main().catch(err => { console.error(err); process.exit(1); });
