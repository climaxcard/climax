// fetch-pos.js (login診断つき最強版)
const { chromium, devices } = require('playwright');
const fs = require('fs'); const path = require('path');
const DEBUG_DIR = path.join(process.cwd(), 'debug'); fs.mkdirSync(DEBUG_DIR, { recursive: true });
function mask(s){ return s ? `(len:${String(s).length})` : '(missing)'; }

async function dumpLoginPage(page, label){
  await page.screenshot({ path: path.join(DEBUG_DIR, `${label}.png`), fullPage: true }).catch(()=>{});
  const html = await page.content().catch(()=> '');
  fs.writeFileSync(path.join(DEBUG_DIR, `${label}.html`), html);
  // 画面内のinput候補を列挙
  const inputs = await page.$$eval('input', els => els.map(e => ({
    type: e.type, name: e.name, id: e.id, ph: e.getAttribute('placeholder') || ''
  }))).catch(()=>[]);
  fs.writeFileSync(path.join(DEBUG_DIR, 'inputs.json'), JSON.stringify(inputs, null, 2));
  console.log('[INFO] inputs on page:', inputs.slice(0,10));
}

async function tryLogin(page, email, password, base, loginUrlEnv){
  const urlCandidates = [
    loginUrlEnv,
    `${base}/auth/login`, `${base}/auth/signin`, `${base}/auth`,
    `${base}/login`, `${base}/signin`, base
  ].filter(Boolean);

  const emailSels = [
    'input[type="email"]','input[name="email"]','input[autocomplete="username"]',
    'input#email','input[name="loginId"]','input[name="username"]',
    'input[placeholder*="メール"]','input[placeholder*="email" i]',
    'input[type="text"][name="email"]'
  ];
  const passSels = [
    'input[type="password"]','input[name="password"]','input[autocomplete="current-password"]',
    'input#password','input[placeholder*="パスワード"]','input[placeholder*="password" i]'
  ];
  const submitSels = [
    'button[type="submit"]','button:has-text("ログイン")','button:has-text("Sign in")','button:has-text("Sign In")',
    'input[type="submit"]','button:has-text("Sign in with Email")'
  ];

  for (const u of urlCandidates){
    try{
      console.log('[STEP] goto', u);
      await page.goto(u, { waitUntil: 'domcontentloaded', timeout: 60000 });
      await page.waitForLoadState('networkidle', { timeout: 15000 }).catch(()=>{});
      await dumpLoginPage(page, 'login-page');

      // もしトップページで「ログイン」ボタンがあるなら開く
      const openBtn = await page.$('a:has-text("ログイン"), button:has-text("ログイン"), a:has-text("Sign in"), button:has-text("Sign in")');
      if (openBtn) { await openBtn.click().catch(()=>{}); await page.waitForTimeout(800); }

      let emailBox=null, passBox=null;
      for (const s of emailSels){ emailBox = await page.$(s); if (emailBox) break; }
      for (const s of passSels){  passBox  = await page.$(s);  if (passBox)  break; }

      if (emailBox && passBox){
        await emailBox.fill(email);
        await passBox.fill(password);

        let clicked=false;
        for (const s of submitSels){
          const btn = await page.$(s);
          if (btn){
            await Promise.all([
              page.waitForNavigation({ waitUntil: 'networkidle', timeout: 30000 }).catch(()=>{}),
              btn.click()
            ]);
            clicked=true; break;
          }
        }
        if (!clicked){
          await passBox.press('Enter');
          await page.waitForLoadState('networkidle', { timeout: 15000 }).catch(()=>{});
        }

        // ログインできているか（/auth/login のままを除外）
        const p = new URL(page.url());
        if (!/\/auth\/?(login|signin)?$/i.test(p.pathname)){
          console.log('[INFO] login ok at', page.url());
          return true;
        }
      }
    }catch(e){ /* try next url */ }
  }
  await dumpLoginPage(page, 'after-login-fail'); // 証跡保存
  return false;
}

(async ()=>{
  const {
    POS_EMAIL, POS_PASSWORD, GAS_WEBHOOK_URL, GAS_SHARED_SECRET,
    GENRE_ID='137', POS_BASE='https://pos.mycalinks.com', LOGIN_URL
  } = process.env;

  console.log('[ENV] POS_EMAIL=%s POS_PASSWORD=%s GAS_WEBHOOK_URL=%s GAS_SHARED_SECRET=%s GENRE_ID=%s POS_BASE=%s LOGIN_URL=%s',
    mask(POS_EMAIL), mask(POS_PASSWORD), mask(GAS_WEBHOOK_URL), mask(GAS_SHARED_SECRET), GENRE_ID, POS_BASE, LOGIN_URL||'(auto)');

  const { chromium, devices } = require('playwright');
  const browser = await chromium.launch({ headless:true });
  const ctx = await browser.newContext({ userAgent: devices['Desktop Chrome'].userAgent, locale:'ja-JP' });
  const page = await ctx.newPage();

  try{
    const ok = await tryLogin(page, POS_EMAIL, POS_PASSWORD, POS_BASE, LOGIN_URL);
    if (!ok) { console.error('Login failed'); process.exit(1); }

    const target = `${POS_BASE}/auth/item?genreId=${GENRE_ID}`;
    console.log('[STEP] goto target', target);
    await page.goto(target, { waitUntil:'domcontentloaded', timeout:60000 });
    await page.waitForLoadState('networkidle', { timeout: 20000 }).catch(()=>{});
    await page.screenshot({ path: path.join(DEBUG_DIR, 'target.png'), fullPage:true });

    console.log('[STEP] waitForResponse');
    const apiResp = await page.waitForResponse(r=>{
      try{
        const url=r.url(); const ct=(r.headers()['content-type']||'').toLowerCase();
        return r.status()===200 && ct.includes('application/json') && (url.includes('/api') || url.includes('genreId='));
      }catch{return false;}
    }, { timeout:60000 });
    const arrData = await apiResp.json();
    let arr = Array.isArray(arrData) ? arrData : [];
    if (!arr.length && arrData && typeof arrData==='object'){
      for (const k of ['data','items','records','result','payload']){
        if (Array.isArray(arrData[k])) { arr = arrData[k]; break; }
        if (arrData[k] && typeof arrData[k]==='object'){
          for (const k2 of ['data','items','records','result','payload']){
            if (Array.isArray(arrData[k][k2])) { arr = arrData[k][k2]; break; }
          }
          if (arr.length) break;
        }
      }
    }
    if (!arr.length){ console.error('No array in JSON. head=', JSON.stringify(arrData).slice(0,300)); process.exit(2); }

    const res = await page.request.post(GAS_WEBHOOK_URL, {
      headers: { 'Content-Type':'application/json', 'Authorization': `Bearer ${GAS_SHARED_SECRET}` },
      data: { items: arr }
    });
    console.log('GAS response:', await res.text());
  } finally {
    await ctx.close(); await browser.close();
  }
})();
