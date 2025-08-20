// fetch-pos.js — Cookie直叩き版（Playwright不要）
/* Node 20+ で動作（グローバル fetch 使用） */
function findArrayPayload(json) {
  if (Array.isArray(json)) return json;
  if (json && typeof json === 'object') {
    const keys = ['data','items','records','result','payload','list'];
    for (const k of keys) {
      if (Array.isArray(json[k])) return json[k];
      if (json[k] && typeof json[k] === 'object') {
        for (const k2 of keys) {
          if (Array.isArray(json[k][k2])) return json[k][k2];
        }
      }
    }
  }
  return [];
}

(async () => {
  const { POS_API_URL, POS_COOKIE, GAS_WEBHOOK_URL, GAS_SHARED_SECRET } = process.env;

  if (!POS_API_URL || !POS_COOKIE || !GAS_WEBHOOK_URL || !GAS_SHARED_SECRET) {
    console.error('Missing env. Need POS_API_URL, POS_COOKIE, GAS_WEBHOOK_URL, GAS_SHARED_SECRET');
    process.exit(1);
  }

  // 1) POS の API を Cookie 付きで直GET
  const res = await fetch(POS_API_URL, {
    headers: {
      'Accept': 'application/json, text/plain, */*',
      'Cookie': POS_COOKIE,
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/123 Safari/537.36'
    }
  });
  const ct = res.headers.get('content-type') || '';
  if (!res.ok) {
    console.error(`POS API error: ${res.status} ${res.statusText}`);
    process.exit(2);
  }
  const text = await res.text();
  let json;
  try { json = JSON.parse(text); } catch {
    console.error('Not JSON. Head:', text.slice(0,200));
    process.exit(2);
  }

  const items = findArrayPayload(json);
  if (!items.length) {
    console.error('No array payload in JSON. Keys:', Object.keys(json||{}));
    process.exit(2);
  }
  console.log('[INFO] fetched items:', items.length, 'sample keys:', Object.keys(items[0]||{}).slice(0,15));

  // 2) GAS へ POST（受け口は既に作成済みの doPost）
  const gas = await fetch(GAS_WEBHOOK_URL, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${GAS_SHARED_SECRET}`
    },
    body: JSON.stringify({ items })
  });
  const gasText = await gas.text();
  console.log('GAS response:', gasText);
  if (!gas.ok || !gasText.includes('"ok":true')) process.exit(3);
})();
