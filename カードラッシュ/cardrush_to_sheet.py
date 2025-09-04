# -*- coding: utf-8 -*-
"""
CardRush（デュエル・マスターズ）買取表 → Googleスプレッドシート直書き込み（堅牢版）
- requests.Session + リトライ、UA付き
- JSON/HTMLのどちらでも解析（__NEXT_DATA__ フォールバック）
- dict配列を再帰探索して抽出（name/model_number/amount を含む配列を優先）
- 0件取得のときはシートを消さず前回データを温存
"""

import argparse
import json
import re
import time
from typing import Any, Dict, Iterable, List, Tuple, Optional

import requests
from bs4 import BeautifulSoup
import gspread
from google.oauth2.service_account import Credentials
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

BASE_URL = "https://cardrush.media/duel_masters/buying_prices"
DEFAULT_LIMIT = 120          # 大き過ぎると弾かれやすいので控えめに
DEFAULT_MAX_PAGES = 200
DEFAULT_SLEEP_MS = 400

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/118.0.0.0 Safari/537.36"
)

HEADERS = {
    "User-Agent": UA,
    "Accept": "application/json,text/html;q=0.9,*/*;q=0.8",
    "Accept-Language": "ja,en;q=0.8",
    "Referer": "https://cardrush.media/duel_masters/buying_prices",
    "X-Requested-With": "XMLHttpRequest",
}

# ---------- HTTP ----------

def make_session() -> requests.Session:
    s = requests.Session()
    s.headers.update(HEADERS)
    retry = Retry(
        total=3,
        backoff_factor=0.6,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset(["GET"]),
        raise_on_status=False,
    )
    s.mount("https://", HTTPAdapter(max_retries=retry))
    return s

def build_params(page: int, limit: int) -> List[Tuple[str, str]]:
    # サーバ負荷になりにくい最小限のパラメータ
    return [
        ("displayMode", "リスト"),
        ("is_hot", "true"),
        ("limit", str(limit)),
        ("page", str(page)),
        ("sort[key]", "amount"),
        ("sort[order]", "desc"),
        # JSON最適化は弾かれる可能性があるので付けない
    ]

# ---------- パース系 ----------

def normalize_amount(v: Any) -> int:
    if v is None:
        return 0
    if isinstance(v, (int, float)):
        return int(v)
    s = re.sub(r"[^\d]", "", str(v))
    return int(s) if s else 0

def _looks_like_item_dict(d: Dict[str, Any]) -> bool:
    # name / model_number / amount のいずれかを持つ辞書を「アイテム辞書」とみなす
    if not isinstance(d, dict):
        return False
    keys = d.keys()
    return any(k in keys for k in ("name", "model_number", "amount"))

def _deep_find_items_array(o: Any) -> List[Dict[str, Any]]:
    """
    与えられたJSON/objから、アイテム辞書(dict)の配列を探索して返す。
    - まずは「要素がdictで、name/model_number/amountのいずれかを含む配列」を優先
    - 見つからなければ他の配列を再帰的に探索
    """
    best: Optional[List[Dict[str, Any]]] = None

    def dfs(x: Any):
        nonlocal best
        if best is not None:
            return
        if isinstance(x, list) and x:
            dicts = [e for e in x if isinstance(e, dict)]
            if dicts:
                if any(_looks_like_item_dict(dd) for dd in dicts):
                    best = dicts
                    return
                for dd in dicts:
                    dfs(dd)
            else:
                for e in x:
                    dfs(e)
        elif isinstance(x, dict):
            # よくあるコンテナキーを優先
            for k in ("items", "list", "data", "records", "products", "result", "buying_prices"):
                if k in x:
                    dfs(x[k])
                    if best is not None:
                        return
            for v in x.values():
                dfs(v)

    dfs(o)
    return best or []

def parse_items_from_json(obj: Any) -> List[Dict[str, Any]]:
    if obj is None:
        return []
    if isinstance(obj, list):
        dicts = [e for e in obj if isinstance(e, dict)]
        if dicts:
            if any(_looks_like_item_dict(dd) for dd in dicts):
                return dicts
            return _deep_find_items_array(obj)
        return _deep_find_items_array(obj)
    if isinstance(obj, dict):
        for k in ("buying_prices", "items", "list", "data", "records", "products", "result"):
            if isinstance(obj.get(k), list):
                return parse_items_from_json(obj[k])
        return _deep_find_items_array(obj)
    return []

def parse_items_from_next_data(html_text: str) -> List[Dict[str, Any]]:
    soup = BeautifulSoup(html_text, "html.parser")
    tag = soup.find("script", id="__NEXT_DATA__")
    if not tag or not tag.text:
        return []
    try:
        data = json.loads(tag.text)
    except Exception:
        return []
    return parse_items_from_json(data)

# ---------- 1ページ取得 ----------

def fetch_page(page: int, limit: int, timeout: float = 20.0, debug: bool=False) -> Tuple[List[Dict[str, Any]], str, int]:
    sess = make_session()
    params = build_params(page, limit)
    r = sess.get(BASE_URL, params=params, timeout=timeout)
    used_url = r.url
    txt = r.text.strip()
    status = r.status_code

    if debug:
        print(f"[debug] page={page} status={status} len={len(txt)} url={used_url}")

    # JSON優先
    if txt.startswith("{") or txt.startswith("["):
        try:
            obj = r.json()
            items = parse_items_from_json(obj)
            if debug:
                print(f"[debug] json-parse items={len(items)}")
            if items:
                return items, used_url, status
        except Exception as e:
            if debug:
                print("[debug] json parse error:", e)

    # HTMLフォールバック（__NEXT_DATA__）
    items = parse_items_from_next_data(txt)
    if debug:
        print(f"[debug] next_data items={len(items)}; head={txt[:300]!r}")
    return items, used_url, status

# ---------- レコード整形 ----------

def items_to_rows(items: Iterable[Dict[str, Any]], used_url: str) -> List[List[Any]]:
    rows: List[List[Any]] = []
    for it in items:
        if not isinstance(it, dict):
            continue
        name = it.get("name", "")
        model = it.get("model_number", "")
        amount = normalize_amount(it.get("amount"))
        rarity = it.get("rarity", "")
        category = it.get("display_category", "")
        if name or model or amount:
            rows.append([name, model, amount, category, rarity, used_url])
    return rows

def scrape_all(limit=DEFAULT_LIMIT, max_pages=DEFAULT_MAX_PAGES, sleep_ms=DEFAULT_SLEEP_MS, debug: bool=False) -> List[List[Any]]:
    all_rows: List[List[Any]] = []
    for page in range(1, max_pages + 1):
        items, used_url, status = fetch_page(page, limit, debug=debug)
        # 403/5xx などで中断
        if status >= 400 and not items:
            if debug:
                print(f"[debug] stop at page={page} (http {status})")
            break
        if not items:
            if debug:
                print(f"[debug] stop at page={page} (no items)")
            break
        all_rows.extend(items_to_rows(items, used_url))
        if len(items) < limit:
            break
        time.sleep(max(0, sleep_ms) / 1000.0)
    if debug:
        print(f"[debug] total rows={len(all_rows)}")
    return all_rows

# ---------- シート書き込み ----------

def write_to_sheet(rows: List[List[Any]], sheet_url: str, sheet_name: str, creds_path: str):
    # 0件なら更新しない（前回データ温存）
    if not rows:
        print("[warn] fetched 0 rows. keep previous data (no overwrite).")
        return

    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(creds_path, scopes=scopes)
    gc = gspread.authorize(creds)

    sh = gc.open_by_url(sheet_url)
    try:
        ws = sh.worksheet(sheet_name)
        ws.clear()
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(sheet_name, rows=max(100, len(rows) + 10), cols=10)

    header = ["カード名", "型番", "買取金額(円)", "カテゴリ", "レア", "取得元URL"]
    ws.update("A1", [header] + rows)
    # 金額列フォーマット
    if rows:
        ws.format(f"C2:C{len(rows)+1}", {"numberFormat": {"type": "NUMBER", "pattern": "#,##0"}})

# ---------- CLI ----------

def main():
    ap = argparse.ArgumentParser(description="CardRush（DM）→ Googleスプレッドシート直書き（堅牢版）")
    ap.add_argument("--sheet-url", required=True, help="書き込み先スプレッドシートのURL")
    ap.add_argument("--sheet-name", default="CardRush_DM", help="シート（タブ）名")
    ap.add_argument("--creds", required=True, help="サービスアカウントJSONのパス")
    ap.add_argument("--limit", type=int, default=DEFAULT_LIMIT)
    ap.add_argument("--max-pages", type=int, default=DEFAULT_MAX_PAGES)
    ap.add_argument("--sleep-ms", type=int, default=DEFAULT_SLEEP_MS)
    ap.add_argument("--debug", action="store_true", help="詳細ログを出す")
    args = ap.parse_args()

    if not args.sheet_name:
        args.sheet_name = "CardRush_DM"

    rows = scrape_all(limit=args.limit, max_pages=args.max_pages, sleep_ms=args.sleep_ms, debug=args.debug)
    write_to_sheet(rows, args.sheet_url, args.sheet_name, args.creds)
    print(f"[OK] {len(rows)} 件を書き込み完了（0なら前回データ維持）: {args.sheet_name}")

if __name__ == "__main__":
    main()
