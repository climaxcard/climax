# -*- coding: utf-8 -*-
"""
CardRush（デュエル・マスターズ）買取表スクレイパー
- 取得項目: カード名(name) / 型番(model_number) / 金額(amount) / カテゴリ(display_category) / レア(rarity)
- 出力: UTF-8 BOM 付き CSV（Excelでも文字化けしにくい）
- 仕様:
  * JSON系パラメータを付与してまずJSONで取得（高速・安定）
  * もしHTMLしか返らなくても __NEXT_DATA__ をパースしてフォールバック
  * ページング（limit, page）でデータが尽きるまで収集
  * 金額は数値(円)に正規化（"¥28,000" -> 28000）
  * 同名カードの混在対策として必ず「型番」を出力
"""

import argparse
import csv
import json
import time
from typing import Dict, Iterable, List, Tuple, Any
import re

import requests
from bs4 import BeautifulSoup

BASE_URL = "https://cardrush.media/duel_masters/buying_prices"

DEFAULT_LIMIT = 100
DEFAULT_MAX_PAGES = 200
DEFAULT_SLEEP_MS = 350

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) "
                  "Chrome/118.0.0.0 Safari/537.36",
    "Accept": "application/json,text/html;q=0.9,*/*;q=0.8",
    "Accept-Language": "ja,en;q=0.8",
}


def build_params(page: int, limit: int) -> Dict[str, str]:
    """
    JSONを強く要求するためのパラメータ群。
    フィールドを絞って軽量化も図る。
    """
    p = {
        "limit": str(limit),
        "page": str(page),
        # 表示モードは任意（サーバ側で意味が無くても害はない）
        "displayMode": "リスト",
        # 並び順（任意）
        "sort[key]": "amount",
        "sort[order]": "desc",
        # JSON返却時にフィールド制限
        "to_json_option[only][]": "id",
    }

    # 複数キーの to_json_option[only][] を付ける
    for k in ["name", "model_number", "amount", "rarity", "display_category"]:
        p.setdefault("to_json_option[only][]", [])
        # requestsは同名キーの複数値をリストで渡せばよい
    # 上のやり方だと最後しか残らないので、下で別途追加
    return p


def build_params_with_multi(page: int, limit: int) -> List[Tuple[str, str]]:
    """
    requests に同名キーを複数渡すため、(key, value) の配列で返す版
    """
    params_list = [
        ("limit", str(limit)),
        ("page", str(page)),
        ("displayMode", "リスト"),
        ("sort[key]", "amount"),
        ("sort[order]", "desc"),
        ("to_json_option[only][]", "id"),
        ("to_json_option[only][]", "name"),
        ("to_json_option[only][]", "model_number"),
        ("to_json_option[only][]", "amount"),
        ("to_json_option[only][]", "rarity"),
        ("to_json_option[only][]", "display_category"),
        # 必要なら関連も
        ("associations[]", "ocha_product"),
        ("to_json_option[include][ocha_product][only][]", "id"),
    ]
    return params_list


def normalize_amount(v: Any) -> int:
    """
    金額を数値(円)へ正規化。空は None の代わりに 0 へ（列の都合で int に寄せる）。
    """
    if v is None:
        return 0
    if isinstance(v, (int, float)):
        return int(v)
    s = re.sub(r"[^\d]", "", str(v))
    return int(s) if s else 0


def parse_as_items_from_json(obj: Any) -> List[Dict[str, Any]]:
    """
    返ったJSONが配列 or {items:[...]} など、素直な配列を取り出す。
    """
    if obj is None:
        return []
    if isinstance(obj, list):
        return obj
    if isinstance(obj, dict):
        # よくあるキーを優先
        for k in ("items", "list", "data", "records", "products"):
            if isinstance(obj.get(k), list):
                return obj[k]
        # 1段深い階層にある場合もあるので探索（浅め）
        for v in obj.values():
            if isinstance(v, list):
                return v
            if isinstance(v, dict):
                for k in ("items", "list", "data", "records", "products"):
                    if isinstance(v.get(k), list):
                        return v[k]
    return []


def parse_items_from_next_data(html_text: str) -> List[Dict[str, Any]]:
    """
    HTML内の __NEXT_DATA__ を探して、それらしい配列を抽出。
    """
    soup = BeautifulSoup(html_text, "html.parser")
    tag = soup.find("script", id="__NEXT_DATA__")
    if not tag or not tag.text:
        return []

    try:
        next_data = json.loads(tag.text)
    except Exception:
        return []

    # 再帰的に配列を探索（最初に見つかった非空の配列を返す簡易版）
    def deep_find_list(o: Any) -> List[Any]:
        if isinstance(o, list):
            return o
        if isinstance(o, dict):
            # 配列がありそうなキーを優先的に見る
            for k in ("items", "list", "data", "records", "products", "result"):
                if isinstance(o.get(k), list) and o.get(k):
                    return o[k]
            # 全走査（最初に見つかった配列を返す）
            for vv in o.values():
                found = deep_find_list(vv)
                if isinstance(found, list) and found:
                    return found
        return []

    arr = deep_find_list(next_data)
    # arr の要素がdictでないと使いにくいので、そのまま返す（dict想定）
    if isinstance(arr, list):
        return arr
    return []


def fetch_page(page: int, limit: int, timeout: float = 20.0) -> Tuple[List[Dict[str, Any]], str]:
    """
    1ページ分を取得して配列化。返り値: (items, used_url)
    """
    params = build_params_with_multi(page, limit)
    r = requests.get(BASE_URL, params=params, headers=HEADERS, timeout=timeout)
    url_used = r.url
    text = r.text.strip()

    # まずJSONとして試す
    if text.startswith("{") or text.startswith("["):
        try:
            obj = r.json()
            items = parse_as_items_from_json(obj)
            return items, url_used
        except Exception:
            pass

    # フォールバック: HTML → __NEXT_DATA__
    items = parse_items_from_next_data(text)
    return items, url_used


def items_to_rows(items: Iterable[Dict[str, Any]], url_used: str) -> List[List[Any]]:
    """
    API/HTMLから得たアイテム dict をCSV行（配列）へ変換。
    """
    rows: List[List[Any]] = []
    for it in items:
        name = it.get("name", "")
        model = it.get("model_number", "")
        amount = normalize_amount(it.get("amount"))
        rarity = it.get("rarity", "")
        category = it.get("display_category", "")

        # name または model または amount があれば採用
        if name or model or amount:
            rows.append([name, model, amount, category, rarity, url_used])
    return rows


def scrape_all(limit: int = DEFAULT_LIMIT,
               max_pages: int = DEFAULT_MAX_PAGES,
               sleep_ms: int = DEFAULT_SLEEP_MS) -> List[List[Any]]:
    """
    ページングで全件（尽きるまで）取得
    """
    all_rows: List[List[Any]] = []
    for page in range(1, max_pages + 1):
        items, url_used = fetch_page(page, limit)
        if not items:
            break

        rows = items_to_rows(items, url_used)
        all_rows.extend(rows)

        # 次ページ判定：limit未満なら打ち切り
        if len(items) < limit:
            break

        time.sleep(max(0, sleep_ms) / 1000.0)

    return all_rows


def save_csv(rows: List[List[Any]], out_csv: str) -> None:
    """
    UTF-8 BOM付きで保存（Excelで開きやすい）
    """
    header = ["カード名", "型番", "買取金額(円)", "カテゴリ", "レア", "取得元URL"]
    with open(out_csv, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(header)
        writer.writerows(rows)


def main():
    ap = argparse.ArgumentParser(description="CardRush（デュエマ）買取表スクレイピング → CSV出力")
    ap.add_argument("-o", "--out", default="cardrush_dm_buylist.csv",
                    help="出力CSVパス（既定: cardrush_dm_buylist.csv）")
    ap.add_argument("--max-pages", type=int, default=DEFAULT_MAX_PAGES,
                    help=f"最大ページ数（既定: {DEFAULT_MAX_PAGES}）")
    ap.add_argument("--limit", type=int, default=DEFAULT_LIMIT,
                    help=f"1ページ件数（既定: {DEFAULT_LIMIT}）")
    ap.add_argument("--sleep-ms", type=int, default=DEFAULT_SLEEP_MS,
                    help=f"ページ間スリープms（既定: {DEFAULT_SLEEP_MS}ms）")
    args = ap.parse_args()

    rows = scrape_all(limit=args.limit, max_pages=args.max_pages, sleep_ms=args.sleep_ms)
    save_csv(rows, args.out)

    print(f"[OK] 取得件数: {len(rows)} 件 -> {args.out}")


if __name__ == "__main__":
    main()
