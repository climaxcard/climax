# -*- coding: utf-8 -*-
"""
2シート比較（型番AND名前, 高速版）
- 類似度: RapidFuzz（高速）
- 3段階マッチ（すべて AND：型番＆名前）
  P1: 型番完全一致 + name >= p1
  P2: 型番digitsOnly一致 + name >= p2
  P3: 型番部分一致候補から modelスコア上位K→ name >= p3
- デバッグ出力シート: 差分比較 / 未一致_Sheet1 / 候補プレビュー
"""
import argparse, re, unicodedata
from typing import Any, Dict, List, Tuple, Optional
import gspread
from google.oauth2.service_account import Credentials
from rapidfuzz import fuzz  # ★ 高速な文字列類似度

# ===== 正規化 =====
def nfkc(s: Any) -> str:
    return unicodedata.normalize("NFKC", str(s)) if s is not None else ""

def normalize_text(s: Any) -> str:
    s = nfkc(s).lower().strip()
    s = re.sub(r"\s+", "", s)
    s = s.replace("－","-").replace("—","-").replace("ー","-").replace("–","-").replace("~","-")
    return s

def strip_brackets(s: str) -> str:
    return re.sub(r"[《「『【\(\[＜<].*?[》」』】\)\]＞>]", "", s)

def normalize_name(s: Any) -> str:
    s = strip_brackets(nfkc(s))
    s = normalize_text(s)
    s = re.sub(r"(再録|プロモ|限定|中古|傷あり|未使用|特価|英語版|日本語版|美品|完品|.*版)", "", s)
    return s

def normalize_model(s: Any) -> str:
    return re.sub(r"[^0-9a-z]", "", normalize_text(s))

def digits_only(s: str) -> str:
    return re.sub(r"\D", "", s)

def price_to_int(v: Any) -> int:
    if v is None: return 0
    if isinstance(v, (int, float)): return int(v)
    return int(re.sub(r"[^\d]", "", str(v)) or 0)

# ===== 類似度（RapidFuzz） =====
def name_score(a: str, b: str) -> float:
    """Token系で順不同や挿入に頑健に。0..1"""
    a2, b2 = normalize_name(a), normalize_name(b)
    if not a2 or not b2: return 0.0
    # token_set_ratio と partial_ratio の平均（0..100）→ 0..1
    s1 = fuzz.token_set_ratio(a2, b2)
    s2 = fuzz.partial_ratio(a2, b2)
    return (s1 + s2) / 200.0

def model_score_quick(a: str, b: str) -> float:
    """型番は厳しめ：部分一致強め + 文字違いに少し寛容"""
    a2, b2 = normalize_model(a), normalize_model(b)
    if not a2 or not b2: return 0.0
    if a2 == b2: return 1.0
    if a2 in b2 or b2 in a2: return 0.95
    # 低コスト: 部分一致寄りの ratio（0..100→0..1）
    return fuzz.partial_ratio(a2, b2) / 100.0

# ===== 列指定 =====
def col_letter_to_index(col: str) -> int:
    col = col.strip().upper(); v = 0
    for ch in col:
        if not ('A' <= ch <= 'Z'): raise ValueError(f"列記号が不正: {col}")
        v = v*26 + (ord(ch)-64)
    return v-1

def find_col_index(header_row: List[str], key: str) -> Optional[int]:
    if not key: return None
    if re.fullmatch(r"[A-Za-z]+", key.strip()):
        try: return col_letter_to_index(key)
        except: pass
    key_n = normalize_text(key)
    for i,h in enumerate(header_row):
        if normalize_text(h) == key_n: return i
    for i,h in enumerate(header_row):
        if key_n and key_n in normalize_text(h): return i
    return None

# ===== シートI/O =====
def read_sheet(gc, sheet_url: str, sheet_name: str) -> List[List[str]]:
    ws = gc.open_by_url(sheet_url).worksheet(sheet_name)
    return ws.get_all_values()

def write_sheet(sheet_url: str, sheet_name: str, rows: List[List[Any]], header: List[str], creds_path: str):
    creds = Credentials.from_service_account_file(creds_path, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    gcw = gspread.authorize(creds); sh = gcw.open_by_url(sheet_url)
    try:
        ws = sh.worksheet(sheet_name); ws.clear()
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(sheet_name, rows=max(100, len(rows)+10), cols=len(header)+5)
    ws.update(values=[header]+rows, range_name="A1")
    if rows and header:
        last = len(rows)+1
        hdr_idx = {h:i for i,h in enumerate(header)}
        def fmt(colname, pattern):
            if colname in hdr_idx:
                col = hdr_idx[colname]+1
                ws.format(f"{chr(64+col)}2:{chr(64+col)}{last}", {"numberFormat":{"type":"NUMBER","pattern":pattern}})
        fmt("価格(1)", "#,##0"); fmt("価格(2)", "#,##0"); fmt("差分(2-1)", "#,##0;[Red]-#,##0")

# ===== インデックス（シート2） =====
def build_index_sheet2(rows: List[List[str]], name_col: int, model_col: int, price_col: int):
    items: List[Dict[str,Any]] = []
    by_model_exact: Dict[str, List[int]] = {}
    by_model_digits: Dict[str, List[int]] = {}
    by_prefix: Dict[str, List[int]] = {}
    for r_idx, row in enumerate(rows[1:], start=2):
        name = row[name_col] if name_col < len(row) else ""
        model = row[model_col] if model_col < len(row) else ""
        price = price_to_int(row[price_col] if price_col < len(row) else 0)
        n_name = normalize_name(name)
        n_model = normalize_model(model)
        dkey   = digits_only(n_model)
        pref6  = n_model[:6]
        item = {"row": r_idx, "name": name, "model": model, "price": price,
                "n_name": n_name, "n_model": n_model, "dkey": dkey, "pref6": pref6}
        i = len(items); items.append(item)
        if n_model: by_model_exact.setdefault(n_model, []).append(i)
        if dkey:    by_model_digits.setdefault(dkey, []).append(i)
        if pref6:   by_prefix.setdefault(pref6, []).append(i)
    return items, by_model_exact, by_model_digits, by_prefix

# ===== 比較（高速3段階 + 候補上限） =====
def compare_fast(
    s1_rows: List[List[str]], s2_rows: List[List[str]],
    s1_name_key: str, s1_model_key: str, s1_price_key: str,
    s2_name_key: str, s2_model_key: str, s2_price_key: str,
    p1_name=0.40, p2_name=0.55, p3_name=0.70,
    p3_topk=120  # Pass3で modelスコア上位から最大K件だけ名前比較
):
    h1 = s1_rows[0] if s1_rows else []; h2 = s2_rows[0] if s2_rows else []
    s1_name_col  = find_col_index(h1, s1_name_key)
    s1_model_col = find_col_index(h1, s1_model_key)
    s1_price_col = find_col_index(h1, s1_price_key)
    s2_name_col  = find_col_index(h2, s2_name_key)
    s2_model_col = find_col_index(h2, s2_model_key)
    s2_price_col = find_col_index(h2, s2_price_key)
    for label,val in {
        "sheet1 name":s1_name_col, "sheet1 model":s1_model_col, "sheet1 price":s1_price_col,
        "sheet2 name":s2_name_col, "sheet2 model":s2_model_col, "sheet2 price":s2_price_col
    }.items():
        if val is None: raise ValueError(f"{label} 列が見つかりません。キーを確認してください。")

    items2, by_exact, by_digits, by_prefix = build_index_sheet2(s2_rows, s2_name_col, s2_model_col, s2_price_col)

    results: List[List[Any]] = []
    unmatched_rows: List[List[Any]] = []
    preview_rows: List[List[Any]] = []

    for r1, row1 in enumerate(s1_rows[1:], start=2):
        name1  = row1[s1_name_col]  if s1_name_col  < len(row1) else ""
        model1 = row1[s1_model_col] if s1_model_col < len(row1) else ""
        price1 = price_to_int(row1[s1_price_col]  if s1_price_col < len(row1) else 0)
        n_name1  = normalize_name(name1)
        n_model1 = normalize_model(model1)
        if not n_model1 or not n_name1:  # 比較できない行はスキップ
            continue

        dkey1   = digits_only(n_model1)
        pref6_1 = n_model1[:6]
        hit = None; reason = ""; nscore = 0.0

        # ---- Pass1: 型番 完全一致 ----
        cand = [items2[i] for i in by_exact.get(n_model1, [])]
        if cand:
            # 先に軽い model スコアで上位を少数に絞る → 名前比較
            cand.sort(key=lambda it: model_score_quick(n_model1, it["n_model"]), reverse=True)
            cand = cand[:min(len(cand), 20)]
            best, best_ns = None, -1.0
            for it in cand:
                ns = name_score(n_name1, it["n_name"])
                if ns > best_ns:
                    best_ns = ns; best = it
            if best and best_ns >= p1_name:
                hit = best; reason = "P1 exact"; nscore = best_ns

        # ---- Pass2: 型番 digitsOnly 一致 ----
        if not hit and dkey1 and dkey1 in by_digits:
            cand = [items2[i] for i in by_digits[dkey1]]
            # digits一致でもノイズが多い可能性→ modelスコアで上位50だけ名前比較
            cand.sort(key=lambda it: model_score_quick(n_model1, it["n_model"]), reverse=True)
            cand = cand[:min(len(cand), 50)]
            best, best_ns = None, -1.0
            for it in cand:
                ns = name_score(n_name1, it["n_name"])
                if ns > best_ns:
                    best_ns = ns; best = it
            if best and best_ns >= p2_name:
                hit = best; reason = "P2 digitsOnly"; nscore = best_ns

        # ---- Pass3: 型番 部分一致候補（上位Kだけ）----
        if not hit:
            cand_idx = set()
            if pref6_1 and pref6_1 in by_prefix:
                cand_idx.update(by_prefix[pref6_1])
            # 先頭4文字の似たprefixも拾う（粗い近傍）
            pref4 = n_model1[:4]
            if pref4:
                for k,v in by_prefix.items():
                    if k.startswith(pref4):
                        cand_idx.update(v)
                        if len(cand_idx) > p3_topk*3:
                            break
            # 候補に対して modelスコアで上位Kに絞る
            cand = [items2[i] for i in cand_idx] if cand_idx else []
            if not cand:
                # 最終保険：全体から modelスコア上位K（K小さめ）
                cand = sorted(items2, key=lambda it: model_score_quick(n_model1, it["n_model"]), reverse=True)[:min(len(items2), p3_topk)]
            else:
                cand.sort(key=lambda it: model_score_quick(n_model1, it["n_model"]), reverse=True)
                cand = cand[:min(len(cand), p3_topk)]

            best, best_ns = None, -1.0
            for it in cand:
                ns = name_score(n_name1, it["n_name"])
                if ns > best_ns:
                    best_ns = ns; best = it
            if best and best_ns >= p3_name:
                hit = best; reason = "P3 contains"; nscore = best_ns

        if hit:
            diff = hit["price"] - price1
            results.append([
                model1, name1, price1,
                hit["model"], hit["name"], hit["price"],
                diff, reason, f"{nscore:.2f}", r1, hit["row"]
            ])
        else:
            # 未一致 & 候補プレビュー（上位3）
            unmatched_rows.append([r1, model1, name1, price1, n_model1, dkey1, n_name1])
            # ざっくり modelスコアで上位を3件
            prev = sorted(items2, key=lambda it: model_score_quick(n_model1, it["n_model"]), reverse=True)[:3]
            for k, it in enumerate(prev, start=1):
                ms = model_score_quick(n_model1, it["n_model"])
                ns = name_score(n_name1, it["n_name"])
                preview_rows.append([
                    r1, model1, name1, f"cand{k}",
                    it["model"], it["name"], it["price"],
                    f"ms={ms:.2f} ns={ns:.2f}", it["row"]
                ])

    results.sort(key=lambda r: abs(r[6]), reverse=True)
    header_main = ["型番(1)","名称(1)","価格(1)","型番(2)","名称(2)","価格(2)","差分(2-1)","一致段階","name_score","row1","row2"]
    header_unm  = ["row1","型番(1)","名称(1)","価格(1)","n_model1","digits1","n_name1"]
    header_prev = ["row1","型番(1)","名称(1)","候補","型番(2)","名称(2)","価格(2)","スコア詳細","row2"]
    return results, header_main, unmatched_rows, header_unm, preview_rows, header_prev

# ===== Main =====
def main():
    ap = argparse.ArgumentParser(description="2シート比較（高速版：RapidFuzz）")
    ap.add_argument("--sheet-url", required=True)
    ap.add_argument("--sheet1", required=True)
    ap.add_argument("--sheet2", required=True)
    ap.add_argument("--sheet1-name-col", default="カード名")
    ap.add_argument("--sheet1-model-col", default="型番")
    ap.add_argument("--sheet1-price-col", default="O")
    ap.add_argument("--sheet2-name-col", default="カード名")
    ap.add_argument("--sheet2-model-col", default="型番")
    ap.add_argument("--sheet2-price-col", default="C")
    ap.add_argument("--out-sheet", default="差分比較")
    ap.add_argument("--creds", default=r"C:\Users\user\OneDrive\Desktop\credentials.json")
    ap.add_argument("--p1-name", type=float, default=0.40)
    ap.add_argument("--p2-name", type=float, default=0.55)
    ap.add_argument("--p3-name", type=float, default=0.70)
    ap.add_argument("--p3-topk", type=int, default=120)
    args = ap.parse_args()

    # 読み取り
    rcreds = Credentials.from_service_account_file(args.creds, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
    gr = gspread.authorize(rcreds)
    s1 = read_sheet(gr, args.sheet_url, args.sheet1)
    s2 = read_sheet(gr, args.sheet_url, args.sheet2)

    results, h_main, um_rows, h_unm, prev_rows, h_prev = compare_fast(
        s1, s2,
        args.sheet1_name_col, args.sheet1_model_col, args.sheet1_price_col,
        args.sheet2_name_col, args.sheet2_model_col, args.sheet2_price_col,
        args.p1_name, args.p2_name, args.p3_name, args.p3_topk
    )

    # 書き込み
    write_sheet(args.sheet_url, args.out_sheet, results, h_main, args.creds)
    write_sheet(args.sheet_url, "未一致_Sheet1", um_rows, h_unm, args.creds)
    write_sheet(args.sheet_url, "候補プレビュー", prev_rows, h_prev, args.creds)

    print(f"[OK] マッチ:{len(results)} / 未一致:{len(um_rows)}  →  出力: {args.out_sheet} / 未一致_Sheet1 / 候補プレビュー")

if __name__ == "__main__":
    main()
