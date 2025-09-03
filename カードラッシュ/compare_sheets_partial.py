# -*- coding: utf-8 -*-
"""
2シート比較（E+F融合キーで突合 / 価格乖離は no data / 高速版）
- 突合キー: シート1の E(Exp) + F(型番) を結合 → fused1
            CardRushは model から [exp][code] を抽出 → fused2
- 3段階マッチ（全て AND：融合キー＆名前）
  P1: 融合キー 完全一致 + name >= p1
  P2: 融合キー digitsOnly 一致 + name >= p2
  P3: 融合キー 近傍（prefix等）から上位K + name >= p3
- 価格の近さも加味（name/modelと重み付けで合成）
- 価格が「大きく乖離」している場合はカードラッシュ側価格を "no data" にして出力
- F(価格2) と G(差分) を赤/青で塗り分け（文字は黒固定）
"""

import argparse, re, unicodedata
from typing import Any, Dict, List, Tuple, Optional
import gspread
from google.oauth2.service_account import Credentials
from rapidfuzz import fuzz

# ========= 正規化 =========
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

def normalize_exp(s: Any) -> str:
    return re.sub(r"[^0-9a-z]", "", normalize_text(s))

def digits_only(s: str) -> str:
    return re.sub(r"\D", "", s)

def price_to_int(v: Any) -> int:
    if v is None: return 0
    if isinstance(v, (int, float)): return int(v)
    return int(re.sub(r"[^\d]", "", str(v)) or 0)

# ========= CardRush model の [exp][code] 抽出 =========
def split_exp_and_code_from_model(raw_model: str) -> Tuple[str, str]:
    """
    例: '24RP4DM1秘/DM1' -> normalize -> '24rp4dm1dm1'
         'dm' 出現位置を優先的に境界とみなし exp='24rp4', code='dm1dm1'
    """
    n = normalize_text(raw_model)
    alnum = re.sub(r"[^0-9a-z]", "", n)
    if not alnum:
        return "", ""
    m = re.search(r"dm\d*", alnum)
    if m:
        idx = m.start()
        return alnum[:idx] or "", alnum[idx:] or ""
    m2 = re.match(r"([0-9a-z]{2,}?)([a-z]{2}\d.*)", alnum)
    if m2:
        return m2.group(1), m2.group(2)
    return alnum[:6], alnum[6:]

# ========= 類似度/スコア =========
def name_score(a: str, b: str) -> float:
    a2, b2 = normalize_name(a), normalize_name(b)
    if not a2 or not b2: return 0.0
    return (fuzz.token_set_ratio(a2, b2) + fuzz.partial_ratio(a2, b2)) / 200.0

def model_score_fused(a_fused: str, b_fused: str) -> float:
    # 完全一致/包含は高スコア、それ以外は partial_ratio
    a2, b2 = normalize_model(a_fused), normalize_model(b_fused)
    if not a2 or not b2: return 0.0
    if a2 == b2: return 1.0
    if a2 in b2 or b2 in a2: return 0.95
    return fuzz.partial_ratio(a2, b2) / 100.0

def price_similarity(p1: int, p2: int, base: int = 500, cap: float = 1.0) -> float:
    denom = max(p1, base)
    d = abs(p2 - p1) / denom
    return max(0.0, 1.0 - min(d, cap))

# ========= 列ヘルパ =========
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

# ========= シートI/O =========
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
        # 数値フォーマット（価格列/差分）
        fmt("価格(1)", "#,##0"); fmt("価格(2)", "#,##0"); fmt("差分(2-1)", "#,##0;[Red]-#,##0")

        # 条件付き書式：F(価格2) / G(差分) を赤・青（文字色は黒固定）
        try:
            sheet_id = ws.id
            reqs = []
            # F列（価格2）
            reqs += [
                {"addConditionalFormatRule": {
                    "rule": {
                        "ranges": [{"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": last, "startColumnIndex": 5, "endColumnIndex": 6}],
                        "booleanRule": {"condition": {"type":"CUSTOM_FORMULA","values":[{"userEnteredValue":"=$F2>$C2"}]},
                                        "format": {"backgroundColor":{"red":1.0,"green":0.8,"blue":0.8},
                                                   "textFormat":{"foregroundColor":{"red":0,"green":0,"blue":0}}}}
                    },"index":0}},
                {"addConditionalFormatRule": {
                    "rule": {
                        "ranges": [{"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": last, "startColumnIndex": 5, "endColumnIndex": 6}],
                        "booleanRule": {"condition": {"type":"CUSTOM_FORMULA","values":[{"userEnteredValue":"=$F2<$C2"}]},
                                        "format": {"backgroundColor":{"red":0.8,"green":0.85,"blue":1.0},
                                                   "textFormat":{"foregroundColor":{"red":0,"green":0,"blue":0}}}}
                    },"index":0}},
            ]
            # G列（差分）
            reqs += [
                {"addConditionalFormatRule": {
                    "rule": {
                        "ranges": [{"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": last, "startColumnIndex": 6, "endColumnIndex": 7}],
                        "booleanRule": {"condition": {"type":"CUSTOM_FORMULA","values":[{"userEnteredValue":"=$G2>0"}]},
                                        "format": {"backgroundColor":{"red":1.0,"green":0.8,"blue":0.8},
                                                   "textFormat":{"foregroundColor":{"red":0,"green":0,"blue":0}}}}
                    },"index":0}},
                {"addConditionalFormatRule": {
                    "rule": {
                        "ranges": [{"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": last, "startColumnIndex": 6, "endColumnIndex": 7}],
                        "booleanRule": {"condition": {"type":"CUSTOM_FORMULA","values":[{"userEnteredValue":"=$G2<0"}]},
                                        "format": {"backgroundColor":{"red":0.8,"green":0.85,"blue":1.0},
                                                   "textFormat":{"foregroundColor":{"red":0,"green":0,"blue":0}}}}
                    },"index":0}},
            ]
            sh.batch_update({"requests": reqs})
        except Exception as e:
            print("[warn] 条件付き書式の追加に失敗:", e)

# ========= インデックス（CardRush_DM 側） =========
def build_index_sheet2(rows: List[List[str]], name_col: int, model_col: int, price_col: int):
    items: List[Dict[str,Any]] = []
    by_fused_exact: Dict[str, List[int]] = {}
    by_fused_digits: Dict[str, List[int]] = {}
    by_fused_prefix: Dict[str, List[int]] = {}
    for r_idx, row in enumerate(rows[1:], start=2):
        name = row[name_col] if name_col < len(row) else ""
        model = row[model_col] if model_col < len(row) else ""
        price = price_to_int(row[price_col] if price_col < len(row) else 0)
        n_name = normalize_name(name)
        exp2, code2 = split_exp_and_code_from_model(model)
        fused2 = normalize_model(exp2 + code2)  # CardRush側の融合キー
        dkey   = digits_only(fused2)
        pref6  = fused2[:6]
        item = {"row": r_idx, "name": name, "model": model, "price": price,
                "n_name": n_name, "fused": fused2, "dkey": dkey, "pref6": pref6}
        i = len(items); items.append(item)
        if fused2: by_fused_exact.setdefault(fused2, []).append(i)
        if dkey:   by_fused_digits.setdefault(dkey, []).append(i)
        if pref6:  by_fused_prefix.setdefault(pref6, []).append(i)
    return items, by_fused_exact, by_fused_digits, by_fused_prefix

# ========= 比較本体（融合キー + 価格乖離 no data） =========
def compare_fast_fused(
    s1_rows: List[List[str]], s2_rows: List[List[str]],
    s1_name_key: str, s1_model_key: str, s1_exp_key: str, s1_price_key: str,
    s2_name_key: str, s2_model_key: str, s2_price_key: str,
    p1_name=0.40, p2_name=0.55, p3_name=0.68, p3_topk=200,
    w_m=0.55, w_n=0.25, w_p=0.20,
    price_base=500, price_cap=1.0,
    no_data_ps=0.25,          # 価格類似度がこの値を下回ったら "no data"
    no_data_mult=1.5          # |diff| > no_data_mult * max(price1, price_base) でも "no data"
):
    h1 = s1_rows[0] if s1_rows else []; h2 = s2_rows[0] if s2_rows else []
    s1_name_col  = find_col_index(h1, s1_name_key)
    s1_model_col = find_col_index(h1, s1_model_key)
    s1_exp_col   = find_col_index(h1, s1_exp_key)
    s1_price_col = find_col_index(h1, s1_price_key)
    s2_name_col  = find_col_index(h2, s2_name_key)
    s2_model_col = find_col_index(h2, s2_model_key)
    s2_price_col = find_col_index(h2, s2_price_key)
    for label,val in {
        "sheet1 name":s1_name_col, "sheet1 model":s1_model_col, "sheet1 exp":s1_exp_col, "sheet1 price":s1_price_col,
        "sheet2 name":s2_name_col, "sheet2 model":s2_model_col, "sheet2 price":s2_price_col
    }.items():
        if val is None: raise ValueError(f"{label} 列が見つかりません。キーを確認してください。")

    # CardRush 側インデックス
    items2, by_exact, by_digits, by_prefix = build_index_sheet2(s2_rows, s2_name_col, s2_model_col, s2_price_col)

    results: List[List[Any]] = []
    unmatched_rows: List[List[Any]] = []
    preview_rows: List[List[Any]] = []

    def composite_total(ms: float, ns: float, p1: int, p2: int) -> float:
        ps = price_similarity(p1, p2, base=price_base, cap=price_cap)
        return w_m*ms + w_n*ns + w_p*ps

    for r1, row1 in enumerate(s1_rows[1:], start=2):
        name1  = row1[s1_name_col]  if s1_name_col  < len(row1) else ""
        model1 = row1[s1_model_col] if s1_model_col < len(row1) else ""
        exp1   = row1[s1_exp_col]   if s1_exp_col  < len(row1) else ""
        price1 = price_to_int(row1[s1_price_col]  if s1_price_col < len(row1) else 0)

        n_name1 = normalize_name(name1)
        fused1  = normalize_model(normalize_exp(exp1) + normalize_model(model1))  # ★ E+F 融合キー
        if not fused1 or not n_name1:
            continue

        dkey1   = digits_only(fused1)
        pref6_1 = fused1[:6]

        hit = None; reason = ""; debug = ""; matched_score = -1.0

        # ---- Pass1: 融合 完全一致 ----
        cand = [items2[i] for i in by_exact.get(fused1, [])]
        if cand:
            best, best_score = None, -1.0
            for it in cand[:30]:
                ns = name_score(n_name1, it["n_name"])
                if ns < p1_name: 
                    continue
                ms = 1.0
                score = composite_total(ms, ns, price1, it["price"])
                if score > best_score:
                    best_score, best, matched_score = score, it, score
                    debug = f"ms={ms:.2f} ns={ns:.2f} ps={price_similarity(price1,it['price'],price_base,price_cap):.2f}"
            if best:
                hit = best; reason = "P1 exact"

        # ---- Pass2: 融合 digitsOnly 一致 ----
        if not hit and dkey1 and dkey1 in by_digits:
            cand = [items2[i] for i in by_digits[dkey1]]
            # 融合キーで model スコア上位から
            cand.sort(key=lambda it: model_score_fused(fused1, it["fused"]), reverse=True)
            best, best_score = None, -1.0
            for it in cand[:80]:
                ns = name_score(n_name1, it["n_name"])
                if ns < p2_name:
                    continue
                ms = model_score_fused(fused1, it["fused"])
                score = composite_total(ms, ns, price1, it["price"])
                if score > best_score:
                    best_score, best, matched_score = score, it, score
                    debug = f"ms={ms:.2f} ns={ns:.2f} ps={price_similarity(price1,it['price'],price_base,price_cap):.2f}"
            if best:
                hit = best; reason = "P2 digitsOnly"

        # ---- Pass3: 融合 近傍（prefix等） ----
        if not hit:
            cand_idx = set()
            if pref6_1 and pref6_1 in by_prefix:
                cand_idx.update(by_prefix[pref6_1])
            # 先頭4が近い prefix も追加
            pref4 = fused1[:4]
            if pref4:
                for k, v in by_prefix.items():
                    if k.startswith(pref4):
                        cand_idx.update(v)
                        if len(cand_idx) > p3_topk*3: break

            cand = [items2[i] for i in cand_idx] if cand_idx else []
            if not cand:
                cand = sorted(items2, key=lambda it: model_score_fused(fused1, it["fused"]), reverse=True)[:min(len(items2), p3_topk)]
            else:
                cand.sort(key=lambda it: model_score_fused(fused1, it["fused"]), reverse=True)
                cand = cand[:min(len(cand), p3_topk)]

            best, best_score = None, -1.0
            for it in cand:
                ns = name_score(n_name1, it["n_name"])
                if ns < p3_name:
                    continue
                ms = model_score_fused(fused1, it["fused"])
                score = composite_total(ms, ns, price1, it["price"])
                if score > best_score:
                    best_score, best, matched_score = score, it, score
                    debug = f"ms={ms:.2f} ns={ns:.2f} ps={price_similarity(price1,it['price'],price_base,price_cap):.2f}"
            if best:
                hit = best; reason = "P3 contains"

        if hit:
            p2 = hit["price"]
            diff = p2 - price1
            ps  = price_similarity(price1, p2, base=price_base, cap=price_cap)
            # ★ 価格乖離が大きい場合は "no data"
            flag_no_data = (ps < no_data_ps) or (abs(diff) > no_data_mult * max(price1, price_base))
            if flag_no_data:
                results.append([
                    model1, name1, price1,
                    hit["model"], hit["name"], "no data",  # 価格(2) は "no data"
                    "",  # 差分は空欄
                    "no data", debug, r1, hit["row"]
                ])
            else:
                results.append([
                    model1, name1, price1,
                    hit["model"], hit["name"], p2,
                    diff, reason, debug, r1, hit["row"]
                ])
        else:
            unmatched_rows.append([r1, exp1, model1, name1, price1, fused1, dkey1, n_name1])
            # プレビュー（融合キーの model スコアで上位3）
            scored = []
            for it in items2:
                ms = model_score_fused(fused1, it["fused"])
                ns = name_score(n_name1, it["n_name"])
                sc = composite_total(ms, ns, price1, it["price"])
                scored.append((sc, ms, ns, it))
            scored.sort(key=lambda x: x[0], reverse=True)
            for k,(sc,ms,ns,it) in enumerate(scored[:3], start=1):
                preview_rows.append([
                    r1, model1, name1, f"cand{k}",
                    it["model"], it["name"], it["price"],
                    f"score={sc:.2f} ms={ms:.2f} ns={ns:.2f} ps={price_similarity(price1,it['price'],price_base,price_cap):.2f}",
                    it["row"]
                ])

    results.sort(key=lambda r: (0 if r[6]=="" else abs(r[6])), reverse=True)  # 差分空は末尾寄せ
    header_main = ["型番(1)","名称(1)","価格(1)","型番(2)","名称(2)","価格(2)","差分(2-1)","一致段階","スコア内訳","row1","row2"]
    header_unm  = ["row1","E(1)","F(1)","名称(1)","価格(1)","fused1","digits1","n_name1"]
    header_prev = ["row1","型番(1)","名称(1)","候補","型番(2)","名称(2)","価格(2)","スコア詳細","row2"]
    return results, header_main, unmatched_rows, header_unm, preview_rows, header_prev

# ========= Main =========
def main():
    ap = argparse.ArgumentParser(description="2シート比較（E+F融合/価格乖離 no data/高速）")
    ap.add_argument("--sheet-url", required=True)
    ap.add_argument("--sheet1", required=True)
    ap.add_argument("--sheet2", required=True)
    ap.add_argument("--sheet1-name-col", default="カード名")
    ap.add_argument("--sheet1-model-col", default="F")  # ★ 型番=F
    ap.add_argument("--sheet1-exp-col",   default="E")  # ★ Exp=E
    ap.add_argument("--sheet1-price-col", default="O")  # ★ 価格=O
    ap.add_argument("--sheet2-name-col", default="カード名")
    ap.add_argument("--sheet2-model-col", default="型番")
    ap.add_argument("--sheet2-price-col", default="C")
    ap.add_argument("--out-sheet", default="差分比較")
    ap.add_argument("--creds", default=r"C:\Users\user\OneDrive\Desktop\デュエマ買取表\カードラッシュ\credentials.json")
    # しきい値・候補幅
    ap.add_argument("--p1-name", type=float, default=0.40)
    ap.add_argument("--p2-name", type=float, default=0.55)
    ap.add_argument("--p3-name", type=float, default=0.65)
    ap.add_argument("--p3-topk", type=int, default=200)
    # 重み
    ap.add_argument("--w-model", type=float, default=0.55)
    ap.add_argument("--w-name",  type=float, default=0.25)
    ap.add_argument("--w-price", type=float, default=0.20)
    # 価格近さ
    ap.add_argument("--price-base", type=int, default=500)
    ap.add_argument("--price-cap",  type=float, default=1.0)
    # no data 判定
    ap.add_argument("--no-data-ps", type=float, default=0.25)
    ap.add_argument("--no-data-mult", type=float, default=1.5)
    args = ap.parse_args()

    # 読み取り
    rcreds = Credentials.from_service_account_file(args.creds, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
    gr = gspread.authorize(rcreds)
    s1 = read_sheet(gr, args.sheet_url, args.sheet1)
    s2 = read_sheet(gr, args.sheet_url, args.sheet2)

    results, h_main, um_rows, h_unm, prev_rows, h_prev = compare_fast_fused(
        s1, s2,
        args.sheet1_name_col, args.sheet1_model_col, args.sheet1_exp_col, args.sheet1_price_col,
        args.sheet2_name_col, args.sheet2_model_col, args.sheet2_price_col,
        args.p1_name, args.p2_name, args.p3_name, args.p3_topk,
        args.w_model, args.w_name, args.w_price,
        args.price_base, args.price_cap,
        args.no_data_ps, args.no_data_mult
    )

    # 書き込み
    write_sheet(args.sheet_url, args.out_sheet, results, h_main, args.creds)
    write_sheet(args.sheet_url, "未一致_Sheet1", um_rows, h_unm, args.creds)
    write_sheet(args.sheet_url, "候補プレビュー", prev_rows, h_prev, args.creds)

    print(f"[OK] マッチ:{len(results)} / 未一致:{len(um_rows)} → 出力: {args.out_sheet} / 未一致_Sheet1 / 候補プレビュー")

if __name__ == "__main__":
    main()
