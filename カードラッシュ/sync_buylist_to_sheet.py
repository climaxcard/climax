# -*- coding: utf-8 -*-
"""
Excel(buylist.xlsx) を Googleスプレッドシートの『シート1』に上書き同期
- Excelの先頭シートを丸ごと A1 から書き込み
- 数値/文字をそのまま転送（空は空文字）
- 依存: pandas, openpyxl, gspread, google-auth
"""

import argparse
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import math

def read_excel_as_values(xlsx_path: str):
    # 先頭シートを読み込み（インデックス列なし）
    df = pd.read_excel(xlsx_path, sheet_name=0, dtype=object, engine="openpyxl")
    # NaN を空文字に
    df = df.where(pd.notnull(df), "")
    # 2次元配列化（ヘッダ行も含める）
    header = list(df.columns)
    data = df.values.tolist()
    rows = [header] + [list(map(_to_cell, r)) for r in data]
    return rows

def _to_cell(v):
    # pandasの型をgspread更新に適した型へ
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        # float が整数っぽいなら int にしておくとキレイ
        if isinstance(v, float) and (math.isfinite(v)) and abs(v - round(v)) < 1e-9:
            return int(round(v))
        return v if math.isfinite(v) else ""
    return str(v) if v is not None else ""

def write_to_sheet(rows, sheet_url, sheet_name, creds_path):
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(creds_path, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_url(sheet_url)
    try:
        ws = sh.worksheet(sheet_name)
        ws.clear()
    except gspread.exceptions.WorksheetNotFound:
        # 行数・列数に余裕を持って作成
        ws = sh.add_worksheet(title=sheet_name, rows=max(100, len(rows)+10), cols=max(26, len(rows[0])+5))

    # 大きい表でも安定するように分割更新
    # ここは丸ごと1回でもOKだが、分割の方がタイムアウトに強い
    BATCH = 5000  # 行バッチ
    start = 0
    r = 1
    while start < len(rows):
        chunk = rows[start:start+BATCH]
        range_name = f"A{r}"
        ws.update(values=chunk, range_name=range_name)
        start += BATCH
        r += len(chunk)

    print(f"[OK] '{sheet_name}' を {len(rows)-1} 行で更新しました。")

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel-path", default=r"C:\Users\user\OneDrive\Desktop\デュエマ買取表\buylist.xlsx")
    ap.add_argument("--sheet-url", required=True)
    ap.add_argument("--sheet-name", default="シート1")
    ap.add_argument("--creds", default=r"C:\Users\user\OneDrive\Desktop\デュエマ買取表\カードラッシュ\credentials.json")
    args = ap.parse_args()

    rows = read_excel_as_values(args.excel_path)
    write_to_sheet(rows, args.sheet_url, args.sheet_name, args.creds)

if __name__ == "__main__":
    main()
