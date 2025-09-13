# -*- coding: utf-8 -*-
"""
サービスアカウント経由でGoogleスプレッドシートを読み取り、
指定の列（0始まりIDX_*）から buylist.xlsx を生成します。

環境変数:
  GOOGLE_CREDS_JSON : サービスアカウントJSON（Secretsに保存した文字列）
  SHEET_ID          : スプレッドシートID
  SHEET_NAME        : タブ名（例: "シート1"）
  OUT_XLSX          : 出力パス（既定: buylist.xlsx）

列インデックス（0始まり）: ※スクリプト内の定数を必要に応じて変更
  IDX_NAME   : C列(2)
  IDX_PACK   : E列(4)
  IDX_CODE   : F列(5)
  IDX_RARITY : G列(6)
  IDX_BOOST  : H列(7)
  IDX_PRICE  : O列(14)
  IDX_IMGURL : Q列(16)

出力の列名は「日本語+英語の二重ヘッダ」を持たせて、
後段の gen_buylist.py がどちらでも解釈できるようにします。
"""
import os
import io
import sys
import json
import re
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ===== 列インデックス（0始まり） =====
IDX_NAME   = 2   # C
IDX_PACK   = 4   # E
IDX_CODE   = 5   # F
IDX_RARITY = 6   # G
IDX_BOOST  = 7   # H
IDX_PRICE  = 14  # O
IDX_IMGURL = 16  # Q

SHEET_ID   = os.getenv("SHEET_ID")
SHEET_NAME = os.getenv("SHEET_NAME", "シート1")
OUT_XLSX   = os.getenv("OUT_XLSX", "buylist.xlsx")
CREDS_JSON = os.getenv("GOOGLE_CREDS_JSON")

if not SHEET_ID:
    print("[ERR] SHEET_ID が未設定です", file=sys.stderr); sys.exit(2)
if not CREDS_JSON:
    print("[ERR] GOOGLE_CREDS_JSON が未設定です", file=sys.stderr); sys.exit(2)

# 認証
info = json.loads(CREDS_JSON)
scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
creds = service_account.Credentials.from_service_account_info(info, scopes=scopes)
sheets = build("sheets", "v4", credentials=creds).spreadsheets()

# ざっくり広い範囲を取得（必要に応じてA1表記を調整）
result = sheets.values().get(spreadsheetId=SHEET_ID, range=f"{SHEET_NAME}!A1:ZZ").execute()
values = result.get("values", [])
if not values:
    print("[ERR] シートが空です", file=sys.stderr); sys.exit(3)

# 1行目をヘッダーっぽく使いつつ、行ごとに安全に長さ合わせ
max_len = max(len(r) for r in values) if values else 0
norm = [r + [""]*(max_len - len(r)) for r in values]
header = norm[0] if norm else []
rows = norm[1:] if len(norm) > 1 else []

def safe_get(row, idx):
    return row[idx] if idx is not None and 0 <= idx < len(row) else ""

# =IMAGE("...") からURLだけを抜く（Q列想定）
IMG_RE = re.compile(r'(?i)^=IMAGE\("([^"]+)"')
def extract_image_url(s: str) -> str:
    if not s:
        return ""
    s = s.strip()
    m = IMG_RE.match(s)
    return m.group(1) if m else s

records = []
for r in rows:
    name   = safe_get(r, IDX_NAME)
    pack   = safe_get(r, IDX_PACK)
    code   = safe_get(r, IDX_CODE)
    rarity = safe_get(r, IDX_RARITY)
    boost  = safe_get(r, IDX_BOOST)
    price  = safe_get(r, IDX_PRICE)
    img    = extract_image_url(safe_get(r, IDX_IMGURL))

    # 価格を数値化（失敗時はNaN→後段で扱いやすい）
    try:
        price_num = float(str(price).replace(",", "").strip()) if str(price).strip() else None
    except:
        price_num = None

    records.append({
        # 日本語ヘッダ
        "カード名": name,
        "封入パックの型番": pack,
        "カードの型番": code,
        "レアリティ": rarity,
        "ブースト": boost,
        "買取金額": price_num,
        "画像URL": img,
        # 英語ヘッダ（後段が英語キー参照でもOKにする二重化）
        "name": name,
        "pack": pack,
        "code": code,
        "rarity": rarity,
        "boost": boost,
        "price": price_num,
        "image_url": img,
    })

df = pd.DataFrame.from_records(records, columns=[
    "カード名","封入パックの型番","カードの型番","レアリティ","ブースト","買取金額","画像URL",
    "name","pack","code","rarity","boost","price","image_url"
])

# Excelへ
with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as w:
    df.to_excel(w, index=False, sheet_name="シート1")

print(f"[OK] Downloaded -> {OUT_XLSX}  rows={len(df)}  sheet='{SHEET_NAME}'")
