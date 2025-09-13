# -*- coding: utf-8 -*-
import os, sys, json, re, pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build

# 0始まりの列インデックス
IDX_NAME, IDX_PACK, IDX_CODE, IDX_RARITY, IDX_BOOST, IDX_PRICE, IDX_IMGURL = 2,4,5,6,7,14,16

SHEET_ID   = os.getenv("SHEET_ID")
SHEET_NAME = os.getenv("SHEET_NAME", "シート1")
OUT_XLSX   = os.getenv("OUT_XLSX", "buylist.xlsx")
CREDS_JSON = os.getenv("GOOGLE_CREDS_JSON")
if not SHEET_ID or not CREDS_JSON:
    print("[ERR] SHEET_ID / GOOGLE_CREDS_JSON 未設定", file=sys.stderr); sys.exit(2)

info = json.loads(CREDS_JSON)
scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
creds  = service_account.Credentials.from_service_account_info(info, scopes=scopes)
sheets = build("sheets", "v4", credentials=creds).spreadsheets()
values = sheets.values().get(spreadsheetId=SHEET_ID, range=f"{SHEET_NAME}!A1:ZZ").execute().get("values", [])
if not values: print("[ERR] 空シート", file=sys.stderr); sys.exit(3)

max_len = max(len(r) for r in values)
rows = [r + [""]*(max_len-len(r)) for r in values][1:]

def g(row,i): return row[i] if 0 <= i < len(row) else ""
import io
import pandas as pd

import re
IMG_RE = re.compile(r'(?i)^=IMAGE\("([^"]+)"')
def imgurl(s):
    s=(s or "").strip(); m=IMG_RE.match(s); return m.group(1) if m else s

recs=[]
for r in rows:
    name, pack, code, rarity, boost = g(r,IDX_NAME), g(r,IDX_PACK), g(r,IDX_CODE), g(r,IDX_RARITY), g(r,IDX_BOOST)
    price_raw = g(r,IDX_PRICE)
    try: price = float(str(price_raw).replace(",","").strip()) if str(price_raw).strip() else None
    except: price=None
    img = imgurl(g(r,IDX_IMGURL))
    recs.append({
        "カード名":name,"封入パックの型番":pack,"カードの型番":code,"レアリティ":rarity,"ブースト":boost,"買取金額":price,"画像URL":img,
        "name":name,"pack":pack,"code":code,"rarity":rarity,"boost":boost,"price":price,"image_url":img
    })

df = pd.DataFrame.from_records(recs, columns=[
    "カード名","封入パックの型番","カードの型番","レアリティ","ブースト","買取金額","画像URL",
    "name","pack","code","rarity","boost","price","image_url"
])

with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as w:
    df.to_excel(w, index=False, sheet_name="シート1")
print(f"[OK] Downloaded -> {OUT_XLSX}  rows={len(df)}  sheet='{SHEET_NAME}'")
