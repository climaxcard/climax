
# -*- coding: utf-8 -*-
r"""
デュエマ買取表 静的ページ生成（完成版・最適化）
- SP: アイコン1行横スクロール / LINE・Xは正方形フィット&1.2倍（PC/SPとも1.2xで中央トリミング）
- SP: 価格8桁→縮小 / 9桁→さらに縮小（はみ出し防止）/ 「買取カートへ」は価格の下で小さめ＆常にカード下辺固定
- PC: 4列固定（281×374.66）/ カード名・価格1.2倍 / 「買取カートへ」は価格の右横
- 見出し「デュエマ買取表」1.2倍 / ロゴとアクション群は中央寄せ
- 画像ON/OFFどちらの状態でも必ず描画されるrender()に刷新
- ページャ中央 / カート・ビューア開閉時のスクロールロック＋フォーカス管理（jump防止）
"""

from typing import Optional, List, Tuple
from pathlib import Path
from urllib.parse import urlparse, parse_qs
import pandas as pd
import html as html_mod
import unicodedata as ud
import base64, mimetypes, os, sys, hashlib, io, json, re, glob

try:
    import requests
    from PIL import Image
except Exception:
    requests = None
    Image = None

# ====== 入力パス ======
DEFAULT_EXCEL = r"C:\Users\user\Desktop\デュエマ買取表\buylist.xlsm"
ALT_EXCEL     = DEFAULT_EXCEL
FALLBACK_WINDOWS = DEFAULT_EXCEL

EXCEL_PATH = os.getenv("EXCEL_PATH", DEFAULT_EXCEL)
SHEET_NAME = os.getenv("SHEET_NAME", "シート1")
if len(sys.argv) > 1 and sys.argv[1]:
    EXCEL_PATH = sys.argv[1]

# ====== 出力・設定 ======
OUT_DIR    = Path(os.getenv("OUT_DIR", "docs"))
PER_PAGE   = int(os.getenv("PER_PAGE", "80"))
BUILD_THUMBS = os.getenv("BUILD_THUMBS", "0") == "1"

LIFF_ID = os.getenv("LIFF_ID", "2007983559-MKopknXd")
OA_ID   = os.getenv("OA_ID",   "@512nwjvn")
TAKUHAI_URL = os.getenv("TAKUHAI_URL", "https://climaxcard.github.io/climax2/#buy")

# ロゴ/アイコン（環境変数でパス指定可）
LOGO_FILE_ENV      = os.getenv("LOGO_FILE", "").strip()
X_ICON_FILE_ENV    = os.getenv("X_ICON_FILE", "").strip()
LINE_ICON_FILE_ENV = os.getenv("LINE_ICON_FILE", "").strip()
IG_ICON_FILE_ENV   = os.getenv("IG_ICON_FILE", "").strip()
TIKTOK_ICON_FILE_ENV = os.getenv("TIKTOK_ICON_FILE", "").strip()

# よく使う固定ディレクトリ（探索用）
FIXED_DIRS = [
    Path(r"C:\Users\user\OneDrive\Desktop\デュエマ買取表"),
    Path(r"C:\Users\user\OneDrive\Desktop\デュエマ買取表"),
    Path.cwd(),
    Path(__file__).parent if '__file__' in globals() else Path.cwd()
]

# 列位置フォールバック（0始まり）
IDX_NAME, IDX_PACK, IDX_CODE, IDX_RARITY, IDX_BOOST, IDX_PRICE, IDX_IMGURL = 2,4,5,6,7,14,16
# P列 = 16列目 → 0始まりで 15
IDX_PROMO = 15  # P列（チェック/フラグ）
THUMB_DIR = OUT_DIR / "assets" / "thumbs"
THUMB_W = 600

# ====== 共有ユーティリティ ======
def _cand_paths(names):
    for d in FIXED_DIRS + [Path.cwd()/ "assets"]:
        for n in names:
            p = d / n
            if p.exists() and p.is_file():
                yield p

def find_logo_path():
    if LOGO_FILE_ENV:
        p = Path(LOGO_FILE_ENV)
        if p.exists(): return p
    names = ["logo.png", "logo.svg", "Logo.png", "ロゴ.png", "shop-logo.png", "assets/logo.png", "assets/logo.svg"]
    return next(_cand_paths(names), None)

def find_icon_path(env_path: str, default_names: List[str]):
    if env_path:
        p = Path(env_path)
        if p.exists() and p.is_file(): return p
    return next(_cand_paths(default_names), None)

def file_to_data_uri(p: Optional[Path]) -> str:
    if not p: return ""
    mime = mimetypes.guess_type(str(p))[0] or ("image/svg+xml" if p.suffix.lower()==".svg" else "image/png")
    try:
        b64 = base64.b64encode(p.read_bytes()).decode("ascii")
        return f"data:{mime};base64,{b64}"
    except Exception:
        return ""

LOGO_URI = file_to_data_uri(find_logo_path())
X_ICON_URI    = file_to_data_uri(find_icon_path(X_ICON_FILE_ENV,    ["X.png","x.png","x-logo.png","assets/X.png"]))
LINE_ICON_URI = file_to_data_uri(find_icon_path(LINE_ICON_FILE_ENV, ["LINE.png","line.png","line-icon.png","assets/LINE.png"]))
INSTAGRAM_ICON_URI = file_to_data_uri(find_icon_path(IG_ICON_FILE_ENV, ["IG.png","Instagram.png","instagram.png","insta.png","assets/instagram.png","assets/IG.png"]))
TIKTOK_ICON_URI = file_to_data_uri(
    find_icon_path(TIKTOK_ICON_FILE_ENV, ["tiktok.png", "TikTok.png", "assets/tiktok.png", "assets/TikTok.png"])
)

# ====== 入力 ======
def _read_csv_auto(path: Path) -> pd.DataFrame:
    for enc in ("utf-8-sig","cp932","utf-8"):
        try: return pd.read_csv(path, encoding=enc)
        except Exception: pass
    return pd.read_csv(path)

def _normalize_two_header_layout(df: pd.DataFrame) -> pd.DataFrame:
    try:
        cand=[]; m=min(12,len(df))
        for i in range(m):
            row=df.iloc[i].astype(str).tolist()
            if "display_name" in row and "cardnumber" in row: cand.append(i)
        if not cand: return df
        hdr=cand[0]; start=hdr+2
        df2=df.iloc[start:].copy(); df2.columns=df.iloc[hdr].tolist()
        return df2.reset_index(drop=True)
    except Exception:
        return df

def _resolve_input(pref: Optional[str]) -> Path:
    cands=[]
    if pref: cands.append(Path(pref))
    cands += [Path("buylist.csv"), Path("buylist.xlsm"), Path(ALT_EXCEL), Path(FALLBACK_WINDOWS)]
    for p in cands:
        try:
            if p.exists() and p.is_file(): return p
        except Exception: pass
    files = sorted([Path(p) for p in glob.glob("*.csv")]+[Path(p) for p in glob.glob("*.xlsm")],
                   key=lambda x: x.stat().st_mtime, reverse=True)
    if files: return files[0]
    raise FileNotFoundError("CSV/Excel が見つかりません。")

def load_buylist_any(path_hint: str, sheet_name: Optional[str]) -> pd.DataFrame:
    p=_resolve_input(path_hint)
    if p.suffix.lower()==".csv":
        df0=_read_csv_auto(p); return _normalize_two_header_layout(df0)
    try:
        if sheet_name:
            df0=pd.read_excel(p, sheet_name=sheet_name, header=None, engine="openpyxl")
        else:
            xls=pd.ExcelFile(p, engine="openpyxl")
            df0=pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=None, engine="openpyxl")
    except Exception:
        xls=pd.ExcelFile(p, engine="openpyxl")
        df0=pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=None, engine="openpyxl")
    return _normalize_two_header_layout(df0)

df_raw = load_buylist_any(EXCEL_PATH, SHEET_NAME)
from datetime import datetime

try:
    # 実際に使われる入力ファイルのパスを解決して更新時刻を取得
    INPUT_FILE = _resolve_input(EXCEL_PATH)
    UPDATED_AT = datetime.fromtimestamp(INPUT_FILE.stat().st_mtime).strftime("%Y/%m/%d %H:%M")
    UPDATED_LABEL = f"最終更新：{UPDATED_AT}"
except Exception:
    UPDATED_LABEL = ""

# ====== テキスト整形 ======
SEP_RE = re.compile(r"[\s\u30FB\u00B7·/／\-_—–−]+")

def clean_text(s: pd.Series) -> pd.Series:
    s=s.astype(str)
    s=s.str.replace(r'(?i)^\s*nan\s*$', '', regex=True)
    s=s.replace({"nan":"","NaN":"","None":"","NONE":"","null":"","NULL":"","nil":"","NIL":""})
    return s.fillna("").str.strip()

def to_int_series(s: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce").round().astype("Int64")
    s=s.astype(str).str.replace(r"[^\d\.\-,]", "", regex=True).str.replace(",", "", regex=False)
    return pd.to_numeric(s, errors="coerce").round().astype("Int64")

def to_bool_series(s: pd.Series) -> pd.Series:
    TRUE_SET  = {"true","1","yes","y","on","☑","✓","✔","◯","○","レ","済"}
    FALSE_SET = {"false","0","no","n","off","","none","nan","null"}

    def _one(v):
        # ネイティブ bool / 数値はそのまま判定
        if isinstance(v, bool):
            return v
        if isinstance(v, (int, float)):
            try:
                return float(v) > 0
            except Exception:
                return False

        # 文字列は正規化して判定
        t = str(v).strip().lower()
        if t in TRUE_SET:
            return True
        if t in FALSE_SET:
            return False

        # Excel の TRUE/FALSE（大文字）など
        if t == "true":
            return True
        if t == "false":
            return False

        # それ以外は安全側で False
        return False

    return s.map(_one)

def detail_to_img(val: str) -> str:
    if not isinstance(val,str): return ""
    s=val.strip().replace("＠","@").replace("＂",'"').replace("＇","'")
    m=re.search(r'@?IMAGE\s*\(\s*["\']\s*(https?://[^"\']+)\s*["\']', s, flags=re.IGNORECASE)
    if m: return m.group(1).strip()
    m=re.search(r'^[=]?\s*["\']\s*(https?://[^"\']+)\s*["\']\s*$', s)
    if m: return m.group(1).strip()
    m=re.search(r'(https?://[^\s"\')]+)', s)
    if m: return m.group(1).strip()
    if s.lower().startswith(("http://","https://")): return s
    parsed=urlparse(s)
    if "id=" in s:
        qs=parse_qs(parsed.query)
        id_val=qs.get("id",[parsed.path.split("/")[-1]])[0]
        if id_val: return f"https://dm.takaratomy.co.jp/wp-content/card/cardimage/{id_val}.jpg"
    slug=(parsed.path.split("/")[-1] or "").strip()
    if slug: return f"https://dm.takaratomy.co.jp/wp-content/card/cardimage/{slug}.jpg"
    return ""

def nfkc_lower(s:str)->str: return ud.normalize("NFKC", s or "").lower()
def kata_to_hira(text:str)->str: return "".join(chr(ord(ch)-0x60) if "ァ"<=ch<="ン" else ch for ch in text)

def normalize_for_search_py(s:str)->str:
    s0 = nfkc_lower(s)
    # ▼ ここを拡張
    s0 = (s0
          .replace("complex","こんぷれっくす")
          .replace("c0br4","こぶら")
          .replace("伝説","でんせつ")
          .replace("真気楼","ぺてんしー")
          .replace("墓碑","えぴたふ")
          .replace("西方","びすまるく")
    )
    s0 = kata_to_hira(s0)
    s0 = SEP_RE.sub("", s0)
    return s0


def searchable_row_py(row: pd.Series) -> str:
    parts=[row.get(k,"") for k in ("name","code","pack","rarity","booster")]
    return normalize_for_search_py(" ".join(map(str, parts)))

def get_col(df: pd.DataFrame, names: List[str], fallback_idx: Optional[int]):
    for nm in names:
        if nm in df.columns: return df[nm]
    if fallback_idx is not None and fallback_idx < df.shape[1]:
        return df.iloc[:, fallback_idx]
    return pd.Series([""]*len(df), index=df.index)

S_NAME   = get_col(df_raw, ["display_name","商品名"],            IDX_NAME)
S_PACK   = get_col(df_raw, ["expansion","エキスパンション"],      IDX_PACK)
S_CODE   = get_col(df_raw, ["cardnumber","カード番号"],           IDX_CODE)
S_RARITY = get_col(df_raw, ["rarity","レアリティ"],               IDX_RARITY)
S_BOOST  = get_col(df_raw, ["pack_name","封入パック","パック名"],  IDX_BOOST)
S_PRICE  = get_col(df_raw, ["buy_price","買取価格"],             IDX_PRICE)
S_IMGURL = get_col(df_raw, ["allow_auto_print_label","画像URL"],  IDX_IMGURL)
S_PROMO  = get_col(df_raw, ["promo","強化","チェック","check","flag"], IDX_PROMO)

df = pd.DataFrame({
    "promo":   to_bool_series(S_PROMO),
    "name":    clean_text(S_NAME),
    "pack":    clean_text(S_PACK),
    "code":    clean_text(S_CODE),
    "rarity":  clean_text(S_RARITY),
    "booster": clean_text(S_BOOST),
    "price":   to_int_series(S_PRICE) if len(S_PRICE) else pd.Series([None]*len(df_raw)),
    "image":   clean_text(S_IMGURL).map(detail_to_img),
})
df = df[~df["name"].str.match(r"^Unnamed", na=False)]
df = df[df["name"].str.strip()!=""].reset_index(drop=True)
df["s"] = df.apply(searchable_row_py, axis=1)

# ====== サムネ生成（任意） ======
def url_to_hash(u:str)->str: return hashlib.md5(u.encode("utf-8")).hexdigest()
def ensure_thumb(url: str) -> Optional[str]:
    if not url: return None
    THUMB_DIR.mkdir(parents=True, exist_ok=True)
    fname = url_to_hash(url) + ".webp"; outp = THUMB_DIR / fname
    if outp.exists(): return f"assets/thumbs/{fname}"
    if not (requests and Image): return None
    try:
        r=requests.get(url, timeout=12, headers={"User-Agent":"Mozilla/5.0"}); r.raise_for_status()
        im=Image.open(io.BytesIO(r.content)).convert("RGB")
        w,h=im.size
        if w<=0 or h<=0: return None
        new_h=int(h * THUMB_W / w)
        im=im.resize((THUMB_W, max(1,new_h)), Image.LANCZOS)
        outp.parent.mkdir(parents=True, exist_ok=True)
        im.save(outp, "WEBP", quality=60, method=6)
        return f"assets/thumbs/{fname}"
    except Exception:
        return None

if BUILD_THUMBS: df["thumb"] = df["image"].map(ensure_thumb)
else:            df["thumb"] = ""

# ====== ペイロード ======
def build_payload(df: pd.DataFrame) -> Tuple[str, str]:
    # 欠損カラムの補完（priceはNone, promoはFalse, 他は空文字）
    for c in ["name","pack","code","rarity","booster","price","image","thumb","s","promo"]:
        if c not in df.columns:
            if c == "price":
                df[c] = None
            elif c == "promo":
                df[c] = False
            else:
                df[c] = ""

    records = []
    for rec in df[["name","pack","code","rarity","booster","price","image","thumb","s","promo"]].to_dict(orient="records"):
        # 価格の正規化
        price = rec.get("price", None)
        try:
            price = None if pd.isna(price) else int(price)
        except Exception:
            price = None

        # 強化フラグの正規化（True/1/☑/✓/✔/◯ などを True 扱い）
        promo_raw = rec.get("promo")
        promo = False
        if isinstance(promo_raw, bool):
            promo = promo_raw
        elif isinstance(promo_raw, (int, float)):
            promo = float(promo_raw) > 0
        else:
            t = str(promo_raw).strip().lower()
            promo = t in {"true","1","yes","y","on","☑","✓","✔","◯","○","レ","済"}

        records.append({
            "n": rec.get("name",""),
            "p": rec.get("pack",""),
            "c": rec.get("code",""),
            "r": rec.get("rarity",""),
            "b": rec.get("booster",""),
            "pr": price,
            "i": rec.get("image",""),
            "t": rec.get("thumb",""),
            "s": rec.get("s",""),
            "k": 1 if promo else 0,   # ← 強化フラグ
        })

    payload = json.dumps(records, ensure_ascii=False, separators=(",",":"))
    ver = hashlib.md5(payload.encode("utf-8")).hexdigest()[:8]
    return ver, payload


CARDS_VER, CARDS_JSON = build_payload(df)

# ====== CSS ======
base_css = """

/* -------- Base / Variables -------- */
*{box-sizing:border-box}
:root{
  --bg:#fff; --panel:#fff; --border:#e5e7eb; --accent:#e11d48; --text:#111; --muted:#6b7280;
  --header-h:132px; --pane-w:760px;

  /* Radius */
  --radius-card:12px; --radius-btn:12px; --radius-img:12px;

  /* Pager */
  --pager-font:14px; --pager-pad-y:8px; --pager-pad-x:12px; --pager-radius:12px;

  /* Price width (base) */
  --price-w:9.5ch;

  /* SP専用（必要箇所だけ上書きされる） */
  --price-w-img:8.7ch; --btn-w-img:100%;
  --price-w-list:8.5ch; --btn-w-list:100%;
  --sp-btn-h:20px; --sp-foot-h:calc(var(--sp-btn-h) + 20px); --sp-body-h:92px; --sp-list-body-h:120px;

  /* Cart mini buttons */
  --cart-btn-h:26px; --cart-btn-font:13px; --cart-btn-pad-x:10px;
}

html{scrollbar-gutter:stable both-edges}
body{
  margin:0; color:var(--text); background:var(--bg);
  font-family:Inter,system-ui,'Noto Sans JP',sans-serif;
  padding-top:calc(var(--header-h) + 8px);
  overscroll-behavior-y:contain;
}

/* -------- Header -------- */
header{
  position:fixed; inset:0 0 auto 0; z-index:1000; background:#fff;
  border-bottom:1px solid var(--border); box-shadow:0 2px 10px rgba(0,0,0,.04);
  padding:10px 12px;
}
.header-wrap{
  max-width:1440px; margin:0 auto; width:100%;
  display:grid; gap:8px;
  grid-template-columns:auto 1fr auto;
  grid-template-areas:"logo title actions";
  align-items:center; place-items:center; place-content:center;
}
.brand-left{ grid-area:logo; display:flex; align-items:center; justify-content:center; margin:0 auto; }
.brand-left img{ height:60px; max-height:60px; width:auto; object-fit:contain; display:block; }

.center-ttl{
  grid-area:title; font-weight:1000; color:#111; text-align:center;
  font-size:clamp(24px,3.2vw,32px); line-height:1.1; margin:0 10px;
  display:flex; flex-direction:column; align-items:center; gap:2px; margin-top:12px !important;
}
.center-ttl .ttl{ margin:0; line-height:1; }
.center-ttl .sub{ font-size:35% !important; margin-top:0 !important; line-height:1; }

/* Actions */
.actions{
  grid-area:actions; display:flex; align-items:center; justify-content:center; gap:8px;
  flex-wrap:nowrap; overflow-x:auto; transform-origin:center;
}
.iconbtn,.iconimg{
  display:inline-flex; align-items:center; gap:8px; flex:0 0 auto;
  border:1px solid var(--border); background:#fff; color:#111;
  border-radius:var(--radius-btn); padding:10px 14px; font-size:14px; line-height:1; height:40px;
  text-decoration:none; transition:transform .12s ease, background .12s ease;
}
.iconbtn:hover,.iconimg:hover{ background:#f9fafb; transform:translateY(-1px) }
.iconbtn svg{ width:18px; height:18px; display:block; color:#111; }
.iconbtn span{ white-space:nowrap; }
.iconimg{ width:40px; height:40px; padding:0; overflow:hidden; border-radius:10px; aspect-ratio:1/1; line-height:0; position:relative; }
.iconimg img{ width:100%; height:100%; object-fit:cover; display:block; aspect-ratio:1/1; }
/* LINE / X は 1.4x 中央トリミング */
.iconimg--line img,
.iconimg--x img {
  width: 140%;
  height: 140%;
  position: absolute;
  top: 50%;
  left: 50%;
  transform: translate(-50%, -50%);
}
.iconimg--tiktok img {
  width: 130%;
  height: 130%;
  position: absolute;
  top: 50%;
  left: 50%;
  transform: translate(-50%, -50%);
}

/* Instagram は 1.1x 中央トリミング */
.iconimg--ig img {
  width: 110%;
  height: 110%;
  position: absolute;
  top: 50%;
  left: 50%;
  transform: translate(-50%, -50%);
}

/* -------- Layout / Notice -------- */
.wrap{ max-width:1440px; margin:0 auto; padding:10px 12px; }
.wrap > .controls, .wrap > nav.simple, .wrap > #grid, .wrap > .note { justify-self:center; }

.notice{
  max-width:var(--pane-w); width:100%; margin:8px auto 12px;
  border:1px solid var(--border); background:#fff; border-radius:12px; overflow:hidden;
  box-shadow:0 4px 18px rgba(0,0,0,.04);
}
.notice-hd{ display:flex; align-items:center; justify-content:space-between; gap:8px; padding:8px 10px; background:#fff; }
.notice-title{ font-weight:800; font-size:14px; }
.notice-toggle{ border:1px solid var(--border); background:#fff; color:#111; font-size:12px; padding:6px 10px; border-radius:10px; cursor:pointer; }
.notice.open .notice-toggle{ background:#fff1f3; border-color:var(--accent); color:var(--accent); box-shadow:0 0 0 2px rgba(225,29,72,.12) inset; }
.notice-body{ max-height:0; overflow:hidden; transition:max-height .25s ease; border-top:1px dashed #e5e7eb; background:#fff; }
.notice.open .notice-body{ max-height:1000px; }
.notice-body ul{ margin:0; padding:10px 14px; display:grid; gap:6px; font-size:13px; line-height:1.5; }
.notice-body li{ list-style:disc inside; }
.notice a{ color:#2563eb; text-decoration:underline; }

/* -------- Controls (search) -------- */
.controls{
  display:grid; grid-template-columns:1fr 1fr; grid-template-areas:"q1 q2" "q3 q4" "acts acts";
  gap:14px; align-items:center; justify-content:center;
  max-width:var(--pane-w); width:100%; margin:12px auto 18px;
  background:#fff; border:1px solid var(--border); border-radius:12px; padding:16px 18px; box-shadow:0 4px 18px rgba(0,0,0,.04);
}
#nameQ{grid-area:q1} #codeQ{grid-area:q2} #packQ{grid-area:q3} #rarityQ{grid-area:q4}
.controls .btns{ grid-area:acts; display:flex; gap:10px; flex-wrap:wrap; justify-content:center; }
input.search{
  background:#fff; border:1px solid var(--border); color:#111; border-radius:var(--radius-btn);
  padding:12px 14px; font-size:16px; outline:none; width:100%;
}
input.search::placeholder{ color:#9ca3af }
.btn{
  background:#fff; border:1px solid var(--border); color:#111; border-radius:var(--radius-btn);
  padding:12px 14px; font-size:16px; cursor:pointer; white-space:nowrap;
}
.btn:hover{ background:#f9fafb; }
.controls .btns .btn.active,
.controls .btns .btn[aria-pressed="true"]{
  border-color:var(--accent) !important; color:var(--accent) !important; background:#fff1f3;
  box-shadow:0 0 0 2px rgba(225,29,72,.15) inset; font-weight:800;
}

/* -------- Grid / Cards -------- */
.grid{ margin:8px auto; width:100%; display:grid; justify-content:center; }
/* ベース（汎用） */
.grid{ margin:8px auto; width:100%; display:grid; justify-content:center; }
.grid.grid-img{ display:grid; gap:8px; justify-content:center; }
.grid.grid-list{ display:grid; gap:8px; justify-content:center; }

/* PC(>=1024px) の列数だけ明示 */
@media (min-width:1024px){
  .grid.grid-img{ grid-template-columns:repeat(4, minmax(0,1fr)); }
  .grid.grid-list{ grid-template-columns:repeat(4, minmax(0,1fr)); }
}

/* スマホ(<=700px) の列数 */
@media (max-width:700px){
  .grid.grid-img{  grid-template-columns:repeat(4, minmax(0,1fr)); } /* 画像ONは4列のままならそのまま */
  .grid.grid-list{ grid-template-columns:repeat(2, minmax(0,1fr)); } /* ← 画像OFFは2列 */
}
.card{ background:var(--panel); border:1px solid var(--border); border-radius:var(--radius-card); overflow:hidden; }
.th{ aspect-ratio:63/88; background:#f3f4f6; cursor:zoom-in }
.th img{ width:100%; height:100%; object-fit:contain !important; display:block; background:#f3f4f6 }

.b{ padding:10px; display:flex; flex-direction:column; min-height:160px; }
.n{ font-size:15px; font-weight:800; line-height:1.25; margin:0 0 6px; color:#111; display:flex; gap:6px; align-items:flex-start; flex-wrap:wrap; word-break:break-word; }
.n .ttl{ flex:1 1 100%; min-width:0; display:-webkit-box; -webkit-line-clamp:3; -webkit-box-orient:vertical; overflow:hidden; }
.n .code{ order:2; margin-top:2px; font-weight:700; font-size:12px; color:#374151; background:#f3f4f6; border:1px solid #e5e7eb; border-radius:8px; padding:1px 6px; white-space:nowrap; }
.meta{ font-size:12px; color:var(--muted); word-break:break-word }

/* Price & Button */
.foot{ margin-top:auto; display:flex; flex-direction:column; gap:8px; align-items:center; justify-content:center; }
.pricebar{ display:flex; align-items:center; gap:10px; justify-content:center; }
.mx{ font-weight:1000; color:var(--accent); font-size:clamp(18px,2.0vw,22px); line-height:1; white-space:nowrap; display:inline-block; max-width:100%; font-variant-numeric:tabular-nums; letter-spacing:-0.02em; }
.mx.shrink{ font-size:clamp(13px,1.8vw,16px); }
.mx.ultra{ font-size:clamp(11px,1.6vw,13px); letter-spacing:-0.04em; }
.mx.tiny{  font-size:clamp(10px,3.0vw,13px); letter-spacing:-0.04em; }
.pricebar .btn.mini{ padding:6px 10px; font-size:12px; border-radius:10px; }

/* Pager */
nav.simple{ margin:16px 0; }
nav.simple .pager{ display:flex; justify-content:center; align-items:center; gap:6px !important; transform:none !important; }
nav.simple .num, nav.simple a.num, nav.simple button.num, nav.simple a, nav.simple button{
  color:#111; background:#fff; border:1px solid var(--border);
  font-size:var(--pager-font) !important;
  padding:var(--pager-pad-y) var(--pager-pad-x) !important;
  border-radius:var(--pager-radius) !important; text-decoration:none; white-space:nowrap;
  line-height:1 !important; cursor:pointer;
}
nav.simple .num[aria-current="page"]{ background:#111; color:#fff; border-color:#111; cursor:default; }
nav.simple .ellipsis{ border:none; background:transparent; cursor:default; padding:0 2px; }

/* -------- Overlays (Viewer / Cart) -------- */
.viewer,.cart-modal{
  position:fixed; inset:0; z-index:2000; display:none; align-items:center; justify-content:center;
  background:rgba(17,24,39,.5); padding:16px;
}
.viewer.show,.cart-modal.show{ display:flex; }
.viewer .vc{ position:relative; max-width:min(960px,95vw); max-height:90vh; background:#fff; border-radius:var(--radius-card); padding:8px; box-shadow:0 10px 30px rgba(0,0,0,.2); }
#viewerImg{ display:block; max-width:100%; max-height:80vh; height:auto; width:auto; margin:auto; }
#viewerClose{ position:absolute; top:8px; right:8px; border:1px solid var(--border); background:#fff; border-radius:var(--radius-btn); padding:6px 10px; font-size:14px; cursor:pointer; }

.cart-sheet{ width:min(680px,96vw); max-height:86vh; background:#fff; color:#111; border:1px solid var(--border); border-radius:var(--radius-card); overflow:hidden; box-shadow:0 12px 40px rgba(0,0,0,.25); display:flex; flex-direction:column; }
.cart-head{ display:flex; align-items:center; justify-content:space-between; padding:12px 14px; border-bottom:1px solid var(--border); background:#fafafa; }
.cart-list{ padding:10px 12px; overflow:auto; flex:1 1 auto; }
.cart-row{ display:grid; grid-template-columns:1fr auto auto auto; gap:8px; align-items:center; border-bottom:1px dashed #e5e7eb; padding:8px 0; }
.cart-row .nm{ font-weight:700; min-width:0; }
.cart-row .pr{ white-space:nowrap; color:#374151; }
.cart-row .qty{ display:flex; align-items:center; gap:8px; }
.cart-foot{ border-top:1px solid var(--border); padding:10px 12px; display:flex; align-items:center; gap:10px; }
.cart-act{ display:flex; gap:8px; flex-wrap:wrap; margin-left:auto; }
#cartTotal{ white-space:nowrap; font-variant-numeric:tabular-nums; }

/* Mini buttons (±/削除 共通) */
.cart-modal .cart-row .btn.mini{
  height:var(--cart-btn-h) !important; line-height:var(--cart-btn-h) !important;
  font-size:var(--cart-btn-font) !important; padding:0 var(--cart-btn-pad-x) !important; border-radius:6px !important;
}
.cart-modal .cart-row .qty{ gap:6px; }
.cart-modal .cart-row .qty .btn.mini{ width:var(--cart-btn-h) !important; padding:0 !important; }
.cart-modal .cart-row .btn.mini.rm{ width:auto !important; min-width:0 !important; white-space:nowrap !important; justify-self:end; padding:0 var(--cart-btn-pad-x) !important; opacity:.9; }
.cart-modal .cart-row .btn.mini.rm:hover{ opacity:1; }

button, .btn, .iconbtn, .notice-toggle{ touch-action:manipulation; -webkit-tap-highlight-color:transparent; user-select:none; }

/* Shared Radius (late) */
.card, .grid .card, .cart-sheet, .viewer .vc{ border-radius:var(--radius-card) !important; }
.btn, .iconbtn, input.search, nav.simple a, nav.simple button, #viewerClose{ border-radius:var(--radius-btn) !important; }
.th, .iconimg{ border-radius:var(--radius-img) !important; }

/* -------- Desktop (>=1024px) -------- */
@media (min-width:1024px){
  .header-wrap{ max-width:1100px; display:grid; grid-template-columns:1fr auto 1fr; grid-template-areas:"logo title actions"; align-items:center; gap:16px; }
  .brand-left{ grid-area:logo; justify-self:end; } .actions{ grid-area:actions; justify-self:start; }
  .center-ttl{ justify-self:center; text-align:center; margin:0; font-size:clamp(48px,4.5vw,64px) !important; line-height:1.15; margin-top:20px !important; }
  .brand-left img{ height:110px; max-height:110px; }

  /* PC: 画像ONは固定 281px × 4列 */
  .grid.grid-img{ grid-template-columns:repeat(4,281px); justify-content:center; gap:12px; }
  .grid.grid-img .card{ width:281px; }
  .grid.grid-img .th{ width:281px; height:374.66px; aspect-ratio:auto; overflow:hidden; background:#f3f4f6; }

  /* PC: 画像OFFは流体4列（1行に収まる） */
  .grid.grid-list{ grid-template-columns:repeat(4,minmax(0,1fr)); max-width:1200px; width:100%; margin:8px auto; gap:10px; }
  .grid.grid-list .card{ width:auto; }
  .grid.grid-list .b{ padding:8px; min-height:auto; }
  .grid.grid-list .n{ font-size:15px; margin:0 0 2px; }
  .grid.grid-list .card .foot{ gap:8px !important; }
  .grid.grid-list .card .btn.mini{ padding:0 10px; font-size:12.5px; width:var(--btn-w-list); box-sizing:border-box; }

  .b{ min-height:118px; } .n{ font-size:18px; margin:0 0 3px; }
  .mx{ font-size:30px; } .mx.shrink{ font-size:21px; }
  .foot{ align-items:center; } .pricebar{ flex-direction:row; gap:12px; }
}

/* -------- Mobile (<=700px) -------- */
@media (max-width:700px){
  :root{
    --header-h:144px;
    --sp-btn-h:20px; --sp-foot-h:calc(var(--sp-btn-h) + 20px); --sp-body-h:88px; --sp-list-body-h:120px;
    --price-w-img:8.7ch; --btn-w-img:100%;
    --price-w-list:10.5ch; /* 価格が見切れないよう少し広め（必要なら調整） */
  }

  /* Header sizing */
  .header-wrap{ grid-template-columns:auto 1fr; grid-template-areas:"logo title" "actions actions"; justify-content:center; justify-items:center; }
  .brand-left{ justify-self:center; }
  .brand-left img{ height:48px; max-height:48px; margin:auto; margin-left:22px; transform:scale(1.1); transform-origin:center; }
  .center-ttl{ margin-top:10px !important; font-size:clamp(24px,7.2vw,34px); }
  .wrap{ padding:6px; }

  .actions{ transform:scale(.9); gap:6px; overflow-x:auto; flex-wrap:nowrap; }
  .iconbtn,.iconimg{ height:32px; }
  .iconbtn{ padding:5px 9px; font-size:12px; } .iconbtn svg{ width:15px; height:15px; }
  .iconimg{ width:32px; height:32px; border-radius:8px; }
  .iconbtn--cart{ flex:0 0 auto; }

  /* Controls: ボタン列は1行固定＋横スワイプ（画像ON/OFFも2行目に落ちない） */
  .controls{ gap:8px; padding:8px 10px; }
  input.search{ padding:10px 12px; font-size:15px; }
  .controls .btns{ gap:6px; flex-wrap:nowrap; overflow-x:auto; white-space:nowrap; -webkit-overflow-scrolling:touch; scrollbar-width:none; }
  .controls .btns::-webkit-scrollbar{ display:none; }
  .controls .btns .btn{ flex:0 0 auto; padding:8px 10px; font-size:14px; }

  /* Grid gaps */
  .grid.grid-img,.grid.grid-list{ column-gap:3px; row-gap:1px; grid-template-columns:repeat(4,minmax(0,1fr)); }
  .card{ transform:none; }
  .th{ aspect-ratio:63/88; }

  /* Pager */
  nav.simple a, nav.simple button{ padding:6px 10px; font-size:12px; }

  /* Card body base (SP) */
  .b{ padding:6px; min-height:auto; }
  .n{ margin-bottom:2px; font-size:12px; line-height:1.2; }
  .n .ttl{ font-size:12px; -webkit-line-clamp:3; }
  .n .code{ font-size:10px; padding:0 4px; }

  /* ▼ 画像ON：本文固定高＋フッター下辺固定（価格↑／ボタン↓） */
  .grid.grid-img .card{ display:flex; flex-direction:column; }
  .grid.grid-img .card .b{ position:relative; height:var(--sp-body-h) !important; padding:6px 6px calc(var(--sp-foot-h) + 4px) !important; }
  .grid.grid-img .card .foot{
    position:absolute; left:6px; right:6px; bottom:4px; height:var(--sp-foot-h) !important;
    display:grid; grid-template-rows:min-content var(--sp-btn-h) !important;
    align-content:end; justify-items:center; gap:2px; margin:0; padding:0;
  }
  .grid.grid-img .card .pricebar{ flex-direction:column; gap:2px; }
  .grid.grid-img .card .pricebar .mx{ display:inline-block; width:var(--price-w-img); text-align:center; font-size:clamp(12px,3.8vw,15px); line-height:1; margin:0; }
  .grid.grid-img .card .btn.mini{ width:var(--btn-w-img); max-width:none; }

  /* 価格桁縮小（SP） */
  .mx{ text-align:center; width:100%; white-space:nowrap; }
  .mx.m7{ font-size:clamp(12px,3.4vw,14.5px) !important; letter-spacing:-0.02em; }
  .mx.shrink{ font-size:clamp(11px,3.2vw,13.5px) !important; letter-spacing:-0.03em; }
  .mx.ultra{ font-size:clamp(10px,2.9vw,12.5px) !important; letter-spacing:-0.04em; }
  .mx.tiny{  font-size:clamp(9px,2.8vw,11px) !important; }

  /* ▼ 画像OFF（list）：“鉄板”配置 ＝ 左下：価格／右横：ボタン（1行・底固定・横スワイプ） */
  .grid.grid-list{ grid-template-columns:repeat(2,minmax(0,1fr)) !important; column-gap:8px; row-gap:8px; }
  .grid.grid-list .card{ display:flex; flex-direction:column; min-width:0 !important; }
  .grid.grid-list .card .b{ position:static !important; height:auto !important; min-height:auto !important; padding:8px !important; flex:1 1 auto; }

  /* タイトル/メタは全文表示 */
  .grid.grid-list .card .n .ttl, .grid.grid-list .card .meta{
    display:block !important; -webkit-line-clamp:unset !important; -webkit-box-orient:unset !important;
    overflow:visible !important; white-space:normal !important;
  }

  /* === ここが要点 === */
  .grid.grid-list .card .foot{
    margin-top:auto !important;
    display:flex !important; align-items:center !important; gap:8px !important;
    white-space:nowrap !important; overflow-x:auto !important; -webkit-overflow-scrolling:touch;
    padding:0 !important; height:auto !important;
    /* 古いgrid指定を無効化 */
    grid-template-columns:unset !important; grid-auto-rows:unset !important;
  }
  .grid.grid-list .card .foot::-webkit-scrollbar{ display:none; }

  /* 左=価格（内容幅／省略しない／左寄せ） */
  .grid.grid-list .card .pricecell{ order:1; flex:0 0 auto !important; min-width:max-content !important; }
  .grid.grid-list .card .pricebar{
    width:auto !important; flex:0 0 auto !important; margin:0 !important;
    flex-direction:row !important; justify-content:flex-start !important; align-items:center !important; gap:6px !important;
  }
  .grid.grid-list .card .pricecell .mx{
    display:inline-block !important; width:auto !important; max-width:none !important;
    white-space:nowrap !important; overflow:visible !important; text-overflow:clip !important;
    text-align:left !important; margin:0 !important; line-height:1 !important;
    font-size:clamp(12px,3.6vw,15px) !important; /* 視認性確保 */
  }

  /* 右=ボタン（内容幅／横長すぎ防止） */
  .grid.grid-list .card .btncell{ order:2; flex:0 0 auto !important; min-width:0 !important; display:flex !important; align-items:center !important; }
  .grid.grid-list .card .btn.mini{
    width:auto !important; max-width:12ch !important; /* 伸びすぎ防止（必要なら調整） */
    height:var(--sp-btn-h) !important; line-height:var(--sp-btn-h) !important;
    padding:0 8px !important; border-radius:6px !important;
    white-space:nowrap !important; overflow:hidden !important; text-overflow:ellipsis !important; font-size:12px !important;
  }

  /* Cart 数量入力 */
  .cart-row .qty .qty-input{ width:48px; }
  /* Cart 行：2×2（右上=価格／右下=削除） */
  .cart-row{
    display:grid; grid-template-columns:1fr auto; grid-template-areas:"nm pr" "qty rm";
    gap:8px; align-items:start;
  }
  .cart-row .nm{ grid-area:nm !important; }
  .cart-row .pr{ grid-area:pr !important; justify-self:end; white-space:nowrap; }
  .cart-row .qty{ grid-area:qty !important; }
  .cart-row .btn.mini.rm{ grid-area:rm !important; justify-self:end; width:auto !important; min-width:0 !important; white-space:nowrap !important; padding:0 var(--cart-btn-pad-x) !important; }
}

/* --- Inputs (both) --- */
.cart-row .qty .qty-input{
  width:56px; text-align:center; padding:4px 6px; border:1px solid var(--border); border-radius:8px; font-size:14px; line-height:1.2;
}
.cart-row .qty .qty-input::-webkit-outer-spin-button,
.cart-row .qty .qty-input::-webkit-inner-spin-button{ -webkit-appearance:none; margin:0; }
.cart-row .qty .qty-input[type=number]{ -moz-appearance:textfield; }

/* === SP専用：画像ONの「買取カートへ」だけ小さく（修正版） === */
@media (max-width:700px){
  /* 画像ON用の専用高さ（必要なら 16〜20 に調整OK） */
  :root{ --sp-btn-h-img: 18px; }

  /* 価格上段 / ボタン下段 の2段レイアウトを固定 */
  .grid.grid-img .card .foot{
    grid-template-rows: min-content var(--sp-btn-h-img) !important;
  }

  /* SP × 画像ON では型番を隠す */
  .grid.grid-img .card .n .code{ display:none !important; }

  /* ボタン小型化＋はみ出し防止 */
  .grid.grid-img .card .btn.mini{
    height: var(--sp-btn-h-img) !important;
    line-height: var(--sp-btn-h-img) !important;
    font-size: 12px !important;
    padding: 0 6px !important;
    max-width: 100% !important;
    white-space: nowrap !important;
    overflow: hidden !important;
    text-overflow: ellipsis !important;
  }
}

/* === FINAL OVERRIDE: SP×画像OFF「左下=価格／右横=カート」強制固定 === */
@media (max-width:700px){
  /* カードは縦並び、本文→フッターが最下段に来る */
  #grid.grid-list .card{ display:flex !important; flex-direction:column !important; min-width:0 !important; }
  #grid.grid-list .card .b{ flex:1 1 auto !important; }

  /* フッターは左寄せ・1行・横スクロール可（古い grid 指定は無効化） */
  #grid.grid-list .card .foot{
    margin-top:auto !important;
    display:flex !important;
    align-items:center !important;
    justify-content:flex-start !important;
    gap:8px !important;
    white-space:nowrap !important;
    overflow-x:auto !important;
    -webkit-overflow-scrolling:touch;
    grid-template-columns:unset !important;
    grid-auto-rows:unset !important;
    padding:0 !important;
  }
  #grid.grid-list .card .foot::-webkit-scrollbar{ display:none; }

  /* 左＝価格：内容幅で左寄せ。7桁以上でも省略させない */
  #grid.grid-list .card .pricecell{
    order:1 !important;
    flex:0 0 auto !important;
    min-width:max-content !important;
  }
  #grid.grid-list .card .pricebar{
    width:auto !important;
    margin:0 !important;
    flex:0 0 auto !important;
    flex-direction:row !important;
    justify-content:flex-start !important; /* ← 中央寄せを打ち消す */
    align-items:center !important;
    gap:6px !important;
  }
  #grid.grid-list .card .pricecell .mx{
    display:inline-block !important;
    width:auto !important;
    max-width:none !important;
    text-align:left !important;
    white-space:nowrap !important;
    overflow:visible !important;
    text-overflow:clip !important;
    margin:0 !important;
    line-height:1 !important;
  }

  /* 右＝ボタン：内容幅。伸びすぎ防止だけ軽く */
  #grid.grid-list .card .btncell{
    order:2 !important;
    flex:0 0 auto !important;
    min-width:0 !important;
    display:flex !important; align-items:center !important;
  }
  #grid.grid-list .card .btn.mini{
    width:auto !important;           /* 100%禁止 */
    max-width:12ch !important;       /* 伸びすぎ防止（必要なら10–14chで調整） */
    height:var(--sp-btn-h) !important;
    line-height:var(--sp-btn-h) !important;
    padding:0 8px !important;
    border-radius:6px !important;
    white-space:nowrap !important;
    overflow:hidden !important;
    text-overflow:ellipsis !important;
    font-size:12px !important;
  }
}
/* ===== notice: PCだけ調整（幅900px以上） ===== */
@media (min-width:900px){
  /* ヘッダーは横並びOK。タイトルだけ絶対中央に重ねる */
  #noticeBox .notice-hd{
    position: relative !important;
    display: flex !important;
    align-items: center !important;
    gap: 8px;
  }
  /* タイトルを“ど真ん中”に固定（ボタン幅の影響を受けない） */
  #noticeBox .notice-title{
    position: absolute !important;
    left: 0; right: 0;
    text-align: center !important;
    font-size: 20px !important;
    line-height: 1.5;
    font-weight: 700;
    pointer-events: none;
  }
  /* ボタンは従来どおり右端に寄せる */
  #noticeBox .notice-toggle{
    margin-left: auto !important;
    position: relative;
    z-index: 1;
  }
  /* 本文（「買取は店舗と郵送で…」など）をPCだけ拡大 */
  #noticeBox .notice-body,
  #noticeBox .notice-body p,
  #noticeBox .notice-body li,
  #noticeBox .notice-body a{
    font-size: 17px !important; /* お好みで18pxまで */
    line-height: 1.8 !important;
  }
  #noticeBox .notice-body a{
    font-weight: 500;
    text-decoration: underline;
  }
}

/* ===== PCヘッダー調整（ロゴは動かさず、それ以外を左へ） ===== */
:root{
  --hdr-shift-pc: -13%;  /* ← 左へ寄せる量。-8%～-12%で微調整可 */
}

@media (min-width:1024px){
  /* 全体の間隔は少しだけ詰める */
  .header-wrap{ gap: 8px !important; }

  /* ロゴ（.brand-left）は触らない：配置そのまま */

  /* タイトルと右アクションだけ左へシフト */
  .center-ttl,
  .actions{
    transform: translateX(var(--hdr-shift-pc));
    will-change: transform;
  }

  /* タイトルは左寄せ＋ロゴとの距離を控えめに */
  .center-ttl{
    justify-self: start !important;
    text-align: left !important;
    margin-left: 8px !important; /* ロゴとの間隔。0～12pxで調整可 */
    margin-right: 0 !important;
  }

  /* ロゴ右にほんの少しだけ余白を残す（任意） */
  .brand-left img{
    margin-right: 4px !important;
  }
}
@media (min-width:1024px){
  /* 全体をまとめて左に 10% シフト */
  header .header-wrap{
    transform: translateX(-5%);
    will-change: transform;
  }

  /* デュエマ買取表と郵送買取アイコンの余白を増やす */
  .center-ttl{
    margin-right: 24px !important; /* ← デフォルトより広め。好みで 20～30px で調整可 */
  }
}
/* === FORCE: 画像OFF(list)はPCもSPも横並び（左：価格 / 右：ボタン） === */
.grid.grid-list .card .foot{
  display:flex !important; flex-direction:row !important;
  align-items:center !important; gap:8px !important;
  flex-wrap:nowrap !important; justify-content:flex-start !important;
  padding:0 !important;
}
.grid.grid-list .card .pricebar{ margin:0 !important; }
.grid.grid-list .card .btn.mini{ width:auto !important; max-width:12ch !important; }
/* === PC: 画像ON(img)は価格の右横にボタン === */
@media (min-width:1024px){
  .grid.grid-img .card .pricebar{
    display:flex !important; flex-direction:row !important;
    align-items:center !important; gap:12px !important;
    justify-content:center !important; /* 左寄せなら flex-start */
  }
}
/* === PATCH A: PCの「買取カートへ」がはみ出す対策 === */
@media (min-width:1024px){
  /* 画像ON/OFFどちらでも、価格の右横ボタンを“入るサイズ”にする */
  .grid .card .pricebar .btn.mini{
    display:inline-flex !important;
    align-items:center !important;
    font-size:14px !important;
    padding:6px 12px !important;
    height:auto !important;
    line-height:1.2 !important;
    max-width:none !important;      /* 枠切れ防止 */
    white-space:nowrap !important;
  }
  /* PCの画像OFFで幅を100%にしていた指定を打ち消し（伸び縮み可） */
  .grid.grid-list .card .btn.mini{
    width:auto !important;
  }
}

/* === PATCH B: SPで“7桁以上”でもはみ出さない（画像ON/画像OFF 両対応） === */
@media (max-width:700px){
  /* 画像ON：価格段が広くても下段ボタンが切れないようフッター高さ＆価格幅を少し増やす */
  :root{
    --price-w-img: 10.5ch;                       /* 7〜9桁想定で拡大 */
    --sp-btn-h-img: 20px;                         /* ボタン行に少し高さを */
    --sp-foot-h: calc(var(--sp-btn-h-img) + 24px);
  }
  .grid.grid-img .card .foot{
    grid-template-rows: min-content var(--sp-btn-h-img) !important;
  }

  /* 画像OFF：1行・横スクロールを強制、ボタンは縮めてもはみ出さない */
  #grid.grid-list .card .foot{
    display:flex !important;
    align-items:center !important;
    gap:8px !important;
    overflow-x:auto !important;
    -webkit-overflow-scrolling:touch;
    white-space:nowrap !important;
  }
  #grid.grid-list .card .pricebar{
    flex:0 0 auto !important;
    margin:0 !important;
  }
  #grid.grid-list .card .btn.mini{
    flex:0 0 auto !important;
    width:auto !important;          /* 100%禁止 */
    max-width:12ch !important;      /* 伸びすぎ防止 */
    font-size:12px !important;
    height:var(--sp-btn-h) !important;
    line-height:var(--sp-btn-h) !important;
    white-space:nowrap !important;
    overflow:hidden !important;
    text-overflow:ellipsis !important;
  }
}

/* === PATCH C: 念のため SP画像OFFは2列固定（競合打ち消し） === */
@media (max-width:700px){
  #grid.grid-list{ grid-template-columns:repeat(2,minmax(0,1fr)) !important; }
}
/* PC：画像OFF（grid-list）の「買取カートへ」を広げる */
@media (min-width:1024px){
  #grid.grid-list .card .btn.mini{
    width: auto !important;
    max-width: none !important;    /* ← ここで 12ch 制限を解除 */
    font-size: 14px !important;
    padding: 8px 14px !important;
    height: auto !important;
    line-height: 1.2 !important;
    white-space: nowrap !important;
  }
  #grid.grid-list .card .foot{ gap:12px !important; } /* 見た目の余白も少し広げる（任意） */
}


"""
# ===== JS =====
base_js = r"""
(function(){
  const LIFF_ID="__LIFF_ID__";
  const OA_ID="__OA_ID__";

  // ★ ユーティリティ
  const MQ_SP = window.matchMedia && window.matchMedia('(max-width:700px)');
  const isMobile = () => !!(MQ_SP && MQ_SP.matches);
  const slowNet = !!(navigator.connection && (navigator.connection.saveData ||
                    ['slow-2g','2g'].includes(navigator.connection.effectiveType)));
  const cores = navigator.hardwareConcurrency || 4;

  // 幅が 700px をまたいだら検索・描画をやり直し
  MQ_SP && MQ_SP.addEventListener('change', () => { apply(); });

  // __PER_PAGE__ を置換して受ける
  const PER_PAGE_ADJ = __PER_PAGE__;

  // Header height
  const header=document.querySelector('header');
  const setHeaderH=()=>{ const h=header?.getBoundingClientRect().height||144; document.documentElement.style.setProperty('--header-h', Math.ceil(h)+'px'); };
  setHeaderH(); addEventListener('resize', setHeaderH); addEventListener('load', setHeaderH);
  document.querySelectorAll('header img').forEach(img=>{ if(!img.complete) img.addEventListener('load', setHeaderH); });

  // Refs
  const nameQ=document.getElementById('nameQ'), codeQ=document.getElementById('codeQ'),
        packQ=document.getElementById('packQ'), rarityQ=document.getElementById('rarityQ'),
        grid=document.getElementById('grid'), navs=[...document.querySelectorAll('nav.simple')];

  // 既存ボタン
  const btnDesc=document.getElementById('btnPriceDesc'), btnAsc=document.getElementById('btnPriceAsc'),
        btnNone=document.getElementById('btnSortClear'), btnImg=document.getElementById('btnToggleImages');
  const btnPromo = document.getElementById('btnPromo');
  let promoOnly = false;

  // 追加：最新弾
  const btnLatest = document.getElementById('btnLatest');
  let latestOnly = false;

  function setLatestBtn(){
    if(!btnLatest) return;
    btnLatest.classList.toggle('active', latestOnly);
    btnLatest.setAttribute('aria-pressed', latestOnly ? 'true' : 'false');
  }
  function setPromoBtn(){
    if(!btnPromo) return;
    btnPromo.classList.toggle('active', promoOnly);
    btnPromo.setAttribute('aria-pressed', promoOnly ? 'true' : 'false');
  }

  const viewer=document.getElementById('viewer'), viewerImg=document.getElementById('viewerImg'), viewerClose=document.getElementById('viewerClose');

  // Notice toggle
  const noticeBox = document.getElementById('noticeBox');
  const noticeBody = document.getElementById('noticeBody');
  const noticeToggle = document.getElementById('noticeToggle');
  const NOTICE_KEY = 'climax_notice_open_v1';

  function setNotice(open){
    if(!noticeBox || !noticeBody || !noticeToggle) return;
    noticeBox.classList.toggle('open', open);
    noticeToggle.setAttribute('aria-expanded', open ? 'true' : 'false');
    noticeToggle.textContent = open ? '閉じる' : 'すべて見る';
    if(open){
      const h = noticeBody.scrollHeight;
      noticeBody.style.maxHeight = h + 'px';
    }else{
      noticeBody.style.maxHeight = '0px';
    }
    try{ localStorage.setItem(NOTICE_KEY, open ? '1':'0'); }catch(_){}
  }

  // 初期状態
  let openNotice = false;
  try{ openNotice = (localStorage.getItem(NOTICE_KEY) === '1'); }catch(_){}
  setNotice(openNotice);

  noticeToggle?.addEventListener('click', (e)=>{
    e.preventDefault();
    setNotice(!noticeBox.classList.contains('open'));
  });

  addEventListener('resize', ()=>{
    if(noticeBox?.classList.contains('open') && noticeBody){
      noticeBody.style.maxHeight = noticeBody.scrollHeight + 'px';
    }
  });

  // Cart
  const fabCart=document.getElementById('fabCart');
  const fabBadge=document.getElementById('fabBadge');
  const cartModal=document.getElementById('cartModal');
  const cartList=document.getElementById('cartList');
  const cartTotal=document.getElementById('cartTotal');
  const cartSend=document.getElementById('cartSend');
  const cartClose=document.getElementById('cartClose');
  const cartClear=document.getElementById('cartClear');

  // Scroll lock
  let _lockScrollY = 0, _locked = false, _prevScrollBehavior = '', _lastActive = null;
  function lockScroll(){
    if (_locked) return;
    _lockScrollY = window.pageYOffset || document.documentElement.scrollTop || 0;
    _prevScrollBehavior = getComputedStyle(document.documentElement).scrollBehavior || '';
    document.documentElement.style.scrollBehavior = 'auto';
    document.body.style.width = '100%';
    document.body.style.position = 'fixed';
    document.body.style.top = `-${_lockScrollY}px`;
    document.body.style.left = '0';
    document.body.style.right = '0';
    document.body.classList.add('modal-open');
    _lastActive = document.activeElement || null;
    _locked = true;
  }
  function unlockScroll(){
    if (!_locked) return;
    document.body.classList.remove('modal-open');
    document.body.style.position = '';
    const y = Math.max(0, -parseInt(document.body.style.top || '0', 10) || 0);
    document.body.style.top = '';
    document.body.style.left = '';
    document.body.style.right = '';
    document.body.style.width = '';
    document.documentElement.style.scrollBehavior = _prevScrollBehavior || '';
    window.scrollTo(0, y);
    if (_lastActive && typeof _lastActive.focus === 'function'){
      try{ _lastActive.focus({preventScroll:true}); }catch(_){}
    }
    _lastActive = null;
    _locked = false;
  }

  // Focus trap
  function trapFocus(container){
    function onKey(e){
      if(e.key !== 'Tab') return;
      const focusables = container.querySelectorAll('a[href],area[href],input,select,textarea,button,[tabindex]:not([tabindex="-1"])');
      const list = [...focusables].filter(el=>!el.hasAttribute('disabled') && el.tabIndex !== -1 && el.offsetParent!==null);
      if(!list.length) return;
      const first = list[0], last = list[list.length-1];
      if(e.shiftKey && document.activeElement === first){ e.preventDefault(); last.focus({preventScroll:true}); }
      else if(!e.shiftKey && document.activeElement === last){ e.preventDefault(); first.focus({preventScroll:true}); }
    }
    container._trapHandler = onKey;
    container.addEventListener('keydown', onKey);
  }
  function untrapFocus(container){
    if(container?._trapHandler){ container.removeEventListener('keydown', container._trapHandler); container._trapHandler = null; }
  }

  // Cart storage
  const CART_KEY="climax_cart_v1";
  let cart=[];
  let _skipQtyBlurOnce = false;
  function loadCart(){ try{ cart=JSON.parse(localStorage.getItem(CART_KEY)||"[]"); }catch{ cart=[]; } updateFab(); }
  function saveCart(){ localStorage.setItem(CART_KEY, JSON.stringify(cart)); updateFab(); }
  function updateFab(){ if(fabBadge) fabBadge.textContent=String(cart.reduce((a,b)=>a+(+b.qty||1),0)); }
  function delFromCart(i){ cart.splice(i,1); saveCart(); renderCartModal(); }

  const NAME_CAP = 10;
  function sameSkuTotal(name, model){
    return cart.reduce((s,it)=> s + ((it.name === name && it.model === model) ? (+it.qty||1) : 0), 0);
  }

  function incQty(i,d){
    const item = cart[i];
    const cur  = +item.qty || 1;

    if (d > 0) {
      const pairQty  = sameSkuTotal(item.name, item.model);
      const remaining = NAME_CAP - pairQty;
      if (remaining <= 0) {
        alert(`${item.name}［${item.model||''}］は合計${NAME_CAP}枚までです。`);
        return;
      }
      item.qty = cur + Math.min(d, remaining);
    } else {
      item.qty = Math.max(1, cur + d);
    }

    saveCart();
    renderCartModal();
  }

  function totalPrice(){ return cart.reduce((s,i)=> s+(i.price||0)*(i.qty||1), 0); }
  function escHtml(s){ return (s==null?'':String(s)).replace(/[&<>"']/g, m=>({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"}[m])); }

  function renderCartModal(){
    if(!cartList||!cartTotal) return;

    const keepY = cartList.scrollTop;

    cartList.innerHTML = cart.map((it,idx)=>`
      <div class="cart-row">
        <div class="nm">${escHtml(it.name)}${it.model?` <small>[${escHtml(it.model)}]</small>`:''}</div>
        <div class="pr">${it.price?`@${Number(it.price).toLocaleString()}円`:''}</div>
        <div class="qty">
          <button class="btn mini" data-a="dec" data-i="${idx}">－</button>
          <input type="number" class="qty-input" min="1" max="${NAME_CAP}" inputmode="numeric" pattern="[0-9]*" data-i="${idx}" value="${it.qty||1}" aria-label="数量">
          <button class="btn mini" data-a="inc" data-i="${idx}">＋</button>
        </div>
        <button class="btn mini rm" data-a="rm" data-i="${idx}">削除</button>
      </div>`).join('') || '<p style="color:#6b7280">カートは空です。</p>';

    cartList.scrollTop = keepY;
    cartTotal.textContent='合計：'+ totalPrice().toLocaleString() +'円';
  }

  // Modal open/close
  function openModal(modal){
    renderCartModal();
    modal?.classList.add('show');
    lockScroll();
    trapFocus(modal);
    modal?.setAttribute('aria-hidden','false');
    const firstBtn = modal.querySelector('button, [href], [tabindex]:not([tabindex="-1"])');
    (firstBtn || modal).focus?.({preventScroll:true});
  }
  function closeModal(modal){
    modal?.classList.remove('show');
    modal?.setAttribute('aria-hidden','true');
    untrapFocus(modal);
    unlockScroll();
  }
  if (fabCart) fabCart.addEventListener('click', (e)=>{ e.preventDefault(); document.activeElement?.blur?.(); openModal(cartModal); });
  cartClose?.addEventListener('click', ()=> closeModal(cartModal));
  cartModal?.addEventListener('click', e=>{ if(e.target===cartModal) closeModal(cartModal); });

  // カート操作
  function onCartAction(e){
  const btn = e.target.closest('button[data-a]');
  if(!btn) return;

  // ▼ 追加：次に発生する blur を1回だけ無視する
  _skipQtyBlurOnce = true;

  e.preventDefault();
  const a = btn.dataset.a, i = +btn.dataset.i;
  const keepY = cartList.scrollTop;

  if(a==='inc') incQty(i,+1);
  if(a==='dec') incQty(i,-1);
  if(a==='rm')  delFromCart(i);

  cartList.scrollTop = keepY;

  // ▼ 追加：イベントバブリングが終わったタイミングで解除
  setTimeout(()=>{ _skipQtyBlurOnce = false; }, 0);
}
  cartList?.addEventListener('click',       onCartAction);

  function onQtyFieldChange(e){
    const inp = e.target.closest('input.qty-input');
    if(!inp) return;

    const i = +inp.dataset.i;
    if (Number.isNaN(i) || !cart[i]) return;

    let v = Math.floor(parseInt(inp.value, 10));
    if (!Number.isFinite(v) || v < 1) v = 1;

    const CAP = (typeof NAME_CAP !== 'undefined') ? NAME_CAP : 10;
    if (v > CAP) v = CAP;

    cart[i].qty = v;
    saveCart();

    const keepY = cartList.scrollTop;
    renderCartModal();
    setTimeout(()=>{
      const el = cartList.querySelector(`input.qty-input[data-i="${i}"]`);
      el?.focus({preventScroll:true});
      el?.select();
      cartList.scrollTop = keepY;
    }, 0);
  }
  cartList?.addEventListener('change', onQtyFieldChange);
  cartList?.addEventListener('blur', (e)=>{
  // 数量欄のフォーカスが +/− ボタンに移る時の blur は処理しない
  if (_skipQtyBlurOnce) return;

  const rt = e.relatedTarget || document.activeElement; // 保険：relatedTarget 取れない環境あり
  if (rt && rt.closest && rt.closest('button[data-a]')) return;

  onQtyFieldChange(e);
}, true);
  cartList?.addEventListener('input', (e)=>{
    const inp = e.target.closest('input.qty-input');
    if(!inp) return;
    if (inp.value.length > 3) inp.value = inp.value.slice(0,3);
  });

  cartClear?.addEventListener('click', ()=>{ if(!cart.length){ alert('すでに空です'); return; } if(!confirm('カートを空にしますか？')) return; cart=[]; saveCart(); renderCartModal(); });

  // === 「LINEに送る」ボタン ===
  cartSend?.addEventListener('click', async () => {
    if (!cart.length) { alert('カートが空です。'); return; }
    const fullText = buildCartText();
    const sent = await sendCartViaLIFFInChunks(fullText);
    if (!sent) await openOAMessageWithText(fullText);
  });

  async function openOAMessageWithText(fullText){
    const RAW = OA_ID || "";
    const IDW = RAW.startsWith("@") ? RAW : "@"+RAW;
    const IDN = RAW.replace(/^@/, "");
    const enc = encodeURIComponent(fullText);
    try { await navigator.clipboard.writeText(fullText); } catch(_) {}
    const scheme = `line://oaMessage/${IDW}/?${enc}`;
    const intent = `intent://oaMessage/%40${encodeURIComponent(IDN)}/?${enc}` +
                   `#Intent;scheme=line;package=jp.naver.line.android;end`;
    let switched = false;
    const onVis = () => { if (document.visibilityState === 'hidden') switched = true; };
    document.addEventListener('visibilitychange', onVis, { once:true });
    location.href = scheme;
    await new Promise(r => setTimeout(r, 800));
    if (!switched && /Android/i.test(navigator.userAgent)) {
      location.href = intent;
      await new Promise(r => setTimeout(r, 600));
    }
    return switched;
  }

  // 画像ON/OFF
  let showImages=(localStorage.getItem('showImages')??'1')==='1';
  function setImgBtn(){ if(!btnImg) return; btnImg.textContent=showImages?'画像OFF':'画像ON'; btnImg.classList.toggle('active',showImages); btnImg.setAttribute('aria-pressed',showImages?'true':'false'); }
  function toggleImages(e){ if(e){e.preventDefault();e.stopPropagation();} showImages=!showImages; localStorage.setItem('showImages',showImages?'1':'0'); setImgBtn(); render(); }

  // 正規化
const SEP_RE = /[\s\u30FB\u00B7·/／\-_—–−]+/g;
function kataToHira(str){
  return (str || '').replace(/[\u30A1-\u30FA]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0x60));
}

// 読み置換マップ（漢字→かな）
const kanjiReadingMap = {
  "伝説": "でんせつ",
  "真気楼": "ぺてんしー",
  "墓碑": "えぴたふ",
  "西方": "びすまるく",
};

// ラテン文字の別名（英字→かな）
const latinAliasMap = {
  "complex": "こんぷれっくす",
  "c0br4":   "こぶら",
  // 任意で拡張
  // "epitaph":  "えぴたふ",
  // "bismarck": "びすまるく",
};

function normalizeForSearch(s){
  s = (s || '').normalize('NFKC').toLowerCase();
  for (const [k, v] of Object.entries(latinAliasMap))  s = s.split(k).join(v);
  for (const [k, v] of Object.entries(kanjiReadingMap)) s = s.split(k).join(v);
  return kataToHira(s).replace(SEP_RE, '');
}

function normalizeLatin(s){
  s = (s || '')
    .normalize('NFKC').toLowerCase()
    .replace(/0/g,'o').replace(/1/g,'l').replace(/3/g,'e').replace(/4/g,'a').replace(/5/g,'s').replace(/7/g,'t')
    .replace(SEP_RE,'')
    .replace(/[^a-z0-9]/g,'');
  return s;
}

  // データ
  function norm(it){ return { name:it.n??it.name??"", pack:it.p??it.pack??"", code:it.c??it.code??"", rarity:it.r??it.rarity??"", booster:it.b??it.booster??"", price:(it.pr??it.price??null), image:it.i??it.image??"", thumb:it.t??it.thumb??"", promo: (it.k===1 || it.k===true) ? 1 : 0, s:it.s??"" }; }
  let ALL=Array.isArray(window.__CARDS__)?window.__CARDS__.map(norm):[];
  if(!ALL.length){
    const hint=document.createElement('p');
    hint.style.cssText='color:#dc2626;padding:10px;margin:10px;border:1px dashed #fecaca;background:#fff5f5';
    hint.textContent='データが0件です。入力CSV/Excelのヘッダと列位置を確認してください。';
    document.querySelector('main')?.prepend(hint);
  }
  ALL=ALL.map(it=>({...it,
    _name:normalizeForSearch(it.name||""), _code:normalizeForSearch(it.code||""), _packbooster:normalizeForSearch([it.pack||"",it.booster||""].join(" ")), _rarity:normalizeForSearch(it.rarity||""),
    _name_lat:normalizeLatin(it.name||""), _code_lat:normalizeLatin(it.code||""), _packbooster_lat:normalizeLatin([it.pack||"",it.booster||""].join(" ")), _rarity_lat:normalizeLatin(it.rarity||"")
  }));

  // 検索 & ソート
  let VIEW=[], page=1, currentSort=__INITIAL_SORT__;
  function matchEither(kana,latin,qK,qL){ if(!qK && !qL) return true; let ok=false; if(qK && kana.includes(qK)) ok=true; if(qL && latin.includes(qL)) ok=true; return ok; }

  function apply(){
    // 正規化済みクエリ
    const qNameK=normalizeForSearch(nameQ.value||''),  qCodeK=normalizeForSearch(codeQ.value||''),  qPackK=normalizeForSearch(packQ.value||''),  qRarityK=normalizeForSearch(rarityQ.value||'');
    const qNameL=normalizeLatin(nameQ.value||''),      qCodeL=normalizeLatin(codeQ.value||''),      qPackL=normalizeLatin(packQ.value||''),      qRarityL=normalizeLatin(rarityQ.value||'');

    const rawName   = nameQ.value||'';
    const nameHasJP = /[\u3040-\u30ff\u3400-\u9fff]/.test(rawName);
    const nameHasEN = /[A-Za-z0-9]/.test(rawName);

    VIEW = ALL.filter(it => {
      let okName = true;
      if (rawName) {
        if (nameHasJP && nameHasEN) {
          okName = it._name.includes(qNameK);
        } else if (nameHasJP) {
          okName = it._name.includes(qNameK);
        } else {
          okName = it._name_lat.includes(qNameL);
        }
      }
      const okCode   = matchEither(it._code,        it._code_lat,        qCodeK,   qCodeL);
      const okPack   = matchEither(it._packbooster, it._packbooster_lat, qPackK,   qPackL);
      const okRarity = matchEither(it._rarity,      it._rarity_lat,      qRarityK, qRarityL);
      const okLatest = !latestOnly || ((it.pack||'').trim().toUpperCase() === 'DM25-RP3');

      return okName && okCode && okPack && okRarity
             && (!promoOnly || it.promo === 1)
             && okLatest;
    });

    if(currentSort==='desc') VIEW.sort((a,b)=>(b.price||0)-(a.price||0));
    else if(currentSort==='asc') VIEW.sort((a,b)=>(a.price||0)-(b.price||0));

    page=1; render();
  }

  // Pager
  function buildPageButtons(cur,total){
    const around=isMobile()?2:3, btns=[]; btns.push({type:'num',page:1});
    const start=Math.max(2,cur-around), end=Math.min(total-1,cur+around);
    if(start>2) btns.push({type:'ellipsis'});
    for(let p=start;p<=end;p++){ if(p>=2 && p<=total-1) btns.push({type:'num',page:p}); }
    if(end<total-1) btns.push({type:'ellipsis'});
    if(total>1) btns.push({type:'num',page:total});
    return btns;
  }
  function renderPager(cur,total){
     const nums=buildPageButtons(cur,total).map(item =>
       item.type==='ellipsis'
         ? `<span class="ellipsis">…</span>`
         : (item.page===cur
             ? `<button class="num" aria-current="page" disabled>${item.page}</button>`
             : `<a class="num" href="#" data-page="${item.page}">${item.page}</a>` )
     ).join(' ');
     return `<div class="pager">${nums}</div>`;
  }
  function scrollTopSmooth(){ window.scrollTo({top:0, behavior:'smooth'}); }

  // 価格フィット（保険）
  function fitPrices(){
    document.querySelectorAll('#grid .mx').forEach(el=>{
      const tooWide = el.scrollWidth > el.clientWidth && !el.classList.contains('tiny');
      el.classList.toggle('shrink', tooWide);
    });
  }

  // --- render ---
  function render(){
    const total = VIEW.length;
    const pages = Math.max(1, Math.ceil(total / PER_PAGE_ADJ));
    if (page > pages) page = pages;
    const start = (page - 1) * PER_PAGE_ADJ;
    const rows  = VIEW.slice(start, start + PER_PAGE_ADJ);

    if (showImages){
      grid.className = 'grid grid-img';

      grid.innerHTML = rows.map(it=>{
        const esc = s => (s==null?'':String(s)).replace(/[&<>\"']/g, m=>({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"}[m]));
        const nameEsc=esc(it.name||''), codeEsc=esc(it.code||'');
        const rawPrice = (it.price==null||it.price==='') ? '' : String(parseInt(it.price,10));
        const priceHtml = rawPrice ? '¥' + parseInt(rawPrice,10).toLocaleString() : '-';
        const digitCount = rawPrice.replace(/\D/g,'').length;

        let mxClass = 'mx';
        if (isMobile()) {
          if (digitCount >= 9)      mxClass += ' ultra';
          else if (digitCount >= 8) mxClass += ' shrink';
          else if (digitCount >= 7) mxClass += ' m7';
        } else {
          if (digitCount >= 9)      mxClass += ' ultra';
          else if (digitCount >= 8) mxClass += ' shrink';
        }

        let thumb=it.thumb||it.image||'';
        const hasHttp=/^https?:\/\//.test(thumb);
        if(thumb && !hasHttp && !thumb.startsWith('../')) thumb='../'+thumb;

        return `
  <article class="card">
    <div class="th" data-full="${it.image||''}">
      <img alt="${nameEsc}" loading="lazy" decoding="async" width="281" height="374" data-src="${thumb}" src=""
           onerror="this.onerror=null;var p=this.closest('.th');this.src=p?p.getAttribute('data-full'):this.src;">
    </div>
    <div class="b">
      <h3 class="n">
        <span class="ttl">${nameEsc}</span>
        ${codeEsc ? `<span class="code">${codeEsc}</span>` : ``}
      </h3>
      <div class="foot">
        ${
          isMobile()
          ? `
            <div class="pricebar"><span class="${mxClass}">${priceHtml||'-'}</span></div>
            <div class="pricebar"><button class="btn mini btn-add" data-name="${nameEsc}" data-model="${codeEsc}" data-price="${it.price||0}">買取カートへ</button></div>
          `
          : `
            <div class="pricebar">
              <span class="${mxClass}">${priceHtml||'-'}</span>
              <button class="btn mini btn-add" data-name="${nameEsc}" data-model="${codeEsc}" data-price="${it.price||0}">買取カートへ</button>
            </div>
          `
        }
      </div>
    </div>
  </article>`;
      }).join('');
    } else {
      // 画像OFF：必ず grid-list（PC/SPとも同一DOM：.pricecell と .btncell を常に出力）
      grid.className = 'grid grid-list';

      grid.innerHTML = rows.map(it=>{
        const esc = s => (s==null?'':String(s)).replace(/[&<>\"']/g, m=>({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"}[m]));
        const nameEsc=esc(it.name||''), codeEsc=esc(it.code||'');
        const rawPrice = (it.price==null||it.price==='') ? '' : String(parseInt(it.price,10));
        const priceHtml = rawPrice ? '¥' + parseInt(rawPrice,10).toLocaleString() : '-';
        const digitCount = rawPrice.replace(/\D/g,'').length;

        let mxClass = 'mx';
        if (isMobile()) {
          if (digitCount >= 9)      mxClass += ' ultra';
          else if (digitCount >= 8) mxClass += ' shrink';
          else if (digitCount >= 7) mxClass += ' m7';
        } else {
          if (digitCount >= 9)      mxClass += ' ultra';
          else if (digitCount >= 8) mxClass += ' shrink';
        }

        const meta = [it.pack||'', it.booster||''].filter(Boolean).join(' / ');

        return `
    <article class="card">
      <div class="b">
        <h3 class="n">
        <span class="ttl">${nameEsc}</span>
        ${codeEsc ? `<span class="code">${codeEsc}</span>` : ``}
      </h3>
        <div class="meta">${esc(meta)}</div>
        <div class="foot">
          <div class="pricebar pricecell"><span class="${mxClass}">${priceHtml||'-'}</span></div>
          <div class="btncell">
            <button class="btn mini btn-add" data-name="${nameEsc}" data-model="${codeEsc}" data-price="${it.price||0}">買取カートへ</button>
          </div>
        </div>
      </div>
    </article>`;
      }).join('');
    }

    // 画像クリック → ビューア
    if(showImages){
      grid.querySelectorAll('.th').forEach(th=>{
        th.addEventListener('click', (ev)=>{
          ev.preventDefault(); ev.stopPropagation();
          document.activeElement?.blur?.();
          const src = th.getAttribute('data-full') || th.querySelector('img')?.getAttribute('data-src') || th.querySelector('img')?.src || '';
          if(!src) return;
          viewerImg.src = src;
          viewer.classList.add('show');
          viewer.setAttribute('role','dialog'); viewer.setAttribute('aria-modal','true'); viewer.setAttribute('aria-hidden','false');
          lockScroll();
          trapFocus(viewer);
          (viewerClose || viewer).focus?.({preventScroll:true});
        });
      });
    }

    // カート追加
    grid.querySelectorAll('.btn-add').forEach(btn=>{
      let busy = false;
      btn.addEventListener('click', (ev)=>{
        if (busy) return;
        busy = true;
        ev.preventDefault();

        const price = Number(btn.dataset.price)||0;
        const name  = btn.dataset.name||'';
        const code  = btn.dataset.model||'';

        const CAP = (typeof NAME_CAP !== 'undefined') ? NAME_CAP : 10;
        const pairQty = sameSkuTotal(name, code);

        if (pairQty >= CAP) {
          alert(`${name}［${code}］は合計${CAP}枚までです。`);
          busy = false;
          return;
        }

        const f = cart.find(x => (x.name||'')===name && (x.model||'')===code);
        if (f) { f.qty = (+f.qty||1) + 1; }
        else   { cart.push({name, model:code, price, qty:1}); }

        saveCart();
        setTimeout(()=> busy = false, 120);
      }, {passive:false});
    });

    // ページャ
    const pagerHtml=renderPager(page, pages);
    navs.forEach(n=>{
      n.innerHTML=pagerHtml;
      n.onclick=(e)=>{
        const a=e.target.closest('a,button'); if(!a || a.matches('.disabled,[disabled]')) return; e.preventDefault();
        const p=parseInt(a.dataset.page||'0',10); if(p && p!==page){ page=p; render(); scrollTopSmooth(); }
      };
    });

    // lazy
    if(showImages){
      const io=new IntersectionObserver((entries)=>{
        entries.forEach(e=>{
          if(e.isIntersecting){
            const img=e.target, ds=img.getAttribute('data-src');
            if(ds && !img.src){ img.src=ds; img.removeAttribute('data-src'); }
            io.unobserve(img);
          }
        });
      }, { rootMargin:(isMobile()||slowNet)?"300px 0px":"600px 0px", threshold:0.01 });
      document.querySelectorAll('#grid img[data-src]').forEach(img=>io.observe(img));
    }

    fitPrices();
  }
  // --- render end ---

  function setActiveSort(){
    btnDesc?.setAttribute('aria-pressed', currentSort==='desc'?'true':'false');
    btnAsc ?.setAttribute('aria-pressed', currentSort==='asc' ?'true':'false');
    btnNone?.setAttribute('aria-pressed', currentSort===null   ?'true':'false');
    btnDesc?.classList.toggle('active', currentSort==='desc');
    btnAsc ?.classList.toggle('active', currentSort==='asc');
    btnNone?.classList.toggle('active', currentSort===null);
  }

  // ビューア閉じる
  function closeViewer(){
    viewer.classList.remove('show');
    viewer.setAttribute('aria-hidden','true');
    untrapFocus(viewer);
    viewerImg.src='';
    unlockScroll();
  }
  viewerClose?.setAttribute('aria-label','閉じる');
  viewerClose?.addEventListener('click', closeViewer);
  viewer?.addEventListener('click', e=>{ if(e.target===viewer) closeViewer(); });
  addEventListener('keydown', e=>{ if(e.key==='Escape' && viewer.classList.contains('show')) closeViewer(); });

  // LIFF SDK（外部）
  (function addLiff(){ if(document.querySelector('script[data-liff]')) return;
    const s=document.createElement('script'); s.src='https://static.line-scdn.net/liff/edge/2/sdk.js'; s.async=true; s.setAttribute('data-liff','1'); document.head.appendChild(s);
  })();

  async function trySendViaLIFF(text){
    if(!window.liff) return false;
    try{
      if(!liff._climaxInited){ await liff.init({liffId:LIFF_ID}); liff._climaxInited=true; }
      if(!liff.isLoggedIn()){ liff.login(); return false; }
      if(!liff.isInClient()) return false;
      await liff.sendMessages([{type:'text', text}]);
      alert('LINEに仮査定リストを送信しました。'); liff.closeWindow?.(); return true;
    }catch(e){ return false; }
  }

  // テキスト生成 & 分割送信
  function buildCartText(){
    const lines=[]; let total=0;
    for(const it of cart){
      const q=+it.qty||1, p=+it.price||0; total+=q*p;
      lines.push(`・${it.name}${it.model?`［${it.model}］`:''} ×${q}${p?` @${p.toLocaleString()}円`:''}`);
    }
    let body=`【仮査定依頼】\n`+lines.join("\n");
    body+=`\n――――\n合計：${total.toLocaleString()}円\n※仮査定です。現物確認後に正式査定となります。`;
    return body;
  }
  function chunkTextByChars(text, maxChars = 5000){
    const chunks=[];
    let rest = text;
    while(rest.length > maxChars){
      let cut = rest.lastIndexOf("\n", maxChars);
      if(cut < Math.floor(maxChars*0.6)) cut = maxChars;
      chunks.push(rest.slice(0, cut));
      rest = rest.slice(cut).replace(/^\n/, "");
    }
    if(rest) chunks.push(rest);
    return chunks;
  }
  document.getElementById("cartSavePNG")?.addEventListener("click", ()=>{
    if(!cart.length){ alert("カートが空です。"); return; }
    const canvas = buildCartReceiptCanvas();
    const ts = new Date();
    const y = ts.getFullYear(), m = String(ts.getMonth()+1).padStart(2,"0"), d = String(ts.getDate()).padStart(2,"0");
    const hh = String(ts.getHours()).padStart(2,"0"), mm = String(ts.getMinutes()).padStart(2,"0");
    const fname = `kaitori_${y}${m}${d}_${hh}${mm}.png`;
    saveCanvasAsPng(canvas, fname);
    alert("レシート画像を保存しました。LINEで画像を添付して送れます。");
  });
  async function sendCartViaLIFFInChunks(fullText){
    if(!window.liff) return false;
    try{
      if(!liff._climaxInited){ await liff.init({liffId:LIFF_ID}); liff._climaxInited = true; }
      if (!liff.isInClient()) return false;
      if (!liff.isLoggedIn()) return false;
      const parts = chunkTextByChars(fullText, 5000);
      const msgs  = parts.map((t,i)=>({ type:'text', text: (parts.length>1 ? `(${i+1}/${parts.length})\n` : '') + t }));
      for (let i=0; i<msgs.length; i+=5) {
        await liff.sendMessages(msgs.slice(i, i+5));
      }
      alert('LINEに仮査定リストを送信しました。');
      liff.closeWindow?.();
      return true;
    }catch(_){ return false; }
  }

  // 入力遅延
  const DEBOUNCE=(isMobile()||slowNet||cores<=4)?240:120;
  function onInputDebounced(el){ el.addEventListener('input', ()=>{ clearTimeout(el._t); el._t=setTimeout(apply,DEBOUNCE); }); }
  [nameQ,codeQ,packQ,rarityQ].forEach(onInputDebounced);

  function initButtons(){
    btnLatest?.addEventListener('click', ()=>{ latestOnly = !latestOnly; setLatestBtn(); apply(); });
    setLatestBtn();

    btnPromo?.addEventListener('click', ()=>{ promoOnly = !promoOnly; setPromoBtn(); apply(); }); setPromoBtn();
    btnDesc?.addEventListener('click', ()=>{ currentSort=(currentSort==='desc')?null:'desc'; setActiveSort(); apply(); });
    btnAsc ?.addEventListener('click', ()=>{ currentSort=(currentSort==='asc' )?null:'asc' ; setActiveSort(); apply(); });
    btnNone?.addEventListener('click', ()=>{ currentSort=null; setActiveSort(); apply(); });
    btnImg ?.addEventListener('click', toggleImages); setImgBtn();
  }

  setActiveSort(); initButtons(); loadCart(); apply();
})();

// ダブルタップでのズームを全体で抑止（ボタン類のみ）
document.addEventListener('dblclick', (e)=>{
  if (e.target.closest('.btn') || e.target.closest('.iconbtn')) {
    e.preventDefault();
  }
}, {passive:false});
"""

# ===== HTML =====
def html_page(title: str, js_source: str, logo_uri: str, cards_json: str, updated_text: str = "") -> str:
    shop_svg   = "<svg viewBox='0 0 24 24' aria-hidden='true' fill='currentColor'><path d='M3 9.5V8l2.2-3.6c.3-.5.6-.7 1-.7h11.6c.4 0 .7.2 .9 .6L21 8v1.5c0 1-.8 1.8-1.8 1.8-.9 0-1.6-.6-1.8-1.4-.2 .8-.9 1.4-1.8 1.4s-1.6-.6-1.8-1.4c-.2 .8-.9 1.4-1.8 1.4C3.8 11.3 3 10.5 3 9.5zM5 12.5h14V20c0 .6-.4 1-1 1H6c-.6 0-1-.4-1-1v-7.5zm4 1.5v5h6v-5H9zM6.3 5.2 5 7.5h14l-1.3-2.3H6.3z'/></svg>"
    login_svg  = "<svg viewBox='0 0 24 24' aria-hidden='true' fill='currentColor'><path d='M12 12a5 5 0 1 0-5-5 5 5 0 0 0 5 5zm0 2c-4.418 0-8 2.239-8 5v2h16v-2c0-2.761-3.582-5-8-5z'/></svg>"
    takuhai_svg= "<svg viewBox='0 0 24 24' aria-hidden='true' fill='currentColor'><path d='M3 6h11a2 2 0 0 1 2 2v1h3l2 3v5a2 2 0 0 1-2 2h-1a2 2 0 0 1-2-2H8a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2Zm0 2v9h2a2 2 0 0 1 2-2h8V8H3Zm16 3h-2v4h4v-3.2L19 11Z'/></svg>"
    parts=[]
    parts.append("<!doctype html><html lang='ja'><head><meta charset='utf-8'>")
    parts.append("<meta name='viewport' content='width=device-width,initial-scale=1'>")
    parts.append(f"<meta name='cards-ver' content='{CARDS_VER}'>")
    parts.append("<style>"); parts.append(base_css); parts.append("</style>")
    parts.append("<script src='https://static.line-scdn.net/liff/edge/2/sdk.js' data-liff></script>")
    parts.append("</head><body>")

    # header
    parts.append("<header><div class='header-wrap'>")
    parts.append("<div class='brand-left'>")
    if logo_uri:
        parts.append(f"<img src='{logo_uri}' alt='Shop Logo'>")
    else:
        parts.append("<div class='brand-fallback' aria-label='Shop Logo' style='font-weight:900;font-size:18px;'>YOUR SHOP</div>")
    parts.append("</div>")
    parts.append("<div class='center-ttl'>")
    parts.append(f"  <div class='ttl'>{html_mod.escape(title)}</div>")
    if updated_text:
        parts.append(f"  <div class='sub'>{html_mod.escape(updated_text)}</div>")
    parts.append("</div>")
    parts.append("<div class='actions'>")

    # 郵送買取ボタン（残す）
    parts.append(f"<a class='iconbtn' href='{html_mod.escape(TAKUHAI_URL)}' target='_blank' rel='noopener'>{takuhai_svg}<span>郵送買取</span></a>")

    # Shopボタン（残す）
    parts.append(f"<a class='iconbtn' href='https://www.climax-card.jp/' target='_blank' rel='noopener'>{shop_svg}<span>Shop</span></a>")

    # Loginボタン（削除）

    # Xアイコン
    if X_ICON_URI:
        parts.append(f"<a class='iconimg iconimg--x' href='https://x.com/climaxcard' target='_blank' rel='noopener'><img src='{X_ICON_URI}' alt='X'></a>")

    # LINEアイコン
    if LINE_ICON_URI:
        parts.append(f"<a class='iconimg iconimg--line' href='https://line.me/R/ti/p/{OA_ID}' target='_blank' rel='noopener'><img src='{LINE_ICON_URI}' alt='LINE'></a>")

    # Instagramアイコン
    if INSTAGRAM_ICON_URI:
        parts.append(
            "<a class='iconimg iconimg--ig' "
            "href='https://www.instagram.com/cardshopclimax?igsh=d3VybGZraHlhZXUy&utm_source=qr' "
            "target='_blank' rel='noopener'>"
            f"<img src='{INSTAGRAM_ICON_URI}' alt='Instagram'></a>"
        )

    # ★ TikTokアイコン追加
    if TIKTOK_ICON_URI:
        parts.append(
            "<a class='iconimg iconimg--tiktok' "
            "href='https://www.tiktok.com/@climax932' "
            "target='_blank' rel='noopener'>"
            f"<img src='{TIKTOK_ICON_URI}' alt='TikTok'></a>"
        )

    # カートボタン
    parts.append("<button id='fabCart' class='iconbtn iconbtn--cart' type='button' aria-haspopup='dialog' aria-controls='cartModal'>🛒 <span id='fabBadge' class='badge'>0</span></button>")
    parts.append("</div></div></header>")



    # ✅ notice（検索欄の“直前”に配置）
    parts.append(
      "<section class='notice' id='noticeBox'>"
        "<div class='notice-hd'>"
          "<div class='notice-title'>買取注意事項 【初めての方はご一読ください】</div>"
          "<button id='noticeToggle' class='notice-toggle' type='button' aria-expanded='false' aria-controls='noticeBody'>すべて見る</button>"
        "</div>"
        "<div id='noticeBody' class='notice-body' role='region' aria-label='買取注意事項の本文'>"
          "<ul>"
            "<li>買取は店舗と郵送で実施しております。（<a href='https://climaxcard.github.io/climax2/' target='_blank' rel='noopener'>店舗情報はこちら</a>）</li>"
            "<li>店頭買取の場合必ずスリーブを外してからお持ち込みお願いします。</li>"
            "<li>郵送買取の場合はスリーブ等に入れて傷がつかないように発送してください。</li>"
            "<li>必ず同名カードごとに整理をお願いします。</li>"
            "<li>買取価格は在庫状況・相場・買取枚数などにより、予告なく変動する場合がございます。</li>"
            "<li>お品物の状態によっては査定額が上下することがあります。</li>"
            "<li>高額カードやコレクション向けのカードは、わずかなキズでも大きく減額となる場合がございます。</li>"
            "<li>最新弾のカードは相場の変動が激しいため、カード到着時地点の買取金額になります。</li>"
            "<li>リストに記載のないカードも査定可能ですが、通常よりお時間をいただくことがございます。</li>"
            "<li>店頭買取受付時間は 【月～土 11:00～19:00 / 日祝 11:00～18:00】 です。<br>※受付時間は予告なく変更となる場合がございます。予めご了承ください。</li>"
            "<li>郵送買取は24時間受付可能です。</li>"
            "<li>VR以上のカードは一律10円です。枚数制限ありません。</li>"
            "<li>大量にお持ち込みいただく際は、お時間に余裕をもってご来店ください。</li>"
            "<li>未成年の方からの買取はお受けできません。</li>"
          "</ul>"
        "</div>"
      "</section>"
    )

    # main
    parts.append("<main class='wrap'>")
    parts.append("<div class='controls'>")
    parts.append("  <input id='nameQ'   class='search' placeholder='名前：部分一致（ひらがな可）' autocomplete='off'>")
    parts.append("  <input id='codeQ'   class='search' placeholder='型番：例 SP4/SP5, 19/100 等' autocomplete='off'>")
    parts.append("  <input id='packQ'   class='search' placeholder='弾：例 DM25RP1, 邪神VS邪神 等' autocomplete='off'>")
    parts.append("  <input id='rarityQ' class='search' placeholder='レアリティ：SR/VR 等' autocomplete='off'>")
    parts.append("  <div class='btns'>")
    parts.append("    <button id='btnLatest'    class='btn' type='button' aria-pressed='false'>最新弾</button>")
    parts.append("    <button id='btnPromo'     class='btn' type='button' aria-pressed='false'>強化買取中!!</button>")
    parts.append("    <button id='btnPriceDesc' class='btn' type='button' aria-pressed='false'>価格高い順</button>")
    parts.append("    <button id='btnPriceAsc'  class='btn' type='button' aria-pressed='false'>価格低い順</button>")
    parts.append("    <button id='btnToggleImages' class='btn' type='button'>画像ON</button>")
    parts.append("  </div></div>")
    parts.append("  <nav class='simple'></nav><div id='grid' class='grid grid-img'></div><nav class='simple'></nav>")
    parts.append("  <small class='note'>画像クリックで拡大。🛒カートから「LINEに送る」でCLIMAX公式LINEへ送信できます（アプリ外でもOK）。</small>")
    parts.append("</main>")

    # viewer
    parts.append("<div id='viewer' class='viewer' tabindex='-1' aria-hidden='true'><div class='vc'><img id='viewerImg' alt='' role='img'><button id='viewerClose' class='close' aria-label='閉じる'>×</button></div></div>")

    # data
    parts.append("<script>window.__CARDS__=" + cards_json + ";</script>")
    # cart modal  ← 関数内にインデント
    parts.append(
    "<div id='cartModal' class='cart-modal' tabindex='-1' role='dialog' aria-modal='true' aria-hidden='true' aria-label='仮査定カート'>"
    "<div class='cart-sheet'>"
      "<div class='cart-head'>"
        "<strong>仮査定カート</strong>"
        "<button id='cartClose' class='btn'>閉じる</button>"
      "</div>"
      "<div id='cartList' class='cart-list'></div>"
      "<div class='cart-foot'>"
        "<div id='cartTotal'>合計：0円</div>"
        "<div class='cart-act'>"
          "<button id='cartClear' class='btn' style='border-color:#fecaca;color:#b91c1c;'>カートを空にする</button>"
          "<button id='cartSend' class='btn' style='background:#06C755;color:#fff;border-color:#06C755;'>LINEに送る</button>"
        "</div>"
      "</div>"
    "</div>"
    "</div>"
)

    # scripts
    parts.append("<script>")
    parts.append(js_source)
    parts.append("</script>")
    parts.append("</body></html>")
    return "".join(parts)


# ===== 出力 =====
OUT_DIR.mkdir(parents=True, exist_ok=True)

def write_mode(dir_name: str, initial_sort_js_literal: str, title_text: str):
    sub = OUT_DIR / dir_name
    sub.mkdir(parents=True, exist_ok=True)
    js = (base_js
          .replace("__PER_PAGE__", str(PER_PAGE))
          .replace("__INITIAL_SORT__", initial_sort_js_literal)
          .replace("__LIFF_ID__", LIFF_ID)
          .replace("__OA_ID__", OA_ID))
    html = html_page(title_text, js, LOGO_URI, CARDS_JSON, UPDATED_LABEL)
    (sub/"index.html").write_text(html, encoding="utf-8")

write_mode("default", "'desc'", "デュエマ買取表")
write_mode("price_desc", "'desc'", "デュエマ買取表（price_desc）")
write_mode("price_asc",  "'asc'",  "デュエマ買取表（price_asc）")

(OUT_DIR/"index.html").write_text("<meta http-equiv='refresh' content='0; url=default/'>", encoding="utf-8")

print(f"[*] Excel/CSV: {EXCEL_PATH!r}")
print(f"[*] PER_PAGE={PER_PAGE}  BUILD_THUMBS={'1' if BUILD_THUMBS else '0'}")
print(f"[LOGO] {'embedded' if LOGO_URI else 'not found (fallback text used)'}")
print(f"[X ICON] {'embedded' if X_ICON_URI else 'not found'}")
print(f"[LINE ICON] {'embedded' if LINE_ICON_URI else 'not found'}")
print(f"[OK] 生成完了 → {OUT_DIR.resolve()} / 総件数{len(df)}")
