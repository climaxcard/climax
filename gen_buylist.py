# -*- coding: utf-8 -*-
r"""
デュエマ買取表 静的ページ生成（完成版：仮査定カート + LINE送信・LIFF/ディープリンク両対応）
- CSV/Excel 自動対応。二重ヘッダ(日本語/英語キー)も自動正規化
- 列は「ヘッダ名優先 → 位置フォールバック(C/E/F/G/H/O/Q)」
- 画像URLは Q列などから自動抽出 (=IMAGE() もOK)
- PC/SP最適化＆高速ページャ／画像lazy
- 🛒 カート機能（LocalStorage）・数量変更・削除・合計
- 「LINEに送る」：
    * LIFF内（LINEアプリ内）：liff.sendMessages() 一発送信
    * 普通のブラウザ：公式OA宛ディープリンクでLINE起動＆下書き付き遷移
"""

from pathlib import Path
from urllib.parse import urlparse, parse_qs
import pandas as pd
import html as html_mod
import unicodedata as ud
import base64, mimetypes, os, sys, hashlib, io, json, re, glob

# ==== 依存（BUILD_THUMBS=1 の場合のみ使う）====
try:
    import requests
    from PIL import Image
except Exception:
    requests = None
    Image = None

# ========= 基本設定 =========
DEFAULT_EXCEL = "buylist.xlsx"
ALT_EXCEL     = "data/buylist.xlsx"
FALLBACK_WINDOWS = r"C:\Users\user\Desktop\デュエマ買取表\buylist.xlsx"

EXCEL_PATH = os.getenv("EXCEL_PATH", DEFAULT_EXCEL)
SHEET_NAME = os.getenv("SHEET_NAME", "シート1")

OUT_DIR    = Path(os.getenv("OUT_DIR", "docs"))
PER_PAGE   = int(os.getenv("PER_PAGE", "80"))
BUILD_THUMBS = os.getenv("BUILD_THUMBS", "0") == "1"

# ====== LIFF / 送信先（必要に応じて変更）======
LIFF_ID = os.getenv("LIFF_ID", "2007983559-MKopknXd")  # あなたのLIFF ID
OA_ID   = os.getenv("OA_ID",   "@512nwjvn")            # 公式LINEのID（@含む）

LOGO_FILE_ENV = os.getenv("LOGO_FILE", "").strip()
X_ICON_FILE_ENV    = os.getenv("X_ICON_FILE", "").strip()
LINE_ICON_FILE_ENV = os.getenv("LINE_ICON_FILE", "").strip()
FIXED_ICON_DIR = Path(r"C:\Users\user\OneDrive\Desktop\デュエマ買取表")

if len(sys.argv) > 1 and sys.argv[1]:
    EXCEL_PATH = sys.argv[1]

# ===== 位置フォールバック用（0始まり） =====
# 指定：C(2), E(4), F(5), G(6), H(7), O(14=買取価格), Q(16=画像URL)
IDX_NAME   = 2
IDX_PACK   = 4
IDX_CODE   = 5
IDX_RARITY = 6
IDX_BOOST  = 7
IDX_PRICE  = 14
IDX_IMGURL = 16

# サムネ設定
THUMB_DIR = OUT_DIR / "assets" / "thumbs"
THUMB_W = 600

# ---- ロゴ探索 ----
def find_logo_path():
    cands = []
    if LOGO_FILE_ENV:
        cands.append(Path(LOGO_FILE_ENV))
    cands += [Path("assets") / "logo.png"]
    cands += [Path(os.getcwd()) / "logo.png", Path(os.getcwd()) / "logo.png.png"]
    try:
        here = Path(__file__).parent
        cands += [here / "assets" / "logo.png", here / "logo.png", here / "logo.png.png"]
    except NameError:
        pass
    for p in cands:
        try:
            if p and p.exists() and p.is_file():
                return p
        except Exception:
            pass
    return None

def file_to_data_uri(p: Path|None) -> str:
    if not p: return ""
    mime = mimetypes.guess_type(str(p))[0] or "image/png"
    try:
        b64  = base64.b64encode(p.read_bytes()).decode("ascii")
        return f"data:{mime};base64,{b64}"
    except Exception:
        return ""

LOGO_URI = file_to_data_uri(find_logo_path())

# ---- X/LINE 画像探索 ----
def find_icon_path(env_path: str, default_names: list[str]):
    cands = []
    if env_path:
        cands.append(Path(env_path))
    cands += [Path("assets") / n for n in default_names]
    cands += [Path(os.getcwd()) / n for n in default_names]
    try:
        here = Path(__file__).parent
        cands += [here / "assets" / n for n in default_names]
        cands += [here / n for n in default_names]
    except NameError:
        pass
    cands += [FIXED_ICON_DIR / n for n in default_names]
    for p in cands:
        try:
            if p.exists() and p.is_file():
                return p
        except Exception:
            pass
    return None

X_ICON_URI    = file_to_data_uri(find_icon_path(X_ICON_FILE_ENV,    ["X.png", "x.png", "x-logo.png"]))
LINE_ICON_URI = file_to_data_uri(find_icon_path(LINE_ICON_FILE_ENV, ["LINE.png", "line.png", "line-icon.png"]))

# ========= 入力ファイル 読み込み・正規化 =========
def _read_csv_auto(path: Path) -> pd.DataFrame:
    for enc in ("utf-8-sig", "cp932", "utf-8"):
        try:
            return pd.read_csv(path, encoding=enc)
        except Exception:
            pass
    return pd.read_csv(path)

def _normalize_two_header_layout(df: pd.DataFrame) -> pd.DataFrame:
    try:
        cand = []
        m = min(12, len(df))
        for i in range(m):
            row = df.iloc[i].astype(str).tolist()
            if "display_name" in row and "cardnumber" in row:
                cand.append(i)
        if not cand:
            return df
        hdr = cand[0]
        start = hdr + 2
        df2 = df.iloc[start:].copy()
        df2.columns = df.iloc[hdr].tolist()
        return df2.reset_index(drop=True)
    except Exception:
        return df

def _resolve_input(pref: str|None) -> Path:
    cands = []
    if pref: cands.append(Path(pref))
    cands += [Path("buylist.csv"), Path("buylist.xlsx"), Path(ALT_EXCEL), Path(FALLBACK_WINDOWS)]
    for p in cands:
        try:
            if p.exists() and p.is_file():
                return p
        except Exception:
            pass
    files = sorted(
        [Path(p) for p in glob.glob("*.csv")] + [Path(p) for p in glob.glob("*.xlsx")],
        key=lambda x: x.stat().st_mtime, reverse=True
    )
    if files: return files[0]
    raise FileNotFoundError("CSV/Excel が見つかりません。")

def load_buylist_any(path_hint: str, sheet_name: str|None) -> pd.DataFrame:
    p = _resolve_input(path_hint)
    if p.suffix.lower()==".csv":
        df0 = _read_csv_auto(p)
        return _normalize_two_header_layout(df0)
    try:
        if sheet_name:
            df0 = pd.read_excel(p, sheet_name=sheet_name, header=None, engine="openpyxl")
        else:
            xls = pd.ExcelFile(p, engine="openpyxl")
            df0 = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=None, engine="openpyxl")
    except Exception:
        xls = pd.ExcelFile(p, engine="openpyxl")
        df0 = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=None, engine="openpyxl")
    return _normalize_two_header_layout(df0)

df_raw = load_buylist_any(EXCEL_PATH, SHEET_NAME)

# ========= ユーティリティ =========
SEP_RE = re.compile(r"[\s\u30FB\u00B7·/／\-_—–−]+")

def clean_text(s: pd.Series) -> pd.Series:
    s = s.astype(str)
    s = s.str.replace(r'(?i)^\s*nan\s*$', '', regex=True)
    s = s.replace({"nan":"","NaN":"","None":"","NONE":"","null":"","NULL":"","nil":"","NIL":""})
    return s.fillna("").str.strip()

def to_int_series(s: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce").round().astype("Int64")
    s = s.astype(str).str.replace(r"[^\d\.\-,]", "", regex=True).str.replace(",", "", regex=False)
    return pd.to_numeric(s, errors="coerce").round().astype("Int64")

def detail_to_img(val: str) -> str:
    if not isinstance(val, str):
        return ""
    s = val.strip().replace("＠", "@").replace("＂", '"').replace("＇", "'")
    m = re.search(r'@?IMAGE\s*\(\s*["\']\s*(https?://[^"\']+)\s*["\']', s, flags=re.IGNORECASE)
    if m: return m.group(1).strip()
    m = re.search(r'^[=]?\s*["\']\s*(https?://[^"\']+)\s*["\']\s*$', s)
    if m: return m.group(1).strip()
    m = re.search(r'(https?://[^\s"\')]+)', s)
    if m: return m.group(1).strip()
    if s.lower().startswith(("http://","https://")):
        return s
    parsed = urlparse(s)
    if "id=" in s:
        qs = parse_qs(parsed.query)
        id_val = qs.get("id", [parsed.path.split("/")[-1]])[0]
        if id_val:
            return f"https://dm.takaratomy.co.jp/wp-content/card/cardimage/{id_val}.jpg"
    slug = (parsed.path.split("/")[-1] or "").strip()
    if slug:
        return f"https://dm.takaratomy.co.jp/wp-content/card/cardimage/{slug}.jpg"
    return ""

def nfkc_lower(s: str) -> str:
    return ud.normalize("NFKC", s or "").lower()

def kata_to_hira(text: str) -> str:
    return "".join(chr(ord(ch) - 0x60) if "ァ" <= ch <= "ン" else ch for ch in text)

def normalize_for_search_py(s: str) -> str:
    s0 = nfkc_lower(s)
    s0 = s0.replace("complex", "こんぷれっくす").replace("c0br4", "こぶら").replace("伝説", "でんせつ")
    s0 = kata_to_hira(s0)
    s0 = SEP_RE.sub("", s0)
    return s0

def searchable_row_py(row: pd.Series) -> str:
    parts = [row.get(k, "") for k in ("name","code","pack","rarity","booster")]
    return normalize_for_search_py(" ".join(map(str, parts)))

# ========= 列アクセス =========
def get_col(df: pd.DataFrame, names: list[str], fallback_idx: int|None):
    for nm in names:
        if nm in df.columns:
            return df[nm]
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

# ========= データ整形 =========
df = pd.DataFrame({
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

# サムネ列
def url_to_hash(u:str)->str:
    return hashlib.md5(u.encode("utf-8")).hexdigest()

def ensure_thumb(url: str) -> str|None:
    if not url: return None
    THUMB_DIR.mkdir(parents=True, exist_ok=True)
    fname = url_to_hash(url) + ".webp"
    outp = THUMB_DIR / fname
    if outp.exists():
        return f"assets/thumbs/{fname}"
    if not (requests and Image):
        return None
    try:
        r = requests.get(url, timeout=12, headers={"User-Agent":"Mozilla/5.0"})
        r.raise_for_status()
        im = Image.open(io.BytesIO(r.content)).convert("RGB")
        w, h = im.size
        if w <= 0 or h <= 0: return None
        new_h = int(h * THUMB_W / w)
        im = im.resize((THUMB_W, max(1,new_h)), Image.LANCZOS)
        outp.parent.mkdir(parents=True, exist_ok=True)
        im.save(outp, "WEBP", quality=60, method=6)
        return f"assets/thumbs/{fname}"
    except Exception:
        return None

if BUILD_THUMBS:
    df["thumb"] = df["image"].map(ensure_thumb)
else:
    df["thumb"] = ""

# ====== データJSON ======
def build_payload(df: pd.DataFrame) -> tuple[str,str]:
    for c in ["name","pack","code","rarity","booster","price","image","thumb","s"]:
        if c not in df.columns:
            df[c] = "" if c!="price" else None
    records = []
    for rec in df[["name","pack","code","rarity","booster","price","image","thumb","s"]].to_dict(orient="records"):
        price = rec.get("price", None)
        try:
            price = None if pd.isna(price) else int(price)
        except Exception:
            price = None
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
        })
    payload = json.dumps(records, ensure_ascii=False, separators=(",",":"))
    ver = hashlib.md5(payload.encode("utf-8")).hexdigest()[:8]
    return ver, payload

CARDS_VER, CARDS_JSON = build_payload(df)

# ========= 見た目 =========
base_css = """
*{box-sizing:border-box}
:root{
  --bg:#ffffff; --panel:#ffffff; --border:#e5e7eb; --accent:#e11d48;
  --text:#111111; --muted:#6b7280; --header-h: 120px;
}
body{ margin:0;color:var(--text);background:var(--bg);
      font-family:Inter,system-ui,'Noto Sans JP',sans-serif;
      padding-top: calc(var(--header-h) + 8px);
}

/* === header === */
header{
  position:fixed;top:0;left:0;right:0;z-index:1000;background:#fff;border-bottom:1px solid var(--border);
  padding:10px 16px; box-shadow:0 2px 10px rgba(0,0,0,.04);
}
.header-wrap{
  max-width:1080px;margin:0 auto;display:grid;gap:8px;width:100%;
  grid-template-columns:auto 1fr auto;
  grid-template-areas: "logo title actions";
  align-items:center;
}
.brand-left{grid-area:logo;display:flex;align-items:center;gap:12px;min-width:0}
.brand-left img{height:80px;display:block}
.center-ttl{
  grid-area:title;font-weight:1000; text-align:center;
  font-size:clamp(28px, 5.2vw, 52px); line-height:1.05; color:#111;
  white-space: normal; writing-mode: horizontal-tb !important;
  text-orientation: mixed !important; word-break: keep-all; overflow-wrap: normal;
}
.right-spacer{display:none}
.actions{grid-area:actions;display:flex;align-items:center;gap:10px;justify-content:flex-end}
.iconbtn{display:inline-flex;align-items:center;gap:8px;border:1px solid var(--border);background:#fff;color:#111;border-radius:12px;padding:9px 12px;text-decoration:none;font-size:13px;transition:transform .12s ease, background .12s ease}
.iconbtn:hover{background:#f9fafb; transform:translateY(-1px)}
.iconbtn svg{width:16px;height:16px;display:block;color:#111}

/* X/LINE画像アイコン */
.iconimg{display:inline-flex;align-items:center;justify-content:center;border:1px solid var(--border);background:#fff;color:#111;border-radius:12px;text-decoration:none;overflow:hidden;transition:transform .12s ease, background .12s ease;line-height:0;}
.iconimg:hover{ background:#f9fafb; transform:translateY(-1px); }
.iconimg--x, .iconimg--line{ width:32px; height:32px; padding:0; }
.iconimg--x img{ width:118%; height:118%; object-fit:cover; }
.iconimg--line img{ width:145%; height:145%; object-fit:cover; }

/* === layout === */
.wrap{max-width:1200px;margin:0 auto;padding:12px 16px}
.controls{display:grid;grid-template-columns:repeat(2, minmax(180px,1fr));grid-template-areas:"q1 q2" "q3 q4" "acts acts";gap:10px;margin:10px 0 14px;align-items:center;}
#nameQ{grid-area:q1} #codeQ{grid-area:q2} #packQ{grid-area:q3} #rarityQ{grid-area:q4}
.controls .btns{ grid-area:acts; display:flex; gap:8px; flex-wrap:wrap }
input.search{background:#fff;border:1px solid var(--border);color:#111;border-radius:12px;padding:11px 12px;font-size:14px;outline:none;min-width:0;transition:box-shadow .12s ease;width:100%}
input.search::placeholder{color:#9ca3af}
input.search:focus{ box-shadow:0 0 0 2px rgba(17,24,39,.08) }
.btn{background:#fff;border:1px solid var(--border);color:#111;border-radius:12px;padding:9px 12px;font-size:13px;cursor:pointer;text-decoration:none;white-space:nowrap;transition:transform .12s ease, background .12s ease}
.btn:hover{background:#f9fafb; transform:translateY(-1px)}
.btn.active{outline:2px solid var(--accent)}

/* === cards === */
.grid{margin:12px 0;width:100%}
.grid.grid-img{display:grid;grid-template-columns:repeat(4, minmax(0,1fr));gap:12px}
.grid.grid-list{display:grid;grid-template-columns:repeat(2, minmax(0,1fr));gap:12px}
.card{background:var(--panel);border:1px solid var(--border);border-radius:14px;overflow:hidden;box-shadow:0 4px 10px rgba(0,0,0,.04);transition:transform .15s ease, box-shadow .15s ease;}
.card:hover{transform:translateY(-2px);box-shadow:0 10px 20px rgba(0,0,0,.06)}
.th{aspect-ratio:3/4;background:#f3f4f6;cursor:zoom-in}
.th img{width:100%;height:100%;object-fit:cover;display:block;background:#f3f4f6}
.b{padding:10px 12px}
.n{font-size:14px;font-weight:800;line-height:1.25;margin:0 0 6px;color:#111;display:flex;gap:6px;align-items:flex-start;flex-wrap:nowrap;word-break:break-word}
.n .ttl{flex:1 1 auto; min-width:0; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; word-break:keep-all; overflow-wrap:anywhere; line-height:1.25;}
.n .code{flex:0 0 auto;margin-left:0;font-weight:700;font-size:12px;color:#374151;background:#f3f4f6;border:1px solid #e5e7eb;border-radius:8px;padding:2px 6px;white-space:nowrap}
.meta{font-size:11px;color:var(--muted);word-break:break-word}
.p{margin-top:6px;display:flex;flex-wrap:nowrap;width:100%;}
.mx{font-weight:1000;color:var(--accent);font-size:clamp(16px, 2.4vw, 22px);line-height:1.05;white-space:nowrap;display:inline-block;max-width:100%;font-variant-numeric: tabular-nums;letter-spacing:-0.02em;}
.card .actrow{display:flex;gap:8px;margin-top:8px}
.card .actrow .mini{padding:6px 10px;font-size:12px;border-radius:10px}

/* スマホの「買取カートへ」が収まるように縮小 */
@media (max-width:700px){
  .card .actrow .mini{ padding:4px 8px; font-size:11px; line-height:1.1; letter-spacing:0; }
}

/* === viewer === */
.viewer{position: fixed; inset: 0; background: rgba(0,0,0,.86); display: none; align-items: center; justify-content: center; z-index: 2000;}
.viewer.show{ display: flex; }
.viewer .vc{ position: relative; max-width: 92vw; max-height: 92vh; }
.viewer img{ display:block; max-width:92vw; max-height:92vh; width:auto; height:auto; object-fit:contain; }
.viewer button.close{ position:absolute; top:-12px; right:-12px; background:#fff; border:1px solid var(--border); color:#111; border-radius:999px; width:38px; height:38px; cursor:pointer; }
body.modal-open{ overflow:hidden; }

/* === pager === */
nav.simple{ margin:14px 0; }
nav.simple .pager{ display:grid; grid-template-columns: 1fr auto 1fr;  align-items:center; gap:12px; }
nav.simple .left, nav.simple .center, nav.simple .right{ display:flex; align-items:center; gap:8px; flex-wrap:wrap; }
nav.simple .left{  justify-content:flex-start; }
nav.simple .center{justify-content:center; } 
nav.simple .right{ justify-content:flex-end; }
nav.simple a, nav.simple button{ color:#111;background:#fff;border:1px solid var(--border); padding:6px 12px;border-radius:10px;text-decoration:none;white-space:nowrap;font-size:13px;line-height:1;cursor:pointer }
nav.simple a.disabled{opacity:.45;pointer-events:none}
nav.simple .num[aria-current="page"]{ background:#111;color:#fff;border-color:#111;cursor:default }
nav.simple .ellipsis{border:none;background:transparent;cursor:default;padding:0 4px}
nav.simple .controls-mobile{ display:none !important; }

/* === SP調整（閉じ忘れ無しの完全版） === */
@media (max-width:700px){
  :root{ --header-h: 144px; }

  .n{ display:flex; flex-wrap:wrap; align-items:flex-start; }
  .n .ttl{ flex:1 1 100%; min-width:0; display:-webkit-box; -webkit-box-orient:vertical; -webkit-line-clamp:2; overflow:hidden; }
  .n .code{ order:2; margin-top:2px; font-size:10.5px; padding:1px 5px; border-radius:6px; }

  .card{ display:flex; flex-direction:column; }
  .th{ flex:0 0 auto; }
  .b{  flex:1 1 auto; display:flex; flex-direction:column; }
  .p{  margin-top:auto; padding-top:6px; justify-content:flex-start; }
  .mx{ font-size:clamp(12px, 4vw, 16px); }

  .header-wrap{
    grid-template-columns:auto 1fr;
    grid-template-areas:"logo title" "actions actions";
    gap:8px; align-items:center;
  }
  .brand-left img{ height:48px; width:auto; }
  .center-ttl{ margin-left:6px; font-size:clamp(22px, 6.2vw, 34px); }

  /* actions を少し左に寄せつつ、最右に🛒（LINEの右隣） */
  .actions{ grid-area:actions; display:flex; align-items:center; gap:8px; justify-content:flex-end; padding-right:8px; }
  .iconbtn--cart{ position:relative; display:inline-flex; align-items:center; gap:6px; }
  .iconbtn--cart .badge{ background:#e11d48; color:#fff; border-radius:999px; padding:0 6px; font-size:11px; font-weight:800; line-height:1.4; }

  /* ページャは1行死守：横スクロール */
  nav.simple .center{
    justify-content:center;
    gap:4px;
    flex-wrap:nowrap;
    overflow-x:auto;
    white-space:nowrap;
    scrollbar-width:none;
  }
  nav.simple .center::-webkit-scrollbar{ display:none; }
  nav.simple .center .num{ min-width:28px; padding:4px 8px; font-size:12px; line-height:1; }
  nav.simple .ellipsis{ padding:0 2px; }
}

/* PC: ロゴ＆タイトル拡大 */
@media (min-width:1024px){
  .brand-left img{ height:110px; }
  .center-ttl{ font-size:clamp(40px,4.6vw,72px); }
  .actions{ flex-wrap:nowrap; }
}

small.note{color:var(--muted)}

/* カートモーダル */
.cart-modal{ position:fixed; inset:0; background:rgba(0,0,0,.55); display:none; z-index:2100; align-items:center; justify-content:center; }
.cart-modal.show{ display:flex; }
.cart-sheet{ background:#fff; width:min(720px, 94vw); max-height:88vh; border-radius:14px; padding:12px; overflow:auto; }
.cart-head{ display:flex; justify-content:space-between; align-items:center; gap:8px; padding-bottom:8px; border-bottom:1px solid var(--border); }
.cart-list{ display:grid; gap:10px; padding:10px 0; }
.cart-row{ display:grid; grid-template-columns: 1fr auto auto auto; gap:8px; align-items:center; border:1px solid var(--border); border-radius:10px; padding:8px; }
.cart-row .nm{ font-weight:800 }
.cart-row .qty{ display:flex; gap:6px; align-items:center }
.cart-row .qty button{ padding:3px 8px }
.cart-row .rm{ color:#b91c1c; border-color:#fecaca }
.cart-foot{ display:flex; justify-content:space-between; align-items:center; gap:8px; padding-top:8px; border-top:1px solid var(--border) }
.cart-act{ display:flex; gap:8px; flex-wrap:wrap }
"""

# ========= JS（検索＋カート＋LINE送信） =========
base_js = r"""
(function(){
  // ====== 設定 ======
  const LIFF_ID = "__LIFF_ID__";
  const OA_ID   = "__OA_ID__";  // 先頭 @ を含む

  // ヘッダー高さ反映
  const header = document.querySelector('header');
  const setHeaderH = () => {
    const h = header?.getBoundingClientRect().height || 144;
    document.documentElement.style.setProperty('--header-h', Math.ceil(h) + 'px');
  };
  setHeaderH(); window.addEventListener('resize', setHeaderH); window.addEventListener('load', setHeaderH);
  document.querySelectorAll('header img').forEach(img => { if (!img.complete) img.addEventListener('load', setHeaderH); });
  if ('ResizeObserver' in window && header) { const ro = new ResizeObserver(() => setHeaderH()); ro.observe(header); }

  // DOM
  const nameQ=document.getElementById('nameQ'), codeQ=document.getElementById('codeQ'),
        packQ=document.getElementById('packQ'), rarityQ=document.getElementById('rarityQ'),
        grid=document.getElementById('grid'), navs=[...document.querySelectorAll('nav.simple')];
  const btnDesc=document.getElementById('btnPriceDesc'), btnAsc=document.getElementById('btnPriceAsc'),
        btnNone=document.getElementById('btnSortClear'), btnImg=document.getElementById('btnToggleImages');
  const viewer=document.getElementById('viewer'), viewerImg=document.getElementById('viewerImg'), viewerClose=document.getElementById('viewerClose');

  // ===== カートUI（重複無し・堅牢） =====
  const fabCart  = document.getElementById('fabCart');
  const fabBadge = document.getElementById('fabBadge');
  const cartModal= document.getElementById('cartModal');
  const cartList = document.getElementById('cartList');
  const cartTotal= document.getElementById('cartTotal');
  const cartSend = document.getElementById('cartSend');
  const cartClose= document.getElementById('cartClose');
  const cartClear= document.getElementById('cartClear');

  const CART_KEY="climax_cart_v1";
  let cart=[];

  function loadCart(){ try{ cart=JSON.parse(localStorage.getItem(CART_KEY)||"[]"); }catch{ cart=[]; } updateFab(); }
  function saveCart(){ localStorage.setItem(CART_KEY, JSON.stringify(cart)); updateFab(); }
  function updateFab(){ if(fabBadge) fabBadge.textContent = String(cart.reduce((a,b)=>a+(+b.qty||1),0)); }

  function addToCart(item){
    const key=(item.name||"")+"|"+(item.code||"");
    const f=cart.find(x=>((x.name||"")+"|"+(x.model||""))===key);
    if(f){ f.qty=(f.qty||1)+1; }
    else { cart.push({name:item.name, model:item.code||"", price:item.price||0, qty:1}); }
    saveCart();
  }
  function delFromCart(idx){ cart.splice(idx,1); saveCart(); renderCartModal(); }
  function incQty(idx,d){ cart[idx].qty=Math.max(1,(cart[idx].qty||1)+d); saveCart(); renderCartModal(); }
  function totalPrice(){ return cart.reduce((s,i)=>s + (i.price||0)*(i.qty||1), 0); }

  function escHtml(s){ return (s||'').replace(/[&<>\"']/g, m => ({ "&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;" }[m])); }
  function fmtYen(n){ return (n==null||n==='')?'-':('¥'+parseInt(n,10).toLocaleString()); }

  function renderCartModal(){
    if(!cartList||!cartTotal) return;
    cartList.innerHTML = cart.map((it,idx)=>`
      <div class="cart-row">
        <div class="nm">${escHtml(it.name)}${it.model?` <small>[${escHtml(it.model)}]</small>`:''}</div>
        <div class="pr">${it.price?`@${Number(it.price).toLocaleString()}円`:''}</div>
        <div class="qty">
          <button class="btn mini" data-a="dec" data-i="${idx}">－</button>
          <span>${it.qty||1}</span>
          <button class="btn mini" data-a="inc" data-i="${idx}">＋</button>
        </div>
        <button class="btn mini rm" data-a="rm" data-i="${idx}">削除</button>
      </div>`).join('') || '<p style="color:#6b7280">カートは空です。</p>';
    cartTotal.textContent = '合計：'+ totalPrice().toLocaleString() +'円';
  }

  // 開閉（直接＋デリゲーションの二段構え）
  if (fabCart) {
    fabCart.addEventListener('click', (e)=>{ e.preventDefault(); renderCartModal(); cartModal?.classList.add('show'); });
  }
  document.addEventListener('click', (e)=>{
    const trg = e.target.closest('#fabCart');
    if(trg){ e.preventDefault(); renderCartModal(); cartModal?.classList.add('show'); }
  });
  cartClose?.addEventListener('click', ()=> cartModal?.classList.remove('show'));
  cartModal?.addEventListener('click', (e)=>{ if(e.target===cartModal) cartModal.classList.remove('show'); });

  // 行内操作
  cartList?.addEventListener('click', (e)=>{
    const btn=e.target.closest('button[data-a]'); if(!btn) return;
    const a=btn.dataset.a, i=+btn.dataset.i;
    if(a==='inc') incQty(i,+1);
    if(a==='dec') incQty(i,-1);
    if(a==='rm')  delFromCart(i);
  });

  // 「カートを空にする」
  cartClear?.addEventListener('click', ()=>{
    if(!cart.length){ alert('すでに空です'); return; }
    if(!confirm('カートを空にしますか？')) return;
    cart=[]; saveCart(); renderCartModal();
  });

  // ===== 端末状況 =====
  const isMobileConst = matchMedia('(max-width: 700px)').matches;
  const netType = navigator.connection?.effectiveType || ''; const slowNet = /^(slow-2g|2g|3g)$/i.test(netType);
  const cores = navigator.hardwareConcurrency || 4;
  const __PER = __PER_PAGE__;
  const PER_PAGE_ADJ = (isMobileConst || slowNet || cores <= 4) ? Math.min(__PER, 48) : __PER;

  // 画像ON/OFF
  let showImages = (localStorage.getItem('showImages') ?? '1') === '1';
  function setImgBtn(){ if(!btnImg) return; btnImg.textContent = showImages?'画像OFF':'画像ON'; btnImg.classList.toggle('active', showImages); btnImg.setAttribute('aria-pressed', showImages?'true':'false'); }
  function toggleImages(e){ if(e){e.preventDefault();e.stopPropagation();} showImages=!showImages; localStorage.setItem('showImages', showImages?'1':'0'); setImgBtn(); render(); }

  // 正規化
  const SEP_RE=/[\s\u30FB\u00B7·/／\-_—–−]+/g;
  function kataToHira(str){ return (str||'').replace(/[\u30A1-\u30FA]/g, ch => String.fromCharCode(ch.charCodeAt(0)-0x60)); }
  const kanjiReadingMap={ "伝説":"でんせつ" }, latinAliasMap={ "complex":"こんぷれっくす", "c0br4":"こぶら" };
  function normalizeForSearch(s){ s=(s||'').normalize('NFKC').toLowerCase(); for(const[k,v]of Object.entries(latinAliasMap)) s=s.split(k).join(v); for(const[k,v]of Object.entries(kanjiReadingMap)) s=s.split(k).join(v); return kataToHira(s).replace(SEP_RE,''); }
  function normalizeLatin(s){ s=(s||'').normalize('NFKC').toLowerCase().replace(/0/g,'o').replace(/1/g,'l').replace(/3/g,'e').replace(/4/g,'a').replace(/5/g,'s').replace(/7/g,'t').replace(SEP_RE,'').replace(/[^a-z0-9]/g,''); return s; }

  // データ
  function norm(it){ return { name:it.n??it.name??"", pack:it.p??it.pack??"", code:it.c??it.code??"", rarity:it.r??it.rarity??"", booster:it.b??it.booster??"", price:(it.pr??it.price??null), image:it.i??it.image??"", thumb:it.t??it.thumb??"", s:it.s??"" }; }
  let ALL = Array.isArray(window.__CARDS__)? window.__CARDS__.map(norm) : [];
  if(!ALL.length){ const hint=document.createElement('p'); hint.style.cssText='color:#dc2626;padding:10px;margin:10px;border:1px dashed #fecaca;background:#fff5f5'; hint.textContent='データが0件です。入力CSV/Excelのヘッダと列位置を確認してください。'; document.querySelector('main')?.prepend(hint); }
  ALL = ALL.map(it=>({ ...it,
    _name:normalizeForSearch(it.name||""), _code:normalizeForSearch(it.code||""), _packbooster:normalizeForSearch([it.pack||"",it.booster||""].join(" ")), _rarity:normalizeForSearch(it.rarity||""),
    _name_lat:normalizeLatin(it.name||""), _code_lat:normalizeLatin(it.code||""), _packbooster_lat:normalizeLatin([it.pack||"",it.booster||""].join(" ")), _rarity_lat:normalizeLatin(it.rarity||"")
  }));

  // ===== 検索・描画 =====
  const SEP_HTML=' ・ ';
  let VIEW=[], page=1, currentSort=__INITIAL_SORT__;

  function matchEither(kana,latin,qK,qL){ if(!qK && !qL) return true; let ok=false; if(qK && kana.includes(qK)) ok=true; if(qL && latin.includes(qL)) ok=true; return ok; }

  function apply(){
    const qNameK=normalizeForSearch(nameQ.value||''), qCodeK=normalizeForSearch(codeQ.value||''), qPackK=normalizeForSearch(packQ.value||''), qRarityK=normalizeForSearch(rarityQ.value||'');
    const qNameL=normalizeLatin(nameQ.value||''), qCodeL=normalizeLatin(codeQ.value||''), qPackL=normalizeLatin(packQ.value||''), qRarityL=normalizeLatin(rarityQ.value||'');

    VIEW=ALL.filter(it =>
      matchEither(it._name,        it._name_lat,        qNameK, qNameL) &&
      matchEither(it._code,        it._code_lat,        qCodeK, qCodeL) &&
      matchEither(it._packbooster, it._packbooster_lat, qPackK, qPackL) &&
      matchEither(it._rarity,      it._rarity_lat,      qRarityK, qRarityL)
    );
    if(currentSort==='desc') VIEW.sort((a,b)=>(b.price||0)-(a.price||0));
    else if(currentSort==='asc') VIEW.sort((a,b)=> (a.price||0)-(b.price||0));

    page=1; render();
  }

  function buildPageButtons(cur,total){
    const around=matchMedia('(max-width: 700px)').matches?2:3, btns=[]; btns.push({type:'num',page:1});
    const start=Math.max(2,cur-around), end=Math.min(total-1,cur+around);
    if(start>2) btns.push({type:'ellipsis'}); for(let p=start;p<=end;p++){ if(p>=2 && p<=total-1) btns.push({type:'num',page:p}); }
    if(end<total-1) btns.push({type:'ellipsis'}); if(total>1) btns.push({type:'num',page:total}); return btns;
  }
  function renderPager(cur,total,isMobile){
    const first=(cur>1)?`<a href="#" data-jump="first">≪ 最初のページ</a>`:`<a class="disabled">≪ 最初のページ</a>`;
    const prev =(cur>1)?`<a href="#" data-jump="prev">← 前のページ</a>`:`<a class="disabled">← 前のページ</a>`;
    const next =(cur<total)?`<a href="#" data-jump="next">次のページ →</a>`:`<a class="disabled">次のページ →</a>`;
    const last =(cur<total)?`<a href="#" data-jump="last">最後のページ ≫</a>`:`<a class="disabled">最後のページ ≫</a>`;
    const nums=buildPageButtons(cur,total).map(item=> item.type==='ellipsis' ? `<span class="ellipsis">…</span>` : ((item.page===cur)?`<button class="num" aria-current="page" disabled>${item.page}</button>`:`<a class="num" href="#" data-page="${item.page}">${item.page}</a>`)).join(' ');
    const mobileRow=isMobile?`<div class="controls-mobile">${first} ${prev} ${next} ${last}</div>`:'';
    return `<div class="pager"><div class="left">${first} ${prev}</div><div class="center">${nums}</div><div class="right">${next} ${last}</div>${mobileRow}</div>`;
  }
  function scrollTopSmooth(){ window.scrollTo({top:0, behavior:'smooth'}); }

  function render(){
    grid.className = showImages ? 'grid grid-img' : 'grid grid-list';
    const total=VIEW.length, pages=Math.max(1, Math.ceil(total/PER_PAGE_ADJ)); if(page>pages) page=pages;
    const start=(page-1)*PER_PAGE_ADJ, rows=VIEW.slice(start, start+PER_PAGE_ADJ);

    grid.innerHTML = rows.map(it => {
      const nameEsc=escHtml(it.name||''), codeEsc=escHtml(it.code||'');
      const codeHtml = codeEsc ? `<span class="code">[${codeEsc}]</span>` : '';
      if (showImages){
        let thumb=it.thumb||it.image||''; const hasHttp=/^https?:\/\//.test(thumb); if(thumb && !hasHttp && !thumb.startsWith('../')) thumb='../'+thumb;
        return `
  <article class="card">
    <div class="th" data-full="${it.image||''}">
      <img alt="${nameEsc}" loading="lazy" decoding="async" width="600" height="800" data-src="${thumb}" src=""
           onerror="this.onerror=null;var p=this.closest(' .th ');this.src=p?p.getAttribute('data-full'):this.src;">
    </div>
    <div class="b">
      <h3 class="n"><span class="ttl">${nameEsc}</span>${codeHtml}</h3>
      <div class="p"><span class="mx">${fmtYen(it.price)}</span></div>
      <div class="actrow"><button class="btn mini btn-add" data-name="${nameEsc}" data-model="${codeEsc}" data-price="${it.price||0}">買取カートへ</button></div>
    </div>
  </article>`;
      }else{
        const meta=[it.code||'', [it.pack||'',it.booster||''].filter(Boolean).join(' / ')].filter(Boolean).join(SEP_HTML);
        return `
  <article class="card">
    <div class="b">
      <h3 class="n"><span class="ttl">${nameEsc}</span>${codeHtml}</h3>
      <div class="meta">${escHtml(meta)}</div>
      <div class="p"><span class="mx">${fmtYen(it.price)}</span></div>
      <div class="actrow"><button class="btn mini btn-add" data-name="${nameEsc}" data-model="${codeEsc}" data-price="${it.price||0}">買取カートへ</button></div>
    </div>
  </article>`;
      }
    }).join('');

    if(showImages){
      grid.querySelectorAll('.th').forEach(th=>{
        th.addEventListener('click', ()=>{ const src=th.getAttribute('data-full') || th.querySelector('img')?.src || ''; if(!src) return; viewerImg.src=src; viewer.classList.add('show'); document.body.classList.add('modal-open'); });
      });
    }
    grid.querySelectorAll('.btn-add').forEach(btn=>{
      btn.addEventListener('click', ()=>{ addToCart({name:btn.dataset.name, code:btn.dataset.model, price:Number(btn.dataset.price)||0}); });
    });

    const isMobileNow = window.matchMedia('(max-width: 700px)').matches;
    const pagerHtml=renderPager(page,pages,isMobileNow);
    navs.forEach(n=>{
      n.innerHTML=pagerHtml; n.onclick=(e)=>{
        const a=e.target.closest('a,button'); if(!a || a.matches('.disabled,[disabled]')) return; e.preventDefault();
        if(a.dataset.jump==='first'){ page=1; render(); scrollTopSmooth(); return; }
        if(a.dataset.jump==='prev'){ page=Math.max(1,page-1); render(); scrollTopSmooth(); return; }
        if(a.dataset.jump==='next'){ page=Math.min(pages,page+1); render(); scrollTopSmooth(); return; }
        if(a.dataset.jump==='last'){ page=pages; render(); scrollTopSmooth(); return; }
        const p=parseInt(a.dataset.page||'0',10); if(p && p!==page){ page=p; render(); scrollTopSmooth(); }
      };
    });

    // 画像lazy
    if(showImages){
      const io=new IntersectionObserver((entries)=>{
        entries.forEach(e=>{
          if(e.isIntersecting){
            const img=e.target; const ds=img.getAttribute('data-src');
            if(ds && !img.src){ img.src=ds; img.removeAttribute('data-src'); }
            io.unobserve(img);
          }
        });
      }, { rootMargin:(isMobileConst||slowNet)?"300px 0px":"600px 0px", threshold:0.01 });
      document.querySelectorAll('#grid img[data-src]').forEach(img=>io.observe(img));
    }
  }

  function setActiveSort(){
    btnDesc?.setAttribute('aria-pressed', currentSort==='desc' ? 'true':'false');
    btnAsc ?.setAttribute('aria-pressed', currentSort==='asc'  ? 'true':'false');
    btnNone?.setAttribute('aria-pressed', currentSort===null    ? 'true':'false');
    btnDesc?.classList.toggle('active', currentSort==='desc'); btnAsc?.classList.toggle('active', currentSort==='asc'); btnNone?.classList.toggle('active', currentSort===null);
  }

  function closeViewer(){ viewer.classList.remove('show'); viewerImg.src=''; document.body.classList.remove('modal-open'); }
  viewerClose?.addEventListener('click', closeViewer);
  viewer?.addEventListener('click', (e)=>{ if(e.target===viewer) closeViewer(); });
  window.addEventListener('keydown', (e)=>{ if(e.key==='Escape') closeViewer(); });

  // LIFF SDK（CDN・一発追加）
  (function addLiff(){ if(document.querySelector('script[data-liff]')) return;
    const s=document.createElement('script'); s.src='https://static.line-scdn.net/liff/edge/2/sdk.js'; s.async=true; s.setAttribute('data-liff','1'); document.head.appendChild(s);
  })();

  // LINE送信：LIFF内→sendMessages / それ以外→oaMessage ディープリンク
  function openOAMessage(text){
    const enc=encodeURIComponent(text);
    const app=`line://oaMessage/${OA_ID}/?${enc}`;
    const web=`https://line.me/R/oaMessage/${OA_ID}/?${enc}`;
    location.href = app; setTimeout(()=>window.open(web,"_blank"), 500);
  }
  async function trySendViaLIFF(text){
    if(!window.liff) return false;
    try{
      if(!liff._climaxInited){ await liff.init({ liffId: LIFF_ID }); liff._climaxInited=true; }
      if(!liff.isLoggedIn()){ liff.login(); return false; }
      if(!liff.isInClient()) return false;
      await liff.sendMessages([{ type:'text', text }]);
      alert('LINEに仮査定リストを送信しました。'); liff.closeWindow?.(); return true;
    }catch(e){ return false; }
  }

  function buildCartText(){
    const lines=[]; let total=0;
    for(const it of cart){ const q=+it.qty||1, p=+it.price||0; total+=q*p; lines.push(`・${it.name}${it.model?`［${it.model}］`:''} ×${q}${p?` @${p.toLocaleString()}円`:''}`); }
    let body=`【仮査定依頼】\n`+lines.join("\n"); body+=`\n――――\n合計：${total.toLocaleString()}円\n※仮査定です。現物確認後に正式査定となります。`;
    if(body.length>900) body=body.slice(0,880)+"\n…（続き省略）";
    return body;
  }

  cartSend?.addEventListener('click', async ()=>{
    if(!cart.length){ alert('カートが空です。'); return; }
    const text=buildCartText();
    if(await trySendViaLIFF(text)) return;
    openOAMessage(text);
  });

  // 入力遅延
  const DEBOUNCE = (isMobileConst||slowNet||cores<=4)?240:120;
  function onInputDebounced(el){ el.addEventListener('input', ()=>{ clearTimeout(el._t); el._t=setTimeout(apply,DEBOUNCE); }); }
  [nameQ,codeQ,packQ,rarityQ].forEach(onInputDebounced);

  // sort/画像ボタン
  function initButtons(){
    btnDesc?.addEventListener('click', ()=>{ currentSort=(currentSort==='desc')?null:'desc'; setActiveSort(); apply(); });
    btnAsc ?.addEventListener('click', ()=>{ currentSort=(currentSort==='asc') ?null:'asc' ; setActiveSort(); apply(); });
    btnNone?.addEventListener('click', ()=>{ currentSort=null; setActiveSort(); apply(); });
    btnImg ?.addEventListener('click', toggleImages);
    setImgBtn();
  }

  setActiveSort(); initButtons(); loadCart(); apply();
})();
"""

# ===== HTML =====
def html_page(title: str, js_source: str, logo_uri: str, cards_json: str) -> str:
    shop_svg = "<svg viewBox='0 0 24 24' aria-hidden='true' fill='currentColor'><path d='M3 9.5V8l2.2-3.6c.3-.5.6-.7 1-.7h11.6c.4 0 .7.2 .9 .6L21 8v1.5c0 1-.8 1.8-1.8 1.8-.9 0-1.6-.6-1.8-1.4-.2 .8-.9 1.4-1.8 1.4s-1.6-.6-1.8-1.4c-.2 .8-.9 1.4-1.8 1.4C3.8 11.3 3 10.5 3 9.5zM5 12.5h14V20c0 .6-.4 1-1 1H6c-.6 0-1-.4-1-1v-7.5zm4 1.5v5h6v-5H9zM6.3 5.2 5 7.5h14l-1.3-2.3H6.3z'/></svg>"
    login_svg= "<svg viewBox='0 0 24 24' aria-hidden='true' fill='currentColor'><path d='M12 12a5 5 0 1 0-5-5 5 5 0 0 0 5 5zm0 2c-4.418 0-8 2.239-8 5v2h16v-2c0-2.761-3.582-5-8-5z'/></svg>"

    parts = []
    parts.append("<!doctype html><html lang='ja'><head><meta charset='utf-8'>")
    parts.append("<meta name='viewport' content='width=device-width,initial-scale=1'>")
    parts.append("<style>"); parts.append(base_css); parts.append("</style>")
    # LIFF SDK（CDN読込）
    parts.append("<script src='https://static.line-scdn.net/liff/edge/2/sdk.js' data-liff></script>")
    parts.append("</head><body>")

    parts.append("<header><div class='header-wrap'>")
    parts.append("<div class='brand-left'>")
    if logo_uri:
        parts.append(f"<img src='{logo_uri}' alt='Shop Logo'>")
    else:
        parts.append("<div class='brand-fallback' aria-label='Shop Logo'>YOUR SHOP</div>")
    parts.append("</div>")
    parts.append(f"<div class='center-ttl'>{html_mod.escape(title)}</div>")
    parts.append("<div class='right-spacer'></div>")
    parts.append("<div class='actions'>")
    parts.append(f"<a class='iconbtn' href='https://www.climax-card.jp/' target='_blank' rel='noopener'>{shop_svg}<span>Shop</span></a>")
    parts.append(f"<a class='iconbtn' href='https://www.climax-card.jp/member-login' target='_blank' rel='noopener'>{login_svg}<span>Login</span></a>")
    if X_ICON_URI:
        parts.append(f"<a class='iconimg iconimg--x' href='https://x.com/climaxcard' target='_blank' rel='noopener'><img src='{X_ICON_URI}' alt='X'></a>")
    if LINE_ICON_URI:
        parts.append(f"<a class='iconimg iconimg--line' href='https://line.me/R/ti/p/{OA_ID}' target='_blank' rel='noopener'><img src='{LINE_ICON_URI}' alt='LINE'></a>")
    # ← LINEアイコンの“右側”にカート（唯一の🛒）
    parts.append("<button id='fabCart' class='iconbtn iconbtn--cart' type='button'>🛒 <span id='fabBadge' class='badge'>0</span></button>")
    parts.append("</div></div></header>")

    parts.append("<main class='wrap'>")
    parts.append("<div class='controls'>")
    parts.append("  <input id='nameQ'   class='search' placeholder='名前：部分一致（ひらがな可）'>")
    parts.append("  <input id='codeQ'   class='search' placeholder='型番：例 SP4/SP5, 19/100 等'>")
    parts.append("  <input id='packQ'   class='search' placeholder='弾：例 DM25RP1, 邪神VS邪神 等'>")
    parts.append("  <input id='rarityQ' class='search' placeholder='レアリティ：SR/VR 等'>")
    parts.append("  <div class='btns'>")
    parts.append("    <button id='btnPriceDesc' class='btn' type='button' aria-pressed='false'>価格高い順</button>")
    parts.append("    <button id='btnPriceAsc'  class='btn' type='button' aria-pressed='false'>価格低い順</button>")
    parts.append("    <button id='btnSortClear' class='btn' type='button' aria-pressed='false'>標準順</button>")
    parts.append("    <button id='btnToggleImages' class='btn' type='button'>画像ON</button>")
    parts.append("  </div></div>")
    parts.append("  <nav class='simple'></nav><div id='grid' class='grid grid-img'></div><nav class='simple'></nav>")
    parts.append("  <small class='note'>画像クリックで拡大。🛒カートから「LINEに送る」でCLIMAX公式LINEへ送信できます（アプリ外でもOK）。</small>")
    parts.append("</main>")

    # 画像ビューワ
    parts.append("<div id='viewer' class='viewer'><div class='vc'><img id='viewerImg' alt=''><button id='viewerClose' class='close'>×</button></div></div>")

    # データ
    parts.append("<script>window.__CARDS__=" + cards_json + ";</script>")

    # カートモーダル
    parts.append("""
<div id="cartModal" class="cart-modal">
  <div class="cart-sheet" role="dialog" aria-modal="true" aria-label="仮査定カート">
    <div class="cart-head">
      <strong>仮査定カート</strong>
      <button id="cartClose" class="btn">閉じる</button>
    </div>
    <div id="cartList" class="cart-list"></div>
    <div class="cart-foot">
      <div id="cartTotal">合計：0円</div>
      <div class="cart-act">
        <button id="cartClear" class="btn" style="border-color:#fecaca;color:#b91c1c;">カートを空にする</button>
        <button id="cartSend" class="btn" style="background:#06C755;color:#fff;border-color:#06C755;">LINEに送る</button>
      </div>
    </div>
  </div>
</div>
""")

    # JS本体
    parts.append("<script>"); parts.append(js_source); parts.append("</script>")
    parts.append("</body></html>")
    return "".join(parts)

# ========= 出力 =========
OUT_DIR.mkdir(parents=True, exist_ok=True)

def write_mode(dir_name: str, initial_sort_js_literal: str, title_text: str):
    sub = OUT_DIR / dir_name
    sub.mkdir(parents=True, exist_ok=True)
    js = (base_js
          .replace("__PER_PAGE__", str(PER_PAGE))
          .replace("__INITIAL_SORT__", initial_sort_js_literal)
          .replace("__LIFF_ID__", LIFF_ID)
          .replace("__OA_ID__", OA_ID)
          )
    html = html_page(title_text, js, LOGO_URI, CARDS_JSON)
    (sub / "index.html").write_text(html, encoding="utf-8")

# 初期表示は価格高い順。3モード生成
write_mode("default", "'desc'", "デュエマ買取表")
write_mode("price_desc", "'desc'", "デュエマ買取表（price_desc）")
write_mode("price_asc",  "'asc'",  "デュエマ買取表（price_asc）")

# ルートは default/ にリダイレクト
(OUT_DIR / "index.html").write_text("<meta http-equiv='refresh' content='0; url=default/'>", encoding="utf-8")

print(f"[*] Excel/CSV: {EXCEL_PATH!r}")
print(f"[*] PER_PAGE={PER_PAGE}  BUILD_THUMBS={'1' if BUILD_THUMBS else '0'}")
print(f"[LOGO] {'embedded' if LOGO_URI else 'not found (fallback text used)'}")
print(f"[X ICON] {'embedded' if X_ICON_URI else 'not found'}")
print(f"[LINE ICON] {'embedded' if LINE_ICON_URI else 'not found'}")
print(f"[OK] 生成完了 → {OUT_DIR.resolve()} / 総件数{len(df)}")
