# -*- coding: utf-8 -*-
"""
デュエマ買取表 静的ページ生成（完成版・中央タイトル＋ロゴ assets/logo.png 対応）
- CSV/Excel 自動対応。二重ヘッダ(日本語/英語キー)も自動正規化
- 列は「ヘッダ名優先 → 位置フォールバック(C/E/F/G/H/O/Q)」
- 画像URLは Q列系（allow_auto_print_label 等）最優先。=IMAGE() 抽出にも対応
- 画像ON時は「カード名＋型番＋買取価格のみ」表示（スマホ最適化：型番はバッジ・nowrap）
- ロゴは LOGO_FILE 環境変数 or assets/logo.png を最優先で埋め込み（base64）
- 見出しは 2段ヘッダー（PC: 1行 / SP: 2行）。高さはJSで測って被り防止
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

# ========= 設定 =========
DEFAULT_EXCEL = "buylist.xlsx"
ALT_EXCEL     = "data/buylist.xlsx"
FALLBACK_WINDOWS = r"C:\Users\user\Desktop\デュエマ買取表\buylist.xlsx"

# 入力は CSV/Excel 自動検出
EXCEL_PATH = os.getenv("EXCEL_PATH", DEFAULT_EXCEL)
SHEET_NAME = os.getenv("SHEET_NAME", "シート1")

# 出力
OUT_DIR    = Path(os.getenv("OUT_DIR", "docs"))
PER_PAGE   = int(os.getenv("PER_PAGE", "80"))
BUILD_THUMBS = os.getenv("BUILD_THUMBS", "0") == "1"

# ロゴ（任意）
LOGO_FILE_ENV = os.getenv("LOGO_FILE", "").strip()

# argvでファイルパス上書き
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

def logo_to_data_uri(p: Path|None) -> str:
    if not p: return ""
    mime = mimetypes.guess_type(str(p))[0] or "image/png"
    try:
        b64  = base64.b64encode(p.read_bytes()).decode("ascii")
        return f"data:{mime};base64,{b64}"
    except Exception:
        return ""

LOGO_URI = logo_to_data_uri(find_logo_path())

# ========= 入力ファイル 読み込み・正規化（CSV/Excel自動対応） =========
def _read_csv_auto(path: Path) -> pd.DataFrame:
    for enc in ("utf-8-sig", "cp932", "utf-8"):
        try:
            return pd.read_csv(path, encoding=enc)
        except Exception:
            pass
    return pd.read_csv(path)

def _normalize_two_header_layout(df: pd.DataFrame) -> pd.DataFrame:
    """
    二重ヘッダ（2行見出し→英語キー行→データ）を、英語キー行を正式ヘッダに整える。
    """
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
    s = val.strip()
    s = s.replace("＠", "@").replace("＂", '"').replace("＇", "'")

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

# ========= 見た目（PC=1行／SP=2行） =========
base_css = """
*{box-sizing:border-box}
:root{
  --bg:#ffffff; --panel:#ffffff; --border:#e5e7eb; --accent:#e11d48;
  --text:#111111; --muted:#6b7280; --header-h: 120px;
}
body{ margin:0;color:var(--text);background:var(--bg);font-family:Inter,system-ui,'Noto Sans JP',sans-serif; padding-top: var(--header-h); }

/* === header（PCは1行：logo | title | actions、SPは2行） === */
header{
  position:fixed;top:0;left:0;right:0;z-index:1000;background:#fff;border-bottom:1px solid var(--border);
  padding:10px 16px; box-shadow:0 2px 10px rgba(0,0,0,.04);
}
.header-wrap{
  max-width:1200px;margin:0 auto;display:grid;gap:8px;width:100%;
  grid-template-columns:auto 1fr auto;
  grid-template-areas: "logo title actions";   /* ← PCは1行に固定 */
  align-items:center;
}
.brand-left{grid-area:logo;display:flex;align-items:center;gap:12px;min-width:0}
.brand-left img{height:80px;display:block}      /* ← ロゴを大きく（PC） */
.center-ttl{
  grid-area:title; font-weight:1000; text-align:center;
  font-size:clamp(28px, 5.2vw, 52px); line-height:1.05; color:#111;
}
.right-spacer{display:none}                     /* PCでは未使用 */
.actions{grid-area:actions;display:flex;align-items:center;gap:10px;justify-content:flex-end}
.iconbtn{display:inline-flex;align-items:center;gap:8px;border:1px solid var(--border);background:#fff;color:#111;border-radius:12px;padding:9px 12px;text-decoration:none;font-size:13px;transition:transform .12s ease, background .12s ease}
.iconbtn:hover{background:#f9fafb; transform:translateY(-1px)}
.iconbtn svg{width:16px;height:16px;display:block;color:#111}

.wrap{max-width:1200px;margin:0 auto;padding:12px 16px}
.controls{
  display:grid;grid-template-columns:repeat(2, minmax(180px,1fr));
  grid-template-areas: "q1 q2" "q3 q4" "acts acts";
  gap:10px;margin:10px 0 14px;align-items:center;
}
#nameQ{grid-area:q1}
#codeQ{grid-area:q2}
#packQ{grid-area:q3}
#rarityQ{grid-area:q4}
.controls .btns{ grid-area:acts; display:flex; gap:8px; flex-wrap:wrap }

input.search{background:#fff;border:1px solid var(--border);color:#111;border-radius:12px;padding:11px 12px;font-size:14px;outline:none;min-width:0;transition:box-shadow .12s ease;width:100%}
input.search::placeholder{color:#9ca3af}
input.search:focus{ box-shadow:0 0 0 2px rgba(17,24,39,.08) }
.btn{background:#fff;border:1px solid var(--border);color:#111;border-radius:12px;padding:9px 12px;font-size:13px;cursor:pointer;text-decoration:none;white-space:nowrap;transition:transform .12s ease, background .12s ease}
.btn:hover{background:#f9fafb; transform:translateY(-1px)}
.btn.active{outline:2px solid var(--accent)}

.grid{margin:12px 0;width:100%}
.grid.grid-img{display:grid;grid-template-columns:repeat(4, minmax(0,1fr));gap:12px}
.grid.grid-list{display:grid;grid-template-columns:repeat(2, minmax(0,1fr));gap:12px}
.card{
  background:var(--panel);border:1px solid var(--border);border-radius:14px;overflow:hidden;box-shadow:0 4px 10px rgba(0,0,0,.04);transition:transform .15s ease, box-shadow .15s ease;
  content-visibility:auto; contain-intrinsic-size: 600px 800px;
}
.card:hover{transform:translateY(-2px);box-shadow:0 10px 20px rgba(0,0,0,.06)}
.th{aspect-ratio:3/4;background:#f3f4f6;cursor:zoom-in}
.th img{width:100%;height:100%;object-fit:cover;display:block;background:#f3f4f6}
.b{padding:10px 12px}

/* 型番のめり込み対策：バッジ化＋nowrap、タイトルはflexで折返し最適化 */
.n{font-size:14px;font-weight:800;line-height:1.25;margin:0 0 6px;color:#111;display:flex;gap:6px;align-items:baseline;flex-wrap:wrap;word-break:break-word}
.n .code{margin-left:0;font-weight:700;font-size:12px;color:#374151;background:#f3f4f6;border:1px solid #e5e7eb;border-radius:8px;padding:2px 6px;white-space:nowrap}

.meta{font-size:11px;color:var(--muted);word-break:break-word}
.p{margin-top:6px;display:flex;flex-wrap:wrap}
.mx{font-weight:1000;color:var(--accent);font-size:clamp(16px, 2.4vw, 22px);line-height:1.05;text-shadow:none;white-space:nowrap;display:inline-block;max-width:100%}
.grid.grid-img .meta{display:none}

nav.simple{display:flex;justify-content:center;align-items:center;margin:14px 0;gap:14px;flex-wrap:wrap}
nav.simple a{color:#111;background:#fff;border:1px solid var(--border);padding:8px 16px;border-radius:12px;text-decoration:none;white-space:nowrap}
nav.simple a.disabled{opacity:.45;pointer-events:none}
nav.simple strong{color:#111;user-select:none}

.viewer{position:fixed; inset:0; background:rgba(0,0,0,.86); display:none; align-items:center; justify-content:center; z-index:2000}
.viewer.show{display:flex}
.viewer .vc{position:relative;max-width:92vw;max-height:92vh}
.viewer img{max-width:92vw;max-height:92vh;display:block}
.viewer button.close{position:absolute;top:-12px;right:-12px;background:#fff;border:1px solid var(--border);color:#111;border-radius:999px;width:38px;height:38px;cursor:pointer}

/* ===== SPは2行（ロゴ/タイトル/スペーサー + アクション）に戻す ===== */
@media (max-width:700px){
  :root{ --header-h: 144px; }
  .header-wrap{
    grid-template-columns:auto 1fr auto;
    grid-template-areas:
      "logo title spacer"
      "actions actions actions";
  }
  .brand-left img{height:56px}
  .right-spacer{display:block; grid-area:spacer;}
  .actions{justify-content:center}
  .center-ttl{ font-size:clamp(24px, 7vw, 36px) }
  .wrap{ padding:4px }
  .grid.grid-img{ gap:2px }
  .b{padding:6px}
  .n{font-size:12px}
  .n .code{font-size:11px;padding:1px 6px;border-radius:6px}
  .mx{ font-size:clamp(12px, 4.2vw, 16px); white-space:nowrap }
  nav.simple{gap:8px; flex-wrap:nowrap; justify-content:space-between}
  nav.simple a{padding:6px 10px; font-size:12px; display:inline-flex}
  nav.simple strong{font-size:12px}
}
small.note{color:var(--muted)}
"""

# ========= JS =========
base_js = r"""
(function(){
  const header = document.querySelector('header');
  const setHeaderH = () => {
    const h = header?.offsetHeight || 144;
    document.documentElement.style.setProperty('--header-h', h + 'px');
  };
  setHeaderH();
  window.addEventListener('resize', setHeaderH);

  const nameQ  = document.getElementById('nameQ');
  const codeQ  = document.getElementById('codeQ');
  const packQ  = document.getElementById('packQ');
  const rarityQ= document.getElementById('rarityQ');
  const grid   = document.getElementById('grid');
  const navs   = [...document.querySelectorAll('nav.simple')];

  const btnDesc  = document.getElementById('btnPriceDesc');
  const btnAsc   = document.getElementById('btnPriceAsc');
  const btnNone  = document.getElementById('btnSortClear');
  const btnImg   = document.getElementById('btnToggleImages');

  const viewer = document.getElementById('viewer');
  const viewerImg = document.getElementById('viewerImg');
  const viewerClose = document.getElementById('viewerClose');

  const isMobile = matchMedia('(max-width: 700px)').matches;
  const netType = navigator.connection?.effectiveType || '';
  const slowNet = /^(slow-2g|2g|3g)$/i.test(netType);
  const cores = navigator.hardwareConcurrency || 4;

  const __PER = __PER_PAGE__;
  const PER_PAGE_ADJ = (isMobile || slowNet || cores <= 4) ? Math.min(__PER, 48) : __PER;

  function pick(nLo, nMd, nHi){
    return (cores <= 4) ? nLo : ((isMobile || slowNet) ? nMd : nHi);
  }
  const eager1 = pick(4, 8, 16);
  const eager2 = pick(8, 16, 32);

  // 初期表示は画像ON
  let showImages;
  let saved = localStorage.getItem('showImages');
  if (saved === 'true')  { localStorage.setItem('showImages','1'); saved = '1'; }
  if (saved === 'false') { localStorage.setItem('showImages','0'); saved = '0'; }
  if (saved === null) showImages = true;
  else showImages = (saved === '1');

  const SEP_RE = /[\s\u30FB\u00B7·/／\-_—–−]+/g;
  function kataToHira(str){ return (str||'').replace(/[\u30A1-\u30FA]/g, ch => String.fromCharCode(ch.charCodeAt(0)-0x60)); }
  const kanjiReadingMap = { "伝説":"でんせつ" };
  const latinAliasMap = { "complex": "こんぷれっくす", "c0br4": "こぶら" };

  function normalizeForSearch(s){
    s = (s||'').normalize('NFKC').toLowerCase();
    for(const [k,v] of Object.entries(latinAliasMap)){ s = s.split(k).join(v); }
    for(const [k,v] of Object.entries(kanjiReadingMap)){ s = s.split(k).join(v); }
    s = kataToHira(s).replace(SEP_RE, '');
    return s;
  }
  function normalizeLatin(s){
    s = (s||'').normalize('NFKC').toLowerCase()
      .replace(/0/g,'o').replace(/1/g,'l').replace(/3/g,'e').replace(/4/g,'a').replace(/5/g,'s').replace(/7/g,'t')
      .replace(SEP_RE, '').replace(/[^a-z0-9]/g,'');
    return s;
  }

  function fmtYen(n){ return (n==null||n==='')?'-':('¥'+parseInt(n,10).toLocaleString()); }
  function escHtml(s){
    return (s||'').replace(/[&<>\"']/g, m => ({ "&":"&amp;", "<":"&lt;", ">":"&gt;", "\"":"&quot;", "'":"&#39;" }[m]));
  }

  function norm(it){
    return {
      name: it.n ?? it.name ?? "",
      pack: it.p ?? it.pack ?? "",
      code: it.c ?? it.code ?? "",
      rarity: it.r ?? it.rarity ?? "",
      booster: it.b ?? it.booster ?? "",
      price: (it.pr ?? it.price ?? null),
      image: it.i ?? it.image ?? "",
      thumb: it.t ?? it.thumb ?? "",
      s: it.s ?? ""
    };
  }
  let ALL = Array.isArray(window.__CARDS__) ? window.__CARDS__.map(norm) : [];
  if (!ALL.length) {
    const hint = document.createElement('p');
    hint.style.cssText='color:#dc2626;padding:10px;margin:10px;border:1px dashed #fecaca;background:#fff5f5';
    hint.textContent = 'データが0件です。入力CSV/Excelのヘッダと列位置を確認してください。';
    document.querySelector('main')?.prepend(hint);
  }

  ALL = ALL.map(it => ({
    ...it,
    _name:        normalizeForSearch(it.name || ""),
    _code:        normalizeForSearch(it.code || ""),
    _packbooster: normalizeForSearch([it.pack || "", it.booster || ""].join(" ")),
    _rarity:      normalizeForSearch(it.rarity || ""),
    _name_lat:        normalizeLatin(it.name || ""),
    _code_lat:        normalizeLatin(it.code || ""),
    _packbooster_lat: normalizeLatin([it.pack || "", it.booster || ""].join(" ")),
    _rarity_lat:      normalizeLatin(it.rarity || "")
  }));

  let VIEW=[]; let page=1; let currentSort=__INITIAL_SORT__;

  function shrinkPrices(root=document){
    const MIN_PX = 10;
    root.querySelectorAll('.mx').forEach(el=>{
      const style = window.getComputedStyle(el);
      let size = parseFloat(style.fontSize) || 14;
      const fits = () => el.scrollWidth <= el.clientWidth;
      if (fits()) return;
      while (!fits() && size > MIN_PX) { size -= 1; el.style.fontSize = size + 'px'; }
    });
  }

  function cardHTML_img(it){
    const nameEsc = escHtml(it.name||'');
    const full = it.image||'';
    const codeEsc = escHtml(it.code||'');
    let thumb = it.thumb || full;
    const hasHttp = /^https?:\/\//.test(thumb);
    if (thumb && !hasHttp && !thumb.startsWith('../')) thumb = '../' + thumb;
    const codeHtml = codeEsc ? `<span class="code">[${codeEsc}]</span>` : '';
    return `
  <article class="card">
    <div class="th" data-full="${full}">
      <img alt="${nameEsc}" loading="lazy" decoding="async"
           width="600" height="800"
           data-src="${thumb}" src=""
           onerror="this.onerror=null;var p=this.closest('.th');this.src=p?p.getAttribute('data-full'):this.src;">
    </div>
    <div class="b">
      <h3 class="n">${nameEsc}${codeHtml}</h3>
      <div class="p"><span class="mx">${fmtYen(it.price)}</span></div>
    </div>
  </article>`;
  }

  function cardHTML_list(it){
    const nameEsc = escHtml(it.name||'');
    const meta = [it.code||'', [it.pack||'', it.booster||''].filter(Boolean).join(' / ')].filter(Boolean).join(' ・ ');
    return `
  <article class="card">
    <div class="b">
      <h3 class="n">${nameEsc}</h3>
      <div class="meta">${escHtml(meta)}</div>
      <div class="p"><span class="mx">${fmtYen(it.price)}</span></div>
    </div>
  </article>`;
  }

  let io;
  function setupIO(){
    if (io) io.disconnect();
    io = new IntersectionObserver((entries)=>{
      entries.forEach(e=>{
        if (e.isIntersecting){
          const img = e.target;
          const ds = img.getAttribute('data-src');
          if (ds && !img.src){ img.src = ds; img.removeAttribute('data-src'); }
          io.unobserve(img);
        }
      });
    }, { rootMargin: (isMobile || slowNet) ? "300px 0px" : "600px 0px", threshold: 0.01 });

    document.querySelectorAll('#grid img[data-src]').forEach(img=>io.observe(img));
    document.querySelectorAll('#grid img').forEach((img, i)=>{ img.setAttribute('fetchpriority', i < 8 ? 'high' : 'low'); });
  }

  function eagerLoad(n){
    const imgs=[...document.querySelectorAll('#grid img[data-src]')].slice(0,n);
    imgs.forEach(img=>{ if(!img.src){ img.src = img.getAttribute('data-src'); img.removeAttribute('data-src'); }});
  }

  function apply(){
    const qNameK   = normalizeForSearch(nameQ.value||'');
    const qCodeK   = normalizeForSearch(codeQ.value||'');
    const qPackK   = normalizeForSearch(packQ.value||'');
    const qRarityK = normalizeForSearch(rarityQ.value||'');

    const qNameL   = normalizeLatin(nameQ.value||'');
    const qCodeL   = normalizeLatin(codeQ.value||'');
    const qPackL   = normalizeLatin(packQ.value||'');
    const qRarityL = normalizeLatin(rarityQ.value||'');

    const matchEither = (kana, latin, qK, qL) => {
      if (!qK && !qL) return true;
      let ok = false;
      if (qK && kana.includes(qK)) ok = true;
      if (qL && latin.includes(qL)) ok = true;
      return ok;
    };

    VIEW = ALL.filter(it =>
      matchEither(it._name,        it._name_lat,        qNameK,   qNameL)   &&
      matchEither(it._code,        it._code_lat,        qCodeK,   qCodeL)   &&
      matchEither(it._packbooster, it._packbooster_lat, qPackK,   qPackL)   &&
      matchEither(it._rarity,      it._rarity_lat,      qRarityK, qRarityL)
    );

    if(currentSort==='desc') VIEW.sort((a,b)=>(b.price||0)-(a.price||0));
    else if(currentSort==='asc') VIEW.sort((a,b)=>(a.price||0)-(b.price||0));

    page=1; render();
  }

  function render(){
    grid.className = showImages ? 'grid grid-img' : 'grid grid-list';
    const total=VIEW.length;
    const pages=Math.max(1, Math.ceil(total/PER_PAGE_ADJ));
    if(page>pages) page=pages;
    const start=(page-1)*PER_PAGE_ADJ;
    const rows=VIEW.slice(start, start+PER_PAGE_ADJ);
    grid.innerHTML = rows.map(showImages ? cardHTML_img : cardHTML_list).join('');

    if(showImages){
      grid.querySelectorAll('.th').forEach(th=>{
        th.addEventListener('click', ()=>{
          const src = th.getAttribute('data-full') || th.querySelector('img')?.src || '';
          if(!src) return; viewerImg.src = src; viewer.classList.add('show');
        });
      });
    }

    const prev = page>1 ? `<a href="#" data-jump="prev">← 前のページ</a>` : `<a class="disabled">← 前のページ</a>`;
    const next = page<pages ? `<a href="#" data-jump="next">次のページ →</a>` : `<a class="disabled">次のページ →</a>`;
    const navHtml = `${prev} &nbsp;&nbsp; <strong>${page}/${pages}</strong> &nbsp;&nbsp; ${next}`;
    navs.forEach(n=>{
      n.innerHTML=navHtml;
      n.onclick=(e)=>{
        const a=e.target.closest('a[data-jump]'); if(!a) return;
        e.preventDefault();
        const j=a.dataset.jump; if(j==='prev') page--; else if(j==='next') page++;
        render(); window.scrollTo({ top: 0, behavior: 'smooth' });
      };
    });

    shrinkPrices(grid);
    if (showImages) {
      setupIO();
      eagerLoad(eager1);
      setTimeout(()=>eagerLoad(eager2), 600);
    }
  }

  function setActiveSort(){
    btnDesc?.setAttribute('aria-pressed', currentSort==='desc' ? 'true':'false');
    btnAsc ?.setAttribute('aria-pressed', currentSort==='asc'  ? 'true':'false');
    btnNone?.setAttribute('aria-pressed', currentSort===null    ? 'true':'false');
    btnDesc?.classList.toggle('active', currentSort==='desc');
    btnAsc ?.classList.toggle('active', currentSort==='asc');
    btnNone?.classList.toggle('active', currentSort===null);
  }
  function setImgBtn(){
    if (!btnImg) return;
    btnImg.textContent = showImages ? '画像OFF' : '画像ON';
    btnImg.classList.toggle('active', showImages);
    btnImg.setAttribute('aria-pressed', showImages ? 'true' : 'false');
  }
  function toggleImages(e){
    if (e) { e.preventDefault(); e.stopPropagation(); }
    showImages = !showImages;
    localStorage.setItem('showImages', showImages ? '1' : '0');
    setImgBtn();
    render();
  }

  btnDesc?.addEventListener('click', ()=>{ currentSort = (currentSort==='desc') ? null : 'desc'; setActiveSort(); apply(); });
  btnAsc ?.addEventListener('click', ()=>{ currentSort = (currentSort==='asc') ? null : 'asc'; setActiveSort(); apply(); });
  btnNone?.addEventListener('click', ()=>{ currentSort = null; setActiveSort(); apply(); });
  btnImg ?.addEventListener('click', toggleImages);

  const DEBOUNCE = (isMobile || slowNet || cores <= 4) ? 240 : 120;
  function onInputDebounced(el){ el.addEventListener('input', ()=>{ clearTimeout(el._t); el._t=setTimeout(apply,DEBOUNCE); }); }
  [nameQ, codeQ, packQ, rarityQ].forEach(onInputDebounced);

  function closeViewer(){ viewer.classList.remove('show'); viewerImg.src=''; }
  viewerClose?.addEventListener('click', closeViewer);
  viewer?.addEventListener('click', (e)=>{ if(e.target===viewer) closeViewer(); });
  window.addEventListener('keydown', (e)=>{ if(e.key==='Escape') closeViewer(); });

  window.addEventListener('resize', () => shrinkPrices(document));

  setActiveSort(); setImgBtn(); apply();
})();
"""

# ===== HTML =====
def html_page(title: str, js_source: str, logo_uri: str, cards_json: str) -> str:
    shop_svg = "<svg viewBox='0 0 24 24' aria-hidden='true' fill='currentColor'><path d='M3 9.5V8l2.2-3.6c.3-.5.6-.7 1-.7h11.6c.4 0 .7.2.9.6L21 8v1.5c0 1-.8 1.8-1.8 1.8-.9 0-1.6-.6-1.8-1.4-.2.8-.9 1.4-1.8 1.4s-1.6-.6-1.8-1.4c-.2.8-.9 1.4-1.8 1.4s-1.6-.6-1.8-1.4c-.2.8-.9 1.4-1.8 1.4s-1.6-.6-1.8-1.4c-.2.8-.9 1.4-1.8 1.4C3.8 11.3 3 10.5 3 9.5zM5 12.5h14V20c0 .6-.4 1-1 1H6c-.6 0-1-.4-1-1v-7.5zm4 1.5v5h6v-5H9zM6.3 5.2 5 7.5h14l-1.3-2.3H6.3z'/></svg>"
    login_svg= "<svg viewBox='0 0 24 24' aria-hidden='true' fill='currentColor'><path d='M12 12a5 5 0 1 0-5-5 5 5 0 0 0 5 5zm0 2c-4.418 0-8 2.239-8 5v2h16v-2c0-2.761-3.582-5-8-5z'/></svg>"

    parts = []
    parts.append("<!doctype html><html lang='ja'><head><meta charset='utf-8'>")
    parts.append("<meta name='viewport' content='width=device-width,initial-scale=1'>")
    parts.append("<style>"); parts.append(base_css); parts.append("</style></head><body>")
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
    parts.append("  <small class='note'>このページ内のデータのみで検索・並び替え・ページングできます。画像クリックで拡大表示。</small>")
    parts.append("</main>")

    parts.append("<script>")
    parts.append("window.__CARDS__=" + cards_json + ";")
    parts.append("</script>")

    parts.append("<div id='viewer' class='viewer'><div class='vc'><img id='viewerImg' alt=''><button id='viewerClose' class='close'>×</button></div></div>")
    parts.append("<script>")
    parts.append(js_source)
    parts.append("</script></body></html>")
    return "".join(parts)

# ========= 出力 =========
OUT_DIR.mkdir(parents=True, exist_ok=True)

def write_mode(dir_name: str, initial_sort_js_literal: str, title_text: str):
    sub = OUT_DIR / dir_name
    sub.mkdir(parents=True, exist_ok=True)
    js = base_js.replace("__PER_PAGE__", str(PER_PAGE)).replace("__INITIAL_SORT__", initial_sort_js_literal)
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
print(f"[OK] 生成完了 → {OUT_DIR.resolve()} / 総件数{len(df)}")
