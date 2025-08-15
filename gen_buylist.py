# -*- coding: utf-8 -*-
"""
デュエマ買取表 静的ページ生成（豪華版・横4列・完全オフライン・高速版）
- 画像は軽量WebPサムネ（320px）を自家ホスト、クリック時のみフル解像度
- IntersectionObserver + data-src + eagerLoad（初回見えてる分は即読み込み）
- サムネ404時は自動でフル画像へフォールバック
- 1モード=1HTMLのみ（p1〜pN廃止）、データは共通JS（assets/cards.min.js）へ分離
- CI対応：EXCEL_PATH / OUT_DIR / PER_PAGE / BUILD_THUMBS を環境変数で上書き可
"""

from pathlib import Path
from urllib.parse import urlparse, parse_qs
import pandas as pd
import html as html_mod
import unicodedata as ud
import math, base64, mimetypes, os, sys, hashlib, io, json

# ==== 依存（BUILD_THUMBS=1 の場合のみ使う）====
try:
    import requests
    from PIL import Image
except Exception:
    # 未インストールでもスクリプト自体は落とさない
    requests = None
    Image = None

# ========= 設定 =========
DEFAULT_EXCEL = "data/buylist.xlsx"
FALLBACK_WINDOWS = r"C:\Users\user\Desktop\buylist.xlsx"

EXCEL_PATH = os.getenv("EXCEL_PATH", DEFAULT_EXCEL)
SHEET_NAME = os.getenv("SHEET_NAME", "シート1")
OUT_DIR    = Path(os.getenv("OUT_DIR", "docs"))
PER_PAGE   = int(os.getenv("PER_PAGE", "80"))  # 画面内ページング用（データは1HTMLに全件）
BUILD_THUMBS = os.getenv("BUILD_THUMBS", "1") == "1"  # サムネ生成ON/OFF

# 列番号（0始まり）
COL_NAME   = 1
COL_PACK   = 2
COL_CODE   = 3
COL_RARITY = 4
COL_BOOST  = 5
COL_PRICE  = 7
COL_IMGURL = 9

# サムネ設定
THUMB_DIR = OUT_DIR / "assets" / "thumbs"
THUMB_W = 320  # 一覧用はこの幅に縮小（縦は自動）

# ---- ロゴ探索 ----
def find_logo_path():
    cands = [Path(os.getcwd()) / "logo.png", Path(os.getcwd()) / "logo.png.png"]
    try:
        here = Path(__file__).parent
        cands += [here / "logo.png", here / "logo.png.png"]
    except NameError:
        pass
    for p in cands:
        if p.exists() and p.is_file():
            return p
    return None

def logo_to_data_uri(p):
    if not p: return ""
    mime = mimetypes.guess_type(str(p))[0] or "image/png"
    b64  = base64.b64encode(p.read_bytes()).decode("ascii")
    return f"data:{mime};base64,{b64}"

LOGO_URI = logo_to_data_uri(find_logo_path())

# ========= 入力 =========
def load_excel(path_str: str, sheet_name: str|None) -> pd.DataFrame:
    p = Path(path_str)
    if not p.exists():
        p_fb = Path(FALLBACK_WINDOWS)
        if p_fb.exists():
            p = p_fb
        else:
            raise FileNotFoundError(f"Excelが見つかりません: {Path(path_str).resolve()}")
    try:
        if sheet_name:
            return pd.read_excel(p, sheet_name=sheet_name, header=None, engine="openpyxl")
        xls = pd.ExcelFile(p, engine="openpyxl")
        return pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=None, engine="openpyxl")
    except Exception:
        xls = pd.ExcelFile(p, engine="openpyxl")
        return pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=None, engine="openpyxl")

try:
    df_raw = load_excel(EXCEL_PATH, SHEET_NAME)
except FileNotFoundError as e:
    print(str(e)); sys.exit(1)

# ========= ユーティリティ =========
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

def detail_to_img(url: str) -> str:
    if not isinstance(url, str) or not url: return ""
    if "cardimage" in url: return url
    if "id=" in url:
        parsed = urlparse(url); qs = parse_qs(parsed.query)
        id_val = qs.get("id", [parsed.path.split("/")[-1]])[0]
        return f"https://dm.takaratomy.co.jp/wp-content/card/cardimage/{id_val}.jpg"
    return url if url.startswith("http") else ""

def nfkc_lower(s: str) -> str:
    return ud.normalize("NFKC", s or "").lower()

def kata_to_hira(text: str) -> str:
    return "".join(chr(ord(ch) - 0x60) if "ァ" <= ch <= "ン" else ch for ch in text)

def searchable_row(row: pd.Series) -> str:
    return " ".join(kata_to_hira(nfkc_lower(str(row.get(k, ""))))
                    for k in ("name","code","pack","rarity","booster"))

# サムネ生成
def url_to_hash(u:str)->str:
    return hashlib.md5(u.encode("utf-8")).hexdigest()

def ensure_thumb(url: str) -> str|None:
    if not url: return None
    THUMB_DIR.mkdir(parents=True, exist_ok=True)
    fname = url_to_hash(url) + ".webp"
    outp = THUMB_DIR / fname
    if outp.exists():
        return f"assets/thumbs/{fname}"  # 相対はJS側で補正するのでここは基準パスで返す

    if not (requests and Image):
        # 依存が無い場合はスキップしてフルURLを使わせる
        return None
    try:
        r = requests.get(url, timeout=10, headers={"User-Agent":"Mozilla/5.0"})
        r.raise_for_status()
        im = Image.open(io.BytesIO(r.content)).convert("RGB")
        w, h = im.size
        if w <= 0 or h <= 0: return None
        new_h = int(h * THUMB_W / w)
        im = im.resize((THUMB_W, max(1,new_h)), Image.LANCZOS)
        outp.parent.mkdir(parents=True, exist_ok=True)
        im.save(outp, "WEBP", quality=70, method=6)
        return f"assets/thumbs/{fname}"
    except Exception:
        return None

# ========= データ整形 =========
def col(df, i, default=""):
    return df.iloc[:, i] if i < df.shape[1] else pd.Series([default]*len(df))

df = pd.DataFrame({
    "name":    clean_text(col(df_raw, COL_NAME)),
    "pack":    clean_text(col(df_raw, COL_PACK)),
    "code":    clean_text(col(df_raw, COL_CODE)),
    "rarity":  clean_text(col(df_raw, COL_RARITY)),
    "booster": clean_text(col(df_raw, COL_BOOST)),
    "price":   to_int_series(col(df_raw, COL_PRICE) if COL_PRICE < df_raw.shape[1] else pd.Series([None]*len(df_raw))),
    "image":   clean_text(col(df_raw, COL_IMGURL)).map(detail_to_img),
})
df = df[~df["name"].str.match(r"^Unnamed", na=False)]
df = df[df["name"].str.strip()!=""].reset_index(drop=True)
df["_s"] = df.apply(searchable_row, axis=1)

# サムネ列
if BUILD_THUMBS:
    df["thumb"] = df["image"].map(ensure_thumb)
else:
    df["thumb"] = ""

# ========= 見た目 =========
base_css = """
*{box-sizing:border-box}
:root{
  --bg:#ffffff; --panel:#ffffff; --border:#e5e7eb; --accent:#e11d48;
  --text:#111111; --muted:#6b7280; --header-h: 72px;
}
body{ margin:0;color:var(--text);background:var(--bg);font-family:Inter,system-ui,'Noto Sans JP',sans-serif; padding-top: var(--header-h); }
.bg-deco{display:none}
header{
  position:fixed;top:0;left:0;right:0;z-index:1000;background:#fff;border-bottom:1px solid var(--border);
  padding:10px 16px; box-shadow:0 2px 10px rgba(0,0,0,.04);
}
.header-wrap{max-width:1200px;margin:0 auto;display:grid;grid-template-columns:1fr auto 1fr;align-items:center;gap:12px;width:100%}
.brand-left{display:flex;align-items:center;gap:12px;min-width:0;justify-self:start}
.brand-left img{height:60px;display:block}
.brand-fallback{font-weight:1000;letter-spacing:.6px;color:#111;font-size:22px}
.center-ttl{justify-self:center;font-weight:900;white-space:nowrap;font-size:clamp(20px, 3.2vw, 30px); color:#111}
.actions{display:flex;align-items:center;gap:10px;justify-self:end}
.iconbtn{display:inline-flex;align-items:center;gap:8px;border:1px solid var(--border);background:#fff;color:#111;border-radius:12px;padding:9px 12px;text-decoration:none;font-size:13px;transition:transform .12s ease, background .12s ease}
.iconbtn:hover{background:#f9fafb; transform:translateY(-1px)}
.iconbtn svg{width:16px;height:16px;display:block;color:#111}
.wrap{max-width:1200px;margin:0 auto;padding:12px 16px}
.controls{
  display:grid;grid-template-columns:repeat(2, minmax(180px,1fr));
  grid-template-areas: "q1 q2" "q3 q4" "acts acts";
  gap:10px;margin:10px 0 14px;align-items:center;
}
#nameQ{grid-area:q1} #codeQ{grid-area:q2} #packQ{grid-area:q3} #rarityQ{grid-area:q4}
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
.card{background:var(--panel);border:1px solid var(--border);border-radius:14px;overflow:hidden;box-shadow:0 4px 10px rgba(0,0,0,.04);transition:transform .15s ease, box-shadow .15s ease}
.card:hover{transform:translateY(-2px);box-shadow:0 10px 20px rgba(0,0,0,.06)}
.th{aspect-ratio:3/4;background:#f3f4f6;cursor:zoom-in}
.th img{width:100%;height:100%;object-fit:cover;display:block;background:#f3f4f6}
.b{padding:10px 12px}
.n{font-size:14px;font-weight:800;line-height:1.35;margin:0 0 6px;color:#111}
.meta{font-size:11px;color:var(--muted);word-break:break-word}
.p{margin-top:6px;display:flex;flex-wrap:wrap}
.mx{font-weight:1000;color:var(--accent);font-size:clamp(16px, 2.4vw, 22px);line-height:1.05;text-shadow:none;word-break:break-word; overflow-wrap:anywhere;font-variant-numeric:tabular-nums;white-space:nowrap;display:inline-block;max-width:100%}
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
@media (max-width:600px){
  .header-wrap{display:grid;grid-template-columns:auto 1fr;grid-template-areas:"logo title" "actions actions";align-items:center;gap:10px}
  .brand-left{ grid-area:logo; justify-content:flex-start }
  .brand-left img{ height:56px }
  .center-ttl{ grid-area:title; font-size:clamp(20px, 7vw, 28px); line-height:1.1; text-align:left; white-space:nowrap }
  .actions{ grid-area:actions; justify-content:center }
  .grid.grid-img{grid-template-columns:repeat(4, minmax(0,1fr)); gap:8px}
  .b{padding:6px}.n{font-size:11px}
  .mx{ font-size:clamp(12px, 4.2vw, 16px); white-space:nowrap }
  nav.simple{gap:8px; flex-wrap:nowrap; justify-content:space-between}
  nav.simple a{padding:6px 10px; font-size:12px; display:inline-flex}
  nav.simple strong{font-size:12px}
}
small.note{color:var(--muted)}
"""

# ========= JS（共通データ __CARDS__ を読む / 遅延読込＋保険） =========
base_js = r"""
(function(){
  const header = document.querySelector('header');
  const setHeaderH = () => {
    const h = header?.offsetHeight || 88;
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

  // 共通データ読込（assets/cards.min.js が window.__CARDS__ を定義）
  let ALL = Array.isArray(window.__CARDS__) ? window.__CARDS__ : [];

  const SEP_RE = /[\s\u30FB\u00B7·/／\-_—–−]+/g;
  function kataToHira(str){ return (str||'').replace(/[\u30A1-\u30FA]/g, ch => String.fromCharCode(ch.charCodeAt(0)-0x60)); }
  const kanjiReadingMap = { "伝説":"でんせつ" };
  const latinAliasMap = { "complex": "こんぷれっくす", "c0br4": "こぶら" };

  function normalizeForSearch(s){
    s = (s||'').normalize('NFKC').toLowerCase();
    for(const [k,v] of Object.entries(latinAliasMap)){ s = s.split(k).join(v); }
    for(const [k,v] of Object.entries(kanjiReadingMap)){ s = s.split(k).join(v); }
    s = kataToHira(s);
    s = s.replace(SEP_RE, '');
    return s;
  }

  function fmtYen(n){ return (n==null||n==='')?'-':('¥'+parseInt(n,10).toLocaleString()); }
  function escHtml(s){ return (s||'').replace(/[&<>\"']/g, m => ({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"}[m])); }

  let VIEW=[]; let page=1; const PER_PAGE=__PER_PAGE__;
  let currentSort=__INITIAL_SORT__;
  let showImages = (localStorage.getItem('showImages') ?? null) === '1';
  if (localStorage.getItem('showImages') === null) { showImages = false; } // 初回はOFF

  function shrinkPrices(root=document){
    const MIN_PX = 10;
    root.querySelectorAll('.mx').forEach(el=>{
      const style = window.getComputedStyle(el);
      const baseSize = parseFloat(style.fontSize) || 14;
      let size = baseSize;
      const fits = () => el.scrollWidth <= el.clientWidth;
      if (fits()) return;
      while (!fits() && size > MIN_PX) { size -= 1; el.style.fontSize = size + 'px'; }
    });
  }

  // 画像カード（相対サムネは ../ を付与、404時はフル画像へフォールバック）
  function cardHTML_img(it){
    const nameEsc = escHtml(it.name||'');
    const full = it.image?it.image:'';
    let thumb = it.thumb || full;
    if (thumb && !/^https?:\/\//.test(thumb) && !thumb.startsWith('../')) {
      thumb = '../' + thumb;  // docs/<mode>/index.html から見て 1階層上
    }
    return `
  <article class="card">
    <div class="th" data-full="${full}">
      <img
        alt="${nameEsc}"
        loading="lazy" decoding="async"
        width="600" height="800"
        data-src="${thumb}" src=""
        onerror="this.onerror=null;var p=this.closest('.th');this.src=p?p.getAttribute('data-full'):this.src;"
      >
    </div>
    <div class="b">
      <h3 class="n">${nameEsc}</h3>
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

  // IO: 可視域に入ったら読み込む
  let io;
  function setupIO(){
    if (io) io.disconnect();
    const opts = { rootMargin: "200px 0px", threshold: 0.01 };
    io = new IntersectionObserver((entries)=>{
      entries.forEach(e=>{
        if (e.isIntersecting){
          const img = e.target;
          const ds = img.getAttribute('data-src');
          if (ds && !img.src){
            img.src = ds;
            img.removeAttribute('data-src');
          }
          io.unobserve(img);
        }
      });
    }, opts);
    document.querySelectorAll('#grid img[data-src]').forEach(img=>io.observe(img));
    // Above-the-fold の優先度
    document.querySelectorAll('#grid img').forEach((img, i)=>{ img.setAttribute('fetchpriority', i < 8 ? 'high' : 'low'); });
  }

  // 見えてる分は即読み込み（IOの保険）
  function eagerLoad(n=16){
    const imgs=[...document.querySelectorAll('#grid img[data-src]')].slice(0,n);
    imgs.forEach(img=>{
      if(!img.src){
        img.src = img.getAttribute('data-src');
        img.removeAttribute('data-src');
      }
    });
  }

  function render(){
    grid.className = showImages ? 'grid grid-img' : 'grid grid-list';
    const total=VIEW.length; const pages=Math.max(1, Math.ceil(total/PER_PAGE)); if(page>pages) page=pages;
    const start=(page-1)*PER_PAGE; const rows=VIEW.slice(start, start+PER_PAGE);
    grid.innerHTML = rows.map(showImages ? cardHTML_img : cardHTML_list).join('');

    if(showImages){
      grid.querySelectorAll('.th').forEach(th=>{
        th.addEventListener('click', ()=>{
          const src = th.getAttribute('data-full') || th.querySelector('img')?.src || '';
          if(!src) return;
          viewerImg.src = src;
          viewer.classList.add('show');
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
        const j=a.dataset.jump;
        if(j==='prev') page--; else if(j==='next') page++;
        render();
        window.scrollTo({ top: 0, behavior: 'smooth' });
      };
    });

    shrinkPrices(grid);
    if (showImages) {
      setupIO();            // IOでの遅延読込
      eagerLoad(16);        // 最初の16枚は即読み込み
      setTimeout(()=>eagerLoad(32), 600); // 追い保険
    }
  }

  function apply(){
    const nameQv   = normalizeForSearch(nameQ.value||'');
    const codeQv   = normalizeForSearch(codeQ.value||'');
    const packQv   = normalizeForSearch(packQ.value||'');
    const rarityQv = normalizeForSearch(rarityQ.value||'');

    VIEW = ALL.filter(it=>{
      const sName = normalizeForSearch(it.name||'');
      const sCode = normalizeForSearch(it.code||'');
      const sPackBooster = normalizeForSearch([it.pack||'', it.booster||''].join(' '));
      const sRarity = normalizeForSearch(it.rarity||'');

      return (!nameQv   || sName.includes(nameQv))
          && (!codeQv   || sCode.includes(codeQv))
          && (!packQv   || sPackBooster.includes(packQv))
          && (!rarityQv || sRarity.includes(rarityQv));
    });

    if(currentSort==='desc') VIEW.sort((a,b)=>(b.price||0)-(a.price||0));
    else if(currentSort==='asc') VIEW.sort((a,b)=>(a.price||0)-(b.price||0));

    page=1; render();
  }

  function setActiveSort(){
    btnDesc.classList.toggle('active', currentSort==='desc');
    btnAsc .classList.toggle('active', currentSort==='asc');
    btnNone.classList.toggle('active', currentSort===null);
  }
  function setImgBtn(){
    btnImg.textContent = showImages ? '画像OFF' : '画像ON';
    btnImg.classList.toggle('active', showImages);
  }

  btnDesc?.addEventListener('click', ()=>{ currentSort = (currentSort==='desc') ? null : 'desc'; setActiveSort(); apply(); });
  btnAsc ?.addEventListener('click', ()=>{ currentSort = (currentSort==='asc') ? null : 'asc'; setActiveSort(); apply(); });
  btnNone?.addEventListener('click', ()=>{ currentSort = null; setActiveSort(); apply(); });

  btnImg?.addEventListener('click', ()=>{ showImages = !showImages; localStorage.setItem('showImages', showImages ? '1' : '0'); setImgBtn(); render(); });

  function onInputDebounced(el){ el.addEventListener('input', ()=>{ clearTimeout(el._t); el._t=setTimeout(apply,120); }); }
  [nameQ, codeQ, packQ, rarityQ].forEach(onInputDebounced);

  function closeViewer(){ viewer.classList.remove('show'); viewerImg.src=''; }
  viewerClose?.addEventListener('click', closeViewer);
  viewer?.addEventListener('click', (e)=>{ if(e.target===viewer) closeViewer(); });
  window.addEventListener('keydown', (e)=>{ if(e.key==='Escape') closeViewer(); });

  window.addEventListener('resize', () => shrinkPrices(document));

  setActiveSort(); setImgBtn(); apply();
})();
"""

# ===== 共通データを書き出し（assets/cards.min.js） =====
def write_cards_js(df: pd.DataFrame) -> str:
    data = df[["name","pack","code","rarity","booster","price","image","thumb","_s"]].to_dict("records")
    payload = json.dumps(data, ensure_ascii=False, separators=(",",":"))
    ver = hashlib.md5(payload.encode("utf-8")).hexdigest()[:8]
    assets = OUT_DIR / "assets"
    assets.mkdir(parents=True, exist_ok=True)
    (assets / "cards.min.js").write_text("window.__CARDS__="+payload, encoding="utf-8")
    return ver

# ===== HTML（JSONは外部JSから読む） =====
def html_page(title: str, js_source: str, logo_uri: str, cards_ver: str) -> str:
    shop_svg = "<svg viewBox='0 0 24 24' aria-hidden='true' fill='currentColor'><path d='M3 9.5V8l2.2-3.6c.3-.5.6-.7 1-.7h11.6c.4 0 .7.2.9.6L21 8v1.5c0 1-.8 1.8-1.8 1.8-.9 0-1.6-.6-1.8-1.4-.2.8-.9 1.4-1.8 1.4s-1.6-.6-1.8-1.4c-.2.8-.9 1.4-1.8 1.4s-1.6-.6-1.8-1.4c-.2.8-.9 1.4-1.8 1.4s-1.6-.6-1.8-1.4c-.2.8-.9 1.4-1.8 1.4C3.8 11.3 3 10.5 3 9.5zM5 12.5h14V20c0 .6-.4 1-1 1H6c-.6 0-1-.4-1-1v-7.5zm4 1.5v5h6v-5H9zM6.3 5.2 5 7.5h14l-1.3-2.3H6.3z'/></svg>"
    login_svg= "<svg viewBox='0 0 24 24' aria-hidden='true' fill='currentColor'><path d='M12 12a5 5 0 1 0-5-5 5 5 0 0 0 5 5zm0 2c-4.418 0-8 2.239-8 5v2h16v-2c0-2.761-3.582-5-8-5z'/></svg>"

    parts = []
    parts.append("<!doctype html><html lang='ja'><head><meta charset='utf-8'>")
    parts.append("<meta name='viewport' content='width=device-width,initial-scale=1'>")
    parts.append("<link rel='preconnect' href='https://dm.takaratomy.co.jp' crossorigin>")
    parts.append("<link rel='dns-prefetch' href='//dm.takaratomy.co.jp'>")
    parts.append("<style>"); parts.append(base_css); parts.append("</style></head><body>")
    parts.append("<div class='bg-deco' aria-hidden='true'></div>")
    parts.append("<header><div class='header-wrap'>")
    parts.append("<div class='brand-left'>")
    if logo_uri:
        parts.append(f"<!-- LOGO embedded -->\n<img src='{logo_uri}' alt='Shop Logo'>")
    else:
        parts.append("<!-- LOGO missing: fallback -->")
        parts.append("<div class='brand-fallback'>YOUR SHOP</div>")
    parts.append("</div>")
    parts.append(f"<div class='center-ttl'>{html_mod.escape(title)}</div>")
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
    parts.append("    <button id='btnPriceDesc' class='btn'>価格高い順</button>")
    parts.append("    <button id='btnPriceAsc'  class='btn'>価格低い順</button>")
    parts.append("    <button id='btnSortClear' class='btn'>標準順</button>")
    parts.append("    <button id='btnToggleImages' class='btn'>画像ON</button>")
    parts.append("  </div></div>")
    parts.append("  <nav class='simple'></nav><div id='grid' class='grid grid-img'></div><nav class='simple'></nav>")
    parts.append("  <small class='note'>このページ内のデータのみで検索・並び替え・ページングできます。画像クリックで拡大表示。</small>")
    parts.append("</main>")
    # 共通データ（各モード配下HTML → ../assets/）
    parts.append(f"<script src='../assets/cards.min.js?v={cards_ver}'></script>")
    parts.append("<div id='viewer' class='viewer'><div class='vc'><img id='viewerImg' alt=''><button id='viewerClose' class='close'>×</button></div></div>")
    parts.append("<script>"); parts.append(base_js); parts.append("</script></body></html>")
    return "".join(parts)

# ========= 出力 =========
OUT_DIR.mkdir(parents=True, exist_ok=True)

# データを共通JSに一回だけ書き出し（バージョン付与でキャッシュ更新）
CARDS_VER = write_cards_js(df)

def write_mode(dir_name: str, initial_sort_js_literal: str, title_text: str, cards_ver: str):
    sub = OUT_DIR / dir_name
    sub.mkdir(parents=True, exist_ok=True)
    js = base_js.replace("__PER_PAGE__", str(PER_PAGE)).replace("__INITIAL_SORT__", initial_sort_js_literal)
    html = html_page(title_text, js, LOGO_URI, cards_ver)
    (sub / "index.html").write_text(html, encoding="utf-8")

write_mode("default", "null", "デュエマ買取表", CARDS_VER)
write_mode("price_desc", "'desc'", "デュエマ買取表（price_desc）", CARDS_VER)
write_mode("price_asc", "'asc'", "デュエマ買取表（price_asc）", CARDS_VER)

# ルートは default/ にリダイレクト
(OUT_DIR / "index.html").write_text(
    "<meta http-equiv='refresh' content='0; url=default/'>",
    encoding="utf-8"
)

print(f"[LOGO] {'embedded' if LOGO_URI else 'not found (fallback text used)'}")
print(f"[OK] 生成完了 → {OUT_DIR.resolve()} / 総件数{len(df)}")
