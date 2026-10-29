import base64
import difflib
import hashlib
import hmac
import html
import inspect
import io
import json
import os
import re
import time
import unicodedata
from datetime import date, datetime

import streamlit as st
from anthropic import Anthropic

import chat_flow
import consistency
import costs
import db
import docx_edit
import export_tools as ex
import diffreport
import feor
import finder
import navtax
import njt
import records
import redflag
import tplfill
import registry
import search_tools as stx

# Frissítés után a futó alkalmazás néha a saját moduljaink régi, memóriában maradt változatát használja
# (ilyenkor pl. „db.get_tasks” hiányzik). Ha ezt észleljük, újratöltjük őket, így nem kell kézzel újraindítani.
if (not hasattr(db, "get_tasks") or not hasattr(db, "delete_summary") or not hasattr(db, "set_document_summary")
        or "handlers" not in inspect.signature(chat_flow.run_conversation).parameters
        or not hasattr(chat_flow, "HANDLERS_MAY_RETURN_FILES") or not hasattr(db, "log_usage")
        or not hasattr(db, "get_documents_meta") or not hasattr(db, "get_textless_documents")
        or "usage_cb" not in inspect.signature(chat_flow.run_conversation).parameters):
    import importlib
    for _module in (stx, ex, db, docx_edit, costs, feor, finder, consistency, njt, navtax, redflag, diffreport, registry, records, tplfill, chat_flow):
        importlib.reload(_module)

# Kényszerített oldal konfiguráció a legelső sorban
st.set_page_config(page_title="Sara Asszisztens", layout="wide")

# --- BEÁLLÍTÁSOK -----------------------------------------------------------
SARA_MODEL_ID = "claude-sonnet-5-5"  # 2026.10.12-től: a Haiku 5.5 helyett (okosabb, a szövegegység ára kb. 20-szoros)
JEAN_MODEL_ID = "claude-haiku-5-5"   # JEAN, Sara olcsó helyettese (egyszerű kérdés, adatgyűjtés, rutinmunka)
# Ki dolgozik most: SARA vagy JEAN (az oldalsáv választója állítja). Felülírható a titkos beállításokban:
# SARA_MODEL, JEAN_MODEL, és hogy belépéskor ki dolgozik: ASSZISZTENS_ALAP = "SARA" vagy "JEAN".
DEFAULT_ASSISTANT = "SARA"

# A beszélgetés és a többi fül görgethető dobozának magassága. A doboz a böngészőablak magasságához
# igazodik: ablakmagasság mínusz az alábbi képpontérték (fejléc, fülek, beviteli mező helye).
# Ha a doboz alatt üres sáv marad, csökkentsd a számot; ha az oldal görgethetővé válik, növeld.
CHAT_OFFSET_PX = 310    # Beszélgetés fül
PANEL_OFFSET_PX = 205   # Dokumentumok és Összefoglalók fül
SEARCH_OFFSET_PX = 155  # keresési találatok (itt nincsenek fülek)
FALLBACK_HEIGHT_PX = 520  # régebbi Streamlit esetén ez a fix magasság érvényes

# a projektdokumentumok kategóriái (Dokumentumok fül, Sara besorolása)
DOC_CATEGORIES = ["Szerződés", "Ajánlat", "Műszaki leírás", "Terv, rajz", "Engedély, hatóság",
                  "Jegyzőkönyv, emlékeztető", "Levelezés", "Pénzügy, számla", "Jogi, cég", "Egyéb"]
UNCATEGORIZED = "Kategória nélkül"

MAX_FILE_MB = 10   # a tartalom Base64-ként az adatbázisban tárolódik, ezért korlátozzuk a méretet
MAX_IMAGE_MB = 5   # ennél nagyobb képet a modell nem fogad el
DOC_TYPES = ["txt", "pdf", "png", "jpg", "jpeg", "docx", "xlsx", "pptx"]  # a dokumentumtárba feltölthető
ATTACH_TYPES = ["png", "jpg", "jpeg", "gif", "webp", "pdf", "txt", "md", "csv", "docx", "xlsx", "pptx"]
EXT_MIME = {
    "png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg", "gif": "image/gif", "webp": "image/webp",
    "pdf": ex.PDF_MIME, "txt": "text/plain", "md": ex.MD_MIME, "csv": "text/csv",
    "docx": ex.DOCX_MIME, "xlsx": ex.XLSX_MIME, "pptx": ex.PPTX_MIME,
}

# A gombok teljes szélességét az újabb Streamlit width="stretch"-ként, a régebbi use_container_width-ként kéri
STRETCH = ({"width": "stretch"} if "width" in inspect.signature(st.button).parameters
           else {"use_container_width": True})

# A legördülő lista újabb Streamlitben saját, új értéket is elfogad (témaválasztáshoz)
ACCEPT_NEW_OPTIONS = "accept_new_options" in inspect.signature(st.selectbox).parameters

NAV_CFG = {}

# --- KULCSOK BIZTONSÁGOS BETÖLTÉSE ---
try:
    SUPABASE_URL = str(st.secrets["SUPABASE_URL"]).strip().strip("'").strip('"')
    SUPABASE_KEY = str(st.secrets["SUPABASE_KEY"]).strip().strip("'").strip('"')
    ANTHROPIC_API_KEY = str(st.secrets["ANTHROPIC_API_KEY"]).strip().strip("'").strip('"')
    SARA_MODEL_ID = str(st.secrets.get("SARA_MODEL", SARA_MODEL_ID)).strip().strip("'").strip('"') or SARA_MODEL_ID
    JEAN_MODEL_ID = str(st.secrets.get("JEAN_MODEL", JEAN_MODEL_ID)).strip().strip("'").strip('"') or JEAN_MODEL_ID
    DEFAULT_ASSISTANT = str(st.secrets.get("ASSZISZTENS_ALAP", DEFAULT_ASSISTANT)).strip().strip("'").strip('"').upper()
    # a cég NAV Online Számla technikai felhasználója (adószám-ellenőrzéshez); ha hiányzik, az eszköz ezt jelzi
    NAV_CFG = {k: str(st.secrets.get(f"NAV_{k.upper()}", "")).strip() for k in ("login", "password", "signkey",
                                                                                 "taxnumber")}
except Exception:
    st.error("Hiba! Hiányzik a Secrets konfiguráció a Streamlit felületén.")
    st.stop()

claude_client = Anthropic(api_key=ANTHROPIC_API_KEY.replace("'", "").replace('"', ""))
db.init(SUPABASE_URL, SUPABASE_KEY)

# --- MEGJELENÉS ------------------------------------------------------------
# Contact-Line 2002 Bt. arculata (arculati kézikönyv, 2026. október): címsor Montserrat, szöveg Open Sans;
# színek: #0A2540 (Deep Corporate Blue: szöveg, elsődleges), #00D4FF (Azure Light Blue: akcentus, kiemelés),
# #F4F6F9 (Soft Background Gray: másodlagos felületek). Az azúr szövegszínnek fehér alapon túl világos,
# ezért csak vonalakhoz, jelölőkhöz és háttérárnyalatként használjuk.
BRAND = "#0A2540"
BRAND_DARK = "#0A2540"
BRAND_LIGHT = "#F4F6F9"
BRAND_ACCENT = "#00D4FF"
BRAND_ACCENT_SOFT = "#B8F1FF"   # az azúr világos árnyalata (kiemelés háttere)

CSS = f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@500;600;700&family=Open+Sans:ital,wght@0,400;0,500;0,600;0,700;1,400&display=swap');

/* betűtípusok (az ikonok saját betűkészlete érintetlen marad, mert azt az ikon eleme maga állítja be) */
html, body, .stApp, .stApp input, .stApp textarea, .stApp button, .stApp select,
[data-testid="stMarkdownContainer"], [data-testid="stMarkdownContainer"] p, [data-testid="stMarkdownContainer"] li,
[data-testid="stWidgetLabel"] p {{ font-family: 'Open Sans', Roboto, Arial, sans-serif !important; }}
.sara-header, .sara-side-title.big,
[data-testid="stMarkdownContainer"] h1, [data-testid="stMarkdownContainer"] h2, [data-testid="stMarkdownContainer"] h3 {{
    font-family: 'Montserrat', Inter, Helvetica, Arial, sans-serif !important; font-weight: 700; color: {BRAND_DARK}; }}

/* logó az oldalsáv tetején, körülötte szabadon hagyott térrel */
/* (A Streamlit a szövegelemek dobozát alul 1rem-mel felhúzza, mert bekezdést vár benne; a saját, bekezdés
   nélküli elemeinknél ezt alsó margóval egyenlítjük ki, különben rájuk csúszik a következő elem.) */
.sara-logo {{ text-align: center; padding: .1rem 0 .3rem; margin-bottom: 1rem; }}
.sara-logo img {{ width: 13rem; max-width: 92%; height: auto; }}

/* az oldalsáv saját címei (a beépített címsorokra rácsúszott az alattuk lévő elem) */
.sara-side-title {{ font-weight: 600; font-size: 1rem; line-height: 1.4; margin: .1rem 0 1rem; color: {BRAND_DARK}; }}
.sara-side-title.big {{ font-size: 1.2rem; font-weight: 500; }}

/* a Streamlit Cloud „Manage app” gombja (csak az alkalmazás tulajdonosa látja) a jobb alsó sarokban ül:
   a beviteli mező jobb szélét, a küldés gombbal együtt, ne takarja el */
[data-testid="stBottomBlockContainer"] {{ padding-right: 8rem !important; }}
[data-testid="stChatInput"] {{ width: calc(100% - 8rem) !important; max-width: calc(100% - 8rem) !important; }}

/* feladatok: a lezárt (kész, elavult, törölt) tételek halványan, áthúzva */
.sara-task-done {{ text-decoration: line-through; opacity: .5; }}

/* kevesebb üres hely az oldal tetején és alján, hogy a beszélgetés kitöltse a képernyőt */
.block-container, [data-testid="stMainBlockContainer"] {{ padding-top: 3.75rem !important; padding-bottom: 1rem !important; }}

/* TÖMÖREBB FELÜLET (az alap betűméretet a .streamlit/config.toml állítja: theme.baseFontSize) */
/* kisebb térköz az egymás alatti elemek között, alacsonyabb gombok, vékonyabb elválasztók */
[data-testid="stVerticalBlock"] {{ gap: .6rem; }}
.stButton button, .stDownloadButton button, .stFormSubmitButton button, [data-testid="stPopover"] button {{
    min-height: 2rem !important; padding: .15rem .75rem !important; }}
hr {{ margin: .5rem 0 !important; }}
[data-testid="stExpander"] summary {{ padding-top: .35rem !important; padding-bottom: .35rem !important; min-height: 0 !important; }}
/* oldalsáv: kevesebb üres hely felül, kisebb címek */
[data-testid="stSidebarHeader"] {{ height: auto !important; min-height: 2rem; padding-top: .5rem !important; padding-bottom: 0 !important; }}
[data-testid="stSidebarUserContent"] {{ padding-top: .5rem !important; }}
[data-testid="stSidebar"] h2 {{ font-size: 1.15rem !important; padding: .2rem 0 .3rem !important; margin: 0 !important; line-height: 1.4 !important; }}
[data-testid="stSidebar"] h3 {{ font-size: 1rem !important; padding: .2rem 0 .3rem !important; margin: 0 !important; line-height: 1.4 !important; }}
[data-testid="stSidebar"] [data-testid="stHeading"] {{ margin: 0 !important; overflow: visible !important; }}
/* az összefoglalók apró szerkesztés- és törlésgombja */
[class*="st-key-sum_editbtn_"] button, [class*="st-key-sum_delbtn_"] button,
[class*="st-key-task_editbtn_"] button, [class*="st-key-task_delbtn_"] button {{
    min-height: 1.7rem !important; padding: 0 .45rem !important; }}

/* keskeny, mindig látható fejlécsor */
.sara-header {{ font-size: 1.25rem; font-weight: 500; line-height: 2.5rem; white-space: nowrap;
               overflow: hidden; text-overflow: ellipsis; }}
.sara-header .sep {{ margin: 0 .6rem; font-weight: 400; color: {BRAND_ACCENT}; }}
/* a fejlécsor alatt vékony azúr vonal, mint a levélpapíron */
.st-key-header_row {{ border-bottom: 2px solid {BRAND_ACCENT}; padding-bottom: .15rem; }}
.sara-header .badge {{ margin-left: .5rem; font-size: .85rem; font-weight: 400; opacity: .7; }}

/* kisebb címsorok Sara válaszaiban és az előnézetekben (a törzsszöveg változatlan) */
[data-testid="stMarkdownContainer"] h1 {{ font-size: 1.3rem !important; font-weight: 700 !important; padding: .55rem 0 .2rem !important; }}
[data-testid="stMarkdownContainer"] h2 {{ font-size: 1.2rem !important; font-weight: 700 !important; padding: .5rem 0 .2rem !important; }}
[data-testid="stMarkdownContainer"] h3 {{ font-size: 1.1rem !important; font-weight: 700 !important; padding: .45rem 0 .15rem !important; }}
[data-testid="stMarkdownContainer"] h4,
[data-testid="stMarkdownContainer"] h5,
[data-testid="stMarkdownContainer"] h6 {{ font-size: 1rem !important; font-weight: 700 !important; padding: .4rem 0 .1rem !important; }}

/* a görgethető dobozok a böngészőablak magasságához igazodnak */
/* (A :has() szabályok külön állnak: amelyik böngésző nem ismeri, az csak azt a blokkot hagyja ki.
   A dvh a táblagépek és telefonok ténylegesen látható ablakmagassága; ahol nincs, a vh érvényes.) */
.st-key-chat_box {{ height: calc(100vh - {CHAT_OFFSET_PX}px) !important; height: calc(100dvh - {CHAT_OFFSET_PX}px) !important; min-height: 260px; }}
.st-key-docs_box, .st-key-sums_box, .st-key-tasks_box {{ height: calc(100vh - {PANEL_OFFSET_PX}px) !important; height: calc(100dvh - {PANEL_OFFSET_PX}px) !important; min-height: 260px; }}
.st-key-search_box {{ height: calc(100vh - {SEARCH_OFFSET_PX}px) !important; height: calc(100dvh - {SEARCH_OFFSET_PX}px) !important; min-height: 260px; }}
div[data-testid="stVerticalBlockBorderWrapper"]:has(> div > .st-key-chat_box) {{
    height: calc(100vh - {CHAT_OFFSET_PX}px) !important; height: calc(100dvh - {CHAT_OFFSET_PX}px) !important; min-height: 260px; }}
div[data-testid="stVerticalBlockBorderWrapper"]:has(> div > .st-key-docs_box),
div[data-testid="stVerticalBlockBorderWrapper"]:has(> div > .st-key-sums_box),
div[data-testid="stVerticalBlockBorderWrapper"]:has(> div > .st-key-tasks_box) {{
    height: calc(100vh - {PANEL_OFFSET_PX}px) !important; height: calc(100dvh - {PANEL_OFFSET_PX}px) !important; min-height: 260px; }}
div[data-testid="stVerticalBlockBorderWrapper"]:has(> div > .st-key-search_box) {{
    height: calc(100vh - {SEARCH_OFFSET_PX}px) !important; height: calc(100dvh - {SEARCH_OFFSET_PX}px) !important; min-height: 260px; }}

/* A fix magasságú dobozban a tartalom görgessen, ne zsugorodjon: enélkül a megadott magasságú elemek
   (pl. a szerkesztőmezők) összenyomódnak, amikor a tartalom nem fér ki. */
.st-key-chat_box > *, .st-key-docs_box > *, .st-key-sums_box > *, .st-key-search_box > *, .st-key-tasks_box > *,
.st-key-tasks_box [data-testid="stElementContainer"],
.st-key-chat_box [data-testid="stElementContainer"], .st-key-docs_box [data-testid="stElementContainer"],
.st-key-sums_box [data-testid="stElementContainer"], .st-key-search_box [data-testid="stElementContainer"] {{
    flex-shrink: 0 !important; }}

/* keresési találatok */
.sara-hit {{ margin: .1rem 0 .2rem 0; }}
.sara-hit .title {{ font-weight: 600; }}
.sara-hit .meta {{ opacity: .6; font-size: .85rem; margin-left: .4rem; }}
.sara-hit .snip {{ font-size: .92rem; opacity: .9; margin-top: .15rem; }}
.sara-hit mark, .sara-header mark {{ background: {BRAND_ACCENT_SOFT}; color: inherit; padding: 0 .12em; border-radius: 3px; }}
</style>
"""


LOGO_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "contact_line_logo.svg")


@st.cache_data(show_spinner=False)
def logo_data_uri(color):
    """A Contact-Line 2002 logó (a mélykék elemek a megadott színben), képként beágyazható formában (üres, ha a fájl hiányzik)."""
    try:
        with open(LOGO_FILE, encoding="utf-8") as fh:
            svg = fh.read()
    except OSError:
        return ""
    svg = svg.replace("#0A2540", color)
    return "data:image/svg+xml;base64," + base64.b64encode(svg.encode("utf-8")).decode("ascii")


if st.session_state.get("assistant") not in ("SARA", "JEAN"):
    st.session_state["assistant"] = DEFAULT_ASSISTANT if DEFAULT_ASSISTANT in ("SARA", "JEAN") else "SARA"
ASSISTANT = st.session_state["assistant"]
MODEL = JEAN_MODEL_ID if ASSISTANT == "JEAN" else SARA_MODEL_ID
ASSISTANT_NAME = "Jean" if ASSISTANT == "JEAN" else "Sara"
ASSISTANT_FROM = "Jeantól" if ASSISTANT == "JEAN" else "Sarától"
ASSISTANT_ACC = "Jeant" if ASSISTANT == "JEAN" else "Sarát"

# JEAN napján a felület az arculat azúr árnyalatát kapja, hogy első ránézésre látszódjon, ki dolgozik
JEAN_CSS = """
<style>
.stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"], [data-testid="stBottom"] > div
  { background-color: #F0FCFF !important; }
[data-testid="stSidebar"] > div:first-child { background-color: #DDF8FF !important; }
.jean-flag { display: inline-block; margin-left: .7rem; padding: .05rem .55rem; border-radius: 1rem;
  font-size: .8rem; font-weight: 600; letter-spacing: .05em; background: #00D4FF; color: #0A2540; vertical-align: middle; }
</style>
"""


def _avatar(name):
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", name)
    return path if os.path.exists(path) else None


# A beszélgetés szimbólumai az arculat színeiben: a felhasználó szürke, Sara (S) mélykék, Jean (J) azúr
AVATARS = {"user": _avatar("avatar_user.svg"),
           "assistant": _avatar("avatar_jean.svg" if ASSISTANT == "JEAN" else "avatar_sara.svg")}


def chat_message(role):
    return st.chat_message(role, avatar=AVATARS.get(role))


def side_title(text, big=False):
    st.markdown(f'<div class="sara-side-title{" big" if big else ""}">{html.escape(text)}</div>',
                unsafe_allow_html=True)


def topic_input(container, label, topics, value=None, key=None):
    """Téma megadása: gépelés közben a meglévő témák közül ajánl, de új is beírható.
    (Régebbi Streamlitnél sima szövegmező.)"""
    if topics and ACCEPT_NEW_OPTIONS:
        options = list(dict.fromkeys(topics + ([value] if value else [])))
        return container.selectbox(label, options=options, index=options.index(value) if value else None,
                                   placeholder="válassz, vagy írj be újat", accept_new_options=True, key=key)
    if value is None:
        return container.text_input(label, placeholder="pl. Vízhálózat", key=key)
    return container.text_input(label, value=value, key=key)


def inject_css():
    # st.html a csak stílust tartalmazó blokkot hely foglalása nélkül illeszti be
    css = CSS + (JEAN_CSS if ASSISTANT == "JEAN" else "")
    if hasattr(st, "html"):
        st.html(css)
    else:
        st.markdown(css, unsafe_allow_html=True)


# --- GYORSÍTÓTÁRAZOTT BETÖLTÉSEK (a tárolt fájlok nem módosulnak) ---
@st.cache_data(show_spinner=False, max_entries=40)
def get_document_content(doc_id):
    return db.get_document_content(doc_id)


@st.cache_data(show_spinner=False, max_entries=60)
def get_attachment_content(attachment_id):
    return db.get_attachment_content(attachment_id)


@st.cache_data(show_spinner=False, max_entries=20)
def export_markdown(doc_id, fmt, title, paper=""):
    """A Sara által készített Markdown dokumentum átalakítása Word vagy PDF formátumra.
    paper: "" = a Contact-Line 2002 beépített arculata, "ures" = fejléc nélkül, különben a fejléces papír dokumentum-azonosítója
    (egy cég-törzsaktába feltöltött .docx)."""
    text = base64.b64decode(get_document_content(doc_id)).decode("utf-8", errors="replace")
    if paper and paper != "ures":
        docx = ex.markdown_to_docx(text, title, base=base64.b64decode(get_document_content(paper)))
        return docx if fmt == "docx" else docx_to_pdf(docx)
    brand = ex.EMPTY_BRAND if paper == "ures" else None
    return ex.markdown_to_docx(text, title, brand) if fmt == "docx" else ex.markdown_to_pdf(text, title, brand)


def docx_to_pdf(docx):
    """Word → PDF LibreOffice-szal (ha a szerveren van); különben None."""
    import shutil
    import subprocess
    import tempfile
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if not exe:
        return None
    with tempfile.TemporaryDirectory() as tmp:
        src = os.path.join(tmp, "irat.docx")
        with open(src, "wb") as fh:
            fh.write(docx)
        try:
            subprocess.run([exe, "--headless", "--convert-to", "pdf", "--outdir", tmp, src], capture_output=True,
                           timeout=120, check=True)
            with open(os.path.join(tmp, "irat.pdf"), "rb") as fh:
                return fh.read()
        except Exception:
            return None


# --- FEJLÉCES PAPÍROK ---
# A cég-törzsaktába feltöltött Word-sablon (fejléc, lábléc, margó, betű) a cég fejléces papírja; az adatlap fülön
# jelölhető ki. A projektnél megadható, melyik cégé (különben a kapcsolt cég papírja, ha van, egyébként a Contact-Line 2002 beépített arculata);
# letöltéskor dokumentumonként átállítható; iratmintánál a meghatalmazó / szerződő cég papírja az alapértelmezés.
def letterheads():
    """[(cég-törzsakta, papír dokumentum-azonosító)] – csak a kijelölt fejléces papírral rendelkező cégek."""
    if not db.has_records():
        return []
    out = []
    for p in globals().get("all_projects") or []:
        if is_record(p) and p.get("record_type") == "ceg" and p.get("status") != "closed":
            paper = str(record_extra(p["id"]).get("fejlec") or "")
            if paper:
                out.append((p, paper))
    return out


@st.cache_data(ttl=300, show_spinner=False)
def project_paper_setting(project_id):
    if not project_id or not db.has_records():
        return ""
    try:
        value = db.get_setting(f"fejlec:{project_id}", "") or ""
    except db.DbError:
        return ""
    return value if isinstance(value, str) else ""


def project_paper(project_id):
    """A projekt papírja: "" (Contact-Line 2002, beépített), "ures", vagy egy fejléces papír azonosítója."""
    papers = letterheads()
    chosen = project_paper_setting(project_id)
    if chosen in ("nf", "ures"):
        return "" if chosen == "nf" else "ures"
    by_rec = {r["id"]: paper for r, paper in papers}
    if chosen in by_rec:
        return by_rec[chosen]
    project = (globals().get("projects_by_id") or {}).get(project_id) or {}
    for rid in project.get("links") or []:
        if rid in by_rec:
            return by_rec[rid]
    return ""


def paper_options():
    """A választható papírok: (érték, felirat)."""
    return ([("", "Contact-Line 2002 (beépített)")] + [(paper, f"{r.get('name')} fejléces papírja") for r, paper in letterheads()]
            + [("ures", "Üres (fejléc nélkül)")])


# --- ÁLTALÁNOS SEGÉDFÜGGVÉNYEK ----------------------------------------------
def parse_ts(value):
    try:
        return datetime.fromisoformat(str(value))
    except Exception:
        return None


def fmt_dt(value):
    """Adatbázis-időbélyeg magyar idő szerint, pl. 2026.10.10 18:40."""
    ts = parse_ts(value)
    if ts is None:
        return str(value or "")[:16].replace("T", " ")
    try:
        from zoneinfo import ZoneInfo
        ts = ts.astimezone(ZoneInfo("Europe/Budapest"))
    except Exception:
        pass
    return ts.strftime("%Y.%m.%d %H:%M")


def guess_mime(name, given=""):
    ext = str(name).rsplit(".", 1)[-1].lower() if "." in str(name) else ""
    return EXT_MIME.get(ext) or given or "application/octet-stream"


def is_markdown_doc(doc):
    return doc["mime_type"] == ex.MD_MIME or doc["name"].lower().endswith(".md")


def doc_icon(doc):
    mime = doc["mime_type"]
    return (":material/image:" if mime.startswith("image/") else ":material/article:" if is_markdown_doc(doc) else ":material/table_chart:" if mime == ex.XLSX_MIME
            else ":material/slideshow:" if mime == ex.PPTX_MIME else ":material/description:")


SUMMARY_MODEL = "claude-haiku-5-5"   # a kivonat gépies feladat: a legolcsóbb modell is jól csinálja
DOC_SUMMARY_SYSTEM = (
    "Egy projekt dokumentumtárába feltöltött dokumentumról írsz rövid tartalmi kivonatot, hogy egy munkatárs "
    "gyorsan eldönthesse, mi van benne és kell-e neki. Magyarul, 5–10 mondatban, folyó szövegben írj. "
    "Írd le: mi ez (fajtája), ki készítette és kinek, mikor; a lényeges számokat, mennyiségeket, összegeket, "
    "határidőket és feltételeket; és mire jó. Csak azt írd, ami a dokumentumban áll, ne találgass. "
    "Ne írj címet, bevezetést, felsorolást vagy záró mondatot.")


def make_doc_summary(name, mime_type, data, text, project_id=None):
    """5–10 mondatos kivonat egy dokumentumról. ValueError-t ad érthető okkal, ha nem készíthető.
    Több úton próbálkozik, mert a táblázatos PDF-ek (pl. Excelből nyomtatott beszámoló) kinyert szövege
    összekevert, és erre a kis modell néha üres választ ad:
    1. a kinyert szöveg a kis modellnek; 2. PDF-nél maga a fájl a kis modellnek (látja az elrendezést);
    3. ugyanez Sara fő modelljének."""
    b64 = lambda: base64.b64encode(data).decode("ascii")  # noqa: E731
    is_pdf = mime_type == "application/pdf" and len(data) < 20 * 1024 * 1024
    attempts = []
    if text and text.strip():
        as_text = [{"type": "text", "text": f"A dokumentum neve: {name}\n\nA dokumentum "
                                            f"szövege (a hosszú dokumentumnak csak az eleje):\n\n{text[:40000]}"}]
        attempts.append((SUMMARY_MODEL, as_text))
        if not mime_type == "application/pdf":
            attempts.append((MODEL, as_text))   # ha a kis modell üres választ ad, a fő modell próbálja
    if is_pdf:
        pdf = [{"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": b64()}},
               {"type": "text", "text": f"A dokumentum neve: {name}"}]
        attempts += [(SUMMARY_MODEL, pdf), (MODEL, pdf)]
    elif mime_type in ("image/png", "image/jpeg") and len(data) < 4 * 1024 * 1024:
        attempts.append((SUMMARY_MODEL, [{"type": "image", "source": {"type": "base64", "media_type": mime_type,
                                                                      "data": b64()}},
                                         {"type": "text", "text": f"A kép neve: {name}"}]))
    if not attempts:
        if mime_type in ("image/png", "image/jpeg"):
            raise ValueError("a kép nagyobb 4 MB-nál, ezt a modell nem fogadja")
        if mime_type == "application/pdf":
            raise ValueError("a PDF nagyobb 20 MB-nál, és nincs kiolvasható szövege")
        raise ValueError("a fájlból nem olvasható ki szöveg")
    reasons = []
    for model, content in attempts:
        try:
            response = _summary_request(content, model=model, project_id=project_id)
        except ValueError as e:
            reasons.append(str(e))
            continue
        summary = "".join(getattr(b, "text", "") for b in response.content
                          if getattr(b, "type", "") == "text").strip()
        if summary:
            return summary
        reasons.append(f"üres válasz ({getattr(response, 'stop_reason', '?')})")
    raise ValueError("a modell egyik próbálkozásra sem adott kivonatot: " + "; ".join(dict.fromkeys(reasons)))


def usage_logger(project_id, kind):
    """A modellhívások tokenjeit a usage_log táblába írja (a gyorsítótár hasznának későbbi elemzéséhez)."""
    return lambda model, usage: db.log_usage(project_id, kind, model, usage)


def _summary_request(content, waits=(5, 15, 30), model=SUMMARY_MODEL, project_id=None):
    """Kérés a modellnek; túlterhelés vagy hálózati hiba esetén vár és újrapróbálja (sok dokumentum egymás
    utáni feldolgozásakor ez gyakori, ezért ment eddig csak a második nekifutásra)."""
    import anthropic
    transient = tuple(t for t in (getattr(anthropic, n, None) for n in (
        "RateLimitError", "APIConnectionError", "APITimeoutError", "InternalServerError")) if isinstance(t, type))
    for wait in (*waits, None):
        try:
            # Gondolkodás nélkül: az új modellek alapból gondolkodhatnak, és az elvitte a teljes keretet,
            # így a válasz üres maradt („üres válasz (max_tokens)”). A keret is bővebb, ha mégis gondolkodna.
            kwargs = dict(model=model, max_tokens=4000, system=DOC_SUMMARY_SYSTEM,
                          messages=[{"role": "user", "content": content}])
            try:
                response = claude_client.messages.create(thinking={"type": "disabled"}, **kwargs)
            except Exception as e:
                if getattr(e, "status_code", None) != 400 or "thinking" not in str(e).lower():
                    raise
                response = claude_client.messages.create(**kwargs)
            db.log_usage(project_id, "kivonat", model, getattr(response, "usage", None))
            return response
        except transient as e:
            if wait is None:
                raise ValueError("a modell többszöri próbálkozásra is túlterhelt vagy elérhetetlen volt; "
                                 "próbáld újra később") from e
            time.sleep(wait)
        except Exception as e:
            status = getattr(e, "status_code", None)
            if status == 400:
                raise ValueError("a modell nem tudta feldolgozni a fájlt (túl nagy vagy túl sok oldalas)") from e
            raise ValueError(f"hiba a modell hívásakor: {e}") from e


# --- Szövegfelismerés: a szkennelt PDF-ek és képek szövegét a Haiku egyszer átírja, utána kereshetők -----------
OCR_SYSTEM = (
    "Egy szkennelt vagy fényképezett dokumentum oldalait kapod. Írd át a teljes szövegét szó szerint, magyarul úgy, "
    "ahogy le van írva: ne foglalj össze, ne javíts, ne hagyj ki semmit. A táblázatokat soronként írd, a cellákat "
    "„ | ” jellel elválasztva. A pecsétet, aláírást, kézírást szögletes zárójelben jelöld, pl. [aláírás], "
    "[bélyegző: Realion2001 Zrt.], [kézírással: 2026.10.14.]. Ami olvashatatlan: [olvashatatlan]. Csak az átírt "
    "szöveget add vissza, bevezetés és megjegyzés nélkül.")
OCR_PAGES_PER_CALL = 10


def needs_ocr(mime_type, data, text):
    """Szkennelt-e: PDF-nél oldalanként átlag 100 karakternél kevesebb kiolvasható szöveg; képnél mindig."""
    if mime_type in ("image/png", "image/jpeg"):
        return True
    if mime_type != ex.PDF_MIME:
        return False
    try:
        from pypdf import PdfReader
        pages = len(PdfReader(io.BytesIO(data)).pages)
    except Exception:
        return False
    return len((text or "").strip()) < 100 * max(pages, 1)


def _ocr_call(content, project_id):
    kwargs = dict(model=SUMMARY_MODEL, max_tokens=16000, system=OCR_SYSTEM,
                  messages=[{"role": "user", "content": content}])
    try:
        with claude_client.messages.stream(thinking={"type": "disabled"}, **kwargs) as stream:
            response = stream.get_final_message()
    except Exception as e:
        if getattr(e, "status_code", None) != 400 or "thinking" not in str(e).lower():
            raise
        with claude_client.messages.stream(**kwargs) as stream:
            response = stream.get_final_message()
    db.log_usage(project_id, "ocr", SUMMARY_MODEL, getattr(response, "usage", None))
    return "".join(getattr(b, "text", "") for b in response.content if getattr(b, "type", "") == "text").strip()


def ocr_document(name, mime_type, data, project_id=None):
    """A szkennelt PDF vagy kép szövege a Haikuval felismerve; a PDF-et tízoldalas részekben, párhuzamosan."""
    from concurrent.futures import ThreadPoolExecutor
    if mime_type in ("image/png", "image/jpeg"):
        if len(data) > 4 * 1024 * 1024:
            raise ValueError("a kép nagyobb 4 MB-nál")
        return _ocr_call([{"type": "image", "source": {"type": "base64", "media_type": mime_type,
                                                       "data": base64.b64encode(data).decode("ascii")}},
                          {"type": "text", "text": f"A kép neve: {name}"}], project_id)
    from pypdf import PdfReader, PdfWriter
    reader = PdfReader(io.BytesIO(data))
    chunks = []
    for start in range(0, len(reader.pages), OCR_PAGES_PER_CALL):
        writer = PdfWriter()
        for page in reader.pages[start:start + OCR_PAGES_PER_CALL]:
            writer.add_page(page)
        buf = io.BytesIO()
        writer.write(buf)
        chunks.append((start, buf.getvalue()))

    def run(chunk):
        start, pdf = chunk
        last = min(start + OCR_PAGES_PER_CALL, len(reader.pages))
        return _ocr_call([{"type": "document", "source": {"type": "base64", "media_type": ex.PDF_MIME,
                                                          "data": base64.b64encode(pdf).decode("ascii")}},
                          {"type": "text", "text": f"{name}, {start + 1}–{last}. oldal"}], project_id)
    with ThreadPoolExecutor(max_workers=4) as pool:
        parts = list(pool.map(run, chunks))
    return "\n\n".join(f"--- {s + 1}. oldaltól ---\n{t}" if len(chunks) > 1 else t
                        for (s, _), t in zip(chunks, parts) if t)


def fill_missing_texts(docs):
    """Szövegfelismerés a szöveg nélküli (szkennelt) dokumentumokhoz. Visszaad: a sikertelenek listája."""
    progress = st.progress(0.0, text="Szövegfelismerés…")
    failed = []
    for n, doc in enumerate(docs, 1):
        try:
            data = base64.b64decode(get_document_content(doc["id"]))
            text = ocr_document(doc["name"], doc["mime_type"], data, doc.get("project_id"))
            if not text:
                raise ValueError("a modell nem adott vissza szöveget")
            db.set_document_text(doc["id"], doc["name"], text)
        except Exception as e:
            failed.append(f"{doc['name']}: {e}")
        progress.progress(n / len(docs), text=f"Szövegfelismerés… {n}/{len(docs)}")
    return failed


def save_document_bytes(project_id, name, mime_type, data):
    """Mentés a dokumentumtárba a kereséshez kinyert szöveggel és a kivonattal együtt. Visszaad: (sikerült, hiba oka).
    Ha a kivonat nem készül el, a dokumentum attól még mentődik; a kivonat később a Dokumentumok fülön pótolható."""
    text = ex.extract_text(name, mime_type, data)
    if needs_ocr(mime_type, data, text):   # szkennelt anyag: egyszer szöveggé alakítjuk, hogy kereshető legyen
        try:
            text = ocr_document(name, mime_type, data, project_id) or text
        except Exception:
            pass   # a dokumentum ettől még mentődik; a szövegfelismerés a Dokumentumok fülön pótolható
    try:
        summary = make_doc_summary(name, mime_type, data, text, project_id)
    except Exception:
        summary = None   # a dokumentum ettől még mentődik; a kivonat a Dokumentumok fülön pótolható
    try:
        db.save_document(project_id, name, mime_type, data, text, summary)
        return True, ""
    except db.DbError as e:
        return False, str(e)


def fill_missing_summaries(docs):
    """Kivonat a még kivonat nélküli dokumentumokhoz. Visszaad: a sikertelenek listája."""
    progress = st.progress(0.0, text="Kivonatok készítése…")
    failed = []
    for n, doc in enumerate(docs, 1):
        try:
            data = base64.b64decode(get_document_content(doc["id"]))
            summary = make_doc_summary(doc["name"], doc["mime_type"], data, db.get_document_text(doc["id"]),
                                       doc.get("project_id"))
            db.set_document_summary(doc["id"], summary)
        except Exception as e:
            failed.append(f"{doc['name']}: {e}")
        progress.progress(n / len(docs), text=f"Kivonatok készítése… {n}/{len(docs)}")
    return failed


# Sara ezzel maga nyithatja meg a projekt bármelyik dokumentumát (a kivonatok alapján dönti el, melyiket)
DOC_TOOLS = [{
    "name": "read_project_document",
    "description": ("A projekt dokumentumtárában lévő dokumentum teljes szövegének beolvasása név szerint. Akkor "
                    "használd, ha a válaszhoz egy dokumentum tényleges tartalma kell, és az még nincs a "
                    "beszélgetésben. A nevet a rendszerüzenet dokumentumlistájából vedd."),
    "input_schema": {"type": "object", "properties": {
        "name": {"type": "string", "description": "A dokumentum pontos neve a dokumentumlistából."}},
        "required": ["name"]},
}]


# Meglévő Word-dokumentum javítása az eredeti formázással (a szerkesztést a docx_edit modul végzi)
EDIT_TOOL = {
    "name": "edit_word_document",
    "description": (
        "Meglévő Word-dokumentum (.docx) javítása úgy, hogy az eredeti formázás (betűk, stílusok, számozás, "
        "élőfej, táblázatok) teljesen megmarad. Ezt használd, ha a felhasználó egy meglévő Word-dokumentum "
        "módosítását, javítását, kiegészítését kéri (például szerződés, ajánlat), NE a create_document-et. "
        "Nem a teljes szöveget adod meg, hanem a módosítások listáját; minden módosítás a dokumentum egyetlen "
        "bekezdésén vagy táblázatcelláján belüli, szó szerinti szövegrészletre hivatkozik. Ha egy találat több "
        "helyen is előfordul, adj meg hosszabb részletet (vagy all=true, ha mindenhol cserélni kell). Ha "
        "valamelyik módosítás nem hajtható végre, nem készül fájl, és megkapod a hibák listáját: javítsd őket, és "
        "hívd meg újra az összes módosítással. Alapból követett változásként (Word „Változások követése”) kerülnek "
        "be a módosítások, amelyeket a felhasználó egyenként elfogadhat; track_changes=false esetén tisztán."),
    "input_schema": {"type": "object", "properties": {
        "document": {"type": "string", "description": (
            "A javítandó .docx pontos neve: a projekt dokumentumlistájából, vagy a beszélgetésbe csatolt fájl neve.")},
        "output_title": {"type": "string", "description": (
            "Az új fájl neve kiterjesztés nélkül; ha elhagyod, az eredeti neve + „_javitott”.")},
        "track_changes": {"type": "boolean", "description": "Követett változásként (alapértelmezés: true)."},
        "edits": {"type": "array", "minItems": 1, "items": {"type": "object", "properties": {
            "action": {"type": "string", "enum": list(docx_edit.ACTIONS), "description": (
                "replace: a find szövegrészt a new_text-re cseréli; delete: törli a find részt (ha a find a "
                "teljes bekezdés, a bekezdést); insert_after / insert_before: a find-ot tartalmazó bekezdés "
                "után / elé új bekezdés(eke)t szúr be a new_text-tel, a szomszéd bekezdés formázásával "
                "(soronként egy bekezdés); comment: Word-megjegyzést fűz a find szövegrészhez, a megjegyzés "
                "szövege a new_text (a dokumentum szövege nem változik).")},
            "find": {"type": "string", "description": (
                "Szó szerinti szövegrészlet a dokumentumból, egyetlen bekezdésen/cellán belül. Csere esetén "
                "elég a változó rész néhány szavas környezetével; a változatlan részeket a program nem írja át.")},
            "new_text": {"type": "string", "description": "Az új szöveg (cserénél és beszúrásnál)."},
            "all": {"type": "boolean", "description": "Minden előfordulásra (alapértelmezés: false)."}},
            "required": ["action", "find"]}}},
        "required": ["document", "edits"]},
}


def name_key(name):
    """Fájlnév összehasonlításhoz: ékezet, kis-nagybetű, szóköz és írásjel nélkül. A Mac a fájlnevek ékezetes
    betűit szétbontva (a + vessző) tárolja, Sara pedig egyben írja, ezért a sima összehasonlítás nem talált."""
    s = unicodedata.normalize("NFKD", str(name or "")).casefold()
    return re.sub(r"[^a-z0-9]+", "", "".join(c for c in s if not unicodedata.combining(c)))


def pick_by_name(items, wanted, get_name):
    """Az items közül a wanted nevű: előbb pontos (ékezettől és írásmódtól függetlenül), aztán kiterjesztés
    nélkül, végül részleges egyezés."""
    w = name_key(wanted)
    w_stem = name_key(str(wanted).rsplit(".", 1)[0])
    if not w:
        return None
    for test in (lambda k, ks: k == w,
                 lambda k, ks: ks == w_stem):
        for item in items:
            name = get_name(item)
            if test(name_key(name), name_key(str(name).rsplit(".", 1)[0])):
                return item
    if len(w_stem) < 4:
        return None
    # részleges egyezés: ha több név is illik (pl. …_hatarozott / …_hatarozatlan, …_1 / …_2), csak akkor
    # választunk, ha egy egyértelműen a legjobb; különben nincs találat, és a hívó a teljes listát adja vissza
    hits = []
    for item in items:
        ks = name_key(str(get_name(item)).rsplit(".", 1)[0])
        if w_stem in ks or ks in w_stem:
            hits.append((difflib.SequenceMatcher(None, w_stem, ks).ratio(), item))
    if not hits:
        return None
    hits.sort(key=lambda h: -h[0])
    if len(hits) == 1 or hits[0][0] - hits[1][0] >= 0.08:
        return hits[0][1]
    return None


def mid_ellipsis(name, limit=46):
    """Hosszú fájlnév rövidítése a közepén, hogy a megkülönböztető vége (_hatarozott, _2, dátum, kiterjesztés)
    látsszon. A határ a legördülő lista szélességéhez igazodik (kb. 50 karakter fér ki)."""
    name = str(name or "")
    if len(name) <= limit:
        return name
    tail = max(22, limit // 2)
    return name[:limit - tail - 1].rstrip(" _-–") + "…" + name[-tail:].lstrip(" _-–")


def make_edit_handler(project_docs, extra_files):
    """extra_files: [(név, mime, base64)] – a beszélgetésbe csatolt munkaanyagok (a mostaniak és a korábbiak)."""
    def edit_word(args):
        wanted = str(args.get("document", "")).strip()
        data, name = None, None
        hit = pick_by_name(extra_files, wanted, lambda f: f[0])   # a csatolt fájl az aktuálisabb
        if hit:
            data, name = base64.b64decode(hit[2]), hit[0]
        else:
            doc = pick_by_name([d for d in project_docs if d["name"].lower().endswith(".docx")], wanted,
                               lambda d: d["name"]) or pick_by_name(project_docs, wanted, lambda d: d["name"])
            if doc:
                data, name = base64.b64decode(get_document_content(doc["id"])), doc["name"]
        if data is None:
            known = list(dict.fromkeys([f[0] for f in extra_files] + [d["name"] for d in project_docs
                                                                     if d["name"].lower().endswith(".docx")]))
            raise ValueError("Nincs ilyen nevű Word-dokumentum. Az elérhető .docx fájlok: "
                             + (", ".join(known) if known else "nincs egy sem") + ".")
        name = unicodedata.normalize("NFC", name)
        if not name.lower().endswith(".docx"):
            raise ValueError(f"A(z) {name} nem Word-dokumentum (.docx); ezt csak új dokumentumként lehet elkészíteni.")
        tracked = args.get("track_changes") is not False
        new_data, report = docx_edit.apply_edits(data, args.get("edits"), tracked=tracked, author="Sara")
        stem = str(args.get("output_title") or "").strip() or (name.rsplit(".", 1)[0] + "_javitott")
        return ex.safe_filename(stem, "docx"), ex.DOCX_MIME, new_data
    return {"edit_word_document": edit_word}


def resolve_doc_text(pid, extra_files, wanted):
    """Egy dokumentum neve és szövege név szerint: előbb a beszélgetésbe csatolt fájlok, aztán a projekt dokumentumai."""
    wanted = str(wanted or "").strip()
    hit = pick_by_name(extra_files, wanted, lambda f: f[0])   # a csatolt fájl az aktuálisabb
    if hit:
        return hit[0], ex.extract_text(hit[0], hit[1], base64.b64decode(hit[2]))
    docs = db.get_documents(pid)
    doc = pick_by_name(docs, wanted, lambda d: d["name"])
    if not doc:
        raise ValueError(f"Nincs „{wanted}” nevű dokumentum. Elérhetők: "
                         + ", ".join(list(dict.fromkeys([f[0] for f in extra_files] + [d["name"] for d in docs]))))
    return doc["name"], db.get_document_text(doc["id"])


DOC_STATUS_TOOL = {
    "name": "set_document_status",
    "description": ("Egy projektdokumentum kategóriájának és/vagy elavult jelölésének beállítása (az elavult "
                    "dokumentumot ezután csak kérésre nézed). CSAK akkor használd, ha a felhasználó kifejezetten "
                    "kéri (pl. „a v1 elavult”, „ez a szerződések közé tartozik”)."),
    "input_schema": {"type": "object", "properties": {
        "document": {"type": "string", "description": "A dokumentum pontos neve."},
        "category": {"type": "string", "enum": DOC_CATEGORIES},
        "obsolete": {"type": "boolean", "description": "true: elavult (régi változat); false: újra érvényes."}},
        "required": ["document"]},
}


def make_doc_status_handler(pid):
    def set_status(args):
        docs = db.get_documents(pid)
        doc = pick_by_name(docs, str(args.get("document") or ""), lambda d: d["name"])
        if not doc:
            raise ValueError("Nincs ilyen nevű dokumentum a projektben: " + str(args.get("document")))
        fields = {}
        if args.get("category") in DOC_CATEGORIES:
            fields["category"] = args["category"]
        if isinstance(args.get("obsolete"), bool):
            fields["obsolete"] = args["obsolete"]
        if not fields:
            raise ValueError("Adj meg kategóriát vagy elavult jelölést.")
        db.set_document_org(doc["id"], **fields)
        parts = []
        if "category" in fields:
            parts.append(f"kategória: {fields['category']}")
        if "obsolete" in fields:
            parts.append("elavult" if fields["obsolete"] else "érvényes")
        return f"Beállítva: {doc['name']} – " + ", ".join(parts)
    return {"set_document_status": set_status}


def make_legal_handlers(pid, extra_files):
    """Larától átvett eszközök: jogszabály (njt.hu), adószám (NAV), Red Flag, verzió-összehasonlítás, redline."""
    def red_flag_review(args):
        name, text = resolve_doc_text(pid, extra_files, args.get("document"))
        return redflag.review(claude_client, MODEL, name, text, str(args.get("client_role") or "a felhasználó"),
                              str(args.get("focus") or ""), str(args.get("context") or ""),
                              log=usage_logger(pid, "redflag"))

    def compare_versions(args):
        old_name, old_text = resolve_doc_text(pid, extra_files, args.get("old_document"))
        new_name, new_text = resolve_doc_text(pid, extra_files, args.get("new_document"))
        if old_name == new_name:
            raise ValueError("A két megadott név ugyanarra a dokumentumra mutat; adj meg két különböző változatot.")
        return diffreport.report(claude_client, MODEL, old_name, old_text, new_name, new_text,
                                 str(args.get("client_role") or ""), str(args.get("context") or ""),
                                 log=usage_logger(pid, "diff"))

    def create_redline(args):
        old_name, old_text = resolve_doc_text(pid, extra_files, args.get("old_document"))
        new_name, new_text = resolve_doc_text(pid, extra_files, args.get("new_document"))
        data, count = diffreport.redline_docx(old_text, new_text, f"Változások: {old_name} → {new_name}",
                                              author="Sara")
        stem = str(args.get("title") or "").strip() or (new_name.rsplit(".", 1)[0] + "_valtozasok")
        return ex.safe_filename(stem, "docx"), ex.DOCX_MIME, data

    return {"red_flag_review": red_flag_review, "compare_versions": compare_versions,
            "create_redline_document": create_redline,
            "adoszam_ellenorzes": lambda args: navtax.report(*navtax.query(NAV_CFG, args.get("adoszam"))),
            "feor_kereso": feor.tool_report,
            "njt_jogszabaly": lambda args: njt.lookup(args.get("jogszabaly"), args.get("paragrafus"),
                                                      args.get("kulcsszo"))}


def make_finder_handler(pid, projects):
    """Sara kereső kollégája (Haiku): adatot keres az aktuális és a nem bizalmas projektek dokumentumaiban."""
    def find_information(args):
        return finder.find(
            claude_client, args.get("question"), pid, projects,
            get_docs_meta=db.get_documents_meta,
            get_text=db.get_document_text,
            keyword_search=lambda ids, stems: db.search(ids, stems)[0],
            log=usage_logger(pid, "kereso"))
    def research_web(args):
        return finder.web_research(claude_client, args.get("question"), log=usage_logger(pid, "kereso_web"))
    return {"find_information": find_information, "research_web": research_web}


def make_consistency_handler(pid, project_docs):
    """Az ellenőrző kolléga (Haiku): több dokumentum egymásnak ellentmondó adatainak kigyűjtése."""
    def consistency_check(args):
        names = [str(n) for n in (args.get("documents") or []) if str(n).strip()]
        if names:
            docs, missing = [], []
            for n in names:
                d = pick_by_name(project_docs, n, lambda x: x["name"])
                (docs if d else missing).append(d or n)
            if missing:
                raise ValueError("Nincs ilyen nevű dokumentum a projektben: " + "; ".join(missing)
                                 + ". A pontos nevet a dokumentumlistából vedd.")
        else:
            docs = current_docs(project_docs)   # alapból csak az érvényes változatok
        docs = list({d["id"]: d for d in docs}.values())[:consistency.MAX_DOCS]
        ref = pick_by_name(docs, str(args.get("reference") or ""), lambda x: x["name"]) \
            if str(args.get("reference") or "").strip() else None
        texts = []
        for d in docs:
            try:
                texts.append((d["name"], db.get_document_text(d["id"]) or ""))
            except db.DbError:
                texts.append((d["name"], ""))
        return consistency.check(claude_client, texts, str(args.get("focus") or "").strip(),
                                 ref["name"] if ref else "", log=usage_logger(pid, "ellenorzo"))
    return {"consistency_check": consistency_check}


def make_doc_handlers(project_docs):
    def read_document(args):
        doc = pick_by_name(project_docs, str(args.get("name", "")), lambda d: d["name"])
        if not doc:
            raise ValueError("Nincs ilyen nevű dokumentum a projektben; a pontos nevet a dokumentumlistából vedd.")
        text = ""
        if doc["mime_type"] in (ex.DOCX_MIME, ex.XLSX_MIME, ex.PPTX_MIME):
            try:   # az irodai fájlokat frissen olvassuk ki, hogy a régebben feltöltötteknél is a teljes szöveg legyen
                data = base64.b64decode(get_document_content(doc["id"]))
                text = (ex.extract_text(doc["name"], doc["mime_type"], data) or "").strip()
            except Exception:
                text = ""
        text = text or db.get_document_text(doc["id"]).strip()
        if not text:
            return (f"A(z) {doc['name']} dokumentumnak nincs kiolvasható szövege (például rajz vagy kép). Kérd meg a "
                    "felhasználót, hogy jelölje ki a kérdéséhez, akkor képként látod.")
        return f"A(z) {doc['name']} teljes szövege:\n\n{text[:150000]}"
    return {"read_project_document": read_document}


def live_context(tasks, project_docs, extra=""):
    """A gyakran változó rész (feladatok, dokumentumlista) nem a rendszerüzenetben, hanem a legutóbbi kérdés
    elején megy el. Így a beszélgetés eleje (eszközök, rendszerüzenet, előzmények) változatlan marad, és a
    gyorsítótárból jön (a tizedénél is olcsóbban); ha a lista a rendszerüzenetben lenne, minden feladat- vagy
    dokumentumváltozás után az egész előzmény újra a tárba íródna, kétszeres áron."""
    text = (str(extra or "") + tasks_for_prompt(tasks) + docs_for_prompt(project_docs)).strip()
    return f"[AKTUÁLIS ÁLLAPOT]\n{text}\n[/AKTUÁLIS ÁLLAPOT]\n\n" if text else ""


# --- TÖRZSAKTÁK: cég, ingatlan, személy, jármű (tartós adatok, adatlappal) ---------------------------------
def is_record(p):
    return bool(p) and p.get("kind") == "torzs"


def is_templates(p):
    return bool(p) and p.get("kind") == "minta"


def is_special(p):
    return is_record(p) or is_templates(p)


def record_label(p):
    label, icon, _ = records.RECORD_TYPES.get(p.get("record_type"), ("Törzsakta", ":material/folder:", ""))
    return f"{icon} {p.get('name')}"


def record_plain(p):
    """Törzsakta neve legördülő listába (ott a :material/…: ikonkód nem jelenik meg, csak szövegként)."""
    return f"{p.get('name')} ({record_kind_name(p).lower()})"


def record_kind_name(p):
    return records.RECORD_TYPES.get(p.get("record_type"), ("Törzsakta",))[0]


def record_sheet_doc(record_id):
    """A törzsakta legfrissebb (nem elavult) adatlapja, vagy None."""
    try:
        docs = db.get_documents(record_id)
    except db.DbError:
        return None
    sheets = [d for d in docs if d["name"].startswith(records.SHEET_PREFIX) and not is_obsolete(d)]
    return sheets[0] if sheets else None   # a lista a legújabbal kezdődik


@st.cache_data(ttl=120, show_spinner=False, max_entries=200)
def record_sheet_text(record_id, doc_id):
    try:
        return (db.get_document_text(doc_id) or "").strip()
    except db.DbError:
        return ""


def linked_records_context(project, projects):
    """A projekthez kapcsolt törzsakták adatlapjai Sarának (a kérdés elejére, a gyorsítótár után)."""
    if not project or not db.has_records():
        return ""
    by_id = {p["id"]: p for p in projects}
    if is_record(project):
        head = f"EZ EGY TÖRZSAKTA ({record_kind_name(project)}: {project['name']}). Az adatlapja a dokumentumai között van.\n\n"
        extra = record_extra_text(project)
        if extra:
            head += extra + "\n\n"
        linked_here = [p for p in projects if project["id"] in (p.get("links") or [])]
        if linked_here:
            head += "Ezekhez a projektekhez kapcsolódik: " + "; ".join(p["name"] for p in linked_here) + "\n\n"
        return head
    parts = []
    for rid in project.get("links") or []:
        rec = by_id.get(rid)
        if not rec:
            continue
        doc = record_sheet_doc(rid)
        text = record_sheet_text(rid, doc["id"]) if doc else ""
        extra = record_extra_text(rec)
        parts.append(f"### {record_kind_name(rec)}: {rec['name']}\n"
                     + (text[:5000] if text else "(még nincs adatlapja – a törzsaktába feltöltött iratból készíthető)")
                     + (f"\n{extra}" if extra else ""))
    if not parts:
        return ""
    return "KAPCSOLT TÖRZSAKTÁK (a projekthez kapcsolt cégek, ingatlanok, személyek, járművek adatlapja):\n" \
        + "\n\n".join(parts) + "\n\n"


RECORD_CONTACTS = [("konyvelo", "Könyvelő"), ("konyvvizsgalo", "Könyvvizsgáló"), ("ugyved", "Ügyvéd")]
FOUNDATION_CONTACTS = [("vagyonellenor", "Vagyonellenőr"), ("szakerto", "Szakértő")]


@st.cache_data(ttl=300, show_spinner=False, max_entries=200)
def record_extra(record_id):
    """A törzsakta kézzel megadott adatai (pl. cégnél könyvelő, könyvvizsgáló, ügyvéd)."""
    try:
        return db.get_setting(f"record:{record_id}", {}) or {}
    except db.DbError:
        return {}


def record_is_foundation(record, extra=None):
    extra = extra if extra is not None else record_extra(record["id"])
    if "alapitvany" in extra:
        return bool(extra["alapitvany"])
    return "alapítvány" in str(record.get("name") or "").lower()


def record_extra_text(record):
    """A kézi adatok szövegként (Sarának és az adatlap fülre)."""
    extra = record_extra(record["id"])
    fields = RECORD_CONTACTS + (FOUNDATION_CONTACTS if record_is_foundation(record, extra) else [])
    lines = []
    for key, label in fields:
        main, contact = str(extra.get(key) or "").strip(), str(extra.get(key + "_kapcs") or "").strip()
        if main or contact:
            lines.append(f"- {label}: {main or '–'}" + (f"; aki ténylegesen intézi: {contact}" if contact else ""))
    if str(extra.get("megjegyzes") or "").strip():
        lines.append(f"- Megjegyzés: {extra['megjegyzes']}")
    return ("Kapcsolattartók és egyéb adatok (kézzel megadva):\n" + "\n".join(lines)) if lines else ""


def render_record_extra(record, readonly):
    """Cégnél (alapítványnál) a könyvelő, könyvvizsgáló, ügyvéd és – alapítványnál – vagyonellenőr, szakértő."""
    if record.get("record_type") != "ceg":
        return
    extra = dict(record_extra(record["id"]))
    with st.expander(":material/contacts: Kapcsolattartók (könyvelő, könyvvizsgáló, ügyvéd…)",
                     expanded=not readonly and not any(str(extra.get(k) or "").strip()
                                                       for k, _ in RECORD_CONTACTS + FOUNDATION_CONTACTS)):
        with st.form(f"rec_extra_{record['id']}", border=False):
            found = st.checkbox("Alapítvány (vagyonellenőr, szakértő is)", value=record_is_foundation(record, extra),
                                disabled=readonly)
            vals = {}
            for key, label in RECORD_CONTACTS + FOUNDATION_CONTACTS:
                if (key, label) in FOUNDATION_CONTACTS and not found:
                    continue
                c1, c2 = st.columns(2)
                vals[key] = c1.text_input(label, value=str(extra.get(key) or ""), disabled=readonly,
                                          placeholder="iroda / bejegyzett személy (pl. a cégkivonat szerint)")
                vals[key + "_kapcs"] = c2.text_input(f"{label} – aki ténylegesen intézi",
                                                     value=str(extra.get(key + "_kapcs") or ""), disabled=readonly,
                                                     placeholder="munkatárs neve, telefon, e-mail")
            vals["megjegyzes"] = st.text_area("Egyéb (pl. bankszámla-kezelő, tárhely, ügyintéző)",
                                              value=str(extra.get("megjegyzes") or ""), height=70, disabled=readonly)
            if st.form_submit_button(":material/save: Mentés", disabled=readonly):
                data = {**extra, **{k: str(v or "").strip() for k, v in vals.items()}, "alapitvany": bool(found)}
                try:
                    db.set_setting(f"record:{record['id']}", data)
                    record_extra.clear()
                    st.session_state["sheet_flash"] = ("success", "Elmentve; Sara a következő kérdéstől látja.")
                    st.rerun()
                except db.DbError as e:
                    st.error(f"A mentés nem sikerült: {e}")


def save_record_sheet(record, markdown, source_name):
    """Az új adatlap mentése a törzsaktába; a korábbi adatlap elavult lesz (ha van rendezés), így a történet megmarad."""
    old = record_sheet_doc(record["id"])
    clean = " ".join(re.sub(r"[\\/:*?\"<>|]+", "-", str(record["name"])).split()).strip(" .") or "törzsakta"
    name = unique_name(f"{records.SHEET_PREFIX}{clean}.md",
                       {d["name"] for d in db.get_documents(record["id"])})
    ok, detail = save_document_bytes(record["id"], name, ex.MD_MIME, (markdown + "\n").encode("utf-8"))
    if not ok:
        raise ValueError(detail)
    if old and db.has_doc_org():
        try:
            db.set_document_org(old["id"], obsolete=True)
        except db.DbError:
            pass
    record_sheet_text.clear()
    return name


def make_record_sheet(record, source_docs):
    """Adatlap a törzsakta típusa szerint a kiválasztott iratokból. Visszaad: (adatlap neve, figyelmeztetések)."""
    rtype = record.get("record_type")
    pid = record["id"]
    if rtype in ("ceg", "ingatlan"):
        doc = source_docs[0]
        text = db.get_document_text(doc["id"]) or ""
        data = base64.b64decode(get_document_content(doc["id"]))
        result = registry.sheet(claude_client, text, doc["name"], log=usage_logger(pid, "adatlap"),
                                pdf=data if doc["mime_type"] == ex.PDF_MIME else None)
        if not result:
            raise ValueError(f"A(z) {doc['name']} nem ismerhető fel " + (
                "cégkivonatként vagy civil szervezeti (alapítványi) kivonatként." if rtype == "ceg"
                else "tulajdoni lapként."))
        kind, md, risks, warns = result
        if kind == "civil" and rtype == "ceg":   # alapítvány: a vagyonellenőr, szakértő mezők is kellenek
            try:
                extra = dict(record_extra(pid))
                if not extra.get("alapitvany"):
                    db.set_setting(f"record:{pid}", {**extra, "alapitvany": True})
                    record_extra.clear()
            except db.DbError:
                pass
        return save_record_sheet(record, md, doc["name"]), list(risks) + list(warns)
    files = [(d["name"], d["mime_type"], base64.b64decode(get_document_content(d["id"]))) for d in source_docs]
    src = ", ".join(d["name"] for d in source_docs)
    if rtype == "szemely":
        data, md = records.person_sheet(claude_client, SARA_MODEL_ID, files, src, log=usage_logger(pid, "adatlap"))
        return save_record_sheet(record, md, src), records.person_checks(data)
    if rtype == "jarmu":
        data, md = records.vehicle_sheet(claude_client, SARA_MODEL_ID, files, src, log=usage_logger(pid, "adatlap"))
        return save_record_sheet(record, md, src), records.vehicle_checks(data)
    raise ValueError("Ismeretlen törzsakta-típus.")


def render_record_sheet(record, project_docs, readonly):
    """Az Adatlap fül: a legfrissebb adatlap, és az adatlap elkészítése a feltöltött iratokból."""
    with st.container(height=FALLBACK_HEIGHT_PX, border=False, key="sheet_box"):
        rtype = record.get("record_type")
        hint = {"ceg": "cégkivonatból (e-cégjegyzék PDF; alapítványnál a civilsznev.birosag.hu kivonata)", "ingatlan": "tulajdoni lapból (TakarNet PDF)",
                "szemely": "okmányból (személyi igazolvány, lakcímkártya – kép vagy PDF, elő- és hátoldal)",
                "jarmu": "forgalmi engedélyből (kép vagy PDF; ha van, törzskönyvvel)"}.get(rtype, "")
        show_flash("sheet_flash")
        render_record_extra(record, readonly)
        paper = str(record_extra(record["id"]).get("fejlec") or "")
        sources = [d for d in project_docs if not d["name"].startswith(records.SHEET_PREFIX) and not is_obsolete(d)
                   and str(d["id"]) != paper   # a fejléces papír nem forrás
                   and not (rtype in ("ceg", "ingatlan") and d["name"].lower().endswith((".docx", ".doc")))]
        if not readonly:
            if not sources:
                st.info(f"Az adatlap {hint} készül. Töltsd fel a Dokumentumok fülön, utána itt egy gombnyomás.")
            else:
                c_sel, c_btn = st.columns([4, 1.6], vertical_alignment="bottom")
                multi = rtype in ("szemely", "jarmu")
                if multi:
                    chosen = c_sel.multiselect(f"Adatlap készítése {hint}:", [d["id"] for d in sources],
                                               format_func=lambda i: next(d["name"] for d in sources if d["id"] == i),
                                               key=f"sheet_src_{record['id']}")
                else:
                    pick = c_sel.selectbox(f"Adatlap készítése {hint}:", [d["id"] for d in sources],
                                           format_func=lambda i: next(d["name"] for d in sources if d["id"] == i),
                                           key=f"sheet_src_{record['id']}")
                    chosen = [pick] if pick else []
                if c_btn.button(":material/badge: Adatlap készítése", key=f"sheet_go_{record['id']}",
                                disabled=not chosen, **STRETCH):
                    with st.spinner("Az adatlap készül…"):
                        try:
                            name, warns = make_record_sheet(record, [d for d in sources if d["id"] in chosen])
                            st.session_state["sheet_flash"] = (
                                "warning" if warns else "success",
                                f"Elkészült: {name}." + (" Ellenőrizendő: " + "; ".join(warns) if warns else ""))
                        except Exception as e:
                            st.session_state["sheet_flash"] = ("error", f"Az adatlap nem készült el: {e}")
                    st.rerun()
        doc = record_sheet_doc(record["id"])
        if doc:
            st.markdown(record_sheet_text(record["id"], doc["id"]) or "_(az adatlap üres)_")
        else:
            st.caption(f"Ennek a törzsaktának még nincs adatlapja. Az adatlap {hint} készül.")


# --- DOKUMENTUMOK RENDEZÉSE: kategória és elavult jelölés -------------------------------------------------

CATEGORIZE_SYSTEM = (
    "Egy projekt dokumentumait sorolod kategóriákba a nevük és a kivonatuk alapján. A kategóriák (csak ezek közül "
    "válassz, pontosan így írva): " + "; ".join(DOC_CATEGORIES) + ". Csak JSON-t adj vissza: "
    '{"<azonosító>": "<kategória>", …}')


def doc_category(doc):
    c = str(doc.get("category") or "").strip()
    return c if c in DOC_CATEGORIES else UNCATEGORIZED


def is_obsolete(doc):
    return bool(doc.get("obsolete"))


def current_docs(project_docs):
    """Az érvényes (nem elavult) dokumentumok."""
    return [d for d in (project_docs or []) if not is_obsolete(d)]


def categorize_docs(project_id, docs):
    """A kategória nélküli dokumentumok besorolása egyetlen olcsó modellhívással (név + kivonat alapján).
    Visszaad: hány dokumentum kapott kategóriát."""
    docs = [d for d in docs if doc_category(d) == UNCATEGORIZED]
    if not docs or not db.has_doc_org():
        return 0
    done = 0
    for start in range(0, len(docs), 60):
        chunk = docs[start:start + 60]
        lines = [f"{d['id']}: {d['name']} – {' '.join(str(d.get('summary') or '').split())[:300]}" for d in chunk]
        kwargs = dict(model=SUMMARY_MODEL, max_tokens=3000, system=CATEGORIZE_SYSTEM,
                      messages=[{"role": "user", "content": "\n".join(lines)}])
        try:
            response = claude_client.messages.create(thinking={"type": "disabled"}, **kwargs)
        except Exception as e:
            if getattr(e, "status_code", None) != 400 or "thinking" not in str(e).lower():
                raise
            response = claude_client.messages.create(**kwargs)
        db.log_usage(project_id, "kivonat", SUMMARY_MODEL, getattr(response, "usage", None))
        text = "".join(getattr(b, "text", "") for b in response.content if getattr(b, "type", "") == "text")
        m = re.search(r"\{.*\}", text, flags=re.DOTALL)
        try:
            data = json.loads(m.group(0)) if m else {}
        except ValueError:
            data = {}
        by_id = {str(d["id"]): d for d in chunk}
        for key, cat in (data or {}).items():
            cat = str(cat or "").strip()
            if str(key) in by_id and cat in DOC_CATEGORIES:
                try:
                    db.set_document_org(by_id[str(key)]["id"], category=cat)
                    done += 1
                except db.DbError:
                    pass
    return done


_VERSION_NOISE = re.compile(r"(?i)(\bv(er(zi[oó])?)?\.?\s*\d+(\.\d+)*\b|\b\d{4}[-._]?\d{2}[-._]?\d{2}\b|"
                            r"\(\d+\)|\bv[eé]gleges\b|\bjav[ií]tott\b|\bfinal\b|\bdraft\b|\btervezet\b|"
                            r"\bm[oó]dos[ií]tott\b|\b[uú]j\b|\br[eé]gi\b)")


def _version_stem(name):
    stem = str(name or "").rsplit(".", 1)[0]
    stem = unicodedata.normalize("NFKD", stem.lower())
    stem = "".join(c for c in stem if not unicodedata.combining(c))
    stem = _VERSION_NOISE.sub(" ", stem)
    return " ".join(re.sub(r"[_\-.]+", " ", stem).split())


def older_versions(new_doc, project_docs):
    """Ugyanannak a dokumentumnak a korábbi, még érvényes változatai (a név a verziójelölésektől eltekintve
    egyezik vagy nagyon hasonló)."""
    stem = _version_stem(new_doc["name"])
    if len(stem) < 4:
        return []
    out = []
    for d in project_docs:
        if d["id"] == new_doc["id"] or is_obsolete(d) or str(d.get("created_at") or "") > str(new_doc.get("created_at") or ""):
            continue
        other = _version_stem(d["name"])
        if other == stem or difflib.SequenceMatcher(None, other, stem).ratio() >= 0.88:
            out.append(d)
    return out


def docs_for_prompt(project_docs):
    """A projekt dokumentumainak listája a kivonatokkal, hogy Sara tudja, mi hol található."""
    if not project_docs:
        return ""
    current = current_docs(project_docs)
    old = [d for d in project_docs if is_obsolete(d)]
    lines, budget = [], 30000
    order = DOC_CATEGORIES + [UNCATEGORIZED]
    grouped = sorted(current, key=lambda d: order.index(doc_category(d)))
    last_cat = None
    for d in grouped:
        if db.has_doc_org() and doc_category(d) != last_cat:
            last_cat = doc_category(d)
            lines.append(f"[{last_cat}]")
        line = f"- {d['name']} ({fmt_dt(d.get('created_at'))[:10]}): {(d.get('summary') or 'kivonat még nincs').strip()}"
        budget -= len(line)
        if budget < 0:
            lines.append("- (a többi dokumentum kivonata helyhiány miatt kimaradt)")
            break
        lines.append(line)
    if old:
        lines.append("[Elavult – régi változatok, csak kifejezett kérésre vagy összevetéshez használd]: "
                     + "; ".join(d["name"] for d in old[:60]))
    return ("\n\nA PROJEKT DOKUMENTUMAI:\nA projekt dokumentumtárában ezek a fájlok vannak, rövid kivonattal. "
            "A kivonat csak tájékozódásra való. Ha egy kérdéshez egy dokumentum tényleges tartalma kell, és a "
            "felhasználó nem jelölte ki, olvasd be a read_project_document eszközzel; pontos számot, feltételt "
            "mindig az eredetiből vegyél. Ha egy kérdésre egy dokumentum adja a választ, mondd meg, melyik.\n"
            + "\n".join(lines))


def render_downloads(doc, key_prefix=""):
    """Letöltőgombok egy dokumentumhoz (a tartalom csak itt, igény szerint töltődik be).
    A key_prefix azért kell, mert ugyanaz a dokumentum több helyen is megjelenhet."""
    try:
        data = base64.b64decode(get_document_content(doc["id"]))
    except Exception:
        st.warning("A dokumentum nem tölthető be.")
        return
    base_name = doc["name"].rsplit(".", 1)[0]
    is_md = is_markdown_doc(doc)
    cols = st.columns(3 if is_md else 1)
    cols[0].download_button(f":material/download: {doc['name']}", data=data, file_name=doc["name"],
                            mime=doc["mime_type"], key=f"{key_prefix}dlb_{doc['id']}")
    if is_md:
        opts = paper_options()
        paper = project_paper(doc.get("project_id"))
        values = [v for v, _ in opts]
        if len(opts) > 2 or paper:
            paper = st.selectbox("Papír", values, index=values.index(paper) if paper in values else 0,
                                 format_func=lambda v: dict(opts)[v], key=f"{key_prefix}paper_{doc['id']}",
                                 label_visibility="collapsed",
                                 help="Milyen fejléces papírra készüljön a Word és a PDF. Az alapértéket a projekt "
                                      "beállítása adja (a projekt Beállítások fülén).")
        for col, (label, fmt, mime) in zip(cols[1:], ((":material/download: Word (.docx)", "docx", ex.DOCX_MIME),
                                                      (":material/download: PDF", "pdf", ex.PDF_MIME))):
            try:
                data_out = export_markdown(doc["id"], fmt, base_name, paper)
                if data_out is None:
                    col.caption("PDF: nyisd meg a Wordöt, és mentsd PDF-be (a szerveren nincs átalakító).")
                    continue
                col.download_button(label, data=data_out, file_name=f"{base_name}.{fmt}", mime=mime,
                                    key=f"{key_prefix}dlb_{fmt}_{doc['id']}")
            except Exception as e:
                col.warning(f"A(z) {fmt.upper()} nem készíthető el: {e}")


def render_doc_preview(doc, text=None):
    """Dokumentum előnézete: kép, Markdown, vagy a kinyert szöveg eleje."""
    try:
        if doc["mime_type"].startswith("image/"):
            st.image(base64.b64decode(get_document_content(doc["id"])), caption=doc["name"], width=400)
        elif is_markdown_doc(doc):
            st.markdown(base64.b64decode(get_document_content(doc["id"])).decode("utf-8", errors="replace"))
        else:
            text = db.get_document_text(doc["id"]) if text is None else text
            if text.strip():
                st.text(text[:6000] + ("\n[…]" if len(text) > 6000 else ""))
            else:
                st.caption("Ehhez a fájlhoz nincs szöveges előnézet; letöltve megnyitható.")
    except Exception:
        st.warning("Az előnézet nem tölthető be.")


# --- A CHATBEN KÉSZÜLT FÁJLOK NYOMON KÖVETÉSE ---
# Ezeket a sorokat az alkalmazás fűzi Sara válasza alá (nem Sara írja).
CREATED_MARK = "📎 **Elkészült:**"
PROBLEM_MARK = "⚠️ **Figyelem:**"
TASK_MARK = "🗒️ **Feladatok:**"


def created_names(content):
    """Az üzenet alá fűzött „Elkészült” sorból a fájlnevek."""
    for line in str(content).splitlines():
        if line.startswith(CREATED_MARK):
            return list(dict.fromkeys(re.findall(r"`([^`]+)`", line)))
    return []


def strip_app_notes(content):
    """Az alkalmazás által hozzáfűzött sorok nélkül adjuk vissza a választ Sarának, különben utánozni kezdi
    őket: „Elkészült”-et ír úgy, hogy valójában nem hívta meg a dokumentumkészítő eszközt."""
    kept = [l for l in str(content).splitlines() if not l.startswith((CREATED_MARK, PROBLEM_MARK, TASK_MARK))]
    return "\n".join(kept).strip() or "(A fájl elkészült.)"


def split_blocks(text):
    """A válasz részekre bontása a megjegyzésekhez: üres sor választja el őket, és a felsorolás minden
    (legfelső szintű) pontja is külön rész – így egy 18 kérdéses listánál minden kérdés alá jut egy mező.
    A kódrészlet egyben marad, a magában álló címsor pedig a következő részhez tapad."""
    blocks, current, in_code = [], [], False
    for line in str(text).splitlines():
        if line.strip().startswith("```"):
            in_code = not in_code
        if not in_code and current and re.match(r"(\d{1,3}[.)]|[-*•])\s+\S", line):
            blocks.append("\n".join(current))   # új felsorolási pont: új rész (a behúzott alpontok maradnak)
            current = []
        if not line.strip() and not in_code:
            if current:
                blocks.append("\n".join(current))
                current = []
            continue
        current.append(line)
    if current:
        blocks.append("\n".join(current))
    merged = []
    for block in blocks:
        if merged and re.fullmatch(r"#{1,6} .*", merged[-1]):
            merged[-1] += "\n\n" + block
        else:
            merged.append(block)
    return merged


def comments_message(answer, notes):
    """A megjegyzések Sarának szóló üzenete: minden megjegyzés fölött az idézett rész."""
    first = re.sub(r"[#*_`>\s]+", " ", answer).strip()
    head = (first[:60] + "…") if len(first) > 60 else first
    parts = [f"Megjegyzéseim a(z) „{head}” kezdetű válaszodhoz:"]
    for n, (block, note) in enumerate(notes, 1):
        quoted = block if len(block) <= 800 else block[:800] + " …"
        parts.append(f"{n}. Ehhez a részhez:\n" + "\n".join("> " + line for line in quoted.splitlines())
                     + f"\n\nMegjegyzés: {note.strip()}")
    parts.append("Kérlek, dolgozd fel a megjegyzéseimet: ahol javítást kérek, add meg a javított részt (ha sok a "
                 "változás, a teljes javított szöveget), ahol kérdezek, válaszolj.")
    return "\n\n".join(parts)


def find_doc(docs, name, not_after=None):
    """A megadott nevű dokumentum; azonos nevűek közül az, amelyik az üzenet előtt utoljára készült."""
    same = [d for d in docs if d["name"] == name]  # a lista a legújabbal kezdődik
    limit = parse_ts(not_after) if not_after else None
    if same and limit:
        for d in same:
            created = parse_ts(d.get("created_at"))
            if created and created <= limit:
                return d
    return same[0] if same else None


def unique_name(name, taken):
    """Ha már van ilyen nevű fájl a projektben, sorszámot kap (így az új változat nem keveredik a régivel)."""
    if name not in taken:
        return name
    stem, _, ext = name.rpartition(".")
    n = 2
    while f"{stem}_{n}.{ext}" in taken:
        n += 1
    return f"{stem}_{n}.{ext}"


# --- A MODELLNEK ÁTADOTT TARTALOM ÖSSZEÁLLÍTÁSA -------------------------------
def file_parts(name, mime, b64_data):
    """Egy fájl a modell számára. Visszaad: (tartalomblokkok, szöveges mellékletek)."""
    if mime.startswith("image/"):
        return [{"type": "image", "source": {"type": "base64", "media_type": mime, "data": b64_data}}], []
    if mime == ex.PDF_MIME:
        return [{"type": "document", "title": name,
                 "source": {"type": "base64", "media_type": ex.PDF_MIME, "data": b64_data}}], []
    text = None
    if mime.startswith("text/"):
        text = base64.b64decode(b64_data).decode("utf-8", errors="replace")
    elif mime in (ex.XLSX_MIME, ex.PPTX_MIME, ex.DOCX_MIME):  # szöveggé alakítva (a képletek is látszanak)
        try:
            text = ex.office_to_text(mime, base64.b64decode(b64_data))
        except Exception:
            text = None
    return [], ([f"[Mellékelt fájl tartalma ({name}):\n{text}]"] if text else [])


def user_content(text, files):
    """Felhasználói üzenet tartalma: files = [(név, mime, base64)], a szöveg a végére kerül."""
    blocks, texts = [], []
    for name, mime, b64_data in files:
        b, t = file_parts(name, mime, b64_data)
        blocks += b
        texts += t
    full_text = "\n\n".join(texts + [text])
    return blocks + [{"type": "text", "text": full_text}] if blocks else full_text


def load_attachment_files(attachments):
    files = []
    for att in attachments:
        try:
            files.append((att["name"], att["mime_type"], get_attachment_content(att["id"])))
        except Exception:
            pass  # a hiányzó munkaanyag ne akassza meg a beszélgetést
    return files


def build_history(messages, atts_by_msg):
    """A tárolt beszélgetés a modell számára, az üzenetekhez csatolt munkaanyagokkal együtt."""
    history = []
    for m in messages:
        if m.get("role") == "assistant":
            history.append({"role": "assistant", "content": strip_app_notes(m["content"])})
        elif m.get("role") == "user":
            files = load_attachment_files(atts_by_msg.get(m.get("id"), []))
            history.append({"role": "user", "content": as_blocks(user_content(m["content"], files))})
    return history


CACHE_MARK = {"type": "ephemeral", "ttl": "1h"}


def as_blocks(content):
    """Egységes alak (blokklista), hogy a mentett kérdés a következő körben betűre ugyanúgy nézzen ki."""
    return content if isinstance(content, list) else [{"type": "text", "text": content}]


def current_message(saved_question, saved_files, extra_text="", extra_files=()):
    """A most küldött kérdés. Az eleje pontosan az, ami az adatbázisba mentődött (a kérdés és a hozzá mentett
    munkaanyagok) – ennek a végére kerül a gyorsítótár-jel, így a következő körben az egész előzmény innen
    olvasható. Ami csak most megy ki (a friss feladat- és dokumentumlista, a kijelölt dokumentumok), az a jel
    UTÁN jön: ha a mentett szövegbe keveredne, a következő körben már nem egyezne, és minden kérdésnél az
    egész beszélgetés újra a tárba íródna (kétszeres áron)."""
    blocks = [dict(b) for b in as_blocks(user_content(saved_question, saved_files))]
    blocks[-1]["cache_control"] = CACHE_MARK
    extra_text = str(extra_text or "").strip()
    if extra_text or extra_files:
        blocks += as_blocks(user_content(extra_text or "(a kérdéshez kijelölt dokumentumok)", list(extra_files)))
    return blocks


MEMORY_MARK = "[Összevont emlékezet]"
MEMORY_WORDS = 3000   # az összevont emlékezet felső határa (szó): így az alap nem hízik minden lezárással


def is_memory(summary):
    return str(summary.get("description") or "").startswith(MEMORY_MARK)


def active_summaries(summaries):
    """Amit a modell megkap: a legutóbbi összevont emlékezet és az utána készült összefoglalók. A régebbiek
    archívak (az Összefoglalók fülön megmaradnak, kérésre a read_old_summary eszközzel beolvashatók). Ha még
    nincs összevont emlékezet (régi projekt), mind aktív."""
    summaries = list(summaries or [])
    last = max((i for i, s in enumerate(summaries) if is_memory(s)), default=None)
    return summaries if last is None else summaries[last:]


def archived_summaries(summaries):
    active_ids = {id(s) for s in active_summaries(summaries)}
    return [s for s in (summaries or []) if id(s) not in active_ids]


OLD_SUMMARY_TOOL = {
    "name": "read_old_summary",
    "description": ("Egy archív (az összevont emlékezetbe már beolvasztott) régi összefoglaló teljes szövege. Csak "
                    "akkor használd, ha egy régi részlet kell, ami az emlékezetben nincs benne. A sorszám az "
                    "ARCHÍV ÖSSZEFOGLALÓK listából."),
    "input_schema": {"type": "object", "properties": {"sorszam": {"type": "integer"}}, "required": ["sorszam"]},
}


def make_old_summary_handler(summaries):
    old = archived_summaries(summaries)
    def read_old(args):
        try:
            s = old[int(args.get("sorszam")) - 1]
        except (TypeError, ValueError, IndexError):
            raise ValueError(f"Nincs ilyen sorszámú archív összefoglaló (1–{len(old)}).")
        return f"Archív összefoglaló – {fmt_dt(s['created_at'])}\n{s['description']}\n\n{s['content']}"
    return {"read_old_summary": read_old}


def archive_for_prompt(summaries):
    old = archived_summaries(summaries)
    if not old:
        return ""
    lines = [f"{n}. {fmt_dt(s['created_at'])[:10]} – {' '.join(str(s.get('description') or '').split())[:160]}"
             for n, s in enumerate(old, 1)]
    return ("\n\nARCHÍV ÖSSZEFOGLALÓK (tartalmuk az összevont emlékezetben van; ha egy régi részlet mégis kell, "
            "a read_old_summary eszközzel olvasd be):\n" + "\n".join(lines))


def summaries_for_prompt(summaries):
    if not summaries:
        return ""
    parts = [f"### Összefoglaló – {fmt_dt(s['created_at'])}\n{s['content']}" for s in summaries]
    return (
        "\n\nKORÁBBI ÖSSZEFOGLALÓK:\n"
        "A projekt korábbi beszélgetéseit a felhasználó összefoglaltatta veled, majd törölte; az alábbi "
        "összefoglalók tartalmazzák, ami azokból megmaradt (időrendben). Tekintsd ezeket a saját korábbi "
        "ismereteidnek, építs rájuk, és ha a felhasználó kéri, mutasd meg őket.\n\n" + "\n\n".join(parts)
    )


# --- FELADATOK: ÁLLAPOTOK, SARA ESZKÖZEI ---------------------------------------
# Az állapot minden feladatnál választható. A pipa ennek gyorsgombja: pipa = lezárva (alapból „kész”),
# pipa nélkül a feladat aktív (nyitott vagy választ várunk).
TASK_STATUS_LABELS = {"open": "nyitott", "waiting": "választ várunk", "done": "kész",
                      "obsolete": "elavult", "cancelled": "törölt"}
TASK_ACTIVE = ("open", "waiting")           # ezek látszanak alapból
TASK_ORDER = ("open", "waiting", "done", "obsolete", "cancelled")   # a választólista sorrendje
NO_TOPIC = "Egyéb"

TASK_TOOLS = [
    {
        "name": "add_task",
        "description": "Új tennivaló felvétele a projekt Feladatok listájára. Csak akkor hívd meg, ha a "
                       "felhasználó kifejezetten kéri a felvételt.",
        "input_schema": {
            "type": "object",
            "properties": {
                "topic": {"type": "string", "description": "A téma rövid neve (pl. Vízhálózat). Ha van már "
                                                           "illő téma a listán, pontosan azt használd."},
                "title": {"type": "string", "description": "A tennivaló egy mondatban (pl. Gépésznek megírni a specifikációt)."},
                "note": {"type": "string", "description": "Rövid megjegyzés, ha van (nem kötelező)."},
            },
            "required": ["topic", "title"],
        },
    },
    {
        "name": "update_task",
        "description": "Meglévő feladat állapotának, megjegyzésének, szövegének vagy témájának módosítása. "
                       "Csak akkor hívd meg, ha a felhasználó kifejezetten kéri.",
        "input_schema": {
            "type": "object",
            "properties": {
                "task_id": {"type": "integer", "description": "A feladat azonosítója a listából (a # utáni szám)."},
                "status": {"type": "string", "enum": list(TASK_STATUS_LABELS),
                           "description": "open = nyitott (nálunk a labda), waiting = választ várunk, done = kész, "
                                          "obsolete = elavult, cancelled = törölt (nem kell elvégezni)."},
                "note": {"type": "string"},
                "title": {"type": "string"},
                "topic": {"type": "string"},
            },
            "required": ["task_id"],
        },
    },
]


def task_status(task):
    return task.get("status") if task.get("status") in TASK_STATUS_LABELS else "open"


def task_topic(task):
    return (task.get("topic") or "").strip() or NO_TOPIC


def task_line(task):
    note = f" (megjegyzés: {task['note']})" if task.get("note") else ""
    return f"#{task['id']} [{TASK_STATUS_LABELS[task_status(task)]}] {task_topic(task)} » {task['title']}{note}"


def match_topic(topic, tasks):
    """Ha már van ilyen téma (kis- és nagybetűtől függetlenül), annak az írásmódját használjuk."""
    topic = str(topic or "").strip()
    for t in tasks:
        if task_topic(t).lower() == topic.lower():
            return task_topic(t)
    return topic or NO_TOPIC


def tasks_for_prompt(tasks):
    if tasks is None:  # a feladatok táblája még nincs létrehozva: Sara nem kap róla sem leírást, sem eszközt
        return ""
    listing = "\n".join(task_line(t) for t in tasks[-200:]) or "(még nincs feladat)"
    return (
        "\n\nFELADATOK:\n"
        "A projekthez feladatlista tartozik (Feladatok fül), két szinten: téma » tennivaló. Állapotok: nyitott "
        "(nálunk a labda), választ várunk, kész, elavult, törölt (nem kell elvégezni; az okot a megjegyzés őrzi). "
        "Feladatot felvenni (add_task) vagy módosítani (update_task) csak akkor szabad, ha a felhasználó "
        "kifejezetten kéri. Ha a beszélgetésben teendő merül fel, magadtól ne vedd fel, csak ajánld fel egy "
        "mondatban, a téma és a tennivaló megnevezésével. Soha ne állítsd, hogy egy feladatot felvettél vagy "
        "módosítottál, ha az eszköz ezt nem igazolta vissza. Feladatot törölni nem tudsz; azt a felhasználó "
        "teheti meg a Feladatok fülön. A jelenlegi lista:\n" + listing
    )


def make_task_handlers(project_id, tasks, events):
    """Sara feladatkezelő eszközeinek végrehajtói. Az events listába kerül, ami ténylegesen megtörtént."""
    by_id = {t["id"]: t for t in tasks}

    def add(inp):
        title = str(inp.get("title") or "").strip()
        if not title:
            raise ValueError("A tennivaló szövege hiányzik.")
        topic = match_topic(inp.get("topic"), list(by_id.values()))
        row = db.add_task(project_id, topic, title, str(inp.get("note") or ""))
        by_id[row["id"]] = row
        events.append(f"felvéve: {topic} » {title}")
        return f"Felvéve a Feladatok listájára: {task_line(row)}"

    def update(inp):
        try:
            task_id = int(inp.get("task_id"))
        except (TypeError, ValueError):
            raise ValueError("Hiányzik a feladat azonosítója (task_id).")
        if task_id not in by_id:
            raise ValueError("Ebben a projektben nincs ilyen azonosítójú feladat.")
        fields = {}
        if inp.get("status") is not None:
            if inp["status"] not in TASK_STATUS_LABELS:
                raise ValueError("Ismeretlen állapot.")
            fields["status"] = inp["status"]
        if isinstance(inp.get("note"), str):
            fields["note"] = inp["note"].strip()
        if isinstance(inp.get("title"), str) and inp["title"].strip():
            fields["title"] = inp["title"].strip()
        if isinstance(inp.get("topic"), str) and inp["topic"].strip():
            fields["topic"] = match_topic(inp["topic"], list(by_id.values()))
        if not fields:
            raise ValueError("Nincs megadva, mit kell módosítani.")
        row = db.update_task(task_id, **fields)
        by_id[task_id] = row
        events.append(f"módosítva: {task_topic(row)} » {row['title']} ({TASK_STATUS_LABELS[task_status(row)]})")
        return f"Módosítva: {task_line(row)}"

    return {"add_task": add, "update_task": update}


def personal_context():
    """A felhasználó személyes adatai Sara számára (név, megszólítás, szakértők). A nyilvános kód helyett a
    Streamlit titkos beállításaiban (SARA_SZEMELY) van, így név és más személyes adat nem kerül a kódtárba."""
    try:
        text = str(st.secrets.get("SARA_SZEMELY", "")).strip()
    except Exception:
        text = ""
    prof = profile_text()
    text = "\n".join(x for x in (text, prof) if x)
    return f"A FELHASZNÁLÓRÓL:\n{text}\n\n" if text else "\n"


PROFILE_FIELDS = [("nev", "Név (ahogy az iratokba kerül)"), ("megszolitas", "Megszólítás (ahogy Sara szólít)"),
                  ("email", "E-mail"), ("telefon", "Telefon"), ("cim", "Cím (levelezési)"),
                  ("egyeb", "Egyéb, amit Sara tudjon rólad (pl. aláírási szokás)")]


@st.cache_data(ttl=300, show_spinner=False)
def load_profile():
    if not db.has_records():
        return {}
    try:
        return db.get_setting("profil", {}) or {}
    except db.DbError:
        return {}


def profile_text():
    """A saját profil Sarának (név, elérhetőség, tisztségek a képviselettel). Stabil, ezért a rendszerüzenetbe kerül."""
    p = load_profile()
    if not p:
        return ""
    lines = [f"{label.split(' (')[0]}: {p[key]}" for key, label in PROFILE_FIELDS if str(p.get(key) or "").strip()]
    roles = [r for r in (p.get("tisztsegek") or []) if str(r.get("szervezet") or "").strip()]
    if roles:
        lines.append("Tisztségei (iratok aláírásához): " + "; ".join(
            f"{r.get('szervezet')} – {r.get('tisztseg') or '?'}"
            + (f" ({r['kepviselet']} képviselet" + (f", együtt: {r['egyutt']}" if r.get("egyutt") else "") + ")"
               if r.get("kepviselet") else "") for r in roles)
                     + ". Ha egy irat egy szervezet nevében készül, az aláírási blokk: „<szervezet> képviseletében: "
                       "<név> <tisztség>”; együttes képviseletnél jelezd, hogy kell a másik aláíró is.")
    return "\n".join(lines)


def render_profile():
    """Saját profil: az adatok, amelyeket Sara az iratokhoz és a megszólításhoz használ."""
    st.caption("Sara ezekből tudja, kinek dolgozik, és hogyan írhatod alá a cégeid nevében készülő iratokat. "
               "Jelszót, bankkártya- vagy okmányszámot ide ne írj; a személyes okmányadatok a saját személyi "
               "törzsaktádba valók.")
    if not db.has_records():
        st.warning("A profilhoz futtasd le a Supabase SQL Editorban a schema.sql-t (az 1/e rész az új).")
        return
    p = dict(load_profile())
    with st.form("profile_form", border=False):
        vals = {key: st.text_input(label, value=str(p.get(key) or "")) for key, label in PROFILE_FIELDS}
        st.markdown("**Tisztségek**")
        rows = p.get("tisztsegek") or []
        edited = st.data_editor(
            rows or [{"szervezet": "", "tisztseg": "", "kepviselet": "önálló", "egyutt": ""}],
            num_rows="dynamic", use_container_width=True, key="profile_roles",
            column_config={"szervezet": st.column_config.TextColumn("Cég / szervezet"),
                           "tisztseg": st.column_config.TextColumn("Tisztség"),
                           "kepviselet": st.column_config.SelectboxColumn("Képviselet", options=["önálló", "együttes"]),
                           "egyutt": st.column_config.TextColumn("Együttes képviseletnél kivel")})
        if st.form_submit_button(":material/save: Mentés", type="primary"):
            data = {k: v.strip() for k, v in vals.items()}
            data["tisztsegek"] = [dict(r) for r in (edited if isinstance(edited, list) else edited.to_dict("records"))
                                  if str(r.get("szervezet") or "").strip()]
            try:
                db.set_setting("profil", data)
                load_profile.clear()
                st.success("A profil elmentve. Sara a következő kérdéstől ezt használja.")
            except db.DbError as e:
                st.error(f"A mentés nem sikerült: {e}")


def build_system(summaries, tasks=None, project_docs=None):
    # Sara szerepe: üzleti szemléletű tanácsadó, aki a jogot és az adózást eszközként használja
    return (
        ("Jeannak hívnak: ma te helyettesíted Sarát, ugyanabban a szerepben és ugyanazokkal a "
         "szabályokkal. Ha egy feladat láthatóan meghaladja a képességeidet (bonyolult elemzés, hosszú szerződés, "
         "kényes döntés-előkészítés), mondd meg egy mondatban, hogy ezt inkább Sara csinálja (oldalsáv: Ki "
         "dolgozik?). " if ASSISTANT == "JEAN" else
         "Sarának hívnak. Van egy olcsóbb helyettesed, Jean (egyszerűbb modell): a felhasználó "
         "az oldalsávban választja ki, ki dolgozik, ezért a beszélgetés korábbi válaszai közül néhányat ő írhatott. "
         "Ha egy korábbi válasz hibás vagy ellentmondásos, javítsd ki nyugodtan. ") +
        "A felhasználó személyes üzleti asszisztense és tanácsadója vagy. "
        "Vegyes projekteken dolgozik veled: pénzügyi, befektetési, műszaki, jogi, adózási, HR és egyéb üzleti "
        "kérdésekben. A dolgod, hogy előkészítsd a döntéseit és előre vidd a munkáját; a döntést mindig ő hozza meg.\n"
        + personal_context() +
        "Képes vagy képek, képernyőképek és dokumentumok elemzésére is. Ha a felhasználó képet vagy fájlt küld, "
        "pontosan elemezd, és arra válaszolj, amit kérdez.\n\n"
        "KINEK AZ OLDALÁN ÁLLSZ:\n"
        "1. Mindig a felhasználó és a cégei oldaláról nézd az ügyet. Mielőtt ajánlaton, szerződésen, levélen vagy más "
        "dokumentumon dolgozol, gondold végig, ki kinek mit ad: ki az ajánlattevő és ki a megrendelő, ki írja és ki "
        "kapja. Ha a felhasználó valakitől ajánlatot vár, akkor nem te készíted el azt az ajánlatot, hanem például az "
        "ajánlatkérést, az elvárások listáját vagy a beérkezett ajánlat értékelését.\n"
        "2. A szerepeket magad állapítsd meg a beszélgetésből, a projekt dokumentumaiból és összefoglalóiból; ne a "
        "felhasználónak kelljen elmagyaráznia. Csak akkor kérdezz rá, ha tényleg nem dönthető el, és rossz irányba "
        "indulva fölösleges munka lenne belőle; ilyenkor is egyetlen rövid kérdéssel.\n\n"
        "MI A STÁTUSZOD:\n"
        "1. Nem vagy a cégcsoport tagja: nem vagy ügyvezető, tulajdonos vagy kuratóriumi tag, és nem vagy ügyvéd, "
        "könyvelő vagy könyvvizsgáló sem. Jogi döntést, aláírást, hatósági bejelentést nem végezhetsz.\n"
        "2. Javaslatot teszel, de a döntést nem veszed át. Ahol jogi vagy adózási döntés múlik a dolgon, jelezd, hogy "
        "szakértőt kell bevonni (ügyvédet, könyvelőt vagy könyvvizsgálót).\n\n"
        "HOGYAN GONDOLKODJ:\n"
        "1. Arra válaszolj, amit a felhasználó kérdezett, és maradj a témánál. Ne nyiss új témát, ne vegyél fel új "
        "szálat, ne vidd el a beszélgetést más irányba. A beszélgetést ő vezeti, te segítesz neki.\n"
        "2. Üzleti szemmel nézd a kérdést: mi a cél, mi a döntés, mennyibe kerül, mi a kockázat. A jog és az adózás "
        "eszköz; jogi részletet csak ott hozz, ahol a döntést befolyásolja, vagy ha a kérdés kifejezetten jogi.\n"
        "3. Ahol van adat, számolj; ahol kell, foglalj állást, és mondd meg, mit javasolsz.\n"
        "4. Ötleteket, alternatívákat, további lehetőségeket csak akkor sorolj, ha a felhasználó kéri; dokumentumot, "
        "táblázatot vagy prezentációt is csak kérésre készíts. Kéretlenül "
        "legfeljebb egy rövid mondatban jelezz valamit, és csak akkor, ha komoly baj (pénzügyi vagy jogi kockázat, "
        "határidő) lenne belőle, hogy nem szólsz.\n\n"
        "HOGYAN VÁLASZOLJ:\n"
        "1. Röviden. Alapból néhány mondat, ami a kérdésre kell, se több, se kevesebb. Hosszabb kifejtést, részletes "
        "elemzést, listát vagy táblázatot csak akkor adj, ha a felhasználó kéri, vagy a kérdés nélküle "
        "megválaszolhatatlan.\n"
        "2. A válasz az első mondatban legyen. Ne ismételd meg a kérdést, ne vezesd fel, ne foglald össze a végén, "
        "és ne kérdezd meg minden válasz végén, hogy segíthetsz-e még.\n"
        "3. Hangnem: laza, közvetlen, tegező, mint egy jó munkatárs, akivel régóta együtt dolgoztok. Belefér egy kis "
        "humor is, ha a helyzet engedi, de a tartalom rovására soha. Hivatalos hangnemre akkor válts, ha a "
        "felhasználó kéri, vagy hivatalos dokumentumot, levelet készítesz.\n"
        "4. Természetes, élő magyar nyelven írj, tükörfordítások, kancelláriai fordulatok és töltelékszavak nélkül.\n"
        "5. Ha valamit nem tudsz, vagy ellenőrizni kell, mondd meg egyenesen.\n\n"
        "INTERNETES KERESÉS:\n"
        f"A mai dátum: {date.today().isoformat()}. Van internetes keresési lehetőséged: elsőként a research_web "
        "eszközzel kérd meg a kereső kollégát, ő keres és a releváns idézeteket adja vissza forrással; a saját "
        "web_search eszközöddel csak ellenőrizz vagy pótolj, ha az ő eredménye nem elég. "
        "A tudásod egy jóval korábbi időpontig tart, ezért a hatályos jogszabályokra, adómértékekre, határidőkre, "
        "díjakra, hatósági szabályokra és minden aktuális adatra keress rá, mielőtt válaszolsz, még ha biztosnak is érzed magad. "
        "Magyar jogi és adózási kérdésekben elsősorban hivatalos forrásokat használj (pl. njt.hu, nav.gov.hu, magyarkozlony.hu). "
        "A válaszban jelöld meg, mely forrásokra támaszkodtál.\n\n"
        "DOKUMENTUMKÉSZÍTÉS:\n"
        "Ha a felhasználó dokumentum, táblázat vagy prezentáció készítését kéri, használd a megfelelő eszközt: "
        "create_document (szöveges dokumentum, amelyet a felhasználó Word, PDF vagy Markdown formátumban tölt le), "
        "create_spreadsheet (Excel táblázat, amelyben a számolt értékek képletek, a számok pedig megfelelő "
        "számformátumúak legyenek, hogy a felhasználó tovább tudjon dolgozni vele), "
        "create_presentation (PowerPoint). Egyszerű kérdésre ne készíts fájlt, arra a chatben válaszolj. "
        "MIELŐTT FÁJLT KÉSZÍTESZ, nézd át, megvan-e minden adat, ami a dokumentum érdemi részéhez kell (a "
        "beszélgetésben, az összefoglalókban és a projekt dokumentumaiban). Ha valami lényeges hiányzik, még ne "
        "készítsd el a fájlt: egyetlen rövid üzenetben, pontokba szedve kérd el, ami kell, és ahol tudsz, javasolj "
        "alapértéket (például: „Határidő? Ha nem mondasz mást: 30 nap.”). A fájl csak a válasz után készüljön el, "
        "kész állapotban. Apróságra (például dátum, aláíró neve) ne kérdezz rá: azt jelöld a dokumentumban "
        "kitöltendő helyként [szögletes zárójelben], és a válaszod végén egy mondatban sorold fel. Ha a felhasználó "
        "kifejezetten tervezetet kér, vagy azt mondja, hogy készüljön el így is, ne kérdezz, hanem a hiányzó "
        "részeket jelöld kitöltendőként. "
        "Fájl kizárólag az eszköz tényleges meghívásával jön létre: a kért dokumentum teljes szövegét ne a "
        "chatbe írd, hanem az eszköznek add át, és soha ne állítsd, hogy egy fájl elkészült, ha az eszköz "
        "ezt nem igazolta vissza. Nagyon hosszú anyagot bonts több fájlra, válaszonként egy eszközhívással. "
        "A fájl elkészítése után röviden írd le, mit tartalmaz és mire kell figyelni; a letöltőgombokat az "
        "alkalmazás teszi a válaszod alá, te ne írj letöltési hivatkozást vagy „Elkészült” sort. "
        "A korábban elkészült fájlok a projekt dokumentumai között vannak; ha a tartalmuk kell, olvasd be őket "
        "a read_project_document eszközzel (csak azt, amelyik tényleg kell); módosítást új fájl készítésével "
        "végezz, és jelezd, hogy új változat készült. "
        "ADATKERESÉS: ha egy adat kell (cégadat, összeg, határidő, feltétel, elérhetőség), amely nincs előtted, "
        "először a find_information eszközzel kérd a kereső kollégát: ő az aktuális projekt és a többi projekt "
        "dokumentumaiban is keres, és szó szerinti idézetet ad vissza. Teljes dokumentumot csak akkor olvass be, ha "
        "a teljes szöveg kell. Ha az adat egy másik projektből származik, ezt mondd meg a felhasználónak. "
        "Ha több, egymástól független adatra van szükséged (például egy cégadatra a dokumentumtárból és egy "
        "jogszabályra a webről, vagy két különböző dokumentumra), az eszközöket egyetlen válaszban, egyszerre hívd "
        "meg: ezek párhuzamosan futnak, így gyorsabb. Egy kérdésen belül legfeljebb 3 keresést kérhetsz a "
        "kollégától. Ha nem talál használhatót, ne fogalmazd át újra és újra: válaszolj azzal, amit tudsz, és kérd "
        "meg a felhasználót, hogy pontosítson vagy adjon több információt. "
        "JOGSZABÁLYOK: a hatályos magyar jogszabályszöveget az njt_jogszabaly eszközzel kéred le a Nemzeti "
        "Jogszabálytárból (paragrafusra vagy kulcsszóra). Ha jogszabályhelyre hivatkozol, előbb kérd le, és a lekért "
        "szöveget idézd a forrás linkjével; ami nincs lekérve, azt jelöld ellenőrizendőnek. A kérésbe kizárólag "
        "jogszabályt, paragrafust és jogi fogalmat írj, soha nevet, céget, címet vagy ügyadatot. "
        "RED FLAG: ha a felhasználó egy szerződés vagy okirat (pl. a másik fél tervezete) átnézését, kockázatait "
        "kéri, a red_flag_review eszközt használd; előbb tisztázd, kinek az oldalán áll (melyik cége, milyen "
        "szerepben), ha nem egyértelmű. VERZIÓK: két változat összevetésére a compare_versions eszköz; "
        "változáskövetéses Word-iratot (redline) csak kérésre készíts a create_redline_document eszközzel. "
        "ADÓSZÁM: cég vagy egyéni vállalkozó adószámának érvényességét az adoszam_ellenorzes eszközzel nézd meg a "
        "NAV-nál, szerződés előtt mindig. "
        "JOGI HATÁR: nem vagy ügyvéd. A jogszabálykeresés, a Red Flag és az összehasonlítás arra való, hogy a "
        "felhasználó felkészülten menjen az ügyvédhez, nem arra, hogy kiváltsa: fontos szerződésnél, vitás "
        "kérdésnél vagy aláírás előtt javasold, hogy egyeztessen ügyvéddel, és foglald össze, mit érdemes "
        "megkérdeznie. "
        "TÖRZSAKTÁK: a felhasználó tartós adatai törzsaktákban vannak (cég, ingatlan, személy, jármű), mindegyik élén "
        "egy adatlappal. A projekthez kapcsolt törzsakták adatlapját az [AKTUÁLIS ÁLLAPOT] részben kapod; iratokhoz "
        "(meghatalmazás, nyilatkozat, járműhasználati engedély) ezekből vedd az adatokat, szó szerint. Ha egy kellő "
        "adat hiányzik, mondd meg, melyik törzsaktába mit kellene feltölteni (pl. friss cégkivonat, forgalmi "
        "engedély), vagy hogy melyik törzsaktát kellene a projekthez kapcsolni. Személyes adatot csak a kapcsolt "
        "törzsaktából vagy a felhasználótól használj. "
        "IRATMINTÁK: ha a felhasználó egy iratot kér, amelyhez van minta az Iratmintákban (meghatalmazás, "
        "járműhasználati engedély, nyilatkozat, igazolás), az irat_mintabol eszközzel nyisd meg az űrlapot; ne írd "
        "meg magad. Ha nincs hozzá minta, készítsd el a create_document eszközzel, és ajánld fel, hogy mintát "
        "csinálhattok belőle (jelölésekkel). "
        "DOKUMENTUMOK RENDJE: a projekt dokumentumait kategóriánként látod; az elavultnak jelöltek régi változatok, "
        "ezekből csak kifejezett kérésre vagy két változat összevetésekor dolgozz, minden más kérdésnél az érvényes "
        "változatból. Ha észreveszed, hogy egy dokumentum egy másik újabb változata, ajánld fel a régi elavultnak "
        "jelölését; a set_document_status eszközt csak a felhasználó kérésére használd. "
        "DOKUMENTUMOK ÖSSZHANGJA: ha a felhasználó azt kéri, hogy több dokumentum adatait hozd szinkronba, "
        "egyeztesd, vagy keresd meg az ellentmondásokat (verziószám, dátum, határidő, összeg, felelős, műszaki adat), "
        "a consistency_check eszközzel kérd az ellenőrző kollégát (a kijelölt dokumentumokra, vagy a projekt "
        "összesére; ha a felhasználó megnevezi a mértékadót, add meg reference-ként). A jelentést mutasd meg; "
        "javítani csak jóváhagyás után javíts, .docx-nél dokumentumonként az edit_word_document eszközzel, "
        "követett változással, hogy a felhasználó lássa, mi változott. "
        "TÖBB KÉRDÉS EGYSZERRE: ha a felhasználótól több dolgot kell megkérdezned (tisztázandó pontok, hiányzó "
        "adatok), számozott listában tedd fel őket, pontonként egy kérdéssel, és a végén egy mondatban írd: „A "
        "Megjegyzések gombbal kérdésenként válaszolhatsz.” Ne kérj kódolt választ (pl. „1b, 2 ok”). Ha a "
        "felhasználó űrlapot kér, ugyanezt csináld: az alkalmazás a válaszod alatti Megjegyzések gombbal minden "
        "pont alá szövegmezőt tesz, és a válaszait egyben küldi vissza. "
        "MEGLÉVŐ WORD-DOKUMENTUM (.docx) javítását, kiegészítését viszont az edit_word_document eszközzel végezd, "
        "ne írd újra create_document-tel: így az eredeti formázás megmarad, és a módosítások követett "
        "változásként látszanak. Előtte olvasd be a dokumentum szövegét, ha még nincs előtted. Ha a felhasználó "
        "tiszta (követés nélküli) változatot kér, track_changes=false. Word-megjegyzést (komment) is tudsz "
        "fűzni egy szövegrészhez a comment művelettel. Ha a felhasználó új fájlnevet kér (például a mai dátummal), "
        "azt az output_title-ben add meg. "
        "Ha egy állítás bizonytalan vagy ellenőrzést igényel, a dokumentumban is jelöld.\n\n"
        "MUNKAANYAGOK:\n"
        "A felhasználó a kérdéseihez képet vagy fájlt csatolhat. Ezek munkaanyagok: a beszélgetés részei, de "
        "nem kerülnek a projekt dokumentumtárába, és az előzmények törlésekor törlődnek."
        + ("\n\nAZ PROJEKT ÁLLAPOTA: a feladatlistát és a dokumentumlistát (kivonatokkal) a felhasználó "
           "legutóbbi üzenetének elején kapod, [AKTUÁLIS ÁLLAPOT] jelöléssel; mindig az a friss, mérvadó állapot.")
        + summaries_for_prompt(active_summaries(summaries)) + archive_for_prompt(summaries)
    )


def summary_system(summaries):
    return (
        "Te Sara vagy, üzleti szemléletű tanácsadó és munkatárs. Most nem kérdésre "
        "válaszolsz, hanem a saját, felhasználóval folytatott beszélgetésedről írsz pontos, tömör, de "
        "hiánytalan összefoglalót, kifogástalan magyar nyelven.\n\n"
        "ÖSSZEVONT EMLÉKEZET: az összefoglalód a korábbi összefoglalókat (lent) és a mostani beszélgetést EGYBEN "
        "tartalmazza, mert a régieket ezután nem kapod meg (archiválódnak). Vedd át belőlük mindazt, ami még "
        "érvényes; amit egy későbbi döntés felülírt, azt csak a végső állapotával írd le (ha fontos, egy félmondat "
        "a változásról). Szerkezet: a cél és a háttér röviden; döntések és érvényes adatok (számok, határidők, "
        "nevek, felelősök) témánként; nyitott kérdések és következő lépések; mely dokumentum mit tartalmaz. "
        f"Tömören: legfeljebb kb. {MEMORY_WORDS} szó; a részletek a dokumentumokban vannak, oda hivatkozz."
        + summaries_for_prompt(active_summaries(summaries))
    )


# --- PROJEKTMŰVELETEK --------------------------------------------------------
def open_project(project_id):
    """Projekt megnyitása; a megnyitás idejét eltároljuk, hogy legközelebb ez induljon."""
    st.session_state["current_project_id"] = project_id
    st.session_state["view"] = "project"
    try:
        db.update_project(project_id, last_opened_at=db.now_iso())
    except db.DbError:
        pass  # ha nem sikerül eltárolni, a projekt attól még megnyílik


def top_sort_order(projects):
    orders = [p["sort_order"] for p in projects if p.get("sort_order") is not None]
    return (min(orders) if orders else 0) - 1


def move_project(active, project_id, delta):
    """A projekt eggyel feljebb (-1) vagy lejjebb (+1) kerül; a sorrendet sorszámként tároljuk."""
    ids = [p["id"] for p in active]
    i = ids.index(project_id)
    j = i + delta
    if not 0 <= j < len(ids):
        return
    ids[i], ids[j] = ids[j], ids[i]
    current = {p["id"]: p.get("sort_order") for p in active}
    for n, pid in enumerate(ids):
        if current[pid] != n:
            db.update_project(pid, sort_order=n)


def dialog(title, on_dismiss, **kwargs):
    """st.dialog; az ablak X-szel bezárásakor az on_dismiss törli a „nyitva van” jelzést (ahol a Streamlit tudja)."""
    if "on_dismiss" in inspect.signature(st.dialog).parameters:
        kwargs["on_dismiss"] = on_dismiss
    return st.dialog(title, **kwargs)


def _close_delete_dialog():
    st.session_state.pop("delete_open", None)


@dialog("Projekt végleges törlése", _close_delete_dialog)
def delete_dialog(project):
    st.write(f"Biztosan törlöd a(z) **{project['name']}** projektet?")
    st.warning("A törlés végleges, és mindent visz: a beszélgetést, a csatolt munkaanyagokat, "
               "a dokumentumtár összes fájlját és az összefoglalókat.")
    c1, c2 = st.columns(2)
    if c1.button(":material/delete: Igen, végleges törlés", type="primary", key="delete_yes", **STRETCH):
        try:
            db.delete_project(project["id"])
        except db.DbError as e:
            st.error(f"A törlés nem sikerült: {e}")
            return
        _close_delete_dialog()
        st.session_state.pop("current_project_id", None)
        st.rerun()
    if c2.button("Mégse", key="delete_no", **STRETCH):
        _close_delete_dialog()
        st.rerun()


# --- ÖSSZEFOGLALÁS ÉS ELŐZMÉNYEK TÖRLÉSE --------------------------------------
def _close_compact_dialog():
    st.session_state.pop("compact_open", None)


@dialog("Összefoglalás és előzmények törlése", _close_compact_dialog, width="large")
def compact_dialog(project, messages, atts_by_msg, summaries, project_docs):
    pid = project["id"]
    message_ids = [m["id"] for m in messages if m.get("id") is not None]
    signature = (pid, len(messages), messages[-1].get("id"))
    draft = st.session_state.get("compact_draft")
    n_atts = sum(len(v) for v in atts_by_msg.values())

    def generate(fix_request="", previous=""):
        return chat_flow.summarize(
            claude_client, MODEL, summary_system(summaries), build_history(messages, atts_by_msg),
            doc_names=[d["name"] for d in project_docs], fix_request=fix_request, previous_draft=previous,
            usage_cb=usage_logger(pid, "osszefoglalo"))

    if not draft or draft["signature"] != signature:
        with st.spinner(f"{ASSISTANT_NAME} összefoglalja a beszélgetést…"):
            try:
                description, content = generate()
            except Exception as e:
                st.error(f"Az összefoglaló nem készült el: {e}")
                st.caption("Semmi nem törlődött. Zárd be az ablakot, és próbáld meg újra.")
                return
        draft = {"signature": signature, "description": description, "content": content, "nonce": 0}
        st.session_state["compact_draft"] = draft

    st.caption(
        f"Ez Sara tervezete. Nézd át, és ha kell, írd át. Jóváhagyás után a projekt {len(messages)} üzenete"
        + (f" és {n_atts} csatolt munkaanyaga" if n_atts else "")
        + " véglegesen törlődik az adatbázisból, és Sara ezután az összefoglalókból dolgozik. "
          "A dokumentumtár nem változik.")
    nonce = draft["nonce"]
    description = st.text_area("Rövid leírás (2–3 mondat, ez látszik a listában)", value=draft["description"],
                               height=90, key=f"cd_desc_{nonce}")
    content = st.text_area("Összefoglaló (Markdown)", value=draft["content"], height=360, key=f"cd_content_{nonce}")
    with st.expander("Előnézet formázva"):
        st.markdown(content)
    fix_request = st.text_input("Mit javítson rajta Sara? (csak az újraíráshoz kell)", key=f"cd_fix_{nonce}",
                                placeholder="pl. írd bele a három leányvállalat pontos számait is")

    c1, c2, c3 = st.columns([2.2, 1.6, 1])
    if c1.button(":material/check: Jóváhagyás és előzmények törlése", type="primary", key="cd_approve", **STRETCH):
        if not content.strip():
            st.error("Az összefoglaló üres, így nem menthető.")
            return
        try:
            db.add_summary(pid, MEMORY_MARK + " " + description.strip().removeprefix(MEMORY_MARK).strip(),
                           content.strip())
        except db.DbError as e:
            st.error(f"Az összefoglaló mentése nem sikerült, ezért semmi nem törlődött: {e}")
            return
        try:
            db.delete_messages(pid, message_ids)
        except db.DbError as e:
            st.session_state["chat_flash"] = ("error", "Az összefoglaló elmentve, de az előzmények törlése "
                                                       f"nem sikerült: {e}")
        else:
            st.session_state["chat_flash"] = ("success", "Az összefoglaló elkészült, az előzmények törölve. "
                                                         "Az Összefoglalók fülön megnézheted és szerkesztheted.")
        st.session_state.pop("compact_draft", None)
        _close_compact_dialog()
        st.rerun()
    if c2.button(":material/refresh: Újraírás Sarával", key="cd_rewrite", **STRETCH):
        with st.spinner("Sara újraírja az összefoglalót…"):
            try:
                new_description, new_content = generate(
                    fix_request.strip() or "Írd meg újra, alaposabban és pontosabban.",
                    f"LEÍRÁS: {description}\n{chat_flow.SUMMARY_SEPARATOR}\n{content}")
            except Exception as e:
                st.error(f"Az újraírás nem sikerült: {e}")
                return
        st.session_state["compact_draft"] = {"signature": signature, "description": new_description,
                                             "content": new_content, "nonce": nonce + 1}
        st.rerun()
    if c3.button("Mégse", key="cd_cancel", **STRETCH):
        # a kézzel átírt szöveget megőrizzük, hogy újranyitáskor ne vesszen el
        st.session_state["compact_draft"] = {**draft, "description": description, "content": content,
                                             "nonce": nonce + 1}
        _close_compact_dialog()
        st.rerun()


def conversation_size_note(pid, messages):
    """A beszélgetés mérete az utolsó kérdés mért adataiból: mennyi ment ki Sarának, és mennyibe került."""
    last_q = next((m for m in reversed(messages) if m.get("role") == "user" and m.get("created_at")), None)
    if not last_q:
        return 0
    try:
        rows = db.get_project_usage(pid, last_q["created_at"])
    except Exception:
        return 0   # a napló még nincs meg: nincs mit kijelezni
    if not rows:
        return 0
    size = max(int(r.get("input_tokens") or 0) + int(r.get("cache_write_tokens") or 0)
               + int(r.get("cache_read_tokens") or 0) for r in rows)
    cost = sum(costs.row_cost(r)[0] for r in rows)
    text = (f"Az utolsó kérdés: Sara {size / 1000:,.0f} ezer tokent kapott ({len(rows)} lépésben), "
            f"költsége kb. ${cost:,.2f}.")
    st.caption(f":orange[{text}]" if size > CLOSE_SUGGEST_TOKENS else text)
    return size


CLOSE_SUGGEST_TOKENS = 100_000   # e fölött a válasz végén pirossal javasoljuk a lezárást


def close_suggestion(project, size):
    """A legutóbbi válasz végén, pirossal: a beszélgetés nagy, a lezárással a folytatás jóval olcsóbb. Minden
    válasz végén ott marad, amíg le nem zárjuk (nem kerül a mentett válaszba, így a modell sem látja)."""
    st.markdown(f":red[**:material/warning: A beszélgetés már nagy ({size / 1000:,.0f} ezer token): minden "
                f"további kérdés ennyit visz magával. Javaslom a lezárást: kérd meg {ASSISTANT_ACC}, hogy a döntéseket "
                "vezesse át a dokumentumokba és a feladatokba, aztán „Összefoglalás és törlés”. A folytatás "
                "lényegesen olcsóbb lesz, és semmi fontos nem vész el.**]")
    compact_button(project, f"compact_btn_suggest_{project['id']}")


def compact_button(project, key):
    if st.button(":material/summarize: Összefoglalás és törlés", key=key,
                 help="Sara összefoglalja az eddigi beszélgetést; jóváhagyásod után az üzenetek törlődnek, "
                      "és Sara az összefoglalóból dolgozik tovább."):
        st.session_state["compact_open"] = project["id"]
        st.rerun()


# --- ELŐZMÉNYEK TÖRLÉSE ÖSSZEFOGLALÓ NÉLKÜL -----------------------------------
def _close_trim_dialog():
    st.session_state.pop("trim_open", None)


def trim_points(messages):
    """Azok a helyek (sorszámok), ahonnan a beszélgetés vége törölhető: a felhasználó kérdései, hogy a
    megmaradó rész mindig teljes kérdés–válasz párokból álljon. Ha nincs ilyen, bármelyik üzenet."""
    with_id = [i for i, m in enumerate(messages) if m.get("id") is not None]
    return [i for i in with_id if messages[i].get("role") == "user"] or with_id


def trim_label(msg):
    text = " ".join(str(msg.get("content") or "").split())
    return f"{fmt_dt(msg.get('created_at'))} – {text[:90]}{'…' if len(text) > 90 else ''}"


@dialog("Előzmények törlése összefoglaló nélkül", _close_trim_dialog, width="large")
def trim_dialog(project, messages, atts_by_msg):
    points = trim_points(messages)
    if not points:
        st.info("Nincs törölhető üzenet.")
        return
    st.caption("Válaszd ki, melyik kérdésedtől induljon a törlés. Az a kérdés és minden, ami utána jött, "
               "törlődik; ami előtte volt, megmarad. Összefoglaló nem készül.")
    start = st.selectbox("A törlés ettől a kérdéstől indul:", options=points, index=len(points) - 1,
                         format_func=lambda i: trim_label(messages[i]), key=f"trim_start_{len(messages)}")
    doomed = [m for m in messages[start:] if m.get("id") is not None]
    n_atts = sum(len(atts_by_msg.get(m["id"], [])) for m in doomed)
    kept = start
    st.warning(
        f"Véglegesen törlődik {len(doomed)} üzenet" + (f" és {n_atts} csatolt munkaanyag" if n_atts else "")
        + (f"; megmarad az előtte lévő {kept} üzenet." if kept else "; a beszélgetésből semmi nem marad meg.")
        + " A törlés nem vonható vissza. A dokumentumtár és az összefoglalók nem változnak.")
    c1, c2 = st.columns(2)
    if c1.button(":material/delete: Igen, törlés", type="primary", key="trim_yes", **STRETCH):
        try:
            db.delete_messages(project["id"], [m["id"] for m in doomed])
        except db.DbError as e:
            st.error(f"A törlés nem sikerült: {e}")
            return
        st.session_state["chat_flash"] = ("success", f"{len(doomed)} üzenet törölve.")
        _close_trim_dialog()
        st.rerun()
    if c2.button("Mégse", key="trim_no", **STRETCH):
        _close_trim_dialog()
        st.rerun()


def trim_button(project, key):
    if st.button(":material/content_cut: Törlés összefoglaló nélkül", key=key,
                 help="Egy kiválasztott kérdésedtől a beszélgetés végéig minden törlődik (pl. vakvágány); "
                      "ami előtte volt, megmarad. Összefoglaló nem készül."):
        st.session_state["trim_open"] = project["id"]
        st.rerun()


# --- BESZÉLGETÉS FÜL ---------------------------------------------------------
def show_flash(key):
    flash = st.session_state.pop(key, None)
    if flash:
        getattr(st, flash[0])(flash[1])


def attachment_caption(attachments):
    names = ", ".join(a["name"] for a in attachments)
    return f":material/attach_file: {names} (munkaanyag, nem kerül a dokumentumtárba)"


def _close_comment_mode(msg_key):
    st.session_state.pop("comment_msg", None)
    for k in [k for k in st.session_state if str(k).startswith(f"cm_{msg_key}_")]:
        del st.session_state[k]


def render_comment_mode(pid, msg_key, content):
    """A válasz részenként, mindegyik alatt egy megjegyzésmezővel; a végén egy üzenetben megy Sarának."""
    answer = strip_app_notes(content)
    blocks = split_blocks(answer)
    st.caption("Írd a megjegyzést annak a résznek az aljára, amelyikre vonatkozik; a többit hagyd üresen.")
    for j, block in enumerate(blocks):
        st.markdown(block)
        st.text_area("Megjegyzés ehhez a részhez", key=f"cm_{msg_key}_{j}", label_visibility="collapsed",
                     placeholder="megjegyzés ehhez a részhez…", height=68)
    c_send, c_cancel = st.columns([2, 1])
    if c_send.button(":material/send: Megjegyzések feldolgozása", key=f"cm_send_{msg_key}", type="primary", **STRETCH):
        notes = [(b, st.session_state.get(f"cm_{msg_key}_{j}", "")) for j, b in enumerate(blocks)]
        notes = [(b, n) for b, n in notes if n.strip()]
        if not notes:
            st.warning("Még egyetlen megjegyzést sem írtál.")
            return
        st.session_state.setdefault("pending_question", {})[pid] = comments_message(answer, notes)
        _close_comment_mode(msg_key)
        st.rerun()
    if c_cancel.button("Mégse", key=f"cm_cancel_{msg_key}", **STRETCH):
        _close_comment_mode(msg_key)
        st.rerun()


# --- IRAT MINTÁBÓL: a Word-minták jelölt helyeinek kitöltése űrlappal (Larától átvéve, egyszerűsítve) -------
TPL_SYNTAX = ("Jelölések a Word-mintában: [Meghatalmazó neve] · [Kelt|dátum] · [Vételár|összeg] és [Vételár|betűvel] · "
              "[Időtartam: 1 év / határozatlan] · [Feladatok|hosszú] · szabad blokk külön bekezdésben: [[BLOKK: Egyéb]]. "
              "Sara a mezőket a kapcsolt törzsaktákból (cég, személy, jármű, ingatlan) és a profilodból tölti elő.")

TPL_FORM_TOOL = {
    "name": "irat_mintabol",
    "description": ("Iratkészítés az Iratminták egy Word-mintájából űrlappal (meghatalmazás, járműhasználati "
                    "engedély, nyilatkozat, igazolás stb.): megnyitja az űrlapot, amelyet a program a projekthez "
                    "kapcsolt törzsakták adatlapjaiból és a felhasználó profiljából előtölt; a felhasználó ellenőrzi, és "
                    "a program kitölti a Word-fájlt (a minta szövege változatlan marad). Add meg a minta nevét (vagy "
                    "egyértelmű részét), és ha a felhasználó megmondta, ki kicsoda, az utasitas mezőben add tovább."),
    "input_schema": {"type": "object", "properties": {
        "minta": {"type": "string"},
        "utasitas": {"type": "string", "description": "A felhasználó utasítása a kitöltéshez, szó szerint."}},
        "required": ["minta"]},
}


def get_templates():
    if not templates_project:
        return []
    try:
        return [d for d in db.get_documents(templates_project["id"])
                if d["name"].lower().endswith(".docx") and not is_obsolete(d)]
    except db.DbError:
        return []


def open_template_form(pid, tpl_doc, instructions=""):
    st.session_state.setdefault("tpl_form", {})[pid] = {
        "doc_id": tpl_doc["id"], "name": tpl_doc["name"], "nonce": int(time.time() * 1000) % 10 ** 9,
        "instructions": instructions}


def close_template_form(pid):
    st.session_state.get("tpl_form", {}).pop(pid, None)


def template_context(project, project_docs, selected=None, limit=60000):
    """Az előtöltés anyaga: a saját profil, a kapcsolt törzsakták adatlapjai, és ha a felhasználó kijelölt iratokat,
    azok szövege (különben a projekt dokumentumainak kivonata). A projekt neve szándékosan nincs benne."""
    parts = []
    prof = profile_text()
    if prof:
        parts.append("=== A felhasználó profilja ===\n" + prof)
    linked = linked_records_context(project, all_projects)
    if linked:
        parts.append("=== Kapcsolt törzsakták ===\n" + linked)
    used = sum(len(p) for p in parts)
    if selected:
        for d in selected:
            text = (db.get_document_text(d["id"]) or "").strip()
            if text and used < limit:
                chunk = f"=== {d['name']} ===\n{text[:15000]}"
                parts.append(chunk)
                used += len(chunk)
    else:
        for d in current_docs(project_docs)[:40]:
            s = " ".join(str(d.get("summary") or "").split())
            if s:
                parts.append(f"=== {d['name']} (kivonat) ===\n{s[:1500]}")
    return "\n\n".join(parts)[:limit + 4000]


def _tpl_set(wkey, value):
    st.session_state[wkey] = value


PRINCIPAL_WORDS = ("meghatalmazó", "megbízó", "munkáltató", "kibocsátó", "nyilatkozó", "igazoló", "engedélyező",
                   "eladó", "bérbeadó", "szolgáltató", "cég", "társaság")


def template_paper(values, papers):
    """Az űrlap kitöltött értékei közül a meghatalmazó / szerződő cég fejléces papírja (az első egyező cég, a
    megbízói szerepű mezők előnyben); üres, ha egyik cég sem szerepel."""
    if not papers:
        return ""
    def rank(key):
        k = key.lower()
        return next((i for i, w in enumerate(PRINCIPAL_WORDS) if w in k), len(PRINCIPAL_WORDS))
    for key, value in sorted(values.items(), key=lambda kv: rank(kv[0])):
        if "|" in key or not str(value or "").strip() or len(str(value)) > 120:
            continue
        for rec, paper in papers:
            if tplfill.names_match(value, rec.get("name") or record_label(rec)):
                return paper
    return ""


def render_template_form(project, project_docs, selected=None):
    """Az űrlap: előtöltött mezők a forrással, figyelmeztetések, szabad blokkok, majd a kész Word a projektbe."""
    pid = project["id"]
    form = st.session_state["tpl_form"][pid]
    nonce = form["nonce"]
    head, close = st.columns([5, 1], vertical_alignment="center")
    head.markdown(f":material/edit_document: **Irat mintából:** {form['name']}")
    if close.button("Mégse", key=f"tplf_close_{nonce}", **STRETCH):
        close_template_form(pid)
        st.rerun()
    try:
        data = base64.b64decode(get_document_content(form["doc_id"]))
        if "info" not in form:
            form["info"] = tplfill.scan(data)
    except Exception as e:
        st.error(f"A minta nem tölthető be: {e}")
        return
    info = form["info"]
    fields = tplfill.form_fields(info)
    if not fields and not info["blocks"]:
        st.info("Ebben a mintában nincsenek jelölt kitöltendő helyek. " + TPL_SYNTAX)
        return
    ikey = f"tplf_instr_{pid}_{form['doc_id']}"
    if ikey not in st.session_state:
        st.session_state[ikey] = form.get("instructions", "")
    form["instructions"] = st.text_input(
        "Utasítás a kitöltéshez (nem kötelező):", key=ikey,
        placeholder="pl. a meghatalmazó a Contact-Line 2002 Bt., a meghatalmazott Kovács Anna, a jármű az ABC-123").strip()
    sig = ",".join(sorted(str(d["id"]) for d in (selected or []))) + "|" + form["instructions"]
    if "prefill" in form and form.get("prefill_sig") != sig:
        form.pop("prefill", None)
        form["nonce"] = nonce = int(time.time() * 1000) % 10 ** 9
    if "prefill" not in form:
        form["prefill"], form["prefill_sig"], form["candidates"] = {}, sig, {}
        with st.spinner(f"{ASSISTANT_NAME} előtölti a mezőket…"):
            try:
                form["prefill"] = tplfill.prefill(claude_client, MODEL, form["name"], info,
                                                  template_context(project, project_docs, selected),
                                                  instructions=form["instructions"],
                                                  log=usage_logger(pid, "mintakitoltes"))
                form["candidates"] = dict(tplfill.LAST_CANDIDATES)
                dropped = list(tplfill.LAST_REJECTED)
                form["note"] = (f"{len(form['prefill'])} mezőt töltöttem elő (a forrás a mező alatt)."
                                + (" Nem vettem át, mert a forrásban nem szerepel: " + ", ".join(dropped) + "."
                                   if dropped else ""))
            except Exception as e:
                form["note"] = f"Az előtöltés most nem sikerült ({e}); töltsd ki kézzel."
    st.caption((form.get("note") or "") + " Ami hiányzik, azt a kapcsolt törzsaktákban érdemes pótolni "
               "(a projekt Beállítások fülén: Kapcsolt törzsakták), vagy írd be kézzel.")
    pre = form["prefill"]
    values = {}
    cols = st.columns(2, gap="medium")
    for i, f in enumerate(fields):
        key = tplfill.value_key(f)
        val, src = pre.get(key, ("", ""))
        if not val and f["kind"] == "date" and "kelt" in f["name"].lower():
            val = tplfill.fmt_date("ma")
        wkey = f"tplv_{pid}_{nonce}_{i}"
        box = st if f["kind"] == "long" else cols[i % 2]
        label = f["name"] + (f" ({f['hint']})" if f.get("hint") else "")
        fhelp = f"A mintában: {f['context']}" if f.get("context") else None
        if f["kind"] == "choice":
            opts = ["– válassz –"] + f["options"]
            v = box.selectbox(label, opts, index=opts.index(val) if val in opts else 0, key=wkey, help=fhelp)
            v = "" if (not v or v == opts[0]) else v
        elif f["kind"] == "long":
            v = box.text_area(label, value=val, key=wkey, height=110, help=fhelp)
        elif f["kind"] in ("amount", "words") and f.get("currency"):
            in_box, cur_box = box.columns([3, 1.1], gap="small")
            codes = list(tplfill.CURRENCIES)
            values[key + "|currency"] = cur_box.selectbox(
                "Pénznem", codes, index=codes.index(f["currency"]) if f["currency"] in codes else 0, key=wkey + "_cur")
            v = in_box.text_input(label, value=tplfill.fmt_amount(val) if val else "", key=wkey, help=fhelp)
        else:
            v = box.text_input(label, value=val, key=wkey, help=fhelp,
                               placeholder={"date": "pl. nov 1. / 2026-11-01 / ma"}.get(f["kind"], ""))
        values[key] = v
        warn = tplfill.validate(f["name"], f["kind"], v)
        cands = (form.get("candidates") or {}).get(key)
        if not str(v).strip() and cands:
            box.caption(":orange[Több lehetőség – melyik?]")
            for n, c in enumerate(cands):
                box.button(c, key=f"{wkey}_c{n}", type="tertiary", on_click=_tpl_set, args=(wkey, c))
        elif not str(v).strip():
            box.caption(":red[hiányzik]")
        elif warn:
            box.caption(f":orange[:material/warning: {warn}]")
        elif src and v == val:
            box.caption(f"forrás: {src}")
    blocks_text = {}
    for j, b in enumerate(info["blocks"]):
        with st.container(border=True):
            st.markdown(f":material/edit_note: **Szabad blokk: {b['name']}**")
            tkey = f"tplb_{pid}_{nonce}_{j}"
            c_in, c_go = st.columns([4, 1.3], vertical_alignment="bottom")
            instr = c_in.text_input(f"Utasítás, ha {ASSISTANT_NAME} fogalmazza:", key=f"tplbi_{pid}_{nonce}_{j}")
            if c_go.button(f":material/auto_awesome: {ASSISTANT_NAME} javasol", key=f"tplbg_{pid}_{nonce}_{j}", **STRETCH):
                with st.spinner("A blokk készül…"):
                    try:
                        st.session_state[tkey] = tplfill.suggest_block(
                            claude_client, MODEL, form["name"], b, values, instr,
                            template_context(project, project_docs, selected, 8000),
                            log=usage_logger(pid, "mintakitoltes"))
                    except Exception as e:
                        st.warning(f"A javaslat most nem készült el: {e}")
            blocks_text[b["name"]] = st.text_area("A blokk szövege (üresen pontozott vonal kerül a helyére):",
                                                  key=tkey, height=150)
    gaps = tplfill.missing(info, values, blocks_text)
    if gaps:
        st.warning("Hiányzik: " + ", ".join(gaps) + ". Ha így készíted el, ezek helyére pontozott vonal kerül.")
    papers = letterheads()
    paper_ids = [""] + [paper for _, paper in papers]
    guess = template_paper(values, papers) or (project_paper(pid) if project_paper(pid) not in ("", "ures") else "")
    paper = st.selectbox(":material/contract: Fejléces papír", paper_ids,
                         index=paper_ids.index(guess) if guess in paper_ids else 0,
                         format_func=lambda v: next((f"{r.get('name')} fejléces papírja" for r, p in papers
                                                     if p == v), "A minta saját formája (nincs fejléces papír)"),
                         key=f"tplf_paper_{nonce}",
                         help="Az elkészült irat fejléce, lábléce és margói erről a papírról jönnek. Alapértelmezés: "
                              "a meghatalmazó / szerződő cég papírja, ha szerepel az űrlapon.") if papers else ""
    if st.button(":material/description: Irat elkészítése", type="primary", key=f"tplf_make_{nonce}"):
        try:
            out = tplfill.fill(data, info, values, blocks_text)
            if paper:
                out = ex.apply_letterhead(out, base64.b64decode(get_document_content(paper)))
        except tplfill.TemplateError as e:
            st.error(f"A kitöltés nem sikerült: {e}")
            return
        stem = form["name"].rsplit(".", 1)[0]
        name = unique_name(f"{stem} – {date.today():%Y-%m-%d}.docx", {d["name"] for d in project_docs})
        ok, detail = save_document_bytes(pid, name, ex.DOCX_MIME, out)
        if not ok:
            st.error(f"A mentés nem sikerült: {detail}")
            return
        msg = (f"Elkészült az irat a(z) „{form['name']}” mintából." + (
            " Kitöltetlen (pontozott vonal): " + ", ".join(gaps) + "." if gaps else "")
            + f"\n\n{CREATED_MARK} `{name}` – itt lent és a **Dokumentumok** fülön is letölthető.")
        try:
            db.save_message(pid, "assistant", msg)
        except db.DbError:
            pass
        close_template_form(pid)
        st.rerun()


def render_template_picker(pid):
    """A beszélgetés fölött: minta választása és az űrlap megnyitása."""
    tpls = get_templates()
    if not tpls:
        return
    with st.popover(":material/edit_document: Irat mintából", **STRETCH):
        pick = st.selectbox("Minta", [d["id"] for d in tpls], index=None, placeholder="válassz mintát",
                            format_func=lambda i: mid_ellipsis(next(d["name"] for d in tpls if d["id"] == i), 40),
                            key=f"tpl_pick_{pid}")
        if st.button(":material/edit_note: Űrlap megnyitása", key=f"tpl_open_{pid}", disabled=pick is None, **STRETCH):
            open_template_form(pid, next(d for d in tpls if d["id"] == pick))
            st.rerun()


def render_chat_tab(project, readonly, project_docs, messages, atts_by_msg, summaries, tasks=None):
    pid = project["id"]
    docs_by_id = {d["id"]: d for d in project_docs}

    # Felső sor: mely projektdokumentumokat lássa Sara, és az összefoglalás gombja
    col_docs, col_compact, col_trim = st.columns([2.6, 2.8, 3.2], gap="small")
    selected_doc_ids = []
    if project_docs and not readonly:
        doc_ids = [d["id"] for d in project_docs]
        # Alapból egy sem: Sara a kivonatokat mindig látja, és ha kell, maga olvassa be a dokumentumot.
        # A kijelölt dokumentumok teljes szövege a kérdéssel együtt megy ki – korábban mind ki volt jelölve,
        # és ez minden kérdésnél több százezer tokent jelentett.
        ctx_key = f"ctx2_{pid}_{'_'.join(map(str, doc_ids))}"
        if st.session_state.pop("ctx_clear", None) == pid:   # a kijelölés egy kérdésre szól: elküldés után üres
            st.session_state.pop(ctx_key, None)
        n_sel = len([i for i in st.session_state.get(ctx_key, []) if i in docs_by_id])
        label = (f":material/attach_file: :orange[Csatolás a kérdéshez ({n_sel} dokumentum, teljes szöveg)]" if n_sel
                 else ":material/attach_file: Csatolás a kérdéshez")
        # kulccsal a felirat változása (darabszám) nem csukja be a nyitott ablakot, ha a Streamlit ezt támogatja
        pop_kw = {"key": f"ctxpop_{pid}"} if "key" in inspect.signature(st.popover).parameters else {}
        with col_docs.popover(label, **pop_kw, help=(
                "A dokumentumokat a Dokumentumok fülre töltsd fel: ez az olcsó út. Onnan Sara a kivonatukat "
                "mindig látja, és ha egy dokumentum tartalma kell, maga beolvassa. Itt csak akkor jelölj ki "
                "egy már feltöltött dokumentumot, ha a következő kérdéshez mindenképp teljes terjedelmében "
                "kell neki (ez drága).")):
            selected_doc_ids = st.multiselect(
                "A következő kérdéshez teljes szövegében csatolva (csak ehhez az egy kérdéshez; elküldés után a "
                "kijelölés törlődik):",
                options=doc_ids, default=[], format_func=lambda i: mid_ellipsis(docs_by_id[i]["name"]),
                key=ctx_key,
            )
    if messages and not readonly:
        with col_compact:
            compact_button(project, "compact_btn_chat")
        with col_trim:
            trim_button(project, "trim_btn_chat")

    convo_size = conversation_size_note(pid, messages) if messages and not readonly else 0
    if not readonly and not is_templates(project):
        if pid in st.session_state.get("tpl_form", {}):
            with st.container(border=True):
                render_template_form(project, project_docs,
                                     [docs_by_id[i] for i in selected_doc_ids if i in docs_by_id])
        else:
            render_template_picker(pid)

    chat_box = st.container(height=FALLBACK_HEIGHT_PX, border=False, key="chat_box")
    with chat_box:
        if not messages:
            st.caption("Az előzményeket összefoglaló váltotta fel (lásd az Összefoglalók fület); Sara azokból dolgozik."
                       if summaries else "Ebben a projektben még nincs üzenet.")
        for i, msg in enumerate(messages):
            with chat_message(msg["role"]):
                msg_key = msg.get("id", i)
                commenting = (not readonly and msg["role"] == "assistant"
                              and st.session_state.get("comment_msg") == f"{pid}:{msg_key}")
                if commenting:
                    render_comment_mode(pid, msg_key, msg["content"])
                    continue
                st.write(msg["content"])
                if msg["role"] == "user":
                    if atts_by_msg.get(msg.get("id")):
                        st.caption(attachment_caption(atts_by_msg[msg["id"]]))
                    continue
                if i == len(messages) - 1 and convo_size > CLOSE_SUGGEST_TOKENS:
                    close_suggestion(project, convo_size)
                if not readonly and len(split_blocks(strip_app_notes(msg["content"]))) > 1:
                    if st.button(":material/rate_review: Megjegyzések", key=f"cm_open_{msg_key}", type="tertiary",
                                 help="A válasz részeihez megjegyzést fűzhetsz, és a végén Sara egyben feldolgozza őket."):
                        st.session_state["comment_msg"] = f"{pid}:{msg_key}"
                        st.rerun()
                # A válaszban készült fájlok letöltése közvetlenül a chatből: a legutóbbi válasznál azonnal
                # látszanak a gombok, a régebbieknél pipára töltődnek be (hogy ne lassítsák az oldalt).
                for name in created_names(msg["content"]):
                    doc = find_doc(project_docs, name, msg.get("created_at"))
                    if not doc:
                        st.caption(f"`{name}` nem található a dokumentumtárban.")
                    elif i == len(messages) - 1 or st.checkbox(f":material/download: {name} letöltése",
                                                               key=f"chat_dl_{msg_key}_{doc['id']}"):
                        render_downloads(doc, key_prefix=f"chat_{msg_key}_")

        # Ha Sara a dokumentumot a chatbe írta (fájl nélkül), egy kattintással letölthető fájl lesz belőle
        last = messages[-1] if messages else None
        if last and not readonly and last["role"] == "assistant" and not created_names(last["content"]):
            saved_answers = st.session_state.setdefault("saved_answers", {})
            answer_key = f"{pid}:{last.get('id', len(messages))}"
            if answer_key in saved_answers:
                saved_doc = find_doc(project_docs, saved_answers[answer_key])
                if saved_doc:
                    render_downloads(saved_doc, key_prefix="saved_")
            elif st.button(":material/save: A legutóbbi válasz mentése dokumentumként", key=f"save_answer_{answer_key}",
                           help="Sara válasza a dokumentumtárba kerül, és Word, PDF vagy Markdown formátumban letölthető."):
                body = strip_app_notes(last["content"])
                heading = re.search(r"^#{1,3}\s+(.+)$", body, flags=re.MULTILINE)
                title = heading.group(1).strip("* ") if heading else f"Sara válasza {date.today().isoformat()}"
                file_name = unique_name(ex.safe_filename(title, "md"), {d["name"] for d in project_docs})
                ok, detail = save_document_bytes(pid, file_name, ex.MD_MIME, (body + "\n").encode("utf-8"))
                if ok:
                    saved_answers[answer_key] = file_name
                    st.rerun()
                else:
                    st.error(f"A mentés nem sikerült: {detail}")

    if readonly:
        st.info("Ez a projekt le van zárva, ezért csak olvasható. Kérdezni az oldalsáv „Újranyitás” gombja után lehet.")
        return

    show_flash("chat_flash")
    submitted = st.chat_input(f"Kérdezz {ASSISTANT_FROM}; képet vagy fájlt is szúrhatsz be a kérdéshez…",
                              accept_file="multiple", file_type=ATTACH_TYPES, key=f"chat_input_{pid}")
    pending = st.session_state.get("pending_question", {}).pop(pid, None)
    resume = None
    if not submitted and not pending:
        resume = orphan_question(pid, messages)
        if not resume:
            return
    if resume:   # Sara válasza elveszett (a futás megszakadt): a kérdést újra feldolgozzuk
        question, uploads = resume["content"], []
    elif pending:   # a Megjegyzések feldolgozása gombbal összeállított üzenet
        question, uploads = pending, []
    else:
        question = (submitted if isinstance(submitted, str) else submitted.text or "").strip()
        uploads = [] if isinstance(submitted, str) else list(submitted.files or [])

    # A csatolt munkaanyagok ellenőrzése
    new_files, rejected = [], []
    for f in uploads:
        mime = guess_mime(f.name, f.type)
        limit_mb = MAX_IMAGE_MB if mime.startswith("image/") else MAX_FILE_MB
        if f.size > limit_mb * 1024 * 1024:
            rejected.append(f"{f.name} (nagyobb, mint {limit_mb} MB)")
        else:
            new_files.append((f.name, mime, f.getvalue()))
    if not question and not new_files:
        st.session_state["chat_flash"] = ("error", "Nem csatoltam: " + "; ".join(rejected) + ".")
        st.rerun()
    question = question or "Kérlek, nézd meg a csatolt fájlt."

    with chat_box:
        if resume:   # a kérdés már látszik és el van mentve; a hozzá csatolt fájlok is megvannak
            saved = None
        else:
            with chat_message("user"):
                st.write(question)
                if new_files:
                    st.caption(attachment_caption([{"name": n} for n, _, _ in new_files]))

            # Mentés: az üzenet, majd hozzá kötve a munkaanyagok (ezek nem kerülnek a dokumentumtárba)
            try:
                saved = db.save_message(pid, "user", question)
            except db.DbError as e:
                st.error(f"Az üzenet mentése nem sikerült: {e}")
                return
        current_files = list(load_attachment_files(atts_by_msg.get(resume.get("id"), []))) if resume else []
        for name, mime, data in new_files:
            try:
                if saved and saved.get("id") is not None:
                    db.save_attachment(pid, saved["id"], name, mime, data)
                current_files.append((name, mime, base64.b64encode(data).decode("ascii")))
            except db.DbError as e:
                rejected.append(f"{name} ({e})")

        saved_question = question   # ez van az adatbázisban (a gyorsítótár ehhez igazodik)
        # Csak az aktuális projekt kiválasztott dokumentumai kerülhetnek a kérdés mellé
        doc_files = []
        for doc_id in selected_doc_ids:
            doc = docs_by_id.get(doc_id)
            if not doc:
                continue
            try:
                doc_files.append((doc["name"], doc["mime_type"], get_document_content(doc_id)))
            except Exception:
                st.error(f"A(z) {doc['name']} dokumentum nem tölthető be, ezért nem csatoltam.")

        if doc_files:   # a kijelölés csak erre a kérdésre szólt (különben minden további kérdéssel újra elmenne)
            st.session_state["ctx_clear"] = pid
        api_messages = build_history(messages[:-1] if resume else messages, atts_by_msg)
        api_messages.append({"role": "user", "content": current_message(
            saved_question, current_files, question[len(saved_question):].strip() + "\n\n"
            + live_context(tasks, project_docs, linked_records_context(project, all_projects)), doc_files)})

        # Eszközök: a beépített webes kereső (az Anthropic szerverein fut) és a dokumentumkészítők (nálunk futnak)
        # A webes keresést elsősorban a kereső kolléga végzi (research_web); Sara saját keresése ellenőrzésre marad
        tools = [{"type": "web_search_20250305", "name": "web_search", "max_uses": 3}] + ex.TOOLS + [finder.WEB_TOOL]
        task_events = []
        task_handlers = {}
        if tasks is not None:  # feladatkezelés csak akkor, ha a feladatok táblája létezik
            tools = tools + TASK_TOOLS
            task_handlers.update(make_task_handlers(pid, tasks, task_events))
        if archived_summaries(summaries):   # a régi, beolvasztott összefoglalók kérésre
            tools = tools + [OLD_SUMMARY_TOOL]
            task_handlers.update(make_old_summary_handler(summaries))
        if project_docs:
            tools = tools + DOC_TOOLS
            task_handlers.update(make_doc_handlers(project_docs))
        _tpls = get_templates() if not is_templates(project) else []
        if _tpls:   # Iratminták: űrlap megnyitása Sara kérésére
            tools = tools + [TPL_FORM_TOOL]

            def _open_form(args, _t=_tpls):
                tpl = pick_by_name(_t, str(args.get("minta") or ""), lambda d: d["name"])
                if not tpl:
                    raise ValueError("Nincs ilyen minta. Az Iratmintákban: " + "; ".join(d["name"] for d in _t))
                open_template_form(pid, tpl, str(args.get("utasitas") or "").strip())
                return (f"Megnyitottam az űrlapot a(z) „{tpl['name']}” mintához: a felhasználó ott látja az előtöltött "
                        "mezőket, kiegészíti, és onnan készül az irat. Az adatokat ne kérdezd meg a chatben; egy mondatban "
                        "mondd meg, hogy az űrlap nyitva van a beszélgetés fölött.")
            task_handlers["irat_mintabol"] = _open_form
        if project_docs and db.has_doc_org():   # kategória és elavult jelölés, kérésre
            tools = tools + [DOC_STATUS_TOOL]
            task_handlers.update(make_doc_status_handler(pid))
        if len(project_docs or []) >= 2:   # az ellenőrző kolléga: dokumentumok összhangja
            tools = tools + [consistency.TOOL]
            task_handlers.update(make_consistency_handler(pid, project_docs))
        try:
            all_projs = db.get_projects()
        except db.DbError:
            all_projs = []
        if all_projs:   # a kereső kolléga: más projektek dokumentumaiban is keres (a bizalmasokat kivéve)
            tools = tools + [finder.TOOL]
        task_handlers.update(make_finder_handler(pid, all_projs))   # a webes kutatás projekt nélkül is megy
        # Larától átvett eszközök (jogszabály, adószám, Red Flag, összehasonlítás, redline)
        tools = tools + [njt.TOOL, njt.WEB_FETCH, navtax.TOOL, feor.TOOL, redflag.TOOL, diffreport.TOOL,
                         diffreport.REDLINE_TOOL]
        _attached_all = list(current_files) + [f for m in messages[::-1]
                                               for f in load_attachment_files(atts_by_msg.get(m.get("id"), []))]
        task_handlers.update(make_legal_handlers(pid, _attached_all))
        attached = [f for m in messages for f in load_attachment_files(atts_by_msg.get(m.get("id"), []))
                    if f[0].lower().endswith(".docx")]
        docx_sources = [f for f in current_files if f[0].lower().endswith(".docx")] + attached[::-1]
        if docx_sources or any(d["name"].lower().endswith(".docx") for d in project_docs):
            tools = tools + [EDIT_TOOL]
            task_handlers.update(make_edit_handler(project_docs, docx_sources))
        taken_names = {d["name"] for d in project_docs}

        def save_created_file(name, mime, data):
            """Sara által készített fájl mentése; siker esetén a tárolt (egyedi) fájlnevet adja vissza."""
            final_name = unique_name(name, taken_names)
            ok, detail = save_document_bytes(pid, final_name, mime, data)
            if ok:
                taken_names.add(final_name)
            return ok, (final_name if ok else detail)

        with chat_message("assistant"):
            with st.spinner(f"{ASSISTANT_NAME} elemzi a tartalmat, keres és gondolkodik… (hosszabb dokumentumnál ez több perc is lehet)"):
                answer, sources, created_files, problems = chat_flow.run_conversation(
                    claude_client, MODEL, build_system(summaries, tasks, project_docs), api_messages, tools,
                    save_created_file,
                    handlers=task_handlers, usage_cb=usage_logger(pid, "chat"))
                answer = answer or f"({ASSISTANT_NAME} most nem adott szöveges választ.)"
                # A hibát az elmentett válaszba is beírjuk, mert az oldal rögtön újratöltődik
                if problems:
                    answer += f"\n\n{PROBLEM_MARK} " + " ".join(problems)
                if task_events:
                    answer += f"\n\n{TASK_MARK} " + "; ".join(task_events) + " – a **Feladatok** fülön látod."
                if created_files:
                    answer += (f"\n\n{CREATED_MARK} " + ", ".join(f"`{n}`" for n in created_files)
                               + " – itt lent és a **Dokumentumok** fülön is letölthető.")
                if sources:
                    answer += "\n\n**Források:**\n" + "\n".join(f"- [{title}]({url})" for url, title in sources.items())
                # AZONNAL mentjük, még mielőtt bármit a képernyőre írnánk: ha közben a felhasználó kattintott
                # (vagy frissült az oldal), a Streamlit a következő képernyőírásnál leállítja a futást, és a
                # kifizetett válasz elveszne
                try:
                    db.save_message(pid, "assistant", answer)
                except db.DbError as e:
                    st.session_state["chat_flash"] = ("error", f"Sara válaszának mentése nem sikerült: {e}")
            st.write(answer)

    if rejected:
        st.session_state["chat_flash"] = ("warning", "Nem csatoltam: " + "; ".join(rejected) + ".")
    st.rerun()


def orphan_question(pid, messages):
    """Ha az utolsó üzenet a felhasználóé, és nincs rá válasz (a futás megszakadt: újraindult az app, elment a
    kapcsolat), egy gombbal újra feldolgozható – nem kell újra beírni. Automatikusan nem indítjuk, mert ha egy
    másik ablakban még fut, kétszer fizetnénk érte."""
    last = messages[-1] if messages else None
    if not last or last.get("role") != "user":
        return None
    mid = str(last.get("id") or len(messages))
    c_note, c_btn = st.columns([4, 1.4], vertical_alignment="center")
    c_note.caption(":material/pending: Az utolsó kérdésre nincs válasz. Ha Sara egy másik ablakban még dolgozik "
                   "rajta, várj, és frissítsd az oldalt; ha a futás megszakadt, folytasd innen.")
    if c_btn.button(":material/play_arrow: Folytatás", key=f"resume_{pid}_{mid}", **STRETCH):
        return last
    return None


def rename_form(doc, taken, key_prefix="ren"):
    """Dokumentum átnevezése helyben (a kiterjesztés marad), hogy ne kelljen törölni és újra feltölteni."""
    stem, dot, ext = doc["name"].rpartition(".")
    if not dot:
        stem, ext = doc["name"], ""
    new = st.text_input("Új név", value=stem, key=f"{key_prefix}_in_{doc['id']}",
                        help=f"A kiterjesztés (.{ext}) marad." if ext else None)
    if st.button(":material/check: Átnevezés", key=f"{key_prefix}_ok_{doc['id']}", type="primary"):
        new = " ".join(str(new or "").replace("/", "-").replace("\\", "-").split()).strip(" .")
        final = f"{new}.{ext}" if ext else new
        if not new:
            st.error("Adj meg egy nevet.")
        elif final == doc["name"]:
            st.info("A név nem változott.")
        elif final in taken:
            st.error("Már van ilyen nevű dokumentum.")
        else:
            try:
                db.rename_document(doc["id"], final)
                try:   # a kereső indexe is az új nevet kapja
                    db.set_document_text(doc["id"], final, db.get_document_text(doc["id"]))
                except db.DbError:
                    pass
                st.rerun()
            except db.DbError as e:
                st.error(f"Az átnevezés nem sikerült: {e}")


# --- DOKUMENTUMOK FÜL --------------------------------------------------------
def render_docs_tab(project, readonly, project_docs):
    pid = project["id"]
    with st.container(height=FALLBACK_HEIGHT_PX, border=False, key="docs_box"):
        if is_templates(project):
            st.info("Ide töltsd fel a Word-mintáidat (.docx). Bármelyik projektben vagy törzsaktában a beszélgetés "
                    "fölötti **Irat mintából** gombbal használhatod őket. " + TPL_SYNTAX)
        else:
            st.caption("A projekt jóváhagyott dokumentumai és a Sara által készített fájlok. Kizárólag ehhez a "
                       "projekthez tartoznak; a beszélgetésbe csatolt munkaanyagok nem ide kerülnek.")
        show_flash("docs_flash")

        if not readonly:
            # A kulcsban lévő számláló nullázza a feltöltőt mentés után, így nem mentjük el kétszer ugyanazt
            nonce = st.session_state.get("uploader_nonce", {}).get(pid, 0)
            uploaded_files = st.file_uploader(
                "Projektdokumentum feltöltése (PDF, Word, Excel, PowerPoint, TXT, PNG, JPG):",
                type=DOC_TYPES,
                accept_multiple_files=True,
                key=f"uploader_{pid}_{nonce}",
            )
            if uploaded_files and st.button(":material/save: Feltöltés a projekt dokumentumtárába", type="primary"):
                saved, failed = 0, []
                for f in uploaded_files:
                    if f.size > MAX_FILE_MB * 1024 * 1024:
                        failed.append(f"{f.name} (nagyobb, mint {MAX_FILE_MB} MB)")
                        continue
                    ok, detail = save_document_bytes(pid, f.name, guess_mime(f.name, f.type), f.getvalue())
                    if ok:
                        saved += 1
                    else:
                        failed.append(f"{f.name} ({detail})")
                if failed:
                    msg = "Nem sikerült feltölteni: " + "; ".join(failed) + "."
                    if saved:
                        msg += f" Sikeresen feltöltve: {saved} fájl."
                    st.session_state["docs_flash"] = ("error", msg)
                else:
                    st.session_state["docs_flash"] = ("success", f"{saved} fájl feltöltve a projekthez.")
                st.session_state.setdefault("uploader_nonce", {})[pid] = nonce + 1
                if saved and is_record(project) and project.get("record_type") in ("ceg", "ingatlan"):
                    try:   # cégkivonatból, tulajdoni lapból magától adatlap
                        for d in db.get_documents(pid):
                            if d["name"] in {f.name for f in uploaded_files} and not d["name"].startswith(records.SHEET_PREFIX):
                                if registry.detect(db.get_document_text(d["id"]) or ""):
                                    name, _w = make_record_sheet(project, [d])
                                    st.session_state["docs_flash"] = ("success", f"Feltöltve; az adatlap is elkészült: {name}.")
                                    break
                    except Exception as e:
                        st.session_state["docs_flash"] = ("warning", f"Feltöltve, de az adatlap nem készült el: {e}")
                if saved and db.has_doc_org():   # az új dokumentumok besorolása és a korábbi változatok keresése
                    try:
                        fresh = db.get_documents(pid)
                        names = {f.name for f in uploaded_files}
                        new_docs = []
                        for d in fresh:   # a legújabb azonos nevű a most feltöltött
                            if d["name"] in names and d["name"] not in {x["name"] for x in new_docs}:
                                new_docs.append(d)
                        categorize_docs(pid, new_docs)
                        sugg = [(n["id"], n["name"], o["id"], o["name"]) for n in new_docs
                                for o in older_versions(n, fresh) if o["id"] not in {x["id"] for x in new_docs}]
                        if sugg:
                            st.session_state.setdefault("ver_sugg", {})[pid] = sugg
                    except Exception:
                        pass   # a rendezés nem akaszthatja meg a feltöltést
                st.rerun()
            st.divider()

        if not project_docs:
            st.info("Ebben a projektben még nincs dokumentum.")
        try:
            textless = db.get_textless_documents(pid) if not readonly else []
        except db.DbError:
            textless = []
        if textless:
            col_note, col_btn = st.columns([5, 2.4], vertical_alignment="center")
            col_note.caption(f"{len(textless)} szkennelt PDF vagy kép szövege még nincs felismerve, ezért a "
                             "kereső és Sara nem tud olvasni belőlük. Egyszeri költség: kb. 1 cent 10 oldalanként.")
            if col_btn.button(":material/document_scanner: Szöveg felismerése", key=f"fill_ocr_{pid}", **STRETCH):
                failed = fill_missing_texts(textless)
                done = len(textless) - len(failed)
                st.session_state["docs_flash"] = (
                    ("warning", f"{done} dokumentum szövege elkészült, {len(failed)} nem:\n\n"
                     + "\n".join(f"- {x}" for x in failed)) if failed
                    else ("success", f"{done} dokumentum szövege elkészült; mostantól kereshetők."))
                st.rerun()
        missing = [d for d in project_docs if not (d.get("summary") or "").strip()]
        if missing and not readonly:
            col_note, col_btn = st.columns([5, 2.4], vertical_alignment="center")
            col_note.caption(f"{len(missing)} dokumentumhoz még nincs kivonat. A kivonatból Sara is látja, "
                             "melyik dokumentumban mit talál.")
            if col_btn.button(":material/summarize: Kivonatok elkészítése", key=f"fill_sums_{pid}", **STRETCH):
                failed = fill_missing_summaries(missing)
                done = len(missing) - len(failed)
                st.session_state["docs_flash"] = (
                    ("warning", f"{done} kivonat elkészült, {len(failed)} nem:\n\n"
                     + "\n".join(f"- {x}" for x in failed)) if failed
                    else ("success", f"{done} kivonat elkészült."))
                st.rerun()
        render_doc_list(pid, project_docs, readonly)


def _set_doc_category(doc_id, key):
    try:
        db.set_document_org(doc_id, category=st.session_state.get(key))
    except db.DbError as e:
        st.session_state["docs_flash"] = ("error", f"A kategória mentése nem sikerült: {e}")


def _set_doc_obsolete(doc_id, key):
    try:
        db.set_document_org(doc_id, obsolete=bool(st.session_state.get(key)))
    except db.DbError as e:
        st.session_state["docs_flash"] = ("error", f"A jelölés mentése nem sikerült: {e}")


def render_doc_row(doc, project_docs, readonly):
    col_info, col_sum, col_dl, col_prev, col_ren, col_del = st.columns([5, 1.5, 1.5, 1.5, 0.8, 0.8])
    size_kb = max(1, round((doc.get("size_bytes") or 0) / 1024))
    name = f"~~{doc['name']}~~" if is_obsolete(doc) else f"**{doc['name']}**"
    col_info.write(f"{doc_icon(doc)} {name} · {size_kb} KB · {fmt_dt(doc.get('created_at'))[:10]}")
    want_summary = col_sum.checkbox("Kivonat", key=f"sum_{doc['id']}")
    want_download = col_dl.checkbox("Letöltés", key=f"dl_{doc['id']}")
    want_preview = col_prev.checkbox("Előnézet", key=f"prev_{doc['id']}")
    if not readonly:
        with col_ren.popover(":material/edit:", help="Átnevezés, kategória, elavult"):
            rename_form(doc, [d["name"] for d in project_docs])
            if db.has_doc_org():
                st.divider()
                cat = doc_category(doc)
                ckey, okey = f"cat_{doc['id']}", f"obs_{doc['id']}"
                st.selectbox("Kategória", DOC_CATEGORIES, key=ckey,
                             index=DOC_CATEGORIES.index(cat) if cat in DOC_CATEGORIES else None,
                             placeholder="válassz", on_change=_set_doc_category, args=(doc["id"], ckey))
                st.checkbox("Elavult (régi változat – Sara csak kérésre nézi)", value=is_obsolete(doc), key=okey,
                            on_change=_set_doc_obsolete, args=(doc["id"], okey))
    if not readonly and col_del.button(":material/delete:", key=f"del_{doc['id']}", help="Dokumentum törlése a projektből"):
        try:
            db.delete_document(doc["id"])
            st.rerun()
        except db.DbError as e:
            st.error(f"A törlés nem sikerült: {e}")
    if want_summary:
        with st.container(border=True):
            st.write((doc.get("summary") or "").strip()
                     or "Ehhez a dokumentumhoz még nincs kivonat; a fenti gombbal elkészíthető.")
    if want_download:
        render_downloads(doc)
    if want_preview:
        with st.container(border=True):
            render_doc_preview(doc)


def render_doc_list(pid, project_docs, readonly):
    """A dokumentumok kategóriánként; az elavultak külön, összecsukva a végén."""
    if not db.has_doc_org():
        if not readonly and project_docs:
            st.caption(":material/info: A dokumentumok rendezéséhez (kategória, elavult jelölés) futtasd le a "
                       "Supabase SQL Editorban a schema.sql-t (az 1/d rész az új); utána ez magától bekapcsol.")
        for doc in project_docs:
            render_doc_row(doc, project_docs, readonly)
        return
    sugg = st.session_state.get("ver_sugg", {}).get(pid) or []
    by_id = {d["id"]: d for d in project_docs}
    sugg = [s for s in sugg if s[0] in by_id and s[2] in by_id and not is_obsolete(by_id[s[2]])]
    for new_id, new_name, old_id, old_name in sugg:
        with st.container(border=True):
            st.markdown(f":material/history: Az új **{new_name}** a(z) **{old_name}** újabb változata?")
            c1, c2 = st.columns(2)
            if c1.button(":material/check: Igen, a régit elavultnak jelölöm", key=f"vs_yes_{new_id}_{old_id}", **STRETCH):
                try:
                    db.set_document_org(old_id, obsolete=True)
                except db.DbError as e:
                    st.session_state["docs_flash"] = ("error", f"A jelölés nem sikerült: {e}")
                st.session_state["ver_sugg"][pid] = [s for s in sugg if s[2] != old_id]
                st.rerun()
            if c2.button("Nem, mindkettő érvényes", key=f"vs_no_{new_id}_{old_id}", **STRETCH):
                st.session_state["ver_sugg"][pid] = [s for s in sugg if not (s[0] == new_id and s[2] == old_id)]
                st.rerun()
    uncategorized = [d for d in project_docs if doc_category(d) == UNCATEGORIZED and not is_obsolete(d)]
    if uncategorized and not readonly:
        col_note, col_btn = st.columns([5, 2.4], vertical_alignment="center")
        col_note.caption(f"{len(uncategorized)} dokumentum még nincs kategóriába sorolva. Sara a kivonatuk alapján "
                         "besorolja őket (egyetlen olcsó hívás); utána a ceruzával bármikor módosíthatod.")
        if col_btn.button(":material/category: Kategorizálás", key=f"categorize_{pid}", **STRETCH):
            try:
                n = categorize_docs(pid, uncategorized)
                st.session_state["docs_flash"] = ("success", f"{n} dokumentum kategóriát kapott.")
            except Exception as e:
                st.session_state["docs_flash"] = ("error", f"A kategorizálás nem sikerült: {e}")
            st.rerun()
    current = current_docs(project_docs)
    for cat in DOC_CATEGORIES + [UNCATEGORIZED]:
        docs = [d for d in current if doc_category(d) == cat]
        if not docs:
            continue
        st.markdown(f"**{cat}** · {len(docs)}")
        for doc in docs:
            render_doc_row(doc, project_docs, readonly)
    old = [d for d in project_docs if is_obsolete(d)]
    if old:
        with st.expander(f":material/inventory_2: Elavult dokumentumok ({len(old)}) – Sara csak kérésre nézi"):
            for doc in old:
                render_doc_row(doc, project_docs, readonly)


# --- ÖSSZEFOGLALÓK FÜL -------------------------------------------------------
def _close_sum_delete_dialog():
    st.session_state.pop("sum_delete_open", None)


@dialog("Összefoglaló végleges törlése", _close_sum_delete_dialog)
def sum_delete_dialog(summary):
    st.write(f"Biztosan törlöd a(z) **{fmt_dt(summary['created_at'])}** időpontú összefoglalót?")
    if summary.get("description"):
        st.caption(summary["description"])
    st.warning("A törlés végleges, nem vonható vissza. Amit ez az összefoglaló őrzött a törölt előzményekből, "
               "azt Sara ezután nem fogja tudni.")
    c1, c2 = st.columns(2)
    if c1.button(":material/delete: Igen, végleges törlés", type="primary", key="sum_delete_yes", **STRETCH):
        try:
            db.delete_summary(summary["id"])
        except db.DbError as e:
            st.error(f"A törlés nem sikerült: {e}")
            return
        if st.session_state.get("sum_editing") == summary["id"]:
            st.session_state["sum_editing"] = None
        st.session_state["sums_flash"] = ("success", "Az összefoglaló törölve.")
        _close_sum_delete_dialog()
        st.rerun()
    if c2.button("Mégse", key="sum_delete_no", **STRETCH):
        _close_sum_delete_dialog()
        st.rerun()


def render_summaries_tab(project, readonly, summaries, messages):
    with st.container(height=FALLBACK_HEIGHT_PX, border=False, key="sums_box"):
        st.caption("Az összefoglalók a törölt előzmények helyett őrzik a projekt lényegét. Sara a legutóbbi "
                   "összevont emlékezetet kapja meg (lezáráskor ez a korábbiakat is magába olvasztja); a régebbiek "
                   "archívak, kérésre beolvassa őket. A dátum melletti gombokkal szerkeszthetők és törölhetők.")
        show_flash("sums_flash")
        if not readonly:
            if messages:
                compact_button(project, "compact_btn_tab")
            else:
                st.caption("Új összefoglaló akkor készíthető, ha van beszélgetés a projektben.")
        if not summaries:
            st.info("Ebben a projektben még nincs összefoglaló.")
        edit_nonce = st.session_state.get("sum_edit_nonce", 0)  # mentés után ez zárja be a szerkesztőt
        for s in reversed(summaries):  # a legfrissebb elöl
            st.divider()
            edited = f" · szerkesztve: {fmt_dt(s['updated_at'])}" if s.get("updated_at") else ""
            editing = not readonly and st.session_state.get("sum_editing") == s["id"]
            # a két kis gomb közvetlenül a dátum mellett áll (szerkesztett összefoglalónál hosszabb a felirat)
            col_title, col_edit, col_del, _ = st.columns([4.6 if edited else 2.1, 0.45, 0.45, 5], gap="small",
                                                         vertical_alignment="center")
            tag = (" · :violet[összevont emlékezet]" if is_memory(s)
                   else " · :gray[archív]" if any(s is o for o in archived_summaries(summaries)) else "")
            col_title.markdown(f"**{fmt_dt(s['created_at'])}**{edited}{tag}")
            if not readonly:
                if col_edit.button(":material/edit:", key=f"sum_editbtn_{s['id']}",
                                   help="Szerkesztés bezárása (mentés nélkül)" if editing else "Szerkesztés"):
                    st.session_state["sum_editing"] = None if editing else s["id"]
                    st.session_state["sum_edit_nonce"] = edit_nonce + 1  # a mezők a tárolt szöveggel induljanak
                    st.rerun()
                if col_del.button(":material/delete:", key=f"sum_delbtn_{s['id']}", help="Összefoglaló törlése"):
                    st.session_state["sum_delete_open"] = s["id"]
                    st.rerun()
            st.write(s.get("description") or "(nincs leírás)")
            with st.expander("Teljes összefoglaló"):
                st.markdown(s["content"])
            if editing:
                new_description = st.text_area("Rövid leírás", value=s.get("description") or "", height=90,
                                               key=f"sum_desc_{s['id']}_{edit_nonce}")
                new_content = st.text_area("Összefoglaló (Markdown)", value=s["content"], height=320,
                                           key=f"sum_content_{s['id']}_{edit_nonce}")
                if st.button(":material/save: Módosítás mentése", key=f"sum_save_{s['id']}", type="primary"):
                    if not new_content.strip():
                        st.error("Az összefoglaló nem lehet üres.")
                    else:
                        try:
                            db.update_summary(s["id"], new_description.strip(), new_content.strip())
                            st.session_state["sum_edit_nonce"] = edit_nonce + 1
                            st.session_state["sum_editing"] = None
                            st.session_state["sums_flash"] = ("success", "Az összefoglaló módosítva.")
                            st.rerun()
                        except db.DbError as e:
                            st.error(f"A mentés nem sikerült: {e}")


# --- FELADATOK FÜL -----------------------------------------------------------
def _task_apply(task_id, **fields):
    try:
        db.update_task(task_id, **fields)
    except db.DbError as e:
        st.session_state["tasks_flash"] = ("error", f"A módosítás nem sikerült: {e}")


def _task_checked(task_id, key):
    # pipa = lezárva (alapból „kész”); a pipa levétele újra nyitottá teszi
    _task_apply(task_id, status="done" if st.session_state.get(key) else "open")


def _task_status_changed(task_id, key):
    value = st.session_state.get(key)
    if value in TASK_STATUS_LABELS:
        _task_apply(task_id, status=value)


def _close_task_delete_dialog():
    st.session_state.pop("task_delete_open", None)


@dialog("Feladat végleges törlése", _close_task_delete_dialog)
def task_delete_dialog(task):
    st.write(f"Biztosan törlöd ezt a feladatot? **{task_topic(task)} » {task['title']}**")
    st.warning("A törlés végleges, a feladat nyom nélkül eltűnik a listáról (ez a tévesen felvett tételekre való). "
               "Ha csak nem kell elvégezni, és ezt később is látni szeretnéd, inkább állítsd „törölt” vagy "
               "„elavult” állapotra, és írd a megjegyzésbe az okát.")
    c1, c2 = st.columns(2)
    if c1.button(":material/delete: Igen, végleges törlés", type="primary", key="task_delete_yes", **STRETCH):
        try:
            db.delete_task(task["id"])
        except db.DbError as e:
            st.error(f"A törlés nem sikerült: {e}")
            return
        st.session_state["tasks_flash"] = ("success", "A feladat törölve.")
        _close_task_delete_dialog()
        st.rerun()
    if c2.button("Mégse", key="task_delete_no", **STRETCH):
        _close_task_delete_dialog()
        st.rerun()


def render_task_row(task, readonly, all_tasks):
    tid = task["id"]
    status = task_status(task)
    nonce = st.session_state.get("task_edit_nonce", 0)
    editing = not readonly and st.session_state.get("task_editing") == tid
    c_chk, c_title, c_status, c_edit, c_del = st.columns([0.35, 7, 1.7, 0.45, 0.45], gap="small",
                                                         vertical_alignment="center")
    # a kulcsban benne van az állapot, így módosítás után a pipa és a lista a friss értéket mutatja
    chk_key = f"task_chk_{tid}_{status}"
    c_chk.checkbox("Lezárva", value=status not in TASK_ACTIVE, key=chk_key, label_visibility="collapsed",
                   disabled=readonly, on_change=_task_checked, args=(tid, chk_key))
    title_class = "sara-task" if status in TASK_ACTIVE else "sara-task-done"
    c_title.markdown(f'<span class="{title_class}">{html.escape(task["title"])}</span>', unsafe_allow_html=True)
    if task.get("note"):
        c_title.caption(f":material/sticky_note_2: {task['note']}")
    st_key = f"task_st_{tid}_{status}"
    c_status.selectbox("Állapot", options=list(TASK_ORDER), index=TASK_ORDER.index(status),
                       format_func=TASK_STATUS_LABELS.get, key=st_key, label_visibility="collapsed",
                       disabled=readonly, on_change=_task_status_changed, args=(tid, st_key))
    if readonly:
        return
    if c_edit.button(":material/edit:", key=f"task_editbtn_{tid}",
                     help="Szerkesztés bezárása (mentés nélkül)" if editing else "Szöveg, téma és megjegyzés szerkesztése"):
        st.session_state["task_editing"] = None if editing else tid
        st.session_state["task_edit_nonce"] = nonce + 1
        st.rerun()
    if c_del.button(":material/delete:", key=f"task_delbtn_{tid}", help="Feladat végleges törlése (tévesen felvett tételhez)"):
        st.session_state["task_delete_open"] = tid
        st.rerun()
    if editing:
        e1, e2 = st.columns([1, 2])
        new_topic = topic_input(e1, "Téma", list(dict.fromkeys(task_topic(t) for t in all_tasks)),
                                value=task_topic(task), key=f"task_e_topic_{tid}_{nonce}")
        new_title = e2.text_input("Tennivaló", value=task["title"], key=f"task_e_title_{tid}_{nonce}")
        new_note = st.text_input("Megjegyzés", value=task.get("note") or "", key=f"task_e_note_{tid}_{nonce}",
                                 placeholder="pl. kire várunk, vagy miért nem kell elvégezni")
        if st.button(":material/save: Mentés", key=f"task_e_save_{tid}", type="primary"):
            if not new_title.strip():
                st.error("A tennivaló szövege nem lehet üres.")
            else:
                others = [t for t in all_tasks if t["id"] != tid]
                try:
                    db.update_task(tid, topic=match_topic(new_topic, others), title=new_title.strip(),
                                   note=new_note.strip())
                except db.DbError as e:
                    st.error(f"A mentés nem sikerült: {e}")
                else:
                    st.session_state["task_editing"] = None
                    st.session_state["task_edit_nonce"] = nonce + 1
                    st.rerun()


def render_settings_tab(project, readonly):
    """A projekt (törzsakta) beállításai egy helyen: név, kapcsolt törzsakták, fejléces papír, bizalmasság."""
    pid = project["id"]
    record = is_record(project)
    if readonly:
        st.info("Lezárt projekt: a beállítások csak olvashatók.")
    new_name = st.text_input("Név", value=project["name"], disabled=readonly or is_special(project),
                             key=f"set_name_{pid}")
    if new_name.strip() and new_name.strip() != project["name"]:
        try:
            db.update_project(pid, name=new_name.strip())
            st.rerun()
        except db.DbError as e:
            st.error(f"Az átnevezés nem sikerült: {e}")
    if not db.has_records():
        st.caption("A kapcsolásokhoz és a fejléces papírhoz futtasd le a Supabase SQL Editorban a schema.sql-t.")
    elif record and project.get("record_type") == "ceg":
        extra = dict(record_extra(pid))
        try:
            words = [d for d in db.get_documents(pid) if d["name"].lower().endswith(".docx")]
        except db.DbError:
            words = []
        ids = [""] + [str(d["id"]) for d in words]
        cur = str(extra.get("fejlec") or "")
        pick = st.selectbox(
            ":material/contract: A cég fejléces papírja (Word-sablon ebből a törzsaktából)", ids,
            index=ids.index(cur) if cur in ids else 0, disabled=readonly, key=f"set_paper_rec_{pid}",
            format_func=lambda i: next((d["name"] for d in words if str(d["id"]) == i),
                                       "— nincs —" if words else "— tölts fel egy .docx sablont a Dokumentumok fülön —"),
            help="A cég projektjeinek dokumentumai és a cég nevében iratmintából készülő iratok erre a papírra kerülnek.")
        if pick != cur:
            try:
                db.set_setting(f"record:{pid}", {**extra, "fejlec": pick})
                record_extra.clear()
                st.rerun()
            except db.DbError as e:
                st.error(f"A beállítás nem menthető: {e}")
    elif not record and not is_special(project):
        recs = [r for r in all_projects if is_record(r) and r.get("status") != "closed"]
        rec_ids = {r["id"] for r in recs}
        linked = [r for r in (project.get("links") or []) if r in rec_ids]
        new_links = st.multiselect(
            ":material/link: Kapcsolt törzsakták", [r["id"] for r in recs], default=linked, disabled=readonly,
            format_func=lambda i: record_plain(next(r for r in recs if r["id"] == i)), key=f"set_links_{pid}",
            placeholder="Válassz törzsaktát…",
            help="A kapcsolt cégek, ingatlanok, személyek, járművek adatlapját Sara minden kérdésnél látja ebben a "
                 "projektben (iratokhoz, ellenőrzéshez).")
        if sorted(new_links) != sorted(linked):
            try:
                db.update_project(pid, links=new_links)
                st.rerun()
            except db.DbError as e:
                st.error(f"A kapcsolás nem menthető: {e}")
        papers = letterheads()
        opts = [""] + [r["id"] for r, _ in papers] + ["nf", "ures"]
        names = {r["id"]: r.get("name") for r, _ in papers}
        names.update(nf="Contact-Line 2002 (beépített)", ures="Üres (fejléc nélkül)")
        cur = project_paper_setting(pid)
        auto = next((names[r] for r in (project.get("links") or []) if r in names and r not in ("nf", "ures")),
                    "Contact-Line 2002")
        new_paper = st.selectbox(
            ":material/contract: Fejléces papír – melyik cégé a projekt", opts,
            index=opts.index(cur) if cur in opts else 0, disabled=readonly, key=f"set_paper_{pid}",
            format_func=lambda i: names.get(i) or f"Automatikus (most: {auto})",
            help="A projekt dokumentumai ennek a cégnek a fejléces papírjára készülnek (Word); letöltéskor "
                 "dokumentumonként átállítható. A papírt a cég törzsaktájának Beállítások fülén jelölöd ki.")
        if new_paper != cur:
            try:
                db.set_setting(f"fejlec:{pid}", new_paper)
                project_paper_setting.clear()
                st.rerun()
            except db.DbError as e:
                st.error(f"A beállítás nem menthető: {e}")
    if not is_special(project):
        confidential = bool(project.get("confidential"))
        new_conf = st.checkbox(
            "Bizalmas törzsakta" if record else "Bizalmas projekt", value=confidential, disabled=readonly,
            key=f"conf_{pid}",
            help=("Más projektből a kereső nem látja; csak oda kapcsolva, ahol kell. Személyi törzsakta alapból "
                  "bizalmas.") if record else (
                "A dokumentumait Sara kereső kollégája más projektekből nem látja. Ebben a projektben kérdezve "
                "a többi (nem bizalmas) projektben továbbra is keres."))
        if new_conf != confidential:
            try:
                db.update_project(pid, confidential=new_conf)
                st.rerun()
            except db.DbError as e:
                st.error(f"A beállítás nem menthető: {e}")


def render_tasks_tab(project, readonly, tasks):
    pid = project["id"]
    with st.container(height=FALLBACK_HEIGHT_PX, border=False, key="tasks_box"):
        if tasks is None:
            st.warning("A Feladatok használatához egyszer frissíteni kell az adatbázist: futtasd le a `schema.sql` "
                       "teljes tartalmát a Supabase **SQL Editor** felületén, majd töltsd újra ezt az oldalt. "
                       "A szkript meglévő adatot nem töröl; az alkalmazás többi része addig is működik.")
            try:
                with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "schema.sql"), encoding="utf-8") as fh:
                    with st.expander("A futtatandó szkript"):
                        st.code(fh.read(), language="sql")
            except OSError:
                pass
            return
        st.caption("Tennivalók témák szerint. Az állapot minden feladatnál választható (nyitott, választ várunk, "
                   "kész, elavult, törölt); a pipa gyorsgomb a lezáráshoz. Sara kérésre vesz fel és módosít feladatot.")
        show_flash("tasks_flash")
        topics = list(dict.fromkeys(task_topic(t) for t in tasks))

        if not readonly:
            with st.expander(":material/add: Új feladat"):
                with st.form(f"task_add_{pid}", clear_on_submit=True):
                    f1, f2 = st.columns([1, 2])
                    topic = topic_input(f1, "Téma", topics)
                    title = f2.text_input("Tennivaló", placeholder="pl. Gépésznek megírni a specifikációt")
                    note = st.text_input("Megjegyzés (nem kötelező)")
                    if st.form_submit_button("Felvétel"):
                        if not title.strip():
                            st.error("Írd be a tennivalót.")
                        else:
                            try:
                                db.add_task(pid, match_topic(topic, tasks), title, note)
                            except db.DbError as e:
                                st.error(f"A felvétel nem sikerült: {e}")
                            else:
                                st.rerun()
                if topics and not ACCEPT_NEW_OPTIONS:
                    st.caption("Meglévő témák: " + ", ".join(topics))

        hidden = [t for t in tasks if task_status(t) not in TASK_ACTIVE]
        show_all = bool(hidden) and st.toggle(f"Kész, elavult és törölt tételek is ({len(hidden)})",
                                              key=f"tasks_all_{pid}")
        visible = tasks if show_all else [t for t in tasks if task_status(t) in TASK_ACTIVE]
        if not tasks:
            st.info("Ebben a projektben még nincs feladat. Vegyél fel egyet fent, vagy kérd meg Sarát a beszélgetésben.")
        elif not visible:
            st.info("Nincs nyitott vagy válaszra váró feladat.")
        for topic in topics:
            rows = [t for t in visible if task_topic(t) == topic]
            if not rows:
                continue
            n_open = sum(1 for t in rows if task_status(t) == "open")
            n_wait = sum(1 for t in rows if task_status(t) == "waiting")
            counts = [f"{n_open} nyitott"] + ([f"{n_wait} válaszra vár"] if n_wait else [])
            if show_all and len(rows) > n_open + n_wait:
                counts.append(f"{len(rows) - n_open - n_wait} lezárt")
            with st.expander(f"**{topic}** · " + ", ".join(counts), expanded=True):
                for task in rows:
                    render_task_row(task, readonly, tasks)


# --- KERESÉS -----------------------------------------------------------------
def hit_opened(kind, item_id):
    """„Megnyitás” jelölő egy találatnál. Az állapotát külön is megjegyezzük, hogy a projektbe ugrás után
    a találati listához visszatérve ugyanazok a találatok legyenek nyitva."""
    open_hits = st.session_state.setdefault("open_hits", set())
    opened = st.checkbox("Megnyitás", value=(kind, item_id) in open_hits, key=f"hit_{kind}_{item_id}")
    (open_hits.add if opened else open_hits.discard)((kind, item_id))
    return opened


def project_label(p):
    """A projekt gombjának felirata; a bizalmas projekt lakat jelet kap."""
    name = str(p.get("name") or "")
    return f":material/shield_lock: {name}" if p.get("confidential") else name


# --- KÖLTSÉGEK -----------------------------------------------------------------
COST_PERIODS = {"today": "Ma", "7d": "Utolsó 7 nap", "month": "E hónap", "all": "Mind"}


def _usd(x):
    sign, x = ("-" if x < 0 else ""), abs(x)
    return f"{sign}${x:,.2f}" if x >= 0.995 else f"{sign}${x:,.3f}"


def _tok(n):
    return f"{n / 1e6:,.1f} M" if n >= 1e6 else f"{n / 1e3:,.0f} e" if n >= 1e3 else str(n)


def render_costs(all_projects):
    st.caption("Sara modellhívásainak költsége dollárban, a usage_log napló és az Anthropic listaárai alapján. "
               "A gyorsítótár bekapcsolása előtti hívások nincsenek a naplóban. Az összeg becslés: a pontos "
               "számla a Claude Console-ban látszik.")
    if hasattr(st, "segmented_control"):
        period = st.segmented_control("Időszak", list(COST_PERIODS), format_func=COST_PERIODS.get, default="7d",
                                      key="cost_period", label_visibility="collapsed") or "7d"
    else:
        period = st.radio("Időszak", list(COST_PERIODS), format_func=COST_PERIODS.get, index=1, horizontal=True,
                          key="cost_period", label_visibility="collapsed")
    start = costs.period_start(period)
    try:
        rows = db.get_usage(start.isoformat() if start else None)
    except db.DbError as e:
        st.warning("A költségnapló még nem érhető el. Futtasd le a schema.sql-t a Supabase SQL Editorban "
                   f"(ez hozza létre a usage_log táblát). Részletek: {e}")
        return
    if not rows:
        st.info("Ebben az időszakban még nincs naplózott hívás.")
        return
    names = {p["id"]: p["name"] for p in all_projects}
    data = costs.summarize(rows, names)
    t = data["total"]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Költség", _usd(t["actual"]), help=f"{t['calls']} modellhívás")
    c2.metric("Gyorsítótár nélkül", _usd(t["without"]))
    c3.metric("Megtakarítás", _usd(t["saved"]),
              f"{t['saved'] / t['without'] * 100:.0f}%" if t["without"] else None)
    c4.metric("Gyorsítótárból", f"{t['hit_rate'] * 100:.0f}%",
              help="A bemenő tokenek ekkora része jött a gyorsítótárból (minél nagyobb, annál olcsóbb).")
    st.caption(f"Bemenet {_tok(t['tokens_in'])} token (ebből gyorsítótárból {_tok(t['cache_read'])}, "
               f"eltárolva {_tok(t['cache_write'])}) · kimenet {_tok(t['tokens_out'])} token · "
               f"{t['web']} webes keresés")

    def table(bucket, label):
        lines = sorted(bucket.items(), key=lambda kv: kv[1]["actual"], reverse=True)
        st.dataframe([{label: k, "Költség ($)": round(v["actual"], 3), "Hívás": v["calls"],
                       "Token": _tok(v["tokens"]),
                       "Arány": f"{v['actual'] / t['actual'] * 100:.0f}%" if t["actual"] else "–"}
                      for k, v in lines], hide_index=True, **STRETCH)

    col_p, col_k = st.columns([3, 2])
    with col_p:
        st.markdown("**Projektenként**")
        table(data["by_project"], "Projekt")
    with col_k:
        st.markdown("**Tevékenységenként**")
        table(data["by_kind"], "Tevékenység")
    if len(data["by_day"]) > 1:
        st.markdown("**Naponta ($)**")
        st.bar_chart({"Költség ($)": dict(sorted(data["by_day"].items()))}, height=200)
    st.markdown("**A legdrágább hívások**")
    st.dataframe([{"Időpont": fmt_dt(c["time"]), "Projekt": c["project"], "Tevékenység": c["kind"],
                   "Bemenet": _tok(c["tokens_in"]), "Kimenet": _tok(c["tokens_out"]),
                   "Költség ($)": round(c["cost"], 3)} for c in data["top_calls"]],
                 hide_index=True, **STRETCH)


def render_search(active, closed):
    query = st.session_state.get("search_q", "").strip()
    terms = stx.query_terms(query)
    include_closed = bool(st.session_state.get("search_closed"))
    scope = active + (closed if include_closed else [])
    with st.container(height=FALLBACK_HEIGHT_PX, border=False, key="search_box"):
        if not terms:
            st.info("Írj be legalább egy, legalább kétbetűs keresőszót az oldalsáv keresőmezőjébe.")
            return
        try:
            docs, sums = db.search([p["id"] for p in scope], stx.filter_stems(terms))
        except db.DbError as e:
            st.error(f"A keresés nem sikerült: {e}")
            return
        # az adatbázis csak előszűr; a pontos, szavankénti egyeztetés itt történik
        docs = [d for d in docs if stx.matches(f"{d['name']}\n{d.get('content_text') or ''}", terms)]
        sums = [x for x in sums if stx.matches(f"{x.get('description') or ''}\n{x['content']}", terms)]
        where = "az aktív és a lezárt projektekben" if include_closed else "az aktív projektekben"
        if not docs and not sums:
            st.info(f"Nincs találat {where}." + ("" if include_closed or not closed else
                                                " A lezárt projektekben nem kerestem; ehhez pipáld be az oldalsávban."))
            return
        hit_projects = {h["project_id"] for h in docs + sums}
        st.caption(f"{len(docs) + len(sums)} találat {len(hit_projects)} projektben ({where}). Több keresőszónál "
                   "mindegyiknek szerepelnie kell; a ragozott alakokat is megtalálja.")

        for p in scope:
            p_sums = [s for s in sums if s["project_id"] == p["id"]]
            p_docs = [d for d in docs if d["project_id"] == p["id"]]
            if not p_sums and not p_docs:
                continue
            st.divider()
            col_name, col_go = st.columns([6, 2])
            col_name.markdown(f"**:material/folder: {p['name']}**" + ("" if p.get("status") == "active" else " (lezárt)"))
            if col_go.button(":material/arrow_forward: Ugrás a projektre", key=f"goto_{p['id']}", **STRETCH):
                open_project(p["id"])
                st.rerun()

            for s in p_sums:
                snips = stx.snippets(s["content"], terms)
                st.markdown(
                    '<div class="sara-hit"><span class="title">Összefoglaló</span>'
                    f'<span class="meta">{fmt_dt(s["created_at"])}</span>'
                    f'<div class="snip">{stx.highlight(s.get("description") or "", terms)}</div>'
                    + "".join(f'<div class="snip">{sn}</div>' for sn in snips) + "</div>",
                    unsafe_allow_html=True)
                if hit_opened("sum", s["id"]):
                    with st.container(border=True):
                        st.markdown(s["content"])

            p_docs.sort(key=lambda d: -stx.count_hits(f"{d['name']}\n{d.get('content_text') or ''}", terms))
            for d in p_docs:
                text = d.get("content_text") or ""
                snips = stx.snippets(text, terms)
                st.markdown(
                    f'<div class="sara-hit"><span class="title">{stx.highlight(d["name"], terms)}</span>'
                    f'<span class="meta">{fmt_dt(d.get("created_at"))[:10]}</span>'
                    + "".join(f'<div class="snip">{sn}</div>' for sn in snips) + "</div>",
                    unsafe_allow_html=True)
                if hit_opened("doc", d["id"]):
                    with st.container(border=True):
                        render_downloads(d, key_prefix="hit_")
                        render_doc_preview(d, text=text)


def index_old_documents(rows):
    """A régebben feltöltött dokumentumok szövegének kinyerése, hogy a keresés megtalálja őket."""
    progress = st.progress(0.0, text="Dokumentumok feldolgozása…")
    failed = []
    for n, row in enumerate(rows, 1):
        try:
            data = base64.b64decode(db.get_document_content(row["id"]))
            db.set_document_text(row["id"], row["name"], ex.extract_text(row["name"], row["mime_type"], data))
        except Exception as e:
            failed.append(f"{row['name']} ({e})")
        progress.progress(n / len(rows), text=f"Dokumentumok feldolgozása… {n}/{len(rows)}")
    return failed


# --- BELÉPÉS ---------------------------------------------------------------
# Az alkalmazás bizalmas anyagokat kezel, ezért jelszó nélkül semmit nem mutat és semmit nem tölt be.
# A jelszó a Streamlit titkos beállításai között van (APP_PASSWORD); ha nincs megadva, az alkalmazás zárva marad.
LOGIN_WINDOW_S = 15 * 60   # ennyi időn belül számoljuk a hibás próbálkozásokat
LOGIN_MAX_FAILS = 10       # ennyi hibás próbálkozás után az ablak végéig senki nem léphet be


@st.cache_resource
def _login_failures():
    return []  # a hibás próbálkozások időpontjai; minden látogatóra közös, így új böngészővel sem kerülhető meg


SESSION_S = 30 * 60   # a belépés ennyi ideig érvényes; minden művelet és a Hosszabbítás gomb újraindítja


def _session_sig(fingerprint, exp):
    return hmac.new(fingerprint.encode("utf-8"), f"sara-session:{exp}".encode("utf-8"), hashlib.sha256).hexdigest()[:32]


def _token_exp(fingerprint, token):
    """A címsorban tárolt munkamenet-jel lejárati ideje, vagy 0, ha a jel hamis vagy hiányzik."""
    try:
        exp_text, sig = str(token or "").split(".", 1)
        exp = int(exp_text)
    except ValueError:
        return 0
    return exp if hmac.compare_digest(sig, _session_sig(fingerprint, exp)) else 0


def _start_session(fingerprint):
    """A belépés (újra)indítása 30 percre. A jel a címsorba is bekerül, így az oldal újratöltése vagy a táblagép
    elalvása után sem kell újra jelszót megadni, amíg le nem jár. A jel a jelszóból képzett kulccsal aláírt,
    hamisítani nem lehet, és jelszócsere után érvénytelen."""
    exp = int(time.time()) + SESSION_S
    st.session_state["auth_ok"] = fingerprint
    st.session_state["auth_exp"] = exp
    st.query_params["s"] = f"{exp}.{_session_sig(fingerprint, exp)}"


def require_login():
    try:
        expected = str(st.secrets.get("APP_PASSWORD", "")).strip()
    except Exception:
        expected = ""
    # A belépés a jelszó lenyomatához kötött: ha a jelszót lecserélik, minden korábbi belépés érvényét veszti.
    fingerprint = hashlib.sha256(("sara:" + expected).encode("utf-8")).hexdigest()
    now = time.time()
    if expected and ((st.session_state.get("auth_ok") == fingerprint and st.session_state.get("auth_exp", 0) > now)
                     or _token_exp(fingerprint, st.query_params.get("s")) > now):
        _start_session(fingerprint)   # minden művelet meghosszabbítja a belépést
        return
    st.session_state.pop("auth_ok", None)
    st.session_state.pop("auth_exp", None)
    if "s" in st.query_params:
        del st.query_params["s"]
    st.markdown('<div class="sara-header">Sara asszisztens</div>', unsafe_allow_html=True)
    if not expected:
        st.error("Az alkalmazás zárva van, mert nincs beállítva belépési jelszó.")
        st.write("A tulajdonos a Streamlit felületén, az alkalmazás **Settings → Secrets** részében adhatja meg, "
                 "egy új sorban: `APP_PASSWORD = \"ide egy hosszú jelszó\"`. Mentés után az alkalmazás újraindul.")
        st.stop()
    failures = _login_failures()
    now = time.time()
    failures[:] = [t for t in failures if now - t < LOGIN_WINDOW_S]
    if len(failures) >= LOGIN_MAX_FAILS:
        st.error("Túl sok hibás próbálkozás. A belépés negyedóráig zárolva van.")
        st.stop()
    with st.form("login_form"):
        password = st.text_input("Jelszó", type="password")
        submitted = st.form_submit_button("Belépés", type="primary")
    if submitted:
        time.sleep(1)  # lassítja a próbálgatást
        if hmac.compare_digest(password.encode("utf-8"), expected.encode("utf-8")):
            _start_session(fingerprint)
            st.rerun()
        failures.append(time.time())
        st.error("Hibás jelszó.")
    st.stop()


# =============================================================================
# AZ OLDAL FELÉPÍTÉSE
# =============================================================================
inject_css()
require_login()

# Az adatbázis ellenőrzése munkamenetenként egyszer: megvannak-e az új oszlopok és táblák
if not st.session_state.get("schema_ok"):
    missing = db.check_schema()
    if missing:
        st.error("Az adatbázist frissíteni kell, mielőtt az alkalmazás használható. Hiányzik vagy nem érhető el: "
                 + "; ".join(missing))
        st.write("Nyisd meg a Supabase **SQL Editor** felületét, illeszd be az alábbi szkriptet (a `schema.sql` "
                 "fájl tartalmát), futtasd le, majd töltsd újra ezt az oldalt. A szkript meglévő adatot nem töröl.")
        try:
            with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "schema.sql"), encoding="utf-8") as fh:
                st.code(fh.read(), language="sql")
        except OSError:
            st.warning("A schema.sql fájl nincs az alkalmazás mellett; töltsd fel a repóba.")
        st.stop()
    st.session_state["schema_ok"] = True

try:
    all_projects = db.get_projects()
except db.DbError as e:
    st.error(f"A projektek nem tölthetők be: {e}")
    st.stop()

active_projects = [p for p in all_projects if p.get("status") == "active" and not is_special(p)]
closed_projects = [p for p in all_projects if p.get("status") != "active" and not is_special(p)]
templates_project = next((p for p in all_projects if is_templates(p)), None)
active_records = [p for p in all_projects if is_record(p) and p.get("status") == "active"]
projects_by_id = {p["id"]: p for p in all_projects}

# Induláskor a legutóbb használt aktív projekt nyílik meg
if st.session_state.get("current_project_id") not in projects_by_id:
    opened = [p for p in active_projects if p.get("last_opened_at")]
    if opened:
        st.session_state["current_project_id"] = max(opened, key=lambda p: str(p["last_opened_at"]))["id"]
    elif active_projects:
        st.session_state["current_project_id"] = active_projects[0]["id"]
    else:
        st.session_state["current_project_id"] = None
current_project = projects_by_id.get(st.session_state.get("current_project_id"))
current_id = current_project["id"] if current_project else None
current_closed = bool(current_project) and current_project.get("status") != "active"


def _search_changed():
    st.session_state["view"] = "search" if st.session_state.get("search_q", "").strip() else "project"
    st.session_state.pop("open_hits", None)


# --- OLDALSÁV (SIDEBAR) ---
with st.sidebar:
    _logo = logo_data_uri(BRAND)
    if _logo:
        st.markdown(f'<div class="sara-logo"><img src="{_logo}" alt="Contact-Line 2002"></div>', unsafe_allow_html=True)
    st.radio("Ki dolgozik?", ["SARA", "JEAN"], key="assistant", horizontal=True,
             format_func=lambda a: "SARA (alapos)" if a == "SARA" else "JEAN (olcsó)",
             help="SARA a pontosabb modellel dolgozik (Sonnet), JEAN a helyettese, sokkal olcsóbban (Haiku). "
                  "Egyszerű kérdésekhez, adatgyűjtéshez, rövid szövegekhez JEAN is elég; elemzéshez, "
                  "szerződéshez, bonyolult dokumentumhoz SARA. A dokumentumkivonatokat és a kereső "
                  "kollégát mindkettőnél ugyanaz az olcsó modell végzi.")
    st.text_input(":material/search: Keresés", key="search_q", on_change=_search_changed,
                  placeholder="keresés dokumentumokban és összefoglalókban",
                  help="Az aktív projektek dokumentumaiban és összefoglalóiban keres. Enterrel indul.")
    st.checkbox("Lezárt projektekben is", key="search_closed", on_change=_search_changed)

    if not st.session_state.get("index_done"):
        try:
            unindexed = db.unindexed_documents()
        except db.DbError:
            unindexed = []
        if not unindexed:
            st.session_state["index_done"] = True
        else:
            st.caption(f"{len(unindexed)} régebbi dokumentum tartalmában még nem lehet keresni.")
            if st.button("Feldolgozás a kereséshez", key="index_btn", **STRETCH):
                failed = index_old_documents(unindexed)
                if failed:
                    st.error("Nem sikerült feldolgozni: " + "; ".join(failed[:5]))
                else:
                    st.rerun()

    st.divider()
    side_title("Projektek", big=True)

    with st.form("project_form", clear_on_submit=True):
        new_project_name = st.text_input("Új téma / projekt neve:")
        submit_button = st.form_submit_button(":material/add: Projekt létrehozása", **STRETCH)
        if submit_button and new_project_name.strip():
            try:
                created = db.create_project(new_project_name, top_sort_order(all_projects))
            except db.DbError as e:
                st.error(f"A projekt létrehozása nem sikerült: {e}")
            else:
                if created:
                    st.session_state["current_project_id"] = created["id"]
                    st.session_state["view"] = "project"
                st.rerun()

    if active_projects:
        side_title("Aktív ügyek")
        for p in active_projects:
            if st.button(project_label(p), key=f"proj_{p['id']}",
                         type="primary" if p["id"] == current_id else "secondary", **STRETCH):
                open_project(p["id"])
                st.rerun()
        if current_project and not current_closed and not is_special(current_project):
            st.caption("A kiválasztott projekt:")
            index = [p["id"] for p in active_projects].index(current_id)
            col_up, col_down, col_close = st.columns([1, 1, 2.2])
            try:
                if col_up.button("▲", key="proj_up", help="Feljebb a listában", disabled=index == 0, **STRETCH):
                    move_project(active_projects, current_id, -1)
                    st.rerun()
                if col_down.button("▼", key="proj_down", help="Lejjebb a listában",
                                   disabled=index == len(active_projects) - 1, **STRETCH):
                    move_project(active_projects, current_id, +1)
                    st.rerun()
                if col_close.button(":material/lock: Lezárás", key="proj_close", **STRETCH,
                                    help="A projekt a lezártak közé kerül: megmarad és olvasható, de nem lehet benne kérdezni."):
                    db.update_project(current_id, status="closed")
                    st.rerun()
            except db.DbError as e:
                st.error(f"A művelet nem sikerült: {e}")
            st.caption(":material/settings: Kapcsolt törzsakták, fejléces papír, bizalmasság: a projekt "
                       "**Beállítások** fülén.")
    else:
        st.info("Nincs aktív projekt.")

    st.divider()
    side_title("Törzsakták", big=True)
    if not db.has_records():
        st.caption("A törzsaktákhoz (cég, ingatlan, személy, jármű) futtasd le a Supabase SQL Editorban a "
                   "schema.sql-t (az 1/e rész az új).")
    else:
        with st.expander(":material/add: Új törzsakta"):
            with st.form("record_form", clear_on_submit=True, border=False):
                rtype = st.selectbox("Fajta", list(records.RECORD_TYPES),
                                     format_func=lambda k: records.RECORD_TYPES[k][0])
                rname = st.text_input("Név", placeholder="pl. Contact-Line 2002 Bt. · Budapest, Fő u. 1. · Kovács Anna · ABC-123")
                if st.form_submit_button(":material/add: Létrehozás", **STRETCH) and rname.strip():
                    try:
                        created = db.create_project(rname, top_sort_order(all_projects), kind="torzs",
                                                    record_type=rtype, confidential=(rtype == "szemely"))
                    except db.DbError as e:
                        st.error(f"A létrehozás nem sikerült: {e}")
                    else:
                        if created:
                            open_project(created["id"])
                        st.rerun()
        for rtype, (label, icon, plural) in records.RECORD_TYPES.items():
            items = sorted([r for r in active_records if r.get("record_type") == rtype], key=lambda r: r["name"].lower())
            if not items:
                continue
            with st.expander(f"{icon} {plural} ({len(items)})", expanded=bool(current_project) and current_project in items):
                for r in items:
                    if st.button(project_label(r), key=f"rec_{r['id']}",
                                 type="primary" if r["id"] == current_id else "secondary", **STRETCH):
                        open_project(r["id"])
                        st.rerun()
        if st.button(":material/library_books: Iratminták", key="tpl_store_btn", **STRETCH,
                     type="primary" if is_templates(current_project) else "secondary"):
            try:
                tp = templates_project or db.create_project("Iratminták", top_sort_order(all_projects), kind="minta")
                if tp:
                    open_project(tp["id"])
                st.rerun()
            except db.DbError as e:
                st.error(f"Az Iratminták nem nyithatók meg: {e}")


    if closed_projects:
        with st.expander(f"Lezárt projektek ({len(closed_projects)})", expanded=current_closed):
            for p in closed_projects:
                if st.button(project_label(p), key=f"proj_{p['id']}",
                             type="primary" if p["id"] == current_id else "secondary", **STRETCH):
                    open_project(p["id"])
                    st.rerun()
            if current_closed:
                st.caption("A kiválasztott lezárt projekt:")
                col_reopen, col_delete = st.columns(2)
                if col_reopen.button(":material/lock_open: Újranyitás", key="proj_reopen", **STRETCH):
                    try:
                        db.update_project(current_id, status="active", sort_order=top_sort_order(all_projects))
                        st.rerun()
                    except db.DbError as e:
                        st.error(f"Az újranyitás nem sikerült: {e}")
                if col_delete.button(":material/delete: Törlés", key="proj_delete", **STRETCH):
                    st.session_state["delete_open"] = current_id
                    st.rerun()

with st.sidebar:
    st.divider()
    with st.popover(":material/verified: Adószám-ellenőrzés", **STRETCH):   # kézi NAV-lekérdezés, AI nélkül
        with st.form("tax_check_form", clear_on_submit=False, border=False):
            tax_in = st.text_input("Adószám", placeholder="12345678-2-42", key="tax_check_in")
            tax_go = st.form_submit_button("Ellenőrzés a NAV-nál", **STRETCH)
        if tax_go:
            try:
                st.session_state["tax_check_out"] = ("ok", navtax.report(*navtax.query(NAV_CFG, tax_in)))
            except Exception as e:
                st.session_state["tax_check_out"] = ("err", str(e))
        out = st.session_state.get("tax_check_out")
        if out:
            (st.markdown if out[0] == "ok" else st.warning)(out[1].replace("\n", "  \n"))
    with st.popover(":material/work: FEOR-kereső", **STRETCH):   # KSH FEOR-08 lista, AI nélkül
        feor_q = st.text_input("Munkakör vagy kód", placeholder="pl. könyvelő, titkár, 4111, 52",
                               key="feor_q", help="Ékezet nélkül is keres. Számmal a kódcsoportot listázza.")
        if feor_q.strip():
            feor_hits = feor.search(feor_q)
            if feor_hits:
                st.markdown("  \n".join(f"**{c}** – [{html.escape(feor.CODES[c])}]({feor.ksh_url(c)})"
                                        for c in feor_hits))
                st.caption("A megnevezésre kattintva a KSH leírása nyílik meg (feladatok, ide tartozó munkakörök).")
            else:
                st.info("Nincs találat. Próbálj rövidebb vagy általánosabb szót.")
    if st.button(":material/badge: Profil", key="profile_btn", **STRETCH):
        st.session_state["view"] = "profile"
        st.rerun()
    if st.button(":material/payments: Költségek", key="costs_btn", **STRETCH):
        st.session_state["view"] = "costs"
        st.rerun()
    if st.button(":material/logout: Kilépés", key="logout_btn", **STRETCH):
        st.session_state.clear()
        st.query_params.clear()
        st.rerun()


@st.fragment(run_every=30)
def session_timer():
    """A belépésből hátralévő idő a fejlécben (félpercenként frissül) és a Hosszabbítás gomb.
    Lejáratkor az egész oldal újrafut, és a jelszókérés jelenik meg."""
    left = int(st.session_state.get("auth_exp", 0) - time.time())
    if left <= 0:
        st.rerun(scope="app")
    c_time, c_btn = st.columns([1, 1.2], vertical_alignment="center", gap="small")
    minutes = (left + 59) // 60
    text = f":material/timer: {minutes} perc"
    c_time.caption(f":red[**{text}**]" if left < 5 * 60 else text)
    if c_btn.button("Hosszabbítás", key="extend_session", type="tertiary",
                    help="Ennyi idő múlva kér újra jelszót a program. Minden művelet és ez a gomb "
                         "újra 30 percre állítja."):
        _start_session(st.session_state["auth_ok"])
        st.rerun(scope="fragment")

# --- FEJLÉC (mindig látható) ---
search_query = st.session_state.get("search_q", "").strip()
in_search = st.session_state.get("view") == "search" and bool(search_query)
in_costs = st.session_state.get("view") == "costs"
in_profile = st.session_state.get("view") == "profile"

if in_costs:
    right_html = "Költségek"
elif in_profile:
    right_html = "Profil"
elif in_search:
    right_html = f'Keresés: <mark>{html.escape(search_query)}</mark>'
elif current_project:
    right_html = ((f'<span class="badge">{html.escape(record_kind_name(current_project))}</span> '
                   if is_record(current_project) else "") + html.escape(current_project["name"])
                  + ('<span class="badge">(bizalmas)</span>' if current_project.get("confidential") else "")
                  + ('<span class="badge">(lezárt)</span>' if current_closed else ""))
else:
    right_html = ""
with st.container(key="header_row"):   # alatta azúr vonal (CSS: .st-key-header_row)
    col_title, col_back, col_timer = st.columns([6, 2, 2.2], vertical_alignment="center")
col_title.markdown(
    f'<div class="sara-header">{ASSISTANT_NAME} asszisztens'
    + ('<span class="jean-flag">Sara helyett</span>' if ASSISTANT == "JEAN" else "")
    + (f'<span class="sep">|</span>{right_html}' if right_html else "") + "</div>",
    unsafe_allow_html=True)
with col_timer:
    session_timer()
if in_costs or in_profile:
    if col_back.button("← Vissza", key="back_from_costs", **STRETCH):
        st.session_state["view"] = "project"
        st.rerun()
elif search_query and not in_search:
    if col_back.button("← Vissza a találatokhoz", key="back_to_search", **STRETCH):
        st.session_state["view"] = "search"
        st.rerun()

# --- FŐKÉPERNYŐ ---
if in_profile:
    render_profile()
elif in_costs:
    render_costs(all_projects)
elif in_search:
    render_search(active_projects + active_records, closed_projects)
elif current_project:
    try:
        project_docs = db.get_documents(current_id)   # csak metaadatok; a tartalom igény szerint töltődik be
        messages = db.get_messages(current_id)
        summaries = db.get_summaries(current_id)
        attachments = db.get_attachments(current_id)
    except db.DbError as e:
        st.error(f"A projekt adatai nem tölthetők be: {e}")
        st.stop()
    try:
        tasks = db.get_tasks(current_id)
    except db.DbError:
        tasks = None  # a feladatok táblája még nincs létrehozva; a Feladatok fül megmondja, mi a teendő
    atts_by_msg = {}
    for att in attachments:
        atts_by_msg.setdefault(att["message_id"], []).append(att)

    if st.session_state.get("delete_open") == current_id and current_closed:
        delete_dialog(current_project)
    elif st.session_state.get("compact_open") == current_id and messages and not current_closed:
        compact_dialog(current_project, messages, atts_by_msg, summaries, project_docs)
    elif st.session_state.get("trim_open") == current_id and messages and not current_closed:
        trim_dialog(current_project, messages, atts_by_msg)
    elif st.session_state.get("sum_delete_open") is not None and not current_closed:
        doomed_summary = next((x for x in summaries if x["id"] == st.session_state["sum_delete_open"]), None)
        if doomed_summary:
            sum_delete_dialog(doomed_summary)
        else:
            _close_sum_delete_dialog()
    elif st.session_state.get("task_delete_open") is not None and tasks and not current_closed:
        doomed_task = next((x for x in tasks if x["id"] == st.session_state["task_delete_open"]), None)
        if doomed_task:
            task_delete_dialog(doomed_task)
        else:
            _close_task_delete_dialog()

    # A darabszám CSS-ből kerül a felirat mögé: a felirat szövege nem változhat, különben a Streamlit
    # a változáskor az első fülre ugrana vissza.
    tab_counts = [None, len(project_docs), len(summaries),
                  None if tasks is None else sum(1 for t in tasks if task_status(t) in TASK_ACTIVE)]
    if not is_record(current_project):
        st.html("<style>" + "".join(
            f'[role="tablist"] [role="tab"]:nth-of-type({i})::after {{ content: "({n})"; margin-left: .3em; }}'
            for i, n in enumerate(tab_counts, 1) if n is not None) + "</style>")
    if is_record(current_project):   # törzsakta: elöl az adatlap
        tab_counts = [None] + tab_counts
        st.html("<style>" + "".join(
            f'[role="tablist"] [role="tab"]:nth-of-type({i})::after {{ content: "({n})"; margin-left: .3em; }}'
            for i, n in enumerate(tab_counts, 1) if n is not None) + "</style>")
        tab_sheet, tab_chat, tab_docs, tab_sums, tab_tasks, tab_set = st.tabs(
            [":material/badge: Adatlap", ":material/chat: Beszélgetés", ":material/description: Dokumentumok",
             ":material/notes: Összefoglalók", ":material/checklist: Feladatok", ":material/settings: Beállítások"])
        with tab_sheet:
            render_record_sheet(current_project, project_docs, current_closed)
    else:
        tab_chat, tab_docs, tab_sums, tab_tasks, tab_set = st.tabs(
            [":material/chat: Beszélgetés", ":material/description: Dokumentumok", ":material/notes: Összefoglalók",
             ":material/checklist: Feladatok", ":material/settings: Beállítások"])  # állandó feliratok, különben fülváltás lenne
    with tab_chat:
        render_chat_tab(current_project, current_closed, project_docs, messages, atts_by_msg, summaries, tasks)
    with tab_docs:
        render_docs_tab(current_project, current_closed, project_docs)
    with tab_sums:
        render_summaries_tab(current_project, current_closed, summaries, messages)
    with tab_tasks:
        render_tasks_tab(current_project, current_closed, tasks)
    with tab_set:
        render_settings_tab(current_project, current_closed)
else:
    st.write("### Kezdéshez válassz vagy hozz létre egy projektet a bal oldali sávban!")
