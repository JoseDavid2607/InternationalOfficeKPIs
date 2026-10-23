# ===========================================================================
#  UASM · Seminars Analytics · multipage app
#  Same structure and visual language as the EIV and International Weeks
#  reports (cover with data sources, floating top navigation with side arrows,
#  English-only UI, axis titles on every chart).
#
#  Pages: Data Center (cover) · Dashboard · Seminar Search · Tables · Manage Data.
#  Data model: Seminarios / Conferencistas / Participantes joined by the
#  "Seminario" ID. Palette per domain: Seminars=teal, Speakers=blue,
#  Participants=pink.
#
#  "Manage Data" writes new rows straight into the three Excel files in Drive
#  (service account with Editor access), so the Excel files stay the single
#  source of truth.
# ===========================================================================
from __future__ import annotations

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import re
import io
import time
import datetime as dt
from typing import Optional, Dict, List, Tuple

try:
    from google.oauth2.service_account import Credentials
    _GSPREAD_OK = True
    _GSPREAD_IMPORT_ERR = None
except ImportError as _e:
    _GSPREAD_OK = False
    _GSPREAD_IMPORT_ERR = str(_e)

import requests
import os
import base64
import hmac
from copy import copy

# ── 1) CONFIGURACIÓN GLOBAL ────────────────────────────────────────────────
st.set_page_config(
    page_title="UASM Seminars Analytics",
    page_icon="🎙️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Paleta por dominio — idéntica al HTML original (Seminars=teal, Speakers=blue,
# Participants=pink), más ámbar para GHE.
SEM = "#0EA5A0"; SEM_SOFT = "#E3F7F5"
SPK = "#3B7DED"; SPK_SOFT = "#EAF2FE"
PAR = "#E0568B"; PAR_SOFT = "#FDEAF3"
AMBER = "#E0A537"
INK = "#1B2333"; INK_SOFT = "#5B6B85"; MUTED = "#8B97AC"
BG = "#F6F8FB"; LINE = "#E4E9F2"

st.markdown(
    "<style>"
    f".suite-header{{display:flex;flex-direction:column;align-items:center;"
    "padding:16px 24px 12px;"
    f"background:linear-gradient(135deg,{INK} 0%,{SEM} 100%);"
    "border-radius:12px;box-shadow:0 2px 8px rgba(14,165,160,.18);margin-bottom:14px;}"
    f".sh-super{{font-size:11px;font-weight:700;letter-spacing:2px;"
    "color:#fff !important;text-transform:uppercase;margin-bottom:2px;}"
    ".sh-title{font-size:26px !important;font-weight:800 !important;color:#fff !important;text-align:center !important;line-height:1.2 !important;margin:0 !important;padding:0 !important;}"
    ".sh-sub{font-size:13px;color:rgba(255,255,255,.85) !important;margin-top:4px;text-align:center;}"
    f".kv{{font-size:26px;font-weight:700;line-height:1.1;font-family:monospace;}}"
    f".kl{{font-size:11px;font-weight:600;color:{MUTED};"
    "text-transform:uppercase;letter-spacing:.5px;margin-top:3px;}"
    f"section[data-testid='stSidebar']{{background:{BG} !important;}}"
    "div[data-testid='stButton'] button{background:#FFFFFF !important;"
    f"border:1px solid {LINE} !important;border-radius:10px !important;"
    "color:#374151 !important;font-size:14px !important;"
    "font-weight:600 !important;height:48px !important;"
    "box-shadow:0 1px 3px rgba(0,0,0,.04) !important;}"
    f"div[data-testid='stButton'] button:hover{{background:{BG} !important;border-color:{SEM} !important;}}"
    "div.stDownloadButton>button{background:transparent !important;"
    "border:none !important;box-shadow:none !important;"
    f"color:{SEM} !important;font-size:13px !important;"
    "padding:0 !important;text-decoration:underline !important;}"
    f"thead th{{background:{SEM_SOFT} !important;color:{INK} !important;"
    "font-weight:700 !important;}"
    ".domain-badge{display:inline-flex;align-items:center;gap:8px;padding:4px 12px;"
    "border-radius:20px;font-size:12px;font-weight:700;margin-bottom:4px;}"
    f".domain-badge.sem{{background:{SEM_SOFT};color:{SEM};}}"
    f".domain-badge.spk{{background:{SPK_SOFT};color:{SPK};}}"
    f".domain-badge.par{{background:{PAR_SOFT};color:{PAR};}}"
    ".pending-card{background:#FAFBFC;border:1px dashed #D8DEE8;border-radius:12px;"
    "padding:22px 24px;margin-top:14px;color:#6B7280;font-size:13.5px;}"
    ".pending-card .tag{display:inline-block;font-family:monospace;font-size:10px;"
    "letter-spacing:.06em;text-transform:uppercase;color:#8B97AC;"
    "background:#EEF1F5;padding:3px 9px;border-radius:5px;margin-bottom:8px;}"
    ".st-key-cover_enter_btn div[data-testid='stButton'] button{"
    f"background:{SEM} !important;border:none !important;color:#fff !important;"
    "font-size:16px !important;font-weight:700 !important;height:52px !important;"
    "box-shadow:0 4px 14px rgba(14,165,160,.28) !important;}"
    f".st-key-cover_enter_btn div[data-testid='stButton'] button:hover{{background:#0B8783 !important;}}"
    ".st-key-nav_toggle{position:fixed !important;top:0.25rem;left:50%;transform:translateX(-50%);"
    "z-index:999999;width:auto !important;max-width:96vw;}"
    ".st-key-nav_toggle div[data-testid='stHorizontalBlock']{"
    "display:flex !important;flex-direction:row !important;flex-wrap:nowrap !important;width:auto !important;"
    "justify-content:center !important;align-items:center !important;gap:14px !important;"
    "overflow:visible;max-width:96vw;}"
    ".st-key-nav_toggle div[data-testid='stHorizontalBlock'] > div{"
    "flex:0 0 auto !important;width:auto !important;min-width:0 !important;}"
    ".st-key-nav_toggle div[data-testid='column'], .st-key-nav_toggle div[data-testid='stColumn']{"
    "width:auto !important;min-width:fit-content !important;flex:0 0 auto !important;}"
    ".st-key-nav_toggle div[data-testid='stPageLink']{width:auto !important;min-width:fit-content !important;overflow:visible !important;}"
    ".st-key-nav_toggle div[data-testid='stPageLink'] a{white-space:nowrap !important;overflow:visible !important;text-overflow:unset !important;width:auto !important;min-width:fit-content !important;font-size:13px !important;padding:6px 4px !important;}"
    ".st-key-nav_toggle div[data-testid='stPageLink'] a p{white-space:nowrap !important;overflow:visible !important;}"
    ".st-key-nav_toggle div[data-testid='stPopover']{width:auto !important;min-width:fit-content !important;}"
    ".st-key-nav_toggle div[data-testid='stPopover'] button{"
    "background:transparent !important;border:none !important;box-shadow:none !important;"
    f"color:{SEM} !important;font-size:13px !important;font-weight:400 !important;"
    "height:auto !important;width:auto !important;min-width:fit-content !important;"
    "padding:6px 4px !important;white-space:nowrap !important;}"
    ".st-key-nav_toggle div[data-testid='stPopover'] button svg{display:none !important;}"
    ".st-key-side_arrows{position:fixed !important;top:50%;left:0;right:0;transform:translateY(-50%);"
    "z-index:999998;width:100%;pointer-events:none;}"
    ".st-key-side_arrows div[data-testid='stHorizontalBlock']{"
    "display:flex !important;flex-direction:row !important;width:100% !important;"
    "justify-content:space-between !important;padding:0 6px;pointer-events:none;}"
    ".st-key-side_arrows div[data-testid='column'], .st-key-side_arrows div[data-testid='stColumn']{"
    "width:auto !important;flex:0 0 auto !important;pointer-events:auto;}"
    ".st-key-side_arrows a{"
    "display:flex !important;align-items:center;justify-content:center;"
    "background:transparent !important;border:none !important;box-shadow:none !important;"
    f"font-size:0 !important;font-weight:400;color:{SEM} !important;opacity:.55;text-decoration:none;"
    "transition:opacity .15s ease;}"
    ".st-key-side_arrows a p{font-size:26px !important;}"
    ".st-key-side_arrows a span:first-child{display:none !important;}"
    ".st-key-side_arrows a:hover{opacity:1;}"
    ".block-container{padding-top:3.2rem !important;}"
    ".st-key-go_to_datacenter_btn a{"
    "display:flex !important;align-items:center;justify-content:center;gap:6px;"
    f"background:{SEM} !important;border:none !important;border-radius:10px !important;"
    "color:#fff !important;font-weight:700 !important;height:44px !important;text-decoration:none !important;}"
    ".st-key-go_to_datacenter_btn a span{color:#fff !important;}"
    ".st-key-go_to_datacenter_btn a:hover{background:#0B8783 !important;}"
    ".hero{display:flex;flex-direction:column;align-items:center;gap:10px;padding:34px 24px 28px;"
    f"background:linear-gradient(135deg,{INK} 0%,{SEM} 100%);border-radius:16px;"
    "box-shadow:0 4px 18px rgba(14,165,160,.22);margin-bottom:6px;}"
    ".hero img{height:62px;max-width:80%;object-fit:contain;}"
    f".hero .h-super{{font-size:11px;font-weight:700;letter-spacing:3px;color:#fff !important;text-transform:uppercase;}}"
    ".hero .h-title{font-size:30px !important;font-weight:800 !important;color:#fff !important;text-align:center;line-height:1.15;}"
    "</style>",
    unsafe_allow_html=True,
)


# ── 2) HELPERS COMPARTIDOS ──────────────────────────────────────────────
def _render_header(title: str, subtitle: str = ""):
    sub = f'<div class="sh-sub">{subtitle}</div>' if subtitle else ""
    st.markdown(
        f'<div class="suite-header"><div class="sh-super">UASM · Seminars Analytics</div>'
        f'<div class="sh-title">{title}</div>{sub}</div>',
        unsafe_allow_html=True,
    )


def _domain_badge(label: str, cls: str):
    st.markdown(f'<span class="domain-badge {cls}">{label}</span>', unsafe_allow_html=True)


def _kpi(label: str, value, color: str = INK):
    st.markdown(
        f'<div class="kv" style="color:{color};">{value}</div><div class="kl">{label}</div>',
        unsafe_allow_html=True,
    )


def _pending_card(label: str, note: str = ""):  # kept for parity with the other reports
    note_html = f"<br>{note}" if note else ""
    st.markdown(
        f'<div class="pending-card"><span class="tag">Coming soon</span>'
        f'<p style="margin-top:8px;">{label}{note_html}</p></div>',
        unsafe_allow_html=True,
    )


def _xlsx_bytes(df: pd.DataFrame, sheet_name: str = "Data") -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf) as w:
        df.to_excel(w, index=False, sheet_name=sheet_name[:31])
    buf.seek(0)
    return buf.getvalue()


def _spacer(px: int):
    st.markdown(f"<div style='height:{px}px;'></div>", unsafe_allow_html=True)


def _axis_titles(fig, x: Optional[str] = None, y: Optional[str] = None):
    """Axis titles on every chart (same convention as the EIV / Weeks reports)."""
    if x:
        fig.update_xaxes(title_text=x, title_font=dict(size=11, color=MUTED), title_standoff=8)
    if y:
        fig.update_yaxes(title_text=y, title_font=dict(size=11, color=MUTED), title_standoff=8)
    return fig


def _find_asset_upwards(rel: str, max_up: int = 8) -> Optional[str]:
    """Finds e.g. 'imagenes/UASM_logo.png' by walking up from this script / the CWD
    (the folder lives at the repo root, several levels above this file)."""
    seen = set()
    for base in (os.path.dirname(os.path.abspath(__file__)), os.getcwd()):
        cur = base
        for _ in range(max_up + 1):
            full = os.path.join(cur, rel)
            if full not in seen:
                seen.add(full)
                if os.path.exists(full):
                    return full
            parent = os.path.dirname(cur)
            if parent == cur:
                break
            cur = parent
    return None


@st.cache_data(show_spinner=False)
def _b64_file(path: str) -> str:
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()


def count_by(items: List[dict], key_fn) -> List[Tuple[str, int]]:
    """Replica countBy() del HTML: cuenta ocurrencias y ordena desc por conteo."""
    m: Dict[str, int] = {}
    for x in items:
        k = key_fn(x) or "Unknown"
        m[k] = m.get(k, 0) + 1
    return sorted(m.items(), key=lambda kv: -kv[1])


def simple_table_df(pairs: List[Tuple[str, int]], label_header: str) -> pd.DataFrame:
    total = sum(n for _, n in pairs)
    rows = [{label_header: l, "Count": n, "Percent": f"{(n/total*100):.1f}%" if total else "0%"} for l, n in pairs]
    return pd.DataFrame(rows)


# ── 3) FILE IDs (Google Drive) — carpeta Reportes/Seminars ─────────────────
SEMINARIOS_FILE_ID = "1YFSUmZ95Md9qHoHvc114Eed-oA_uqjST"     # BD_seminarios.xlsx
CONFERENCISTAS_FILE_ID = "1oEgAJg2pXC6U1arLe2daMp7gPdRKgARm"  # BD_conferencistas.xlsx
PARTICIPANTES_FILE_ID = "18HDQDbaPqdTeEHFcPUDCRTw_bahhcNWa"   # BD_participantes.xlsx

_GSPREAD_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def _get_gspread_access_token() -> Optional[str]:
    if not _GSPREAD_OK or "gcp_service_account" not in st.secrets:
        return None
    try:
        from google.auth.transport.requests import Request as _GoogleAuthRequest
        creds = Credentials.from_service_account_info(
            dict(st.secrets["gcp_service_account"]), scopes=_GSPREAD_SCOPES
        )
        creds.refresh(_GoogleAuthRequest())
        return creds.token
    except Exception:
        return None


@st.cache_data(ttl=300)
def _download_drive_file_bytes(file_id: str) -> bytes:
    token = _get_gspread_access_token()
    if not token:
        if not _GSPREAD_OK:
            st.error(
                "📦 `google-auth` is not installed in this environment "
                f"(import failed with: `{_GSPREAD_IMPORT_ERR}`)."
            )
        elif "gcp_service_account" not in st.secrets:
            st.error(
                "🔑 `st.secrets['gcp_service_account']` was not found. Check "
                "Streamlit Cloud → your app → Settings → Secrets."
            )
        else:
            st.error("🔑 The `gcp_service_account` credentials could not be used to authenticate.")
        st.stop()

    url = f"https://www.googleapis.com/drive/v3/files/{file_id}?alt=media"
    headers = {"Authorization": f"Bearer {token}"}

    resp = None
    last_err = None
    for _ in range(3):
        try:
            resp = requests.get(url, timeout=60, headers=headers)
            if resp.status_code == 200:
                return resp.content
            last_err = RuntimeError(f"HTTP {resp.status_code}: {resp.text[:200]}")
        except Exception as e:
            last_err = e
        time.sleep(2)

    st.error(
        f"🌐 Could not download the Drive file ({file_id}) after several attempts: "
        f"{last_err}\n\nMake sure the file is shared with the service account email "
        "(Editor access is needed for the Manage Data page)."
    )
    st.stop()


# ── 4) NORMALIZACIÓN (misma lógica que normModalidad/normTipo/normGenero) ──
def norm_modalidad(m) -> str:
    if pd.isna(m) or str(m).strip() == "":
        return "Unknown"
    m = str(m).strip()
    low = m.lower()
    if low == "on-campus":
        return "On-Campus"
    if low == "online":
        return "Online"
    if low == "hybrid":
        return "Hybrid"
    return m


def norm_tipo(t) -> str:
    if pd.isna(t) or str(t).strip() == "":
        return "Unknown"
    t = str(t).strip()
    if t == "Externo":
        return "External"
    if t == "NA":
        return "Unknown"
    return t


def norm_genero(g) -> str:
    if pd.isna(g) or str(g).strip() == "":
        return "Unknown"
    g = str(g).strip().upper()
    if g in ("M", "MALE"):
        return "M"
    if g in ("F", "FEMALE"):
        return "F"
    return "Unknown"


def semester_of(date_val) -> Optional[str]:
    if date_val is None or pd.isna(date_val):
        return None
    try:
        month = pd.Timestamp(date_val).month
    except Exception:
        return None
    return "S1" if month <= 6 else "S2"


def fmt_date(date_val) -> Optional[str]:
    if date_val is None or pd.isna(date_val):
        return None
    try:
        return pd.Timestamp(date_val).strftime("%Y-%m-%d")
    except Exception:
        return None


# Source files mix Spanish and English country names ("Alemania", "Brazil", "UK"...).
# Everything is shown in English; the map is keyed by the English name.
COUNTRY_EN = {
    "alemania": "Germany", "brasil": "Brazil", "canada": "Canada", "canadá": "Canada",
    "dinamarca": "Denmark", "egipto": "Egypt", "eslovenia": "Slovenia", "españa": "Spain",
    "estados unidos": "United States", "eeuu": "United States", "usa": "United States",
    "francia": "France", "grecia": "Greece", "holanda": "Netherlands", "países bajos": "Netherlands",
    "inglaterra": "United Kingdom", "uk": "United Kingdom", "reino unido": "United Kingdom",
    "irlanda": "Ireland", "italia": "Italy", "mexico": "Mexico", "méxico": "Mexico", "peru": "Peru",
    "perú": "Peru", "polonia": "Poland", "turquía": "Turkey", "turquia": "Turkey", "suiza": "Switzerland",
    "suecia": "Sweden", "bélgica": "Belgium", "belgica": "Belgium", "japón": "Japan", "japon": "Japan",
    "corea del sur": "South Korea", "rusia": "Russia", "noruega": "Norway", "finlandia": "Finland",
    "hungría": "Hungary", "rumania": "Romania", "singapur": "Singapore", "sudáfrica": "South Africa",
    "panamá": "Panama", "panama": "Panama", "iran": "Iran", "irán": "Iran", "estonia": "Estonia",
    "australia": "Australia", "austria": "Austria", "israel": "Israel", "uruguay": "Uruguay",
    "ecuador": "Ecuador", "venezuela": "Venezuela", "bolivia": "Bolivia", "paraguay": "Paraguay",
}


def country_en(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)) or not str(v).strip():
        return "Unknown"
    t = str(v).strip()
    return COUNTRY_EN.get(t.lower(), t)


# English country name -> centroid, for the speaker bubble map.
COUNTRY_GEO = {
    "Germany": (51.16, 10.45), "Argentina": (-38.42, -63.62), "Brazil": (-14.24, -51.93),
    "Chile": (-35.68, -71.54), "China": (35.86, 104.20), "Colombia": (4.57, -74.30),
    "Costa Rica": (9.75, -83.75), "Denmark": (56.26, 9.50), "Egypt": (26.82, 30.80),
    "Slovenia": (46.15, 14.99), "Spain": (40.46, -3.75), "United States": (37.09, -95.71),
    "France": (46.23, 2.21), "Greece": (39.07, 21.82), "Netherlands": (52.13, 5.29),
    "India": (20.59, 78.96), "Iran": (32.43, 53.69), "Ireland": (53.14, -7.69),
    "Italy": (41.87, 12.57), "Mexico": (23.63, -102.55), "Peru": (-9.19, -75.02),
    "Poland": (51.92, 19.15), "Portugal": (39.40, -8.22), "Turkey": (38.96, 35.24),
    "United Kingdom": (55.38, -3.44), "Canada": (56.13, -106.35), "Switzerland": (46.82, 8.23),
    "Sweden": (60.13, 18.64), "Belgium": (50.50, 4.47), "Japan": (36.20, 138.25),
    "Australia": (-25.27, 133.78), "Austria": (47.52, 14.55), "Israel": (31.05, 34.85),
    "Estonia": (58.60, 25.01), "Uruguay": (-32.52, -55.77), "Ecuador": (-1.83, -78.18),
    "Venezuela": (6.42, -66.59), "Russia": (61.52, 105.32), "Norway": (60.47, 8.47),
    "Finland": (61.92, 25.75), "South Korea": (35.91, 127.77), "Singapore": (1.35, 103.82),
    "South Africa": (-30.56, 22.94), "Panama": (8.54, -80.78), "Romania": (45.94, 24.97),
    "Hungary": (47.16, 19.50),
}


def location_en(v) -> str:
    """'Hibrido' -> 'Hybrid', 'SD 809 y Zoom' -> 'SD 809 and Zoom'."""
    if v is None or (isinstance(v, float) and pd.isna(v)) or not str(v).strip():
        return "TBD"
    t = str(v).strip()
    t = re.sub(r"\bh[ií]brido\b", "Hybrid", t, flags=re.I)
    t = re.sub(r"\s+y\s+", " and ", t)
    return t


def norm_spk_gender(g) -> str:
    if g is None or (isinstance(g, float) and pd.isna(g)):
        return "Unknown"
    t = str(g).strip().lower()
    if t in ("m", "male", "masculino", "hombre"):
        return "Male"
    if t in ("f", "female", "femenino", "mujer"):
        return "Female"
    return "Unknown"


# Color por categoría de seminario — idéntico a CAT_COLORS del HTML.
CAT_COLORS = {
    "Recruitment research seminars": SEM,
    "Visiting Scholars Seminar Series": SPK,
    "Ágora": PAR,
    "GHE": AMBER,
    "Grupo de Investigación Historia y Empresariado": AMBER,
    "Other": MUTED,
}


def cat_color(c: str) -> str:
    return CAT_COLORS.get(c, MUTED)


# Logo por categoría — idéntico a CATEGORY_LOGOS del HTML (archivos en PICS/,
# junto a este script, igual que en el reporte original).
CATEGORY_LOGOS = {
    "Recruitment research seminars": "recruitment.png",
    "Visiting Scholars Seminar Series": "visiting.png",
    "Ágora": "agora.png",
    "GHE": "GHE.png",
    "Grupo de Investigación Historia y Empresariado": "GHE.png",
}


@st.cache_data(ttl=3600)
def _img_b64(filename: str) -> Optional[str]:
    path = os.path.join(os.path.dirname(__file__), "PICS", filename)
    try:
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except Exception:
        return None


def cat_logo_tag(category: str, height: int = 22) -> str:
    """Devuelve un <img> en base64 para el logo de la categoría, o '' si no
    existe (equivalente a catLogo() + <img src="PICS/..."> en el HTML)."""
    fname = CATEGORY_LOGOS.get(category)
    if not fname:
        return ""
    b64 = _img_b64(fname)
    if not b64:
        return ""
    return f'<img src="data:image/png;base64,{b64}" style="height:{height}px;vertical-align:middle;margin-right:8px;">'



# ── 5) CARGA + CONSTRUCCIÓN DEL DATASET (equivalente a buildDataset()) ─────
@st.cache_data(ttl=300)
def load_raw_sheets():
    """Reads the three source workbooks (each opened once) from Drive."""
    def _read(file_id: str, sheet: str) -> pd.DataFrame:
        df = pd.ExcelFile(io.BytesIO(_download_drive_file_bytes(file_id))).parse(sheet)
        df.columns = df.columns.astype(str).str.strip()
        return df
    return (_read(SEMINARIOS_FILE_ID, "seminarios"), _read(CONFERENCISTAS_FILE_ID, "conferencistas"),
            _read(PARTICIPANTES_FILE_ID, "participantes"))


def _clean(v, default=None):
    """Text cell -> stripped string; NaN / empty / 'NA' / Excel errors (#REF!) -> default."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return default
    t = str(v).strip()
    if not t or t.upper() == "NA" or t.startswith("#"):
        return default
    return t


def _parse_date(v):
    """Excel date or text ('19/3/2024', day first) -> Timestamp, else None."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    try:
        if isinstance(v, str):
            ts = pd.to_datetime(v.strip(), dayfirst=True, errors="coerce")
        else:
            ts = pd.Timestamp(v)
        return None if pd.isna(ts) else ts
    except Exception:
        return None


def _time_str(v) -> Optional[str]:
    """Excel time / datetime / '9:00 AM' text / day fraction -> 'HH:MM'."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    try:
        if isinstance(v, dt.time):
            return f"{v.hour:02d}:{v.minute:02d}"
        if isinstance(v, (int, float)) and 0 <= float(v) < 1:
            mins = round(float(v) * 24 * 60)
            return f"{mins // 60:02d}:{mins % 60:02d}"
        t = pd.Timestamp(v) if not isinstance(v, str) else pd.to_datetime(v.strip(), errors="coerce")
        if pd.isna(t):
            return _clean(v)
        return f"{t.hour:02d}:{t.minute:02d}"
    except Exception:
        return _clean(v)


@st.cache_data(ttl=300)
def build_dataset() -> List[dict]:
    """Joins seminars / speakers / participants by the 'Seminario' ID. Duplicated IDs
    (data-entry repeats) are disambiguated automatically: the second occurrence becomes
    'ID (2)' and keeps no participants (they can't be split)."""
    sem_df, conf_df, part_df = load_raw_sheets()

    participants_by_sem: Dict[str, List[dict]] = {}
    for r in part_df.to_dict("records"):
        sem = _clean(r.get("Seminario al que asistió"))
        nombre = _clean(r.get("Nombre"))
        if not sem or not nombre:
            continue
        participants_by_sem.setdefault(sem, []).append({
            "nombre": nombre, "correo": _clean(r.get("Correo"), ""),
            "genero": norm_genero(r.get("Género")), "tipo": norm_tipo(r.get("Tipo Asistente")),
            "modalidad": norm_modalidad(r.get("Modalidad")),
        })

    speakers_by_sem: Dict[str, List[dict]] = {}
    for r in conf_df.to_dict("records"):
        sem = _clean(r.get("Seminario que dirigió"))
        # 'Nombre Completo' is a formula in the workbook (it can show #REF!), so fall back to Nombre + Apellido.
        full = _clean(r.get("Nombre Completo")) or " ".join(
            x for x in (_clean(r.get("Nombre")), _clean(r.get("Apellido"))) if x)
        if not sem or not full:
            continue
        speakers_by_sem.setdefault(sem, []).append({
            "nombre": full, "genero": norm_spk_gender(r.get("Genero")),
            "nacionalidad": country_en(r.get("Nacionalidad")),
            "universidad": _clean(r.get("Nombre Universidad origen"), "Unknown"),
            "paisUniversidad": country_en(r.get("País Universidad origen")),
            "region": _clean(r.get("Ubicación"), "Unknown"),
        })

    seen_ids: Dict[str, int] = {}
    seminars: List[dict] = []
    for r in sem_df.to_dict("records"):
        raw_id = _clean(r.get("Seminario"))
        if not raw_id:
            continue
        seen_ids[raw_id] = seen_ids.get(raw_id, 0) + 1
        sid, data_note = raw_id, None
        if seen_ids[raw_id] > 1:
            sid = f"{raw_id} ({seen_ids[raw_id]})"
            data_note = (
                f'Source data reused ID "{raw_id}" for more than one seminar — participant '
                f"records could not be split automatically and are shown only under the "
                f"first occurrence."
            )
        speakers = [x for x in (_clean(r.get(f"Conferencista {i}")) for i in range(1, 6)) if x]
        ts = _parse_date(r.get("Fecha"))
        year = r.get("Año")
        try:
            year = int(year) if year is not None and not pd.isna(year) else (ts.year if ts is not None else None)
        except (TypeError, ValueError):
            year = ts.year if ts is not None else None
        plist = participants_by_sem.get(raw_id, []) if seen_ids[raw_id] == 1 else []
        seminars.append({
            "id": sid, "rawId": raw_id, "occ": seen_ids[raw_id], "lugarRaw": _clean(r.get("Lugar")),
            "categoria": _clean(r.get("Categoría"), "Other"),
            "nombre": _clean(r.get("Nombre del seminario"), "Untitled seminar"),
            "speakers": speakers, "fecha": ts.strftime("%Y-%m-%d") if ts is not None else None,
            "hora": _time_str(r.get("Hora")), "anio": year, "semestre": semester_of(ts),
            "modalidad": norm_modalidad(r.get("Modalidad")), "lugar": location_en(r.get("Lugar")),
            "area": _clean(r.get("Área"), "Unknown"), "participantsCount": len(plist), "participants": plist,
            "speakerDetails": speakers_by_sem.get(raw_id, []), "dataNote": data_note,
        })
    return seminars


def all_speakers(seminars: List[dict]) -> List[dict]:
    """Unique speakers (a speaker who led several seminars counts once)."""
    seen = set()
    out = []
    for s in seminars:
        for sp in s["speakerDetails"]:
            key = sp["nombre"].strip().lower()
            if key in seen:
                continue
            seen.add(key)
            out.append(sp)
    return out


def all_participants(seminars: List[dict]) -> List[dict]:
    out = []
    for s in seminars:
        out.extend(s["participants"])
    return out


def avg_held_participants(sems: List[dict]) -> float:
    today = dt.date.today().isoformat()
    held = [s for s in sems if s["fecha"] and s["fecha"] <= today and s["participantsCount"] > 0]
    if not held:
        return 0.0
    return round(sum(s["participantsCount"] for s in held) / len(held), 1)


def filtered_seminars(seminars: List[dict], year: str, semester: str) -> List[dict]:
    out = []
    for s in seminars:
        if year != "all" and str(s["anio"]) != str(year):
            continue
        if semester != "all" and s["semestre"] != semester:
            continue
        out.append(s)
    return out


def _donut(pairs: List[Tuple[str, int]], colors: List[str], show_all_labels: bool = False, height: int = 220):
    if not pairs:
        st.info("No data for this filter.")
        return
    labels = [p[0] for p in pairs]
    values = [p[1] for p in pairs]
    fig = go.Figure(go.Pie(
        labels=labels, values=values, hole=0.55,
        marker=dict(colors=colors[:len(labels)] if len(colors) >= len(labels) else colors * (len(labels) // len(colors) + 1),
                    line=dict(color="#fff", width=2)),
        textinfo="label+value", textposition="outside" if show_all_labels else "inside",
        textfont=dict(size=11, color="#fff" if not show_all_labels else INK),
    ))
    fig.update_layout(
        margin=dict(t=10, r=10, b=10, l=10), showlegend=show_all_labels,
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", height=height,
    )
    st.plotly_chart(fig, use_container_width=True)


def _hbar(pairs: List[Tuple[str, int]], color: str, height: int = 220,
          x_title: str = "Count", y_title: Optional[str] = None):
    if not pairs:
        st.info("No data for this filter.")
        return
    labels = [p[0] for p in pairs]
    values = [p[1] for p in pairs]
    fig = go.Figure(go.Bar(
        x=values, y=labels, orientation="h", marker=dict(color=color),
        text=[str(v) for v in values], textposition="outside", cliponaxis=False,
        textfont=dict(size=10.5, color=INK),
    ))
    fig.update_layout(
        margin=dict(t=6, r=30, b=44, l=150),
        xaxis=dict(gridcolor=LINE, tickfont=dict(color=INK_SOFT, size=9)),
        yaxis=dict(tickfont=dict(color=INK, size=10.5), automargin=True, autorange="reversed"),
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", height=height + 20,
    )
    _axis_titles(fig, x_title, y_title)
    st.plotly_chart(fig, use_container_width=True)


def _world_map(speakers: List[dict], focus_country: Optional[str] = None):
    by_nat = count_by(speakers, lambda s: s.get("nacionalidad"))
    unmapped = sum(c for n, c in by_nat if n not in COUNTRY_GEO)
    by_nat = [(n, c) for n, c in by_nat if n in COUNTRY_GEO]
    if not by_nat:
        st.info("No mappable nationalities for this filter.")
        return
    lat = [COUNTRY_GEO[n][0] for n, _ in by_nat]
    lon = [COUNTRY_GEO[n][1] for n, _ in by_nat]
    size = [10 + (c ** 0.5) * 7 for _, c in by_nat]
    fig = go.Figure(go.Scattergeo(
        lat=lat, lon=lon, mode="markers+text", text=[str(c) for _, c in by_nat],
        hovertext=[f"{n}: {c}" for n, c in by_nat], hoverinfo="text",
        textposition="middle center", textfont=dict(size=9, color="#fff"),
        marker=dict(size=size, color=SPK, line=dict(color="#fff", width=1), opacity=0.92),
    ))
    geo_kwargs = dict(
        scope="world", projection_type="natural earth", showland=True, landcolor="#EEF2F8",
        showcountries=True, countrycolor="#D8E0EC", showocean=True, oceancolor=BG, bgcolor="rgba(0,0,0,0)",
    )
    if focus_country and focus_country in COUNTRY_GEO:
        geo_kwargs["center"] = dict(lat=COUNTRY_GEO[focus_country][0], lon=COUNTRY_GEO[focus_country][1])
        geo_kwargs["projection_scale"] = 6
    else:
        geo_kwargs["center"] = dict(lat=15, lon=0)
        geo_kwargs["projection_scale"] = 1
    fig.update_geos(**geo_kwargs)
    fig.update_layout(margin=dict(t=6, r=6, b=6, l=6), paper_bgcolor="rgba(0,0,0,0)", height=380)
    st.plotly_chart(fig, use_container_width=True)
    if unmapped:
        st.caption(f"{unmapped} speaker(s) with a nationality that is not drawn on the map (no coordinates defined).")


# ── 6) PAGE — Data Center (cover) ────────────────────────────────────────
def page_datacenter():
    """Cover, same pattern as the EIV / Weeks reports: hero, title, 'Enter the Report',
    and the list of data sources (files are fetched from Drive only when requested)."""
    seminars = build_dataset()
    logo = _find_asset_upwards("imagenes/UASM_logo_blanco.png")
    logo_html = (f'<img src="data:image/png;base64,{_b64_file(logo)}" alt="UASM">' if logo else "")
    st.markdown(
        f'<div class="hero">{logo_html}<div class="h-super">UASM · Seminars Analytics</div>'
        '<div class="h-title">International Seminars</div></div>',
        unsafe_allow_html=True,
    )
    col_l, col_mid, col_r = st.columns([1, 2, 1])
    with col_mid:
        st.markdown(
            f'<div style="text-align:center;font-size:18px;color:{MUTED};margin:14px auto 0;line-height:1.5;">'
            "Seminars, speakers and participants across the Recruitment research, Visiting Scholars, "
            "Ágora and GHE series — plus a Manage Data page to keep the Excel sources up to date.</div>",
            unsafe_allow_html=True,
        )
        _spacer(26)
        with st.container(key="cover_enter_btn"):
            if st.button("Enter the Report →", key="cover_enter", use_container_width=True):
                st.switch_page(pages[1])

        st.markdown(
            f'<div style="text-align:center;margin-top:34px;font-family:monospace;'
            f'font-size:10.5px;letter-spacing:.06em;text-transform:uppercase;color:{MUTED};">'
            "Data sources used</div>",
            unsafe_allow_html=True,
        )
        files = [
            ("BD_seminarios.xlsx", "Seminar catalog — category, speakers, date, modality, area", SEMINARIOS_FILE_ID),
            ("BD_conferencistas.xlsx", "Speaker profiles — gender, nationality, origin university", CONFERENCISTAS_FILE_ID),
            ("BD_participantes.xlsx", "Attendance roster — name, email, gender, type, modality", PARTICIPANTES_FILE_ID),
        ]
        for fname, desc, fid in files:
            fc1, fc2 = st.columns([5, 1])
            with fc1:
                st.markdown(
                    f'<span style="font-size:12.5px;color:{INK};">{fname}</span><br>'
                    f'<span style="font-size:11px;color:{MUTED};">{desc}</span>',
                    unsafe_allow_html=True,
                )
            with fc2:
                # Fetched from Drive only when requested (keeps the cover fast).
                if st.session_state.get(f"cover_ready_{fname}"):
                    st.download_button("", data=_download_drive_file_bytes(fid), file_name=fname,
                                       key=f"cover_dl_{fname}", use_container_width=True,
                                       icon=":material/download:")
                elif st.button("", key=f"cover_get_{fname}", use_container_width=True,
                               icon=":material/download:", help="Fetch this file from Drive"):
                    st.session_state[f"cover_ready_{fname}"] = True
                    st.rerun()
        st.markdown(
            f'<div style="text-align:center;margin-top:14px;font-size:11.5px;color:{MUTED};">'
            f"{len(files)} sources loaded · {len(seminars)} seminars</div>",
            unsafe_allow_html=True,
        )


# ── 7) PAGE — Dashboard ────────────────────────────────────────────────
def page_dashboard():
    seminars = build_dataset()
    years = sorted(set(str(s["anio"]) for s in seminars if s["anio"] is not None))

    with st.sidebar:
        st.markdown("#### Filters")
        year = st.selectbox("Year", ["all"] + years[::-1], index=1 if years else 0, key="dash_year")
        semester = st.radio("Semester", ["all", "S1", "S2"], horizontal=True, key="dash_semester")

    _render_header(
        "International Seminars — Dashboard",
        "Seminars, speakers, and participants across the Recruitment, Visiting Scholars, "
        "Ágora, and GHE series.",
    )

    sems = filtered_seminars(seminars, year, semester)
    speakers = all_speakers(sems)
    parts = all_participants(sems)
    avg_part = avg_held_participants(sems)

    # ---------------- Seminars domain ----------------
    _domain_badge("📅 Seminars", "sem")
    col1, col2 = st.columns(2)
    with col1:
        _kpi("Seminars", len(sems), SEM)
    with col2:
        _kpi("Delivery Formats", len(set(s["modalidad"] for s in sems)), SEM)

    col_cat, col_mod, col_area = st.columns(3)
    cats = count_by(sems, lambda s: s["categoria"])
    with col_cat:
        st.markdown("##### By Category")
        max_cat = max((n for _, n in cats), default=1)
        for label, n in cats:
            pct = n / max_cat * 100 if max_cat else 0
            st.markdown(
                f'<div style="font-size:12.5px;display:flex;justify-content:space-between;">'
                f'<span>{label}</span><span style="font-family:monospace;color:{cat_color(label)};'
                f'font-weight:600;">{n}</span></div>'
                f'<div style="height:7px;border-radius:4px;background:{BG};overflow:hidden;margin-bottom:10px;">'
                f'<div style="height:100%;border-radius:4px;width:{pct}%;background:{cat_color(label)};"></div></div>',
                unsafe_allow_html=True,
            )
    with col_mod:
        st.markdown("##### Delivery Format")
        _donut(count_by(sems, lambda s: s["modalidad"]), [SEM, SPK, PAR, AMBER], height=230)
    with col_area:
        st.markdown("##### Area of Study")
        _hbar(count_by(sems, lambda s: s["area"]), SEM, height=230, x_title="Seminars", y_title="Area")

    st.markdown("---")

    # ---------------- Speakers domain ----------------
    _domain_badge("🎤 Speakers", "spk")
    col1, col2 = st.columns(2)
    with col1:
        _kpi("Speakers", len(speakers), SPK)
    with col2:
        _kpi("Universities Represented", len(set(s["universidad"] for s in speakers if s["universidad"] != "Unknown")), SPK)

    col_g, col_r = st.columns(2)
    with col_g:
        st.markdown("##### By Gender")
        _donut(count_by(speakers, lambda s: s["genero"]), [SPK, PAR, MUTED], height=200)
    with col_r:
        st.markdown("##### By Origin Region")
        _hbar(count_by(speakers, lambda s: s["region"]), SPK, height=200, x_title="Speakers", y_title="Origin region")

    st.markdown("##### Speaker Nationalities & Universities")
    st.caption("Choose a country to focus the map and see its universities.")
    all_nats = sorted(set(s["nacionalidad"] for s in speakers if s.get("nacionalidad") in COUNTRY_GEO))
    col_map, col_list = st.columns(2)
    with col_map:
        focus = st.selectbox("Focus country", ["All Nationalities"] + all_nats, key="dash_focus_country")
        focus_val = None if focus == "All Nationalities" else focus
        _world_map(speakers, focus_val)
    with col_list:
        scoped = [s for s in speakers if s["nacionalidad"] == focus_val] if focus_val else speakers
        uni_pairs = count_by(scoped, lambda s: s["universidad"])
        st.markdown(f"**{focus}**")
        uni_df = pd.DataFrame(uni_pairs, columns=["University", "Speakers"])
        st.dataframe(uni_df, use_container_width=True, hide_index=True, height=330)

    st.markdown("---")

    # ---------------- Participants domain ----------------
    _domain_badge("👥 Participants", "par")
    col1, col2 = st.columns(2)
    with col1:
        _kpi("Participants", len(parts), PAR)
    with col2:
        _kpi("Avg. Participants / Seminar", avg_part, PAR)

    gender_counts = count_by(parts, lambda p: p["genero"])
    icon_html = ""
    label_map = {"M": "Male", "F": "Female", "Unknown": "Unknown"}
    for g, n in gender_counts:
        color = PAR if g == "F" else (SPK if g == "M" else MUTED)
        symbol = "♀" if g == "F" else ("♂" if g == "M" else "?")
        icon_html += (
            f'<div style="display:inline-flex;align-items:center;gap:8px;background:#fff;'
            f'border:1px solid {LINE};border-radius:10px;padding:8px 14px;margin-right:10px;">'
            f'<span style="width:26px;height:26px;border-radius:50%;background:{color};color:#fff;'
            f'display:flex;align-items:center;justify-content:center;font-size:13px;">{symbol}</span>'
            f'<span style="font-family:monospace;font-weight:700;font-size:14px;">{n}</span>'
            f'<span style="font-size:11px;color:{MUTED};">{label_map.get(g, g)}</span></div>'
        )
    st.markdown(icon_html, unsafe_allow_html=True)

    col_pm, col_pt = st.columns(2)
    with col_pm:
        st.markdown("##### By Modality")
        _donut(count_by(parts, lambda p: p["modalidad"]), [SEM, SPK, AMBER], height=200)
    with col_pt:
        st.markdown("##### By Attendee Type")
        _hbar(count_by(parts, lambda p: p["tipo"]), PAR, height=200, x_title="Participants", y_title="Attendee type")

    st.markdown("---")
    _render_pdf_download(sems, speakers, parts, avg_part,
                         sig=f"{year}|{semester}|{len(sems)}|{len(parts)}|{len(speakers)}")


# ── 8) PAGE — Seminar Search ──────────────────────────────────────────
def page_search():
    seminars = build_dataset()
    _render_header(
        "Seminar Search",
        "Search by name or speaker, and narrow down with category, year, or date if helpful. "
        "Only seminars already held are searchable.",
    )

    cats = sorted(set(s["categoria"] for s in seminars))
    years = sorted(set(str(s["anio"]) for s in seminars if s["anio"] is not None))
    days = sorted(set(s["fecha"] for s in seminars if s["fecha"]))
    if days:
        st.caption(f"{len(days)} distinct seminar days on record ({days[0]} to {days[-1]})")

    query = st.text_input("Search by name, speaker…", key="search_query")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        cat_sel = st.selectbox("Category", ["All"] + cats, key="search_cat")
    with col2:
        year_sel = st.selectbox("Year", ["All"] + years[::-1], key="search_year")
    with col3:
        date_from = st.date_input("From", value=None, key="search_from")
    with col4:
        date_to = st.date_input("To", value=None, key="search_to")

    has_filter = bool(query or cat_sel != "All" or year_sel != "All" or date_from or date_to)

    if not has_filter:
        st.info("Type a name/speaker, or use the filters above, to find seminars.")
        return

    today = dt.date.today().isoformat()
    q = query.strip().lower()
    results = [s for s in seminars if s["fecha"] and s["fecha"] <= today]
    if q:
        results = [
            s for s in results
            if q in (s["nombre"] or "").lower() or q in (s["categoria"] or "").lower()
            or q in " ".join(s["speakers"]).lower()
        ]
    if cat_sel != "All":
        results = [s for s in results if s["categoria"] == cat_sel]
    if year_sel != "All":
        results = [s for s in results if str(s["anio"]) == year_sel]
    if date_from:
        results = [s for s in results if s["fecha"] and s["fecha"] >= date_from.isoformat()]
    if date_to:
        results = [s for s in results if s["fecha"] and s["fecha"] <= date_to.isoformat()]
    results.sort(key=lambda s: s["fecha"] or "", reverse=True)

    if not results:
        st.warning("No seminars match this search.")
    else:
        st.caption(f"{len(results)} seminar{'s' if len(results) != 1 else ''} found.")
        options = [f'{s["nombre"]}  ·  {s["categoria"]} · {s["fecha"] or "TBD"} · {s["participantsCount"]} attendees'
                   for s in results]
        idx = st.radio("Results", options, key="search_result_pick", label_visibility="collapsed")
        selected = results[options.index(idx)]
        _render_seminar_detail(selected)

    # ---- Upcoming seminars ----
    upcoming = [s for s in seminars if s["fecha"] and s["fecha"] > today]
    upcoming.sort(key=lambda s: s["fecha"])
    if upcoming:
        st.markdown("---")
        st.markdown("### Upcoming Seminars")
        st.caption("Not yet held — attendance isn't recorded, so these are shown separately from search.")
        for s in upcoming:
            speaker_str = f' · {", ".join(s["speakers"])}' if s["speakers"] else ""
            st.markdown(
                f'**{s["nombre"]}**  \n<span style="font-family:monospace;font-size:11px;color:{MUTED};">'
                f'{s["categoria"]} · {s["fecha"]}{speaker_str}</span>',
                unsafe_allow_html=True,
            )


def _render_seminar_detail(s: dict):
    st.markdown("---")
    st.markdown(f'<span class="domain-badge sem">{s["categoria"]}</span>', unsafe_allow_html=True)
    st.markdown(f"## {s['nombre']}")

    fcol1, fcol2, fcol3, fcol4 = st.columns(4)
    with fcol1:
        st.markdown(f"**Speaker(s)**  \n{', '.join(s['speakers']) or 'TBD'}")
        st.markdown(f"**Date**  \n{s['fecha'] or 'TBD'}")
    with fcol2:
        st.markdown(f"**Time**  \n{s['hora'] or 'TBD'}")
        st.markdown(f"**Location**  \n{s['lugar'] or 'TBD'}")
    with fcol3:
        st.markdown(f"**Modality**  \n{s['modalidad']}")
        st.markdown(f"**Area**  \n{s['area']}")
    with fcol4:
        st.markdown(f"**Attendees**  \n{s['participantsCount']}")

    if s["dataNote"]:
        st.warning(f"⚠️ {s['dataNote']}")

    col_g, col_t, col_m = st.columns(3)
    with col_g:
        st.markdown("##### By Gender")
        _donut(count_by(s["participants"], lambda p: p["genero"]), [SPK, PAR, MUTED], show_all_labels=True, height=230)
    with col_t:
        st.markdown("##### By Attendee Type")
        _donut(count_by(s["participants"], lambda p: p["tipo"]), [SEM, SPK, AMBER, PAR, MUTED], show_all_labels=True, height=230)
    with col_m:
        st.markdown("##### By Modality")
        _donut(count_by(s["participants"], lambda p: p["modalidad"]), [SEM, SPK, AMBER], show_all_labels=True, height=230)

    st.markdown(f"##### Participants ({len(s['participants'])})")
    if s["participants"]:
        part_df = pd.DataFrame(s["participants"])[["nombre", "correo", "genero", "tipo", "modalidad"]]
        part_df.columns = ["Name", "Email", "Gender", "Attendee Type", "Modality"]
    else:
        part_df = pd.DataFrame(columns=["Name", "Email", "Gender", "Attendee Type", "Modality"])
    st.dataframe(part_df, use_container_width=True, hide_index=True)
    st.download_button(
        "Download as Excel", data=_xlsx_bytes(part_df, "Participants"),
        file_name=f"{s['id']}_participants.xlsx".replace(" ", "_"), key=f"dl_sem_{s['id']}",
        icon=":material/download:",
    )


# ── 9) PAGE — Tables ───────────────────────────────────────────────────
def page_tables():
    seminars = build_dataset()
    years = sorted(set(str(s["anio"]) for s in seminars if s["anio"] is not None))

    with st.sidebar:
        st.markdown("#### Filters")
        year = st.selectbox("Year", ["all"] + years[::-1], index=1 if years else 0, key="tables_year")
        semester = st.radio("Semester", ["all", "S1", "S2"], horizontal=True, key="tables_semester")

    _render_header(
        "Tables",
        "Full seminar listing, plus grouped summaries for seminars, speakers, and participants "
        "— each downloadable as Excel.",
    )

    sems = filtered_seminars(seminars, year, semester)
    speakers = all_speakers(sems)
    parts = all_participants(sems)
    sorted_sems = sorted(sems, key=lambda s: s["fecha"] or "", reverse=True)

    # ---------------- Seminars domain ----------------
    _domain_badge("📅 Seminars", "sem")
    st.markdown("##### Category, Speaker(s) & Total Participants")
    cat_groups = count_by(sems, lambda s: s["categoria"])
    rows_disp = []
    for cat, _ in cat_groups:
        in_cat = [s for s in sorted_sems if s["categoria"] == cat]
        cat_total = sum(s["participantsCount"] for s in in_cat)
        rows_disp.append({"Category": f"— {cat} ({len(in_cat)} seminars) —", "Seminar": "", "Speaker(s)": "",
                           "Date": "", "Modality": "", "Area": "", "Participants": ""})
        for s in in_cat:
            rows_disp.append({
                "Category": "", "Seminar": s["nombre"], "Speaker(s)": ", ".join(s["speakers"]) or "TBD",
                "Date": s["fecha"] or "TBD", "Modality": s["modalidad"], "Area": s["area"],
                "Participants": str(s["participantsCount"]),
            })
        rows_disp.append({"Category": "", "Seminar": "", "Speaker(s)": "", "Date": "", "Modality": "",
                           "Area": f"Total participants — {cat}", "Participants": str(cat_total)})
    df_seminars = pd.DataFrame(rows_disp)
    st.dataframe(df_seminars, use_container_width=True, hide_index=True, height=420)

    df_seminars_flat = pd.DataFrame([{
        "Category": s["categoria"], "Seminar": s["nombre"], "Speakers": ", ".join(s["speakers"]) or "TBD",
        "Date": s["fecha"] or "TBD", "Modality": s["modalidad"], "Area": s["area"],
        "Participants": s["participantsCount"],
    } for s in sorted_sems])
    st.download_button(
        "Download as Excel", data=_xlsx_bytes(df_seminars_flat, "Seminars"),
        file_name="tblSeminars.xlsx", key="dl_tbl_seminars",
    )

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("##### By Modality")
        df = simple_table_df(count_by(sems, lambda s: s["modalidad"]), "Modality")
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.download_button("Download as Excel", data=_xlsx_bytes(df, "Modality"),
                            file_name="tblSemModality.xlsx", key="dl_tbl_sem_mod")
    with col2:
        st.markdown("##### By Category")
        df = simple_table_df(count_by(sems, lambda s: s["categoria"]), "Category")
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.download_button("Download as Excel", data=_xlsx_bytes(df, "Category"),
                            file_name="tblSemCategory.xlsx", key="dl_tbl_sem_cat")
    with col3:
        st.markdown("##### By Area")
        df = simple_table_df(count_by(sems, lambda s: s["area"]), "Area")
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.download_button("Download as Excel", data=_xlsx_bytes(df, "Area"),
                            file_name="tblSemArea.xlsx", key="dl_tbl_sem_area")

    st.markdown("---")

    # ---------------- Speakers domain ----------------
    _domain_badge("🎤 Speakers", "spk")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("##### By Origin Region")
        df = simple_table_df(count_by(speakers, lambda s: s["region"]), "Region")
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.download_button("Download as Excel", data=_xlsx_bytes(df, "Region"),
                            file_name="tblSpkRegion.xlsx", key="dl_tbl_spk_region")
    with col2:
        st.markdown("##### By Gender")
        df = simple_table_df(count_by(speakers, lambda s: s["genero"]), "Gender")
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.download_button("Download as Excel", data=_xlsx_bytes(df, "Gender"),
                            file_name="tblSpkGender.xlsx", key="dl_tbl_spk_gender")
    with col3:
        st.markdown("##### By Nationality")
        df = simple_table_df(count_by(speakers, lambda s: s["nacionalidad"]), "Nationality")
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.download_button("Download as Excel", data=_xlsx_bytes(df, "Nationality"),
                            file_name="tblSpkNationality.xlsx", key="dl_tbl_spk_nat")

    st.markdown("##### By University")
    df = simple_table_df(count_by(speakers, lambda s: s["universidad"]), "University")
    st.dataframe(df, use_container_width=True, hide_index=True)
    st.download_button("Download as Excel", data=_xlsx_bytes(df, "University"),
                        file_name="tblSpkUniversity.xlsx", key="dl_tbl_spk_univ")

    st.markdown("---")

    # ---------------- Participants domain ----------------
    _domain_badge("👥 Participants", "par")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("##### By Modality")
        df = simple_table_df(count_by(parts, lambda p: p["modalidad"]), "Modality")
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.download_button("Download as Excel", data=_xlsx_bytes(df, "Modality"),
                            file_name="tblPartModality.xlsx", key="dl_tbl_part_mod")
    with col2:
        st.markdown("##### By Gender")
        df = simple_table_df(count_by(parts, lambda p: p["genero"]), "Gender")
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.download_button("Download as Excel", data=_xlsx_bytes(df, "Gender"),
                            file_name="tblPartGender.xlsx", key="dl_tbl_part_gender")
    with col3:
        st.markdown("##### By Type")
        df = simple_table_df(count_by(parts, lambda p: p["tipo"]), "Type")
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.download_button("Download as Excel", data=_xlsx_bytes(df, "Type"),
                            file_name="tblPartType.xlsx", key="dl_tbl_part_type")


# ── 10) PDF Summary Report ───────────────────────────────────────────────
def _build_summary_pdf(sems: List[dict], speakers: List[dict], parts: List[dict], avg_part: float) -> bytes:
    """Reporte PDF de resumen — mismo contenido numérico que la versión HTML
    (KPIs + desgloses + listado completo de seminarios por categoría). Sin
    miniaturas de las gráficas incrustadas: requerirían Chrome headless vía
    kaleido, una dependencia fea de mantener en un despliegue en la nube."""
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.lib import colors as rl_colors
    from reportlab.pdfgen import canvas as rl_canvas

    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=A4)
    W, H = A4
    mg = 14 * mm

    def header_band(title: str, color_hex: str):
        c.setFillColor(rl_colors.HexColor(color_hex))
        c.rect(0, H - 26 * mm, W, 26 * mm, fill=1, stroke=0)
        c.setFillColor(rl_colors.white)
        c.setFont("Helvetica-Bold", 15)
        c.drawString(mg, H - 15 * mm, title)
        c.setFont("Helvetica", 9)
        c.drawString(mg, H - 21 * mm, "School of Management · Universidad de los Andes")

    def section_title(text: str, color_hex: str, y: float) -> float:
        c.setFillColor(rl_colors.HexColor(color_hex))
        c.setFont("Helvetica-Bold", 12)
        c.drawString(mg, y, text)
        return y - 8 * mm

    def kpi_box(x, y, w, h, value, label, color_hex):
        c.setFillColor(rl_colors.HexColor("#F6F8FB"))
        c.roundRect(x, y - h, w, h, 3, fill=1, stroke=0)
        c.setFillColor(rl_colors.HexColor(color_hex))
        c.setFont("Helvetica-Bold", 15)
        c.drawCentredString(x + w / 2, y - h / 2 - 2, str(value))
        c.setFillColor(rl_colors.HexColor("#8B97AC"))
        c.setFont("Helvetica", 7)
        c.drawCentredString(x + w / 2, y - h / 2 - 9, label)

    def pdf_list(pairs, x, y, w) -> float:
        total = sum(n for _, n in pairs)
        for label, n in pairs:
            pct = f"{(n/total*100):.1f}" if total else "0"
            c.setFillColor(rl_colors.HexColor("#1B2333"))
            c.setFont("Helvetica", 8.5)
            c.drawString(x, y, str(label))
            c.setFont("Helvetica-Bold", 8.5)
            c.drawRightString(x + w, y, f"{n} ({pct}%)")
            y -= 6 * mm
        return y

    cW = W - 2 * mg

    # ---- Page 1: Seminars ----
    header_band("International Seminars — Summary Report", SEM)
    y = H - 36 * mm
    y = section_title("SEMINARS", SEM, y)
    kpi_box(mg, y, (cW - 8) / 2, 20 * mm, len(sems), "Total Seminars", SEM)
    kpi_box(mg + (cW - 8) / 2 + 8, y, (cW - 8) / 2, 20 * mm,
            len(set(s["modalidad"] for s in sems)), "Delivery Formats", SEM)
    y -= 28 * mm
    y = pdf_list(count_by(sems, lambda s: s["categoria"]), mg, y, cW)

    # ---- Page 2: Speakers ----
    c.showPage()
    header_band("Speakers", SPK)
    y = H - 36 * mm
    y = section_title("SPEAKERS", SPK, y)
    kpi_box(mg, y, (cW - 8) / 2, 20 * mm, len(speakers), "Total Speakers", SPK)
    kpi_box(mg + (cW - 8) / 2 + 8, y, (cW - 8) / 2, 20 * mm,
            len(set(s["universidad"] for s in speakers if s["universidad"] != "Unknown")), "Universities", SPK)
    y -= 28 * mm
    y = pdf_list(count_by(speakers, lambda s: s["region"]), mg, y, cW)

    # ---- Page 3: Participants ----
    c.showPage()
    header_band("Participants", PAR)
    y = H - 36 * mm
    y = section_title("PARTICIPANTS", PAR, y)
    kpi_box(mg, y, (cW - 8) / 2, 20 * mm, len(parts), "Total Participants", PAR)
    kpi_box(mg + (cW - 8) / 2 + 8, y, (cW - 8) / 2, 20 * mm, avg_part, "Avg. / Seminar", PAR)
    y -= 28 * mm
    y = pdf_list(count_by(parts, lambda p: p["tipo"]), mg, y, cW)

    # ---- Page 4+: Full seminar listing by category ----
    c.showPage()
    y = H - 20 * mm
    c.setFillColor(rl_colors.HexColor(INK))
    c.setFont("Helvetica-Bold", 11)
    c.drawString(mg, y, "SEMINARS — CATEGORY, SPEAKER(S) & TOTAL PARTICIPANTS")
    y -= 9 * mm

    sorted_sems = sorted(sems, key=lambda s: s["fecha"] or "", reverse=True)

    def check_page(needed):
        nonlocal y
        if y - needed < 16 * mm:
            c.showPage()
            y = H - 20 * mm

    for cat, _ in count_by(sems, lambda s: s["categoria"]):
        in_cat = [s for s in sorted_sems if s["categoria"] == cat]
        cat_total = sum(s["participantsCount"] for s in in_cat)
        check_page(14 * mm)
        c.setFillColor(rl_colors.HexColor("#E6F7F5"))
        c.rect(mg, y - 5 * mm, cW, 7 * mm, fill=1, stroke=0)
        c.setFillColor(rl_colors.HexColor(INK))
        c.setFont("Helvetica-Bold", 8.5)
        c.drawString(mg + 3, y - 0.5 * mm, f"{cat.upper()} — {len(in_cat)} seminar{'s' if len(in_cat) != 1 else ''}")
        y -= 10 * mm
        for s in in_cat:
            check_page(12 * mm)
            c.setFont("Helvetica", 7.5)
            c.setFillColor(rl_colors.HexColor(INK))
            c.drawString(mg, y, str(s["nombre"])[:95])
            c.drawRightString(mg + cW, y, str(s["participantsCount"]))
            y -= 4.2 * mm
            c.setFont("Helvetica-Oblique", 6.8)
            c.setFillColor(rl_colors.HexColor("#787878"))
            speaker_text = "Speaker(s): " + (", ".join(s["speakers"]) or "TBD")
            c.drawString(mg + 2, y, speaker_text[:110])
            y -= 6 * mm
        check_page(9 * mm)
        c.setStrokeColor(rl_colors.HexColor("#DCDCDC"))
        c.line(mg, y, mg + cW, y)
        y -= 4 * mm
        c.setFillColor(rl_colors.HexColor(SEM))
        c.setFont("Helvetica-Bold", 8)
        c.drawRightString(mg + cW - 2 * mm, y, f"Total participants — {cat}: {cat_total}")
        y -= 9 * mm

    c.setFont("Helvetica", 7)
    c.setFillColor(rl_colors.HexColor("#8B97AC"))
    c.drawString(mg, 10 * mm, f"Generated {dt.date.today().isoformat()}")
    c.save()
    buf.seek(0)
    return buf.getvalue()


def _render_pdf_download(sems, speakers, parts, avg_part, sig: str):
    """The PDF is built only when requested; the last one is kept in session_state
    (keyed by the current filters + data) so reruns don't rebuild it."""
    cache = st.session_state.get("_sem_pdf_cache", {})
    if cache.get("sig") == sig:
        st.download_button(
            "Download PDF report", data=cache["data"],
            file_name="International_Seminars_Report.pdf", mime="application/pdf",
            key="dl_summary_pdf", icon=":material/picture_as_pdf:",
        )
    elif st.button("Prepare PDF report", key="prep_summary_pdf", icon=":material/picture_as_pdf:"):
        with st.spinner("Preparing report…"):
            try:
                st.session_state["_sem_pdf_cache"] = {
                    "sig": sig, "data": _build_summary_pdf(sems, speakers, parts, avg_part)}
            except Exception as e:
                st.error(f"PDF export unavailable: {e}")
                return
        st.rerun()


# ── 11) WRITE-BACK TO THE EXCEL FILES IN DRIVE ───────────────────────────
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# One entry per source workbook: sheet name + how a row is identified (to avoid duplicates).
WRITE_TARGETS = {
    "seminar":     {"file_id": SEMINARIOS_FILE_ID,     "sheet": "seminarios"},
    "speaker":     {"file_id": CONFERENCISTAS_FILE_ID, "sheet": "conferencistas"},
    "participant": {"file_id": PARTICIPANTES_FILE_ID,  "sheet": "participantes"},
}


def _row_key(kind: str, rec: dict) -> tuple:
    g = lambda k: str(rec.get(k) if rec.get(k) is not None and not (isinstance(rec.get(k), float) and pd.isna(rec.get(k))) else "").strip().lower()
    if kind == "seminar":
        return (g("Seminario"),)
    if kind == "speaker":
        full = g("Nombre Completo") if g("Nombre Completo") and not g("Nombre Completo").startswith("#") else f'{g("Nombre")} {g("Apellido")}'.strip()
        return (full, g("Seminario que dirigió"))
    return (g("Nombre"), g("Seminario al que asistió"))


def _drive_token_or_raise() -> str:
    token = _get_gspread_access_token()
    if not token:
        raise RuntimeError("Could not authenticate with the Google service account (check st.secrets['gcp_service_account']).")
    return token


def _drive_download_fresh(file_id: str) -> bytes:
    """Uncached download — used right before writing so we never overwrite newer edits."""
    r = requests.get(f"https://www.googleapis.com/drive/v3/files/{file_id}?alt=media&supportsAllDrives=true",
                     headers={"Authorization": f"Bearer {_drive_token_or_raise()}"}, timeout=60)
    if r.status_code != 200:
        raise RuntimeError(f"Drive download failed (HTTP {r.status_code}): {r.text[:200]}")
    return r.content


def _drive_upload_xlsx(file_id: str, data: bytes) -> None:
    """Replaces the content of an existing Drive file (same ID; Drive keeps the previous version
    in File → Version history)."""
    r = requests.patch(
        f"https://www.googleapis.com/upload/drive/v3/files/{file_id}?uploadType=media&supportsAllDrives=true",
        headers={"Authorization": f"Bearer {_drive_token_or_raise()}", "Content-Type": XLSX_MIME},
        data=data, timeout=120,
    )
    if r.status_code != 200:
        hint = " The service account needs Editor access to this file." if r.status_code in (403, 404) else ""
        raise RuntimeError(f"Drive upload failed (HTTP {r.status_code}): {r.text[:200]}.{hint}")


def _set_cell_value(cell, val) -> None:
    """Writes a value; real dates / times get a readable format only if the cell has none."""
    cell.value = val
    if cell.number_format == "General":
        if isinstance(val, dt.datetime) or (isinstance(val, dt.date) and not isinstance(val, dt.datetime)):
            cell.number_format = "d/m/yyyy"
        elif isinstance(val, dt.time):
            cell.number_format = "h:mm AM/PM"


def _update_row(xlsx: bytes, sheet: str, key_col: str, key_val: str, occurrence: int, values: dict,
                freeze_formulas: bool = False) -> bytes:
    """Changes only the given cells of ONE existing row (found by `key_col` == `key_val`,
    nth occurrence). Everything else in the workbook is left as it is."""
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(xlsx))
    ws = wb[sheet]
    headers = {str(ws.cell(1, c).value).strip(): c for c in range(1, ws.max_column + 1) if ws.cell(1, c).value is not None}
    missing = (set(values) | {key_col}) - set(headers)
    if missing:
        raise RuntimeError(f"These columns are not in sheet '{sheet}': {', '.join(sorted(missing))}")
    target, seen = None, 0
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(r, headers[key_col]).value or "").strip().lower() == key_val.strip().lower():
            seen += 1
            if seen == occurrence:
                target = r
                break
    if target is None:
        raise RuntimeError(f"'{key_val}' was not found in the file (it may have been changed or deleted by someone else).")
    if freeze_formulas:
        vals = load_workbook(io.BytesIO(xlsx), data_only=True)[sheet]
        for r in range(2, ws.max_row + 1):
            for c in headers.values():
                cell = ws.cell(r, c)
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    cached = vals.cell(r, c).value
                    if cached is not None and not (isinstance(cached, str) and cached.startswith("#")):
                        cell.value = cached
    for col, val in values.items():
        _set_cell_value(ws.cell(target, headers[col]), val)
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def _finish_write(fid: str, sheet: str, old: bytes, build, n_new: int, allowed=()) -> bool:
    """build(freeze) -> new workbook bytes. Verifies nothing the dashboard reads is lost
    (retrying once with formulas frozen to their cached values), uploads, clears caches.
    Returns True if formulas had to be frozen."""
    new = build(False)
    lost = [c for c in _lost_columns(old, new, sheet, n_new) if c not in allowed]
    frozen = False
    if lost:
        new = build(True)
        lost = [c for c in _lost_columns(old, new, sheet, n_new) if c not in allowed]
        frozen = True
        if lost:
            raise RuntimeError(f"Saving would lose values in column(s): {', '.join(lost)}. Nothing was saved.")
    _drive_upload_xlsx(fid, new)
    for fn in (_download_drive_file_bytes, load_raw_sheets, build_dataset):
        fn.clear()
    return frozen


def _append_rows(xlsx: bytes, sheet: str, rows: List[dict], freeze_formulas: bool = False) -> bytes:
    """Appends `rows` (dicts keyed by the sheet's header text) below the last data row, copying the
    previous row's cell style. Existing cells are never modified, except when `freeze_formulas`
    is set: formula cells are then replaced by their cached values (used only if saving would
    otherwise lose values that the dashboard reads)."""
    from openpyxl import load_workbook
    from openpyxl.utils import range_boundaries, get_column_letter

    wb = load_workbook(io.BytesIO(xlsx))
    if sheet not in wb.sheetnames:
        raise RuntimeError(f"Sheet '{sheet}' not found in the workbook (found: {', '.join(wb.sheetnames)}).")
    ws = wb[sheet]
    headers = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(1, c).value
        if v is not None and str(v).strip():
            headers[str(v).strip()] = c
    unknown = {k for r in rows for k in r} - set(headers)
    if unknown:
        raise RuntimeError(f"These columns are not in sheet '{sheet}': {', '.join(sorted(unknown))}")
    ncols = max(headers.values())
    last = 1
    for r in range(ws.max_row, 1, -1):
        if any(ws.cell(r, c).value not in (None, "") for c in range(1, ncols + 1)):
            last = r
            break

    if freeze_formulas:
        vals = load_workbook(io.BytesIO(xlsx), data_only=True)[sheet]
        for r in range(2, last + 1):
            for c in range(1, ncols + 1):
                cell = ws.cell(r, c)
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    cached = vals.cell(r, c).value
                    if cached is not None and not (isinstance(cached, str) and cached.startswith("#")):
                        cell.value = cached

    for i, row in enumerate(rows, start=1):
        r = last + i
        for h, c in headers.items():
            cell = ws.cell(r, c)
            if last > 1:
                cell._style = copy(ws.cell(last, c)._style)
            val = row.get(h)
            if val is None or val == "":
                continue
            _set_cell_value(cell, val)

    # Excel Tables / autofilters that end at the old last row must grow with the data.
    for t in list(ws.tables.values()):
        c1, r1, c2, r2 = range_boundaries(t.ref)
        if r2 == last:
            t.ref = f"{get_column_letter(c1)}{r1}:{get_column_letter(c2)}{last + len(rows)}"
            if t.autoFilter is not None:
                t.autoFilter.ref = t.ref
    if ws.auto_filter and ws.auto_filter.ref:
        c1, r1, c2, r2 = range_boundaries(ws.auto_filter.ref)
        if r2 == last:
            ws.auto_filter.ref = f"{get_column_letter(c1)}{r1}:{get_column_letter(c2)}{last + len(rows)}"

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def _lost_columns(old: bytes, new: bytes, sheet: str, n_new: int) -> List[str]:
    """Columns whose values (as pandas/the dashboard reads them) would be lost or the row count wrong."""
    a = pd.read_excel(io.BytesIO(old), sheet_name=sheet)
    b = pd.read_excel(io.BytesIO(new), sheet_name=sheet)
    if len(b) != len(a) + n_new:
        raise RuntimeError(f"Row count check failed (expected {len(a) + n_new}, got {len(b)}). Nothing was saved.")
    return [c for c in a.columns if c in b.columns and a[c].notna().sum() > b[c].iloc[:len(a)].notna().sum()]


def commit_rows(kind: str, rows: List[dict]) -> dict:
    """Appends rows to the Excel file of `kind` ('seminar' | 'speaker' | 'participant') in Drive.
    Always re-reads the file first (fresh), skips rows that already exist, verifies the result
    is readable with no lost values, uploads, and clears the dashboard caches."""
    tgt = WRITE_TARGETS[kind]
    fid, sheet = tgt["file_id"], tgt["sheet"]
    old = _drive_download_fresh(fid)
    existing = {_row_key(kind, rec) for rec in pd.read_excel(io.BytesIO(old), sheet_name=sheet).to_dict("records")}
    fresh, skipped, seen = [], [], set()
    for row in rows:
        k = _row_key(kind, row)
        if k in existing or k in seen:
            skipped.append(" / ".join(x for x in k if x))
            continue
        seen.add(k)
        fresh.append(row)
    if not fresh:
        return {"written": 0, "skipped": skipped, "frozen": False}

    frozen = _finish_write(fid, sheet, old, lambda fz: _append_rows(old, sheet, fresh, freeze_formulas=fz), len(fresh))
    return {"written": len(fresh), "skipped": skipped, "frozen": frozen}


def commit_update(kind: str, key_col: str, key_val: str, occurrence: int, values: dict) -> dict:
    """Edits cells of one existing row (always on a fresh copy of the file)."""
    tgt = WRITE_TARGETS[kind]
    fid, sheet = tgt["file_id"], tgt["sheet"]
    old = _drive_download_fresh(fid)
    frozen = _finish_write(fid, sheet, old, lambda fz: _update_row(old, sheet, key_col, key_val, occurrence, values, fz),
                           0, allowed=set(values))
    return {"written": 1, "skipped": [], "frozen": frozen}


# ── 12) PAGE — Manage Data ───────────────────────────────────────────────
EDITOR_SECRET = "seminars_editor_password"
SEMINAR_PREFIX = {   # category -> ID prefix used in the workbook ("Visiting Scholars 12")
    "Visiting Scholars Seminar Series": "Visiting Scholars",
    "Recruitment research seminars": "Recruitment research",
    "Ágora": "Ágora", "GHE": "GHE",
}
MODALITIES = ["On-Campus", "Online", "Hybrid"]
ATTENDEE_TYPES = ["Professor", "PhD Student", "Student", "Administrative", "Visiting professor", "External"]
REGIONS = ["Europe", "North America", "South America", "Asia", "Africa", "Oceania"]


def _editor_unlocked() -> bool:
    """If st.secrets['seminars_editor_password'] is set, the page asks for it (optional but recommended)."""
    try:
        pw = st.secrets.get(EDITOR_SECRET)
    except Exception:
        pw = None
    if not pw or st.session_state.get("_sem_editor_ok"):
        return True
    with st.form("sem_editor_gate", border=False):
        typed = st.text_input("Editor password", type="password")
        go = st.form_submit_button("Unlock")
    if go:
        if hmac.compare_digest(str(typed), str(pw)):
            st.session_state["_sem_editor_ok"] = True
            st.rerun()
        st.error("Wrong password.")
    return False


def _suggest_seminar_id(category: str, raw_ids: List[str]) -> str:
    prefix = SEMINAR_PREFIX.get(category, category)
    nums = []
    for i in raw_ids:
        m = re.fullmatch(re.escape(prefix) + r"\s+(\d+)", str(i).strip(), flags=re.I)
        if m:
            nums.append(int(m.group(1)))
    return f"{prefix} {max(nums, default=0) + 1}"


def _uniq(values) -> List[str]:
    return sorted({str(v).strip() for v in values if v is not None and str(v).strip() and not str(v).startswith("#")
                   and str(v).strip().upper() != "NA"})


PART_COLS = ["Nombre", "Correo", "Género", "Tipo Asistente", "Modalidad"]   # as in BD_participantes (+ the seminar column)


def _show_saved(kind: str):
    done = st.session_state.pop(f"saved_{kind}", None)
    if done:
        st.success(done["msg"])
        for w in done.get("warnings", []):
            st.warning(w)


def _save(kind: str, action, ok_msg: str, warnings=None) -> bool:
    """Runs a Drive write, remembers the result for the next run (the page reruns right after)."""
    with st.spinner("Saving to Drive…"):
        try:
            res = action()
        except Exception as e:
            st.error(f"Not saved: {e}")
            return False
    warns = list(warnings or [])
    if res.get("skipped"):
        warns.append("Skipped (already in the file): " + "; ".join(res["skipped"]))
    if res.get("frozen"):
        warns.append("Some formula cells in this workbook were converted to their current values so nothing is lost.")
    msg = ok_msg if res.get("written") else "Nothing new to save — everything was already in the file."
    st.session_state[f"saved_{kind}"] = {"msg": msg, "warnings": warns if res.get("written") else warns}
    return True


def _participants_template(types: List[str]) -> bytes:
    """Empty .xlsx with exactly the columns the Drive file asks for (the seminar is chosen on the page),
    with dropdowns for Género / Tipo Asistente / Modalidad."""
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from openpyxl.worksheet.datavalidation import DataValidation
    wb = Workbook()
    ws = wb.active
    ws.title = "participantes"
    ws.append(PART_COLS)
    for c in range(1, len(PART_COLS) + 1):
        ws.cell(1, c).font = Font(bold=True)
        ws.column_dimensions[chr(64 + c)].width = 30 if c <= 2 else 18
    lists = wb.create_sheet("Lists")
    for i, v in enumerate(types, start=1):
        lists.cell(i, 1).value = v
    for i, v in enumerate(MODALITIES, start=1):
        lists.cell(i, 2).value = v
    for i, v in enumerate(["M", "F"], start=1):
        lists.cell(i, 3).value = v
    specs = {"C": f"=Lists!$C$1:$C$2", "D": f"=Lists!$A$1:$A${max(len(types), 1)}", "E": f"=Lists!$B$1:$B${len(MODALITIES)}"}
    for col, ref in specs.items():
        dv = DataValidation(type="list", formula1=ref, allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"{col}2:{col}1000")
    lists.sheet_state = "hidden"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _id_number(raw_id: str) -> int:
    m = re.search(r"(\d+)\s*$", str(raw_id))
    return int(m.group(1)) if m else -1


def page_manage():
    _render_header(
        "Manage Data",
        "Add or edit a seminar, add a speaker, or add participants. Changes are saved straight into the Excel "
        "files in Drive (BD_seminarios / BD_conferencistas / BD_participantes).",
    )
    if not _editor_unlocked():
        return
    st.caption("Drive keeps the previous version of each file (File → Version history), so any save can be rolled back there.")

    seminars = build_dataset()
    sem_df, conf_df, part_df = load_raw_sheets()
    raw_ids = [str(x).strip() for x in sem_df["Seminario"].dropna()]
    firsts = {}
    for sm in seminars:                      # one entry per raw ID (first occurrence)
        firsts.setdefault(sm["rawId"].lower(), sm)
    sem_opts = [(sm["rawId"], f'{sm["rawId"]} — {sm["nombre"][:70]} ({sm["fecha"] or "TBD"})')
                for sm in sorted(firsts.values(), key=lambda x: x["fecha"] or "", reverse=True)]
    sem_ids = [o[0] for o in sem_opts]
    sem_label = dict(sem_opts)
    last_sem = st.session_state.get("last_seminar")      # the seminar created just before (pre-selected below)
    if last_sem not in sem_label:
        last_sem = None

    tab_sem, tab_spk, tab_par = st.tabs(["📅 Seminar", "🎤 Speaker", "👥 Participants"])

    # ───────────────── Seminar (new or edit) ─────────────────
    with tab_sem:
        _show_saved("seminar")
        nv = st.session_state.get("ns_ver", 0)     # bumped after a save so the form starts clean
        cats = _uniq(list(sem_df["Categoría"]) + list(SEMINAR_PREFIX))
        top1, top2 = st.columns(2)
        category = top1.selectbox("Category", cats, accept_new_options=True, key="ns_cat")
        category = (category or "").strip()

        # Seminar ID right at the top: the suggested new ID first, then the earlier seminars of this category.
        in_cat = sorted([sm for sm in firsts.values() if sm["categoria"].lower() == category.lower()],
                        key=lambda x: (_id_number(x["rawId"]), x["fecha"] or ""), reverse=True)
        new_id = _suggest_seminar_id(category, raw_ids)
        id_labels = {new_id: f"{new_id}  ·  new seminar"}
        id_labels.update({sm["rawId"]: f'{sm["rawId"]}  ·  {sm["nombre"][:55]} ({sm["fecha"] or "TBD"})' for sm in in_cat})
        sel_id = top2.selectbox(
            "Seminar ID", list(id_labels), format_func=lambda x: id_labels.get(x, x), accept_new_options=True,
            key=f"ns_id_{category}_{len(raw_ids)}_{nv}",
            help="The first option is a new seminar. Pick an earlier ID of this category to edit it. You can also type a custom ID.")
        sel_id = (sel_id or "").strip()
        cur = firsts.get(sel_id.lower())
        editing = cur is not None and cur["categoria"].lower() == category.lower()
        if editing:
            st.caption(f'Editing **{cur["rawId"]}** — the ID can\'t be changed (speakers and participants are linked to it). '
                       "Only the fields you change are written.")
        ek = f"{category}|{sel_id}|{nv}"

        c1, c2 = st.columns(2)
        with c1:
            name = st.text_input("Seminar name", value=cur["nombre"] if editing else "", key=f"ns_name_{ek}")
            known_spk = _uniq([n for sm in seminars for n in sm["speakers"]])
            default_spk = list(cur["speakers"]) if editing else []
            speakers = st.multiselect("Speaker(s) — up to 5", _uniq(known_spk + default_spk), default=default_spk,
                                      accept_new_options=True, max_selections=5, key=f"ns_spk_{ek}")
            areas = _uniq(list(sem_df["Área"]) + ([cur["area"]] if editing else []))
            area = st.selectbox("Area", areas, accept_new_options=True, key=f"ns_area_{ek}",
                                index=areas.index(cur["area"]) if editing and cur["area"] in areas else None,
                                placeholder="Select or type")
        with c2:
            d1, d2 = st.columns(2)
            d0 = dt.date.fromisoformat(cur["fecha"]) if editing and cur["fecha"] else dt.date.today()
            date = d1.date_input("Date", value=d0, key=f"ns_date_{ek}", format="DD/MM/YYYY")
            t0 = dt.time(*map(int, cur["hora"].split(":"))) if editing and cur["hora"] else dt.time(12, 30)
            time_val = d2.time_input("Time", value=t0, key=f"ns_time_{ek}")
            mod0 = MODALITIES.index(cur["modalidad"]) if editing and cur["modalidad"] in MODALITIES else 0
            modality = st.selectbox("Modality", MODALITIES, index=mod0, key=f"ns_mod_{ek}")
            loc0 = (cur["lugarRaw"] or "") if editing else ("Virtual" if modality == "Online" else "")
            place = st.text_input("Location", value=loc0, key=f"ns_loc_{ek}_{modality}", placeholder="e.g. SD 704")

        if st.button("Save seminar", key="ns_save", icon=":material/save:"):
            errs = []
            if not category: errs.append("Category is required.")
            if not sel_id: errs.append("Seminar ID is required.")
            if not name.strip(): errs.append("Seminar name is required.")
            if not (area or "").strip(): errs.append("Area is required.")
            for e in errs:
                st.error(e)
            if not errs:
                spk5 = (speakers + [None] * 5)[:5]
                known = {sp["nombre"].strip().lower() for sm in seminars for sp in sm["speakerDetails"]}
                missing = [n for n in speakers if n.strip().lower() not in known]
                spk_warn = ([f'No speaker profile yet for: {", ".join(missing)}. Add it in the "Speaker" tab.'] if missing else [])
                if editing:
                    upd = {}
                    if name.strip() != cur["nombre"]: upd["Nombre del seminario"] = name.strip()
                    if list(speakers) != list(cur["speakers"]):
                        upd.update({f"Conferencista {i + 1}": (spk5[i] or "NA") for i in range(5)})
                    if date.isoformat() != (cur["fecha"] or ""):
                        upd["Fecha"] = dt.datetime.combine(date, dt.time()); upd["Año"] = date.year
                    if (cur["hora"] is None and time_val != dt.time(12, 30)) or (cur["hora"] is not None and time_val.strftime("%H:%M") != cur["hora"]):
                        upd["Hora"] = time_val
                    if modality != cur["modalidad"]: upd["Modalidad"] = modality
                    if place.strip() != (cur["lugarRaw"] or ""): upd["Lugar"] = place.strip() or "TBD"
                    if area.strip() != cur["area"]: upd["Área"] = area.strip()
                    if not upd:
                        st.info("No changes to save.")
                    elif _save("seminar", lambda: commit_update("seminar", "Seminario", cur["rawId"], cur["occ"], upd),
                               f'Saved changes to {cur["rawId"]}: {", ".join(upd)}.', spk_warn):
                        st.session_state["ns_ver"] = nv + 1
                        st.rerun()
                else:
                    row = {
                        "Seminario": sel_id, "Categoría": category, "Nombre del seminario": name.strip(),
                        **{f"Conferencista {i + 1}": (spk5[i] or "NA") for i in range(5)},
                        "Fecha": dt.datetime.combine(date, dt.time()), "Hora": time_val, "Año": date.year,
                        "Modalidad": modality, "Lugar": place.strip() or "TBD", "Área": area.strip(),
                    }
                    if _save("seminar", lambda: commit_rows("seminar", [row]),
                             f'Saved seminar "{sel_id}". It is now pre-selected in the Speaker and Participants tabs.', spk_warn):
                        st.session_state["last_seminar"] = sel_id
                        st.session_state["ns_ver"] = nv + 1
                        st.rerun()

    # ───────────────── Speaker ─────────────────
    with tab_spk:
        _show_saved("speaker")
        sv = st.session_state.get("sp_ver", 0)
        mode = st.radio("What do you want to do?", ["New speaker", "Existing speaker — add another seminar"],
                        horizontal=True, key="sp_mode")
        led = st.multiselect("Seminar(s) led", sem_ids, default=[last_sem] if last_sem else [],
                             format_func=lambda x: sem_label.get(x, x), key=f"sp_led_{sv}_{last_sem}")
        profile = {}
        if mode.startswith("New"):
            c1, c2 = st.columns(2)
            with c1:
                first = st.text_input("First name", key=f"sp_first_{sv}")
                last = st.text_input("Last name", key=f"sp_last_{sv}")
                gender = st.selectbox("Gender", ["Female", "Male"], key="sp_gender")
                nat = st.selectbox("Nationality", _uniq(conf_df["Nacionalidad"]), accept_new_options=True,
                                   format_func=country_en, key="sp_nat", index=None, placeholder="Select or type")
                has_birth = st.checkbox("Add birth date", key="sp_has_birth")
                birth = st.date_input("Birth date", value=dt.date(1980, 1, 1), min_value=dt.date(1920, 1, 1),
                                      max_value=dt.date.today(), key="sp_birth", disabled=not has_birth, format="DD/MM/YYYY")
            with c2:
                uni = st.selectbox("Origin university", _uniq(conf_df["Nombre Universidad origen"]), accept_new_options=True,
                                   key="sp_uni", index=None, placeholder="Select or type")
                uni_c = st.selectbox("University country", _uniq(conf_df["País Universidad origen"]), accept_new_options=True,
                                     format_func=country_en, key="sp_uni_c", index=None, placeholder="Select or type")
                region = st.selectbox("Origin region", _uniq(list(conf_df["Ubicación"]) + REGIONS), accept_new_options=True,
                                      key="sp_region", index=None, placeholder="Select or type")
            full = f"{first.strip()} {last.strip()}".strip()
        else:
            known_spk = _uniq([sp["nombre"] for sm in seminars for sp in sm["speakerDetails"]])
            chosen = st.selectbox("Speaker", known_spk, index=None, placeholder="Select a speaker", key=f"sp_existing_{sv}")
            full = chosen or ""
            if chosen:
                prof = conf_df[conf_df.apply(lambda r: _row_key("speaker", r.to_dict())[0] == chosen.lower(), axis=1)]
                profile = prof.iloc[0].to_dict() if len(prof) else {}
                if profile:
                    st.caption(f'Profile reused from the file: {_clean(profile.get("Nombre Universidad origen"), "—")} · '
                               f'{country_en(profile.get("Nacionalidad"))} · {_clean(profile.get("Genero"), "—")}')

        if st.button("Save speaker", key="sp_save", icon=":material/save:"):
            errs = []
            if mode.startswith("New"):
                if not (first.strip() and last.strip()): errs.append("First and last name are required.")
                if not (uni or "").strip(): errs.append("Origin university is required.")
                if not (nat or "").strip(): errs.append("Nationality is required.")
            elif not full:
                errs.append("Select the speaker.")
            if not led: errs.append("Select at least one seminar.")
            for e in errs:
                st.error(e)
            if not errs:
                rows = []
                for sem_id in led:
                    if mode.startswith("New"):
                        base = {"Nombre": first.strip(), "Apellido": last.strip(), "Nombre Completo": full, "Genero": gender,
                                "Nacionalidad": nat.strip(),
                                "Fecha nacimiento": (dt.datetime.combine(birth, dt.time()) if has_birth else "NA"),
                                "Nombre Universidad origen": uni.strip(), "País Universidad origen": (uni_c or "").strip() or "NA",
                                "Ubicación": (region or "").strip() or "NA"}
                    else:
                        def _v(k):
                            v = profile.get(k)
                            return "NA" if v is None or (isinstance(v, float) and pd.isna(v)) else v
                        base = {k: _v(k) for k in ["Nombre", "Apellido", "Genero", "Nacionalidad", "Fecha nacimiento",
                                                   "Nombre Universidad origen", "País Universidad origen", "Ubicación"]}
                        base["Nombre Completo"] = full
                    rows.append({**base, "Seminario que dirigió": sem_id})
                if _save("speaker", lambda: commit_rows("speaker", rows),
                         f'Saved {full} for {len(rows)} seminar(s): {", ".join(led)}.'):
                    st.session_state["sp_ver"] = sv + 1
                    st.rerun()

    # ───────────────── Participants ─────────────────
    with tab_par:
        _show_saved("participant")
        sem_pick = st.selectbox("Seminar attended", sem_ids, format_func=lambda x: sem_label.get(x, x),
                                index=sem_ids.index(last_sem) if last_sem else None, placeholder="Select the seminar",
                                key=f"pa_sem_{last_sem}")
        st.session_state.setdefault("pa_ver", 0)
        blank = lambda: pd.DataFrame({c: [None] * 6 for c in PART_COLS})
        cur_df = st.session_state.get("pa_df")
        if not isinstance(cur_df, pd.DataFrame) or list(cur_df.columns) != PART_COLS:   # e.g. state left by an older version
            st.session_state["pa_df"] = blank()
            st.session_state["pa_ver"] += 1
        types = _uniq(list(part_df["Tipo Asistente"]) + ATTENDEE_TYPES + [v for v in st.session_state["pa_df"]["Tipo Asistente"] if v])

        u1, u2 = st.columns([3, 2])
        with u1:
            up = st.file_uploader("Upload the participants list (.xlsx / .csv)", type=["xlsx", "csv"], key="pa_upload")
        with u2:
            st.markdown("<div style='height:28px;'></div>", unsafe_allow_html=True)
            st.download_button("Download template (.xlsx)", data=_participants_template(_uniq(list(part_df["Tipo Asistente"]) + ATTENDEE_TYPES)),
                               file_name="BD_participantes_template.xlsx", mime=XLSX_MIME, icon=":material/download:", key="pa_template")
            st.caption("Same columns as BD_participantes (the seminar is the one selected above).")

        if up is not None and st.session_state.get("pa_upload_sig") != (up.name, up.size):
            try:
                raw = pd.read_csv(up) if up.name.lower().endswith(".csv") else pd.read_excel(up)
                alias = {"name": "Nombre", "nombre": "Nombre", "email": "Correo", "correo": "Correo", "e-mail": "Correo",
                         "gender": "Género", "género": "Género", "genero": "Género", "sexo": "Género",
                         "attendee type": "Tipo Asistente", "tipo asistente": "Tipo Asistente", "type": "Tipo Asistente",
                         "modality": "Modalidad", "modalidad": "Modalidad"}
                raw.columns = [alias.get(str(c).strip().lower(), str(c).strip()) for c in raw.columns]
                if "Nombre" not in raw.columns:
                    raise ValueError("the file needs a 'Nombre' column (use the template)")
                for c in PART_COLS:
                    if c not in raw.columns:
                        raw[c] = None
                raw = raw[PART_COLS].astype(object).where(raw[PART_COLS].notna(), None)
                raw["Género"] = raw["Género"].map(lambda g: {"M": "M", "F": "F"}.get(norm_genero(g)) if g is not None else None)
                canon = {t.lower(): t for t in types}
                raw["Tipo Asistente"] = raw["Tipo Asistente"].map(lambda t: canon.get(str(t).strip().lower(), str(t).strip()) if t is not None else None)
                raw["Modalidad"] = raw["Modalidad"].map(lambda m: norm_modalidad(m) if m is not None else None).replace({"Unknown": None})
                st.session_state["pa_df"] = raw.reset_index(drop=True)
                st.session_state["pa_ver"] += 1
                st.session_state["pa_upload_sig"] = (up.name, up.size)
                st.rerun()
            except Exception as e:
                st.error(f"Could not read that file: {e}")

        edited = st.data_editor(
            st.session_state["pa_df"], num_rows="dynamic", use_container_width=True, hide_index=True,
            key=f"pa_editor_{st.session_state['pa_ver']}",
            column_config={
                "Nombre": st.column_config.TextColumn("Nombre", required=True),
                "Correo": st.column_config.TextColumn("Correo"),
                "Género": st.column_config.SelectboxColumn("Género", options=["M", "F"]),
                "Tipo Asistente": st.column_config.SelectboxColumn("Tipo Asistente", options=types),
                "Modalidad": st.column_config.SelectboxColumn("Modalidad", options=MODALITIES),
            },
        )
        st.caption("You can edit the table, add rows, or paste a block of rows from Excel. It refreshes whenever you upload a list.")

        if st.button("Save participants", key="pa_save", icon=":material/save:"):
            errs, rows, names_seen = [], [], set()
            if not sem_pick: errs.append("Select the seminar.")
            for i, rec in enumerate(edited.to_dict("records"), start=1):
                nm = _clean(rec.get("Nombre"))
                if not nm:
                    continue
                email = _clean(rec.get("Correo"), "")
                if email and not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", email):
                    errs.append(f"Row {i}: '{email}' is not a valid email.")
                if nm.lower() in names_seen:
                    errs.append(f"Row {i}: '{nm}' appears twice in this list.")
                names_seen.add(nm.lower())
                typ, mod = _clean(rec.get("Tipo Asistente")), _clean(rec.get("Modalidad"))
                if not typ or not mod:
                    errs.append(f"Row {i} ({nm}): Tipo Asistente and Modalidad are required.")
                rows.append({"Nombre": nm, "Correo": email or None, "Género": _clean(rec.get("Género")),
                             "Tipo Asistente": typ, "Modalidad": mod, "Seminario al que asistió": sem_pick})
            if not rows: errs.append("Add at least one participant.")
            for e in errs:
                st.error(e)
            if not errs and _save("participant", lambda: commit_rows("participant", rows),
                                  f'Saved {len(rows)} participant(s) for "{sem_pick}".'):
                st.session_state["pa_df"] = blank()
                st.session_state["pa_ver"] += 1
                st.rerun()


# ── 13) NAVIGATION ────────────────────────────────────────────────────────
# Same skeleton as the EIV / Weeks reports: the cover hides the sidebar; the other pages get a
# floating top navigation plus fixed side arrows to move between sections.
pages = [
    st.Page(page_datacenter, title="Data Center", icon=":material/dashboard:", url_path="datacenter", default=True),
    st.Page(page_dashboard, title="Dashboard", icon="📊", url_path="dashboard"),
    st.Page(page_search, title="Seminar Search", icon="🔎", url_path="search"),
    st.Page(page_tables, title="Tables", icon="📋", url_path="tables"),
    st.Page(page_manage, title="Manage Data", icon="✏️", url_path="manage"),
]
pg = st.navigation(pages, position="hidden")
IS_DATACENTER = pg is pages[0]
nav_pages = pages[1:]
_NAV_VISIBLE = 6

if IS_DATACENTER:
    st.markdown(
        "<style>section[data-testid='stSidebar']{display:none !important;}"
        "div[data-testid='stSidebarCollapsedControl']{display:none !important;}</style>",
        unsafe_allow_html=True,
    )
else:
    with st.sidebar:
        st.markdown(
            f'<div style="color:{INK};font-size:22px;font-weight:800;line-height:1.1;">UASM Seminars</div>',
            unsafe_allow_html=True,
        )
        st.caption("School of Management · International Seminars")
        with st.container(key="go_to_datacenter_btn"):
            st.page_link(pages[0], label="Go to Data Center", icon=":material/home:")
        st.markdown("---")

    with st.container(key="nav_toggle"):
        visible_pages = nav_pages[:_NAV_VISIBLE]
        overflow_pages = nav_pages[_NAV_VISIBLE:]
        nav_cols = st.columns(len(visible_pages) + (1 if overflow_pages else 0))
        for col, page_obj in zip(nav_cols, visible_pages):
            with col:
                st.page_link(page_obj)
        if overflow_pages:
            with nav_cols[-1]:
                with st.popover("More...", use_container_width=False):
                    for page_obj in overflow_pages:
                        st.page_link(page_obj)

    _idx = nav_pages.index(pg)
    with st.container(key="side_arrows"):
        arrow_l, arrow_r = st.columns(2)
        with arrow_l:
            st.page_link(nav_pages[_idx - 1], label="‹")
        with arrow_r:
            st.page_link(nav_pages[(_idx + 1) % len(nav_pages)], label="›")

pg.run()

st.markdown(
    f'<div style="text-align:center;padding:40px 0 10px;font-size:11.5px;color:{MUTED};'
    'font-family:monospace;">International Seminars · School of Management · '
    'Universidad de los Andes</div>',
    unsafe_allow_html=True,
)
