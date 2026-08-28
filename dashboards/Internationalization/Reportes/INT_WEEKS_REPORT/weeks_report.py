# ===========================================================================
#  UASM · International Weeks Analytics · App multipágina
#  Traducción 1:1 del reporte HTML "International Weeks 2026" a Streamlit —
#  misma estructura (Data Center, Overview, una vista por semana, Financial
#  Detail), mismos cálculos (ingresos/egresos/balance/margen desde
#  BD_presupuesto, roster y pagos desde BD_listas, encuesta desde
#  BD_encuesta_curso), y misma identidad visual por socio (color, logo).
#
#  Las 4 fuentes NO exponen directamente la agregación que mostraba el HTML
#  original (esa vino de un cálculo hecho aparte); esta app reconstruye esa
#  agregación en Python a partir de los archivos crudos, para que quede
#  siempre sincronizada con Drive en vez de datos congelados.
# ===========================================================================
from __future__ import annotations

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import io
import time
import functools
from typing import Optional, Dict, List, Tuple

try:
    from google.oauth2.service_account import Credentials
    _GSPREAD_OK = True
    _GSPREAD_IMPORT_ERR = None
except ImportError as _e:
    _GSPREAD_OK = False
    _GSPREAD_IMPORT_ERR = str(_e)

import requests
import os as _os
import base64

# ── 1) CONFIGURACIÓN GLOBAL ────────────────────────────────────────────────
st.set_page_config(
    page_title="International Weeks Analytics",
    page_icon="🌍",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Paleta — misma arquitectura visual que EIV_report.py (header degradado,
# KPI cards, botones, nav superior con flechas), con un acento propio de
# International Weeks: azul elegante (no menta). Los colores propios de
# cada semana (KLU/NOVA/FGV/BABSON) se conservan tal cual — son identidad
# del socio, no "color principal".
INK = "#241420"; INK_SOFT = "#6E5C68"; PAPER = "#FAF8FA"
LINE = "#E9E2E7"; MUTED = "#8C7F87"
ACCENT = "#2C4A6B"           # azul marino elegante — reemplaza el rol de "PINK" de EIV
ACCENT_DARK = "#152C43"      # hover / degradado oscuro
ACCENT_LIGHT = "#5B84AC"     # degradado claro (azul acero)
ACCENT_SOFT = "#E4ECF3"      # fondo suave para tablas / chips
GOLD = "#AD8A32"             # acento cálido secundario (detalles puntuales, no botones)
INCOME = "#2E8B57"; INCOME_SOFT = "#E3F3EA"
EXPENSE = "#C23B3B"; EXPENSE_SOFT = "#FBE6E6"
BALANCE = "#1C6FA5"; BALANCE_SOFT = "#E1EEF7"

st.markdown(
    "<style>"
    f".suite-header{{display:flex;flex-direction:column;align-items:center;"
    "padding:16px 24px 12px;"
    f"background:linear-gradient(135deg,{ACCENT_DARK} 0%,{ACCENT} 55%,{ACCENT_LIGHT} 100%);"
    f"border-radius:12px;box-shadow:0 2px 8px rgba(44,74,107,.22);margin-bottom:14px;}}"
    f".sh-super{{font-size:11px;font-weight:700;letter-spacing:2px;"
    "color:#D8E4EF !important;text-transform:uppercase;margin-bottom:2px;}}"
    ".sh-title{font-size:26px !important;font-weight:800 !important;color:#fff !important;"
    "text-align:center !important;line-height:1.2 !important;margin:0 !important;padding:0 !important;}"
    ".sh-sub{font-size:13px;color:rgba(255,255,255,.80) !important;margin-top:4px;text-align:center;}"
    "div[data-testid='stMarkdownContainer'] .sh-title{font-size:26px !important;font-weight:800 !important;"
    "color:#fff !important;text-align:center !important;}"
    f".kv{{font-size:26px;font-weight:700;line-height:1.1;font-family:monospace;color:{INK};}}"
    f".kv.accent{{color:{ACCENT};}}"
    f".kv.income{{color:{INCOME};}} .kv.expense{{color:{EXPENSE};}} .kv.balance{{color:{BALANCE};}}"
    f".kl{{font-size:11px;font-weight:600;color:{MUTED};"
    "text-transform:uppercase;letter-spacing:.5px;margin-top:3px;}}"
    f"section[data-testid='stSidebar']{{background:{PAPER} !important;}}"
    "div[data-testid='stButton'] button{background:#FFFFFF !important;"
    f"border:1px solid {LINE} !important;border-radius:10px !important;"
    "color:#374151 !important;font-size:14px !important;"
    "font-weight:600 !important;height:48px !important;"
    "box-shadow:0 1px 3px rgba(0,0,0,.04) !important;}"
    f"div[data-testid='stButton'] button:hover{{background:{PAPER} !important;border-color:{ACCENT} !important;}}"
    "div.stDownloadButton>button{background:transparent !important;"
    "border:none !important;box-shadow:none !important;"
    f"color:{ACCENT} !important;font-size:13px !important;"
    "padding:0 !important;text-decoration:underline !important;}"
    f"thead th{{background:{ACCENT_SOFT} !important;color:{INK} !important;"
    "font-weight:700 !important;}}"
    ".pending-card{background:#FAFAFA;border:1px dashed #DCD3D8;border-radius:12px;"
    "padding:22px 24px;margin-top:14px;color:#6B7280;font-size:13.5px;}"
    ".pending-card .tag{display:inline-block;font-family:monospace;font-size:10px;"
    "letter-spacing:.06em;text-transform:uppercase;color:#8C7F87;"
    "background:#F1EAEE;padding:3px 9px;border-radius:5px;margin-bottom:8px;}"
    ".partner-card{border-radius:14px;padding:20px;border:1px solid var(--line,#E9E2E7);"
    "background:#fff;border-top:5px solid var(--wc,#7A1F3D);}"
    ".partner-card .loc{font-family:monospace;font-size:11px;color:#8C7F87;margin-bottom:10px;}"
    ".partner-card .n{font-size:28px;font-weight:700;}"
    ".partner-card .n-label{font-size:11.5px;color:#6E5C68;margin-bottom:10px;}"
    ".margin-pill{display:inline-block;margin-top:8px;padding:3px 10px;border-radius:20px;"
    "font-size:11px;font-weight:600;}"
    f".survey-quote{{background:#FAF8FA;border-left:3px solid {ACCENT};"
    "border-radius:6px;padding:8px 12px;margin-bottom:8px;font-size:12.5px;"
    "font-style:italic;color:#3B2C34;}}"
    ".st-key-cover_enter_btn div[data-testid='stButton'] button{"
    f"background:{ACCENT} !important;border:none !important;color:#fff !important;"
    "font-size:16px !important;font-weight:700 !important;height:52px !important;"
    "box-shadow:0 4px 14px rgba(44,74,107,.28) !important;}"
    f".st-key-cover_enter_btn div[data-testid='stButton'] button:hover{{background:{ACCENT_DARK} !important;}}"
    ".st-key-cover_banner{display:flex !important;justify-content:center !important;width:100% !important;}"
    ".st-key-cover_banner div[data-testid='stImage']{margin:0 auto !important;width:auto !important;"
    "display:flex !important;justify-content:center !important;}"
    ".st-key-cover_banner img{margin:0 auto !important;display:block !important;}"
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
    f"color:{ACCENT} !important;font-size:13px !important;font-weight:400 !important;"
    "height:auto !important;width:auto !important;min-width:fit-content !important;"
    "padding:6px 4px !important;white-space:nowrap !important;}"
    f".st-key-nav_toggle div[data-testid='stPopover'] button:hover{{color:{INK} !important;}}"
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
    f"font-size:0 !important;font-weight:400;color:{ACCENT} !important;opacity:.55;text-decoration:none;"
    "transition:opacity .15s ease;}"
    ".st-key-side_arrows a p{font-size:26px !important;}"
    ".st-key-side_arrows a span:first-child{display:none !important;}"
    ".st-key-side_arrows a:hover{opacity:1;}"
    ".block-container{padding-top:3.2rem !important;}"
    ".st-key-go_to_datacenter_btn a{"
    f"display:flex !important;align-items:center;justify-content:center;gap:6px;"
    f"background:{ACCENT} !important;border:none !important;border-radius:10px !important;"
    "color:#fff !important;font-weight:700 !important;height:44px !important;text-decoration:none !important;}"
    ".st-key-go_to_datacenter_btn a span{color:#fff !important;}"
    f".st-key-go_to_datacenter_btn a:hover{{background:{ACCENT_DARK} !important;}}"
    "div[class*='st-key-gen_report_btn_'] div[data-testid='stDownloadButton'] button{"
    f"background:{ACCENT} !important;color:#fff !important;border:none !important;"
    "text-decoration:none !important;font-weight:700 !important;height:50px !important;"
    f"box-shadow:0 3px 10px rgba(44,74,107,.28) !important;}}"
    "div[class*='st-key-gen_report_btn_'] div[data-testid='stDownloadButton'] button:hover{"
    f"background:{ACCENT_DARK} !important;}}"
    "</style>",
    unsafe_allow_html=True,
)


# ── 2) HELPERS COMPARTIDOS ──────────────────────────────────────────────
def _render_header(title: str, subtitle: str = ""):
    sub = f'<div class="sh-sub">{subtitle}</div>' if subtitle else ""
    st.markdown(
        f'<div class="suite-header"><div class="sh-super">International Weeks · Analytics</div>'
        f'<div class="sh-title">{title}</div>{sub}</div>',
        unsafe_allow_html=True,
    )


def _kpi(label: str, value, cls: str = ""):
    st.markdown(
        f'<div class="kv {cls}">{value}</div><div class="kl">{label}</div>',
        unsafe_allow_html=True,
    )


def _pending_card(label: str, note: str = ""):
    note_html = f"<br>{note}" if note else ""
    st.markdown(
        f'<div class="pending-card"><span class="tag">Coming soon</span>'
        f'<p style="margin-top:8px;">{label}{note_html}</p></div>',
        unsafe_allow_html=True,
    )


def _fmt_money(n) -> str:
    if n is None or (isinstance(n, float) and pd.isna(n)):
        return "$0"
    sign = "-" if n < 0 else ""
    return f"{sign}${abs(round(n)):,.0f}"


def _xlsx_bytes(df: pd.DataFrame, sheet_name: str = "Data") -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf) as w:
        df.to_excel(w, index=False, sheet_name=sheet_name[:31])
    buf.seek(0)
    return buf.getvalue()


def tick(v) -> str:
    return "✅" if v else "❌"


# ── Identidad visual: logos de socios, banner de portada y fotos de
#    profesores — mismo patrón que _show_eiv_logo/_show_eiv_banner/
#    _photo_path en EIV_report.py, adaptado a las rutas propias de Weeks. ──
_LOGO_DIR_CANDIDATES = ["PICS/LOGOS/2026", "PICS/LOGO/2026", "pics/LOGOS/2026"]
_BANNER_CANDIDATES = [
    "PICS/LOGOS/2026/Banner_Semanas.jpg", "PICS/LOGOS/2026/banner_semanas.jpg",
    "pics/LOGOS/2026/Banner_Semanas.jpg", "PICS/LOGOS/2026/Banner_Semanas.JPG",
]
PHOTO_FILENAME_OVERRIDES: Dict[str, str] = {}


def _partner_logo_path(week_key: str) -> Optional[str]:
    """Ruta al logo del socio (KLU.png / NOVA.png / FGV.png / BABSON.png)."""
    fname = WEEK_VISUAL_META.get(week_key, {}).get("logo")
    if not fname:
        return None
    base = _os.path.dirname(_os.path.abspath(__file__))
    for d in _LOGO_DIR_CANDIDATES:
        full = _os.path.join(base, d, fname)
        if _os.path.exists(full):
            return full
    return None


def _show_partner_logo(week_key: str, width: int = 90):
    path = _partner_logo_path(week_key)
    if path:
        try:
            st.image(path, width=width)
        except Exception:
            pass


def _show_weeks_banner(width: Optional[int] = None, use_container_width: bool = False):
    """Banner de portada (Banner_Semanas.jpg), centrado con HTML+base64ing —
    igual criterio que _show_eiv_banner: st.image no se centraba de forma
    confiable con CSS puro."""
    base = _os.path.dirname(_os.path.abspath(__file__))
    path = None
    for rel in _BANNER_CANDIDATES:
        full = _os.path.join(base, rel)
        if _os.path.exists(full):
            path = full
            break
    if not path:
        for d in _LOGO_DIR_CANDIDATES:
            logo_dir = _os.path.join(base, d)
            if _os.path.isdir(logo_dir):
                for fname in sorted(_os.listdir(logo_dir)):
                    if fname.lower().endswith((".jpg", ".jpeg", ".png")) and "banner" in fname.lower():
                        path = _os.path.join(logo_dir, fname)
                        break
            if path:
                break
    if not path:
        return
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
    ext = "jpeg" if path.lower().endswith((".jpg", ".jpeg")) else "png"
    width_css = "100%" if use_container_width else f"{width}px"
    st.markdown(
        f'<div style="width:100%;text-align:center;margin-bottom:10px;">'
        f'<img src="data:image/{ext};base64,{b64}" style="width:{width_css};max-width:100%;border-radius:10px;">'
        f'</div>',
        unsafe_allow_html=True,
    )


def _name_tokens(name: str) -> set:
    """Tokens en minúsculas y sin acentos (p.ej. 'Marcus Thiell' -> {marcus,
    thiell}), usados para emparejar el profesor con su foto sin importar si
    el archivo usa guion o underscore, o si trae solo parte del nombre."""
    import unicodedata, re as _re
    n_ascii = unicodedata.normalize("NFKD", str(name)).encode("ascii", "ignore").decode()
    return set(t for t in _re.split(r"[^a-zA-Z]+", n_ascii.lower()) if t)


def _photo_path(professor: str) -> Optional[str]:
    """Busca la foto del profesor en PICS/PROFES/2026 emparejando por
    tokens del nombre (tolera 'eduardo-boada.jpg' vs 'Eduardo Boada', y
    'marcus_thiell.jpg' vs 'Marcus Thiell'; también resuelve el caso de
    varios profesores en un mismo campo, p.ej. 'Marcus Thiell / Natalia
    Franco', probando cada nombre por separado)."""
    if not professor:
        return None
    base = _os.path.dirname(_os.path.abspath(__file__))
    photo_dir = None
    for d in ("PICS/PROFES/2026", "PICS/PROFES", "pics/PROFES/2026"):
        full = _os.path.join(base, d)
        if _os.path.isdir(full):
            photo_dir = full
            break
    if not photo_dir:
        return None

    files = [f for f in _os.listdir(photo_dir) if f.lower().endswith((".jpg", ".jpeg", ".png"))]
    if not files:
        return None

    import re as _re
    individual_names = [n for n in _re.split(r"[,/&]| y | and ", professor) if n.strip()] or [professor]
    if professor in PHOTO_FILENAME_OVERRIDES:
        for f in files:
            if _os.path.splitext(f)[0] == PHOTO_FILENAME_OVERRIDES[professor]:
                return _os.path.join(photo_dir, f)

    best_file, best_score = None, 0
    for person in individual_names:
        p_tokens = _name_tokens(person)
        if not p_tokens:
            continue
        for f in files:
            f_tokens = _name_tokens(_os.path.splitext(f)[0])
            score = len(p_tokens & f_tokens)
            if score > best_score:
                best_score, best_file = score, f
    if best_file and best_score > 0:
        return _os.path.join(photo_dir, best_file)
    return None


# ── 3) FILE IDs (Google Drive) — carpeta Reportes/INT_WEEKS ─────────────────
LISTAS_FILE_ID = "1Gxev_FWI_mfav3dVWaZczC49F_68Qj9f"       # BD_listas.xlsx (roster + pagos)
SEMANAS_FILE_ID = "1UXmTsOp1X9DKA_OpFy7kmg4fz2_uQT-W"      # BD_semanas.xlsx (metadata por semana)
PRESUPUESTO_FILE_ID = "1835F0CN4QoDgmCoShnbB7xvAnXOwexKO"  # BD_presupuesto.xlsx (budget + TRM)
ENCUESTA_FILE_ID = "1C2xFwqH3FRgBpcXbzs7_NEj1eVllL4xs"     # BD_encuesta_curso.xlsx (frecuencias + comentarios)

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
                "📦 Falta instalar `google-auth` en el entorno "
                f"(el import falló con: `{_GSPREAD_IMPORT_ERR}`)."
            )
        elif "gcp_service_account" not in st.secrets:
            st.error(
                "🔑 No encuentro `st.secrets['gcp_service_account']`. Revisa en "
                "Streamlit Cloud → tu app → Settings → Secrets."
            )
        else:
            st.error("🔑 Las credenciales de `gcp_service_account` no se pudieron usar para autenticar.")
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
        f"🌐 No se pudo descargar el archivo de Drive ({file_id}) tras varios intentos: "
        f"{last_err}\n\nVerifica que el archivo esté compartido con el correo de la "
        "service account, con permiso de Editor."
    )
    st.stop()


# ── 4) METADATA VISUAL POR SOCIO ────────────────────────────────────────
# Estos datos (color, logo, coordenadas) no viven en ninguna de las 4 hojas
# fuente — son identidad visual fija por socio, igual que CATEGORY_LOGOS en
# el reporte de Seminars. Si el próximo año hay un socio nuevo, agrégalo aquí
# (fallback abajo asigna color/posición genérica para no romper el reporte).
WEEK_VISUAL_META = {
    "KLU":    {"color": "#F2994A", "colorSoft": "#FCEADA", "lat": 53.5511, "lon": 9.9937,
               "countryCode": "DEU", "geo_scope": "europe", "logo": "KLU.png"},
    "NOVA":   {"color": "#9B6FD1", "colorSoft": "#EDE3F8", "lat": 38.6979, "lon": -9.3389,
               "countryCode": "PRT", "geo_scope": "europe", "logo": "NOVA.png"},
    "FGV":    {"color": "#4FA8D8", "colorSoft": "#DCEEF9", "lat": -23.5505, "lon": -46.6333,
               "countryCode": "BRA", "geo_scope": "south america", "logo": "FGV.png"},
    "BABSON": {"color": "#5FB88A", "colorSoft": "#DEF3E6", "lat": 42.2967, "lon": -71.2924,
               "countryCode": "USA", "geo_scope": "usa", "logo": "BABSON.png"},
}
_FALLBACK_PALETTE = ["#4C6EF5", "#F2994A", "#9B6FD1", "#4FA8D8", "#5FB88A", "#E61166"]

# BD_listas usa "Producto" con variantes (p.ej. "KLU con hotel"/"KLU sin hotel")
# que colapsan a la misma semana — mapa producto -> key canónico.
PRODUCT_TO_KEY = {
    "KLU": "KLU", "KLU CON HOTEL": "KLU", "KLU SIN HOTEL": "KLU",
    "NOVA": "NOVA", "FGV": "FGV", "BABSON": "BABSON",
}

# Nota: los emojis de bandera (🇩🇪 🇵🇹 🇧🇷 🇺🇸) son en realidad 2 "regional
# indicator symbols" combinados — sin una fuente de emoji con esa ligadura,
# el navegador/Streamlit cae al código de país en texto plano (p.ej. "BR",
# "PT"), que fue justo lo que se veía en la navegación. Se reemplazan por un
# emoji de un solo glifo, distinto por semana, que siempre renderiza igual.
FLAGS = {"Germany": "🇩🇪", "Portugal": "🇵🇹", "Brazil": "🇧🇷", "USA": "🇺🇸"}

BOGOTA = (4.711, -74.0721)

LIKERT_SCORE = {
    "Totalmente en desacuerdo": 1, "En desacuerdo": 2, "Ni de acuerdo ni en desacuerdo": 3,
    "De acuerdo": 4, "Totalmente de acuerdo": 5,
}


def _week_key_from_name(name: str) -> str:
    """De 'International Week' / 'international week' (BD_semanas /
    BD_presupuesto) a la clave canónica (KLU/NOVA/FGV/BABSON/...)."""
    n = str(name).strip().upper()
    for k in WEEK_VISUAL_META:
        if n.startswith(k):
            return k
    if "GETULIO" in n or "VARGAS" in n:
        return "FGV"
    return n.split()[0] if n.split() else n


def _week_meta(key: str, idx: int = 0) -> dict:
    if key in WEEK_VISUAL_META:
        return WEEK_VISUAL_META[key]
    color = _FALLBACK_PALETTE[idx % len(_FALLBACK_PALETTE)]
    return {"color": color, "colorSoft": color + "22", "lat": BOGOTA[0], "lon": BOGOTA[1],
            "countryCode": None, "geo_scope": "world", "logo": None}


# ── 5) CARGA DE HOJAS CRUDAS ─────────────────────────────────────────────
@st.cache_data(ttl=300)
def load_raw_sheets():
    raw_sem = io.BytesIO(_download_drive_file_bytes(SEMANAS_FILE_ID))
    df_semanas = pd.read_excel(raw_sem, sheet_name="Semanas Internacionales")
    df_semanas.columns = df_semanas.columns.str.strip()

    raw_lis = io.BytesIO(_download_drive_file_bytes(LISTAS_FILE_ID))
    df_listas = pd.read_excel(raw_lis, sheet_name="listas")
    df_listas.columns = df_listas.columns.str.strip()

    raw_pre = io.BytesIO(_download_drive_file_bytes(PRESUPUESTO_FILE_ID))
    df_budget = pd.read_excel(raw_pre, sheet_name="budget")
    df_budget.columns = df_budget.columns.str.strip()
    df_trm = pd.read_excel(raw_pre, sheet_name="TRM")
    df_trm.columns = df_trm.columns.str.strip()

    raw_enc = io.BytesIO(_download_drive_file_bytes(ENCUESTA_FILE_ID))
    df_frec = pd.read_excel(raw_enc, sheet_name="frecuencias")
    df_frec.columns = df_frec.columns.str.strip()
    df_com = pd.read_excel(raw_enc, sheet_name="comentarios")
    df_com.columns = df_com.columns.str.strip()

    return df_semanas, df_listas, df_budget, df_trm, df_frec, df_com


COMMENT_QID_LABELS = {
    "P216": "Course — Valued Most", "P216V01": "Faculty — Valued Most",
    "P217": "Course — Could Improve", "P217V01": "Faculty — Could Improve",
}


# ── 6) CONSTRUCCIÓN DEL DATASET (equivalente al DATA embebido en el HTML) ──
@st.cache_data(ttl=300)
def build_weeks() -> Tuple[List[dict], dict]:
    """Reconstruye, por semana (KLU/NOVA/FGV/BABSON), exactamente la misma
    agregación que el HTML original tenía precomputada: roster + pagos
    (BD_listas), metadata (BD_semanas), presupuesto (BD_presupuesto), y
    encuesta (BD_encuesta_curso), unidos por nombre de semana / profesor."""
    df_semanas, df_listas, df_budget, df_trm, df_frec, df_com = load_raw_sheets()

    trm_map = dict(zip(df_trm["Moneda"].astype(str).str.strip(), df_trm["TRM"]))

    df_listas = df_listas.copy()
    df_listas["_key"] = df_listas["Producto"].astype(str).str.strip().str.upper().map(PRODUCT_TO_KEY)
    df_listas["_key"] = df_listas["_key"].fillna(df_listas["Producto"].astype(str).str.strip().str.upper())

    df_semanas = df_semanas.copy()
    df_semanas["_key"] = df_semanas["International Week"].map(_week_key_from_name)

    df_budget = df_budget.copy()
    df_budget["_key"] = df_budget["international week"].map(_week_key_from_name)

    weeks: List[dict] = []
    keys_ordered = df_semanas["_key"].tolist()

    for idx, (_, srow) in enumerate(df_semanas.iterrows()):
        key = srow["_key"]
        meta = _week_meta(key, idx)
        rows = df_listas[df_listas["_key"] == key].copy()

        participants = []
        for _, p in rows.iterrows():
            _codigo_raw = p.get("Código de Estudiante")
            # Excel a veces entrega el código como float (202112345.0) cuando
            # la columna tiene celdas vacías mezcladas — se limpia aquí para
            # que en toda la app (tabla y Excel) salga sin decimales.
            if pd.isna(_codigo_raw):
                _codigo = None
            elif isinstance(_codigo_raw, float) and _codigo_raw.is_integer():
                _codigo = str(int(_codigo_raw))
            else:
                _codigo = str(_codigo_raw).strip()
                if _codigo.endswith(".0"):
                    _codigo = _codigo[:-2]
            participants.append({
                "nombre": p.get("Nombre"), "genero": p.get("Género"),
                "codigo": _codigo, "programa": p.get("Programa") or "Unspecified",
                "tipoPrograma": p.get("Tipo programa"),
                "pago1": pd.notna(p.get("Primer Pago")),
                "pago2": pd.notna(p.get("Segundo Pago")),
                "matricula": pd.notna(p.get("Matricula")),
            })

        registered_count = len(rows)
        abroad_count = sum(1 for p in participants if p["matricula"])
        gender = {"F": 0, "M": 0}
        programas: Dict[str, int] = {}
        programa_gender: Dict[str, Dict[str, int]] = {}
        for p in participants:
            if not p["matricula"]:
                continue
            g = str(p["genero"]).strip().upper() if p["genero"] else "Unknown"
            if g in gender:
                gender[g] += 1
            prog = p["programa"]
            programas[prog] = programas.get(prog, 0) + 1
            programa_gender.setdefault(prog, {"F": 0, "M": 0})
            if g in ("F", "M"):
                programa_gender[prog][g] += 1

        # ---- Presupuesto ----
        wb = df_budget[df_budget["_key"] == key]
        budget_lines = []
        for _, b in wb.iterrows():
            budget_lines.append({
                "concepto": b.get("concepto"), "tipo": b.get("tipo"), "precio": b.get("precio"),
                "moneda": b.get("moneda"), "cantidad": b.get("cantidad"),
                "presupuestado": b.get("presupuestado"),
                "ejecutado": b.get("ejecutado") if pd.notna(b.get("ejecutado")) else None,
            })
        ingresos = sum(b["presupuestado"] for b in budget_lines if b["tipo"] == "ingreso")
        egresos = sum(abs(b["presupuestado"]) for b in budget_lines if b["tipo"] == "gasto")
        balance = ingresos - egresos
        margin_pct = round((balance / ingresos * 100), 2) if ingresos else 0.0

        # ---- Encuesta: frecuencias ----
        prof_name = str(srow.get("Profesor/s") or "").strip()
        f_rows = df_frec[df_frec["nombre_profesor"].astype(str).str.strip().str.upper() == prof_name.upper()]

        questions = []
        aspects_map: Dict[str, Dict[str, float]] = {}
        general_options: Dict[str, int] = {}
        carga_options: Dict[str, int] = {}
        cantidad_options: Dict[str, int] = {}

        for qid, qrows in f_rows.groupby("id_pregunta"):
            qtext = qrows["pregunta"].iloc[0]
            aspecto = qrows["aspecto_evaluado"].iloc[0]
            qtype = qrows["opcion_respuesta_pregunta"].iloc[0]
            options: Dict[str, int] = {}
            for _, r in qrows.iterrows():
                resp = str(r["respuesta"])
                options[resp] = options.get(resp, 0) + int(r["respuestas_por_opcion"])
            questions.append({"text": qtext, "aspect": aspecto, "type": qtype, "options": options})

            n_q = sum(options.values())
            if qtype == "acuerdo":
                # "No aplica" (y cualquier opción sin score Likert) no cuenta
                # ni en el promedio ni en el denominador — así coincide
                # exactamente con el cálculo original.
                valid = {k: v for k, v in options.items() if k in LIKERT_SCORE}
                n_valid = sum(valid.values())
                score = sum(LIKERT_SCORE[k] * v for k, v in valid.items())
                d = aspects_map.setdefault(aspecto, {"score_sum": 0.0, "n": 0})
                d["score_sum"] += score
                d["n"] += n_valid
            elif qtype == "escala010":
                for k, v in options.items():
                    general_options[k] = general_options.get(k, 0) + v
            elif qtype == "carga":
                for k, v in options.items():
                    carga_options[k] = carga_options.get(k, 0) + v
            elif qtype == "cantidad":
                for k, v in options.items():
                    cantidad_options[k] = cantidad_options.get(k, 0) + v

        aspects = [{"aspect": a, "avg": round(d["score_sum"] / d["n"], 2) if d["n"] else 0.0, "n": d["n"]}
                   for a, d in aspects_map.items()]
        aspects.sort(key=lambda a: -a["avg"])

        n_total = sum(d["n"] for d in aspects_map.values())
        score_total = sum(d["score_sum"] for d in aspects_map.values())
        avg_overall = round(score_total / n_total, 2) if n_total else None

        respondents = sum(general_options.values())
        nps = round(sum(int(k) * v for k, v in general_options.items()) / respondents, 2) if respondents else None
        response_rate = round(respondents / abroad_count * 100, 1) if abroad_count else 0.0

        workload_n = sum(carga_options.values())
        wl_score_map = {"Nada": 1, "Poco": 2, "Mucho": 3, "Demasiado": 4}
        workload_avg = (round(sum(wl_score_map.get(k, 0) * v for k, v in carga_options.items()) / workload_n, 2)
                         if workload_n else None)

        objectives_n = sum(cantidad_options.values())
        objectives_all_pct = (round(cantidad_options.get("Todos", 0) / objectives_n * 100, 1)
                               if objectives_n else None)

        # ---- Encuesta: comentarios ----
        # Cada fila de 'comentarios' cuenta como una respuesta, incluso si
        # viene en blanco/NaN — el reporte original las conserva como "NA"
        # en vez de descartarlas (así el conteo de comentarios coincide con
        # el número real de estudiantes que respondieron esa pregunta).
        c_rows = df_com[df_com["nombre_profesor"].astype(str).str.strip().str.upper() == prof_name.upper()]
        comment_groups = []
        for qid, label in COMMENT_QID_LABELS.items():
            items = c_rows[c_rows["id_pregunta"] == qid]["respuesta"].fillna("NA").tolist()
            items = [str(x).strip() if str(x).strip() else "NA" for x in items]
            comment_groups.append({"label": label, "items": items, "n": len(items)})

        survey = {
            "n": n_total, "avg": avg_overall, "nps": nps, "npsN": respondents,
            "aspects": aspects, "commentGroups": comment_groups, "respondents": respondents,
            "responseRate": response_rate, "workloadAvg": workload_avg, "workloadN": workload_n,
            "objectivesAllPct": objectives_all_pct, "objectivesN": objectives_n,
            "questions": questions,
        }

        weeks.append({
            "key": key, "name": srow.get("International Week"), "course": srow.get("Nombre"),
            "professor": prof_name, "location": srow.get("Ubicación"), "country": srow.get("País"),
            "dates": srow.get("fechas"), "invoiceCurrency": srow.get("Moneda"),
            "invoiceValue": srow.get("Valor factura"), "invoiceAttendees": srow.get("Asistentes"),
            "registeredCount": registered_count, "abroadCount": abroad_count,
            "gender": gender, "programas": programas, "programaGender": programa_gender,
            "participantsList": participants, "budgetLines": budget_lines,
            "ingresos": ingresos, "egresos": egresos, "balance": balance, "marginPct": margin_pct,
            "survey": survey, **meta,
        })

    totals = {
        "weeks": len(weeks), "countries": len(set(w["country"] for w in weeks)),
        "studentsAbroad": sum(w["abroadCount"] for w in weeks),
        "totalIncome": sum(w["ingresos"] for w in weeks),
        "totalExpense": sum(w["egresos"] for w in weeks),
        "totalBalance": sum(w["balance"] for w in weeks),
    }
    totals["marginPct"] = round(totals["totalBalance"] / totals["totalIncome"] * 100, 2) if totals["totalIncome"] else 0.0

    return weeks, totals


# ── 7) PÁGINA — Cover / Data Center ──────────────────────────────────────
def page_cover():
    """Portada + Data Center, mismo patrón que page_cover() en EIV_report.py:
    banner, título, botón 'Enter the Report', y lista de fuentes con enlace
    de descarga directa de cada archivo de Drive."""
    weeks, totals = build_weeks()
    st.markdown("<div style='height:0.2vh;'></div>", unsafe_allow_html=True)
    with st.container(key="cover_banner"):
        try:
            _show_weeks_banner(width=777)
        except Exception:
            pass

    col_l, col_mid, col_r = st.columns([1, 2, 1])
    with col_mid:
        st.markdown(
            f'<div style="text-align:center;font-size:38px;font-weight:800;color:{INK};margin-top:18px;">'
            'International Weeks</div>'
            f'<div style="text-align:center;font-size:18px;color:{MUTED};margin:10px auto 0;line-height:1.5;">'
            "Analytics for the 2026 edition of International Weeks — enrollment, roster & payments, "
            "course evaluation, and financial performance across every partner-school immersion.</div>",
            unsafe_allow_html=True,
        )

        st.markdown("<div style='height:26px;'></div>", unsafe_allow_html=True)
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
            ("BD_semanas.xlsx", "Week metadata — professor, dates, location, invoice", SEMANAS_FILE_ID),
            ("BD_listas.xlsx", "Enrollment, payments & program catalog per week", LISTAS_FILE_ID),
            ("BD_presupuesto.xlsx", "Budget by concept — income & expenses, per week", PRESUPUESTO_FILE_ID),
            ("BD_encuesta_curso.xlsx", "Course evaluation survey — frecuencias & comentarios", ENCUESTA_FILE_ID),
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
                st.download_button(
                    "", data=_download_drive_file_bytes(fid), file_name=fname,
                    key=f"cover_dl_{fname}", use_container_width=True,
                    icon=":material/download:",
                )
        st.markdown(
            f'<div style="text-align:center;margin-top:14px;font-size:11.5px;color:{MUTED};">'
            f"{len(files)} sources loaded · {totals['weeks']} international weeks</div>",
            unsafe_allow_html=True,
        )


# ── 8) PÁGINA — Overview ─────────────────────────────────────────────────
def _partner_card(w: dict, show_link: bool = True):
    gender_txt = "   ".join(f'{"♀" if g == "F" else "♂"} {n}' for g, n in w["gender"].items())
    flag = FLAGS.get(w["country"], "")
    st.markdown(
        f'<div class="partner-card" style="--wc:{w["color"]};">'
        f'<div class="loc">{flag} {w["location"]}, {w["country"]} · {w["dates"]}</div>'
        f'<div class="n">{w["abroadCount"]}</div><div class="n-label">Students Abroad</div>'
        f'<div style="font-family:monospace;font-size:16px;font-weight:600;margin-bottom:10px;">{gender_txt}</div>'
        f'<div style="display:flex;justify-content:space-between;border-top:1px solid {LINE};'
        f'padding-top:8px;font-size:12px;"><span>Income</span>'
        f'<span style="font-family:monospace;color:{INCOME};font-weight:600;">{_fmt_money(w["ingresos"])}</span></div>'
        f'<div style="display:flex;justify-content:space-between;font-size:12px;"><span>Expenses</span>'
        f'<span style="font-family:monospace;color:{EXPENSE};font-weight:600;">{_fmt_money(w["egresos"])}</span></div>'
        f'<div style="display:flex;justify-content:space-between;font-size:12px;"><span>Balance</span>'
        f'<span style="font-family:monospace;font-weight:600;">{_fmt_money(w["balance"])}</span></div>'
        f'<span class="margin-pill" style="background:{w["colorSoft"]};color:{w["color"]};">'
        f'{w["marginPct"]}% margin</span></div>',
        unsafe_allow_html=True,
    )
    if show_link:
        btn_key = f"open_week_btn_{w['key']}"
        st.markdown(
            f"<style>.st-key-{btn_key} a{{display:flex !important;align-items:center;"
            "justify-content:center;gap:6px;margin-top:10px;"
            f"background:{w['color']} !important;border:none !important;border-radius:10px !important;"
            "color:#fff !important;font-weight:700 !important;height:44px !important;"
            f"text-decoration:none !important;box-shadow:0 3px 10px rgba(0,0,0,.15) !important;}}"
            f".st-key-{btn_key} a span{{color:#fff !important;}}"
            f".st-key-{btn_key} a:hover{{opacity:.88;}}</style>",
            unsafe_allow_html=True,
        )
        with st.container(key=btn_key):
            st.page_link(week_page_by_key[w["key"]], label=f"Open {w['key']} Report →")


def page_overview():
    weeks, totals = build_weeks()
    _render_header(
        "2026 International Weeks — Program Summary",
        "Graduate international immersions this year, each pairing a partner business school abroad with a School of Management cohort.",
    )

    st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    with col1:
        _kpi("International Weeks", totals["weeks"])
    with col2:
        _kpi("Countries", totals["countries"])
    with col3:
        _kpi("Total Students Abroad", totals["studentsAbroad"])

    col_fin, col_margin = st.columns([1, 2])
    with col_fin:
        st.markdown("##### Financial Balance")
        _kpi("Total Income", _fmt_money(totals["totalIncome"]), "income")
        st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)
        _kpi("Total Expenses", _fmt_money(totals["totalExpense"]), "expense")
        st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)
        _kpi("Total Balance", _fmt_money(totals["totalBalance"]), "balance")
        st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)
        _kpi("Total Margin", f"{totals['marginPct']}%")
    with col_margin:
        st.markdown("##### Margin by Week")
        st.caption(f"Dashed line = blended total margin ({totals['marginPct']}%).")
        fig = go.Figure(go.Bar(
            x=[w["marginPct"] for w in weeks], y=[w["key"] for w in weeks], orientation="h",
            marker=dict(color="#4C8A3F"), text=[f'{w["marginPct"]}%' for w in weeks],
            textposition="outside", cliponaxis=False,
        ))
        fig.add_vline(x=totals["marginPct"], line_dash="dot", line_color=INK,
                       annotation_text=f'Total Margin: {totals["marginPct"]}%', annotation_font_size=9.5)
        fig.update_layout(margin=dict(t=10, r=46, b=26, l=70), xaxis=dict(ticksuffix="%", gridcolor="#EFEBEE"),
                           plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", height=180)
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")
    st.markdown("### Where Students Went")
    st.caption("Click a marker on the map to see that week's partner card, roster, and financials.")

    # Lee la selección del click anterior sobre el mapa (antes de construirlo) —
    # mismo patrón que el mapa clicable de Summary en EIV_report.py.
    _sel = st.session_state.get("ov_map", {})
    _sel_points = _sel.get("selection", {}).get("points", []) if _sel else []
    focus_key = None
    if _sel_points:
        cd = _sel_points[0].get("customdata")
        if cd:
            focus_key = cd[0] if isinstance(cd, (list, tuple)) else cd

    col_map, col_card = st.columns([2, 1])
    with col_map:
        fig_map = go.Figure()
        for w in weeks:
            highlight = not focus_key or focus_key == w["key"]
            fig_map.add_trace(go.Scattergeo(
                lat=[BOGOTA[0], w["lat"]], lon=[BOGOTA[1], w["lon"]], mode="lines",
                line=dict(color=w["color"], width=1.6), opacity=0.7 if highlight else 0.12,
                showlegend=False, hoverinfo="skip",
            ))
        fig_map.add_trace(go.Scattergeo(
            lat=[w["lat"] for w in weeks], lon=[w["lon"] for w in weeks], mode="markers+text",
            text=[w["key"] for w in weeks], customdata=[w["key"] for w in weeks],
            textposition="top center",
            textfont=dict(size=11, color=[INK if (not focus_key or focus_key == w["key"]) else "#C9BEC5" for w in weeks]),
            marker=dict(size=13, color=[w["color"] for w in weeks],
                        opacity=[1 if (not focus_key or focus_key == w["key"]) else 0.25 for w in weeks],
                        line=dict(color="#fff", width=1.5)),
            hovertext=[f'{w["location"]}, {w["country"]} — {w["abroadCount"]} students — click to select' for w in weeks],
            hoverinfo="text",
        ))
        fig_map.add_trace(go.Scattergeo(
            lat=[BOGOTA[0]], lon=[BOGOTA[1]], mode="markers+text", text=["Bogotá"],
            textposition="bottom center", textfont=dict(size=11, color=INK),
            marker=dict(size=11, color=INK, symbol="diamond"), showlegend=False, hoverinfo="skip",
        ))
        fig_map.update_geos(showland=True, landcolor="#F1EDEF", showcountries=True, countrycolor="#DDD4D9",
                             showocean=True, oceancolor="#FAF8FA", bgcolor="rgba(0,0,0,0)",
                             projection_type="natural earth")
        fig_map.update_layout(margin=dict(t=6, r=6, b=6, l=6), height=380, showlegend=False,
                               paper_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig_map, use_container_width=True, on_select="rerun",
                         selection_mode="points", key="ov_map")
        if focus_key:
            st.caption("Click the marker again (or double-click the map) to clear the selection.")
    with col_card:
        if focus_key:
            w_sel = next((w for w in weeks if w["key"] == focus_key), None)
            if w_sel:
                _partner_card(w_sel)
        else:
            st.info("Click a week's marker on the map to see its partner card here.")

    # ---- Participants by Program ----
    st.markdown("##### Participants by Program")
    scoped_weeks = [w for w in weeks if not focus_key or w["key"] == focus_key]
    program_totals: Dict[str, int] = {}
    by_prog_week: Dict[str, Dict[str, int]] = {}
    for w in scoped_weeks:
        for prog, n in w["programas"].items():
            program_totals[prog] = program_totals.get(prog, 0) + n
            by_prog_week.setdefault(prog, {})[w["key"]] = by_prog_week.get(prog, {}).get(w["key"], 0) + n
    programs_sorted = sorted(program_totals, key=lambda p: -program_totals[p])
    if programs_sorted:
        fig_p = go.Figure()
        for w in scoped_weeks:
            fig_p.add_trace(go.Bar(
                x=[by_prog_week.get(p, {}).get(w["key"], 0) for p in programs_sorted],
                y=programs_sorted, orientation="h", name=w["key"], marker=dict(color=w["color"]),
            ))
        fig_p.update_layout(
            barmode="stack", margin=dict(t=6, r=16, b=26, l=200),
            legend=dict(orientation="h", x=0, y=1.12),
            yaxis=dict(tickfont=dict(size=10), automargin=True, autorange="reversed"),
            xaxis=dict(gridcolor="#EFEBEE"),
            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
            height=max(260, 34 * len(programs_sorted)),
        )
        st.plotly_chart(fig_p, use_container_width=True)
    else:
        st.info("No program data available for this filter.")

    st.markdown("---")
    st.markdown("### By Partner")
    st.caption("Open a week for the full roster, payment status, and budget breakdown.")
    cols = st.columns(min(4, len(weeks)) or 1)
    for i, w in enumerate(weeks):
        with cols[i % len(cols)]:
            _partner_card(w)


# ── 9) PÁGINA — Week Detail (una por socio) ─────────────────────────────
def page_week(key: str):
    weeks, _ = build_weeks()
    w = next((x for x in weeks if x["key"] == key), None)
    if w is None:
        st.error(f"No data found for week '{key}'.")
        return

    _render_header(w["name"], w["course"])

    # ---- Banner de color de la semana: logo del socio + info (izquierda),
    # nombre del profesor + foto (derecha) — TODO con estilos inline (no
    # depende de la clase .week-band) para que el layout de una sola fila
    # no dependa del orden de carga del <style> global.
    logo_path = _partner_logo_path(w["key"])
    logo_html = ""
    if logo_path:
        with open(logo_path, "rb") as f:
            _logo_b64 = base64.b64encode(f.read()).decode()
        logo_html = (
            f'<img src="data:image/png;base64,{_logo_b64}" '
            f'style="height:80px;max-width:200px;object-fit:contain;'
            f'flex:0 0 auto;margin-right:20px;">'
        )

    prof_photo = _photo_path(w["professor"])
    if prof_photo:
        with open(prof_photo, "rb") as f:
            _prof_b64 = base64.b64encode(f.read()).decode()
        _prof_ext = "jpeg" if prof_photo.lower().endswith((".jpg", ".jpeg")) else "png"
        prof_photo_html = (
            f'<img src="data:image/{_prof_ext};base64,{_prof_b64}" '
            'style="width:86px;height:86px;border-radius:50%;object-fit:cover;'
            'flex:0 0 auto;border:2px solid #fff;box-shadow:0 2px 8px rgba(0,0,0,.18);">'
        )
    else:
        prof_photo_html = ""

    prof_html = (
        '<div style="margin-left:auto;flex:0 0 auto;display:flex;'
        'flex-direction:row;align-items:center;gap:14px;">'
        '<div style="text-align:right;">'
        f'<div style="font-weight:700;font-size:15px;color:{INK};white-space:nowrap;">{w["professor"]}</div>'
        f'<div style="font-size:10.5px;color:{MUTED};text-transform:uppercase;'
        'font-family:monospace;letter-spacing:.04em;">Faculty Lead</div></div>'
        f'{prof_photo_html}</div>'
    )

    band_style = (
        "display:flex;flex-direction:row;flex-wrap:nowrap;align-items:center;"
        f"border-radius:12px;padding:20px 24px;background:{w['colorSoft']};"
    )
    info_html = (
        '<div style="min-width:0;flex:1 1 auto;overflow:hidden;">'
        f'<div style="font-family:monospace;font-size:11.5px;text-transform:uppercase;'
        f'letter-spacing:.06em;font-weight:600;color:{w["color"]};margin-bottom:0;'
        f'line-height:1;">{w["name"]}</div>'
        '<h2 style="font-size:19px;margin:0;line-height:1.05;white-space:nowrap;'
        f'overflow:hidden;text-overflow:ellipsis;">{w["course"]}</h2>'
        '<div style="font-size:13px;color:#6E5C68;margin-top:0;line-height:1.05;white-space:nowrap;'
        f'overflow:hidden;text-overflow:ellipsis;">{w["location"]}, {w["country"]} · {w["dates"]}</div>'
        '</div>'
    )

    st.markdown(
        f'<div style="{band_style}">{logo_html}{info_html}{prof_html}</div>',
        unsafe_allow_html=True,
    )

    st.markdown("<div style='height:22px;'></div>", unsafe_allow_html=True)
    col_fin, col_att, col_map = st.columns([1, 1, 2])
    with col_fin:
        _kpi("Income", _fmt_money(w["ingresos"]), "income")
        st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)
        _kpi("Expenses", _fmt_money(w["egresos"]), "expense")
        st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)
        _kpi("Balance", _fmt_money(w["balance"]), "balance")
        st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)
        _kpi("Margin", f'{w["marginPct"]}%')
    with col_att:
        _kpi("Students Abroad", w["abroadCount"])
        st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)
        _kpi("Registered, Did Not Travel", w["registeredCount"] - w["abroadCount"])
    with col_map:
        fig_geo = go.Figure()
        if w.get("countryCode"):
            fig_geo.add_trace(go.Choropleth(
                locations=[w["countryCode"]], z=[1], locationmode="ISO-3",
                colorscale=[[0, w["color"]], [1, w["color"]]], showscale=False,
                marker_line_color="#fff", marker_line_width=0.5,
            ))
        fig_geo.add_trace(go.Scattergeo(
            lat=[w["lat"]], lon=[w["lon"]], mode="markers+text", text=[w["location"]],
            textposition="bottom center", textfont=dict(size=11, color=INK),
            marker=dict(size=12, color=INK, line=dict(color="#fff", width=1.5)),
        ))
        fig_geo.update_geos(scope=w.get("geo_scope", "world"), showland=True, landcolor="#F1EDEF",
                             showcountries=True, countrycolor="#DDD4D9", showocean=True,
                             oceancolor="#FAF8FA", bgcolor="rgba(0,0,0,0)")
        fig_geo.update_layout(margin=dict(t=6, r=6, b=6, l=6), height=220, showlegend=False,
                               paper_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig_geo, use_container_width=True)

    # ---- Pyramid: participants by program, gender split ----
    st.markdown("### Participants by Program")
    st.caption("Women (left) vs. men (right); total shown at right.")
    programs = sorted(w["programas"], key=lambda p: -w["programas"][p])
    if programs:
        women = [-(w["programaGender"].get(p, {}).get("F", 0)) for p in programs]
        men = [w["programaGender"].get(p, {}).get("M", 0) for p in programs]
        totals_p = [abs(a) + b for a, b in zip(women, men)]
        fig_pyr = go.Figure()
        fig_pyr.add_trace(go.Bar(x=women, y=programs, orientation="h", name="Women",
                                  marker=dict(color=w["colorSoft"]),
                                  text=[str(-v) for v in women], textposition="inside",
                                  textfont=dict(color=w["color"])))
        fig_pyr.add_trace(go.Bar(x=men, y=programs, orientation="h", name="Men",
                                  marker=dict(color=w["color"]),
                                  text=[str(v) for v in men], textposition="inside",
                                  textfont=dict(color="#fff")))
        for p, t in zip(programs, totals_p):
            fig_pyr.add_annotation(xref="paper", x=1.03, xanchor="left", y=p, showarrow=False,
                                    text=f"Total: {t}", font=dict(size=9.5, color=MUTED))
        fig_pyr.update_layout(
            barmode="overlay", margin=dict(t=10, r=90, b=26, l=190),
            legend=dict(orientation="h", x=0, y=1.12),
            xaxis=dict(gridcolor="#EFEBEE", zeroline=True, zerolinecolor="#D8CFD5"),
            yaxis=dict(tickfont=dict(size=10.5), automargin=True, autorange="reversed"),
            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", height=280,
        )
        st.plotly_chart(fig_pyr, use_container_width=True)
    else:
        st.info("No enrolled/traveling participants for this week yet.")

    # ---- Roster ----
    # Cualquier estudiante con X en 1st Payment, 2nd Payment o Enrolled
    # (i.e. no completó ese trámite) queda al final de la lista y en gris
    # tenue, como si no hubiera viajado. La columna auxiliar usada para
    # ordenar/colorear NUNCA se muestra: se calcula aparte y no entra al
    # DataFrame que se despliega, para no depender de Styler.hide().
    st.markdown(f'### Roster ({w["registeredCount"]} registered · {w["abroadCount"]} traveled)')
    incomplete = lambda p: not (p["pago1"] and p["pago2"] and p["matricula"])
    participants_sorted = sorted(w["participantsList"], key=lambda p: 1 if incomplete(p) else 0)
    incomplete_flags = [incomplete(p) for p in participants_sorted]

    roster_df = pd.DataFrame([{
        "Name": p["nombre"], "Code": p["codigo"] if p["codigo"] else "—", "Program": p["programa"],
        "Type": p["tipoPrograma"], "1st Payment": tick(p["pago1"]), "2nd Payment": tick(p["pago2"]),
        "Enrolled": tick(p["matricula"]),
    } for p in participants_sorted])

    def _style_roster(row: pd.Series) -> List[str]:
        if incomplete_flags[row.name]:
            return ["color:#ADA6AB;"] * len(row)
        return [""] * len(row)

    styled_roster = roster_df.style.apply(_style_roster, axis=1)
    st.dataframe(styled_roster, use_container_width=True, hide_index=True, height=420)
    st.download_button(
        "Download as Excel",
        data=_xlsx_bytes(pd.DataFrame([{
            "Name": p["nombre"], "Code": p["codigo"], "Program": p["programa"], "Type": p["tipoPrograma"],
            "1st Payment": "Yes" if p["pago1"] else "No", "2nd Payment": "Yes" if p["pago2"] else "No",
            "Enrolled": "Yes" if p["matricula"] else "No",
        } for p in participants_sorted]), "Roster"),
        file_name=f'{w["key"]}_Roster.xlsx', key=f'dl_roster_{w["key"]}',
        icon=":material/download:",
    )

    # ---- Survey ----
    st.markdown("### Course Evaluation Survey")
    sv = w["survey"]
    st.caption(f'{sv["respondents"]} of {w["abroadCount"]} students abroad responded to the survey.')
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        _kpi("Students Responded", sv["respondents"])
    with c2:
        _kpi("Response Rate", f'{sv["responseRate"]}%')
    with c3:
        _kpi("Avg. Satisfaction", f'{sv["avg"]} / 5.0' if sv["avg"] is not None else "—")
    with c4:
        _kpi("Would Recommend", f'{sv["nps"]} / 10' if sv["npsN"] else "—")
    with c5:
        _kpi("Objectives Met", f'{sv["objectivesAllPct"]}%' if sv["objectivesN"] else "—")

    if sv["aspects"]:
        st.markdown("##### By Aspect")
        fig_asp = go.Figure(go.Bar(
            x=[a["avg"] for a in sv["aspects"]], y=[a["aspect"] for a in sv["aspects"]], orientation="h",
            marker=dict(color=w["color"]), text=[f'{a["avg"]:.2f}' for a in sv["aspects"]],
            textposition="outside", cliponaxis=False,
        ))
        fig_asp.update_layout(
            margin=dict(t=10, r=50, b=26, l=270), xaxis=dict(range=[1, 5.5], gridcolor="#EFEBEE"),
            yaxis=dict(tickfont=dict(size=10.5), automargin=True, autorange="reversed"),
            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", height=280,
        )
        st.plotly_chart(fig_asp, use_container_width=True)

    st.markdown("##### Comments")
    ccols = st.columns(2)
    for i, g in enumerate(sv["commentGroups"]):
        with ccols[i % 2]:
            st.markdown(f'**{g["label"]} ({g["n"]})**')
            if g["items"]:
                with st.container(height=200):
                    quotes_html = "".join(
                        f'<div class="survey-quote">“{item}”</div>' for item in g["items"]
                    )
                    st.markdown(quotes_html, unsafe_allow_html=True)
            else:
                st.caption("No comments for this question.")

    st.markdown("<div style='height:18px;'></div>", unsafe_allow_html=True)
    # Misma lógica de EIV_report.py: cachea el PDF en session_state por
    # semana, así solo se reconstruye cuando cambian los datos de esa
    # semana en vez de en cada rerun de Streamlit.
    pdf_key = w["key"]
    cache = st.session_state.get("_weeks_pdf_cache", {})
    if cache.get("key") != pdf_key:
        with st.spinner("Preparing report…"):
            pdf_bytes = _build_week_survey_pdf(w)
        st.session_state["_weeks_pdf_cache"] = {"key": pdf_key, "data": pdf_bytes}
        cache = st.session_state["_weeks_pdf_cache"]
    col_l, col_mid, col_r = st.columns([1, 2, 1])
    with col_mid:
        with st.container(key=f"gen_report_btn_{w['key']}"):
            st.download_button(
                "Generate Evaluation Report", data=cache["data"],
                file_name=f'ISS_{w["key"]}_Evaluation_Report.pdf', mime="application/pdf",
                key=f'dl_pdf_{w["key"]}', use_container_width=True,
                icon=":material/picture_as_pdf:",
            )


# ── 10) PÁGINA — Financial Detail ────────────────────────────────────────
def page_financial():
    weeks, totals = build_weeks()
    _render_header(
        "Financial Detail",
        "Full budget by concept for every week — sourced from BD_presupuesto.xlsx. "
        "\"Executed\" fills in once a line item is reconciled.",
    )

    st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        _kpi("Total Income", _fmt_money(totals["totalIncome"]), "income")
    with c2:
        _kpi("Total Expenses", _fmt_money(totals["totalExpense"]), "expense")
    with c3:
        _kpi("Total Balance", _fmt_money(totals["totalBalance"]), "balance")
    with c4:
        _kpi("Total Margin", f'{totals["marginPct"]}%')

    st.markdown("##### Income vs. Expenses by Week")
    fig = go.Figure()
    fig.add_trace(go.Bar(x=[w["key"] for w in weeks], y=[w["ingresos"] for w in weeks],
                          name="Income", marker=dict(color=INCOME)))
    fig.add_trace(go.Bar(x=[w["key"] for w in weeks], y=[w["egresos"] for w in weeks],
                          name="Expenses", marker=dict(color=EXPENSE)))
    fig.update_layout(
        barmode="group", margin=dict(t=20, r=10, b=30, l=90),
        legend=dict(orientation="h", x=0, y=1.14), yaxis=dict(tickformat="$,.0f", gridcolor="#EFEBEE"),
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", height=320,
    )
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("##### Comparison")
    compare_df = pd.DataFrame([{
        "Week": w["name"], "Income": _fmt_money(w["ingresos"]), "Expenses": _fmt_money(w["egresos"]),
        "Balance": _fmt_money(w["balance"]), "Margin": f'{w["marginPct"]}%',
    } for w in weeks] + [{
        "Week": "Total", "Income": _fmt_money(totals["totalIncome"]), "Expenses": _fmt_money(totals["totalExpense"]),
        "Balance": _fmt_money(totals["totalBalance"]), "Margin": f'{totals["marginPct"]}%',
    }])
    st.dataframe(compare_df, use_container_width=True, hide_index=True)

    compare_raw = pd.DataFrame([{
        "Week": w["name"], "Income": w["ingresos"], "Expenses": w["egresos"],
        "Balance": w["balance"], "Margin %": w["marginPct"],
    } for w in weeks] + [{
        "Week": "Total", "Income": totals["totalIncome"], "Expenses": totals["totalExpense"],
        "Balance": totals["totalBalance"], "Margin %": totals["marginPct"],
    }])
    st.download_button(
        "Download as Excel", data=_xlsx_bytes(compare_raw, "Comparison"),
        file_name="International_Weeks_Financial_Comparison.xlsx", key="dl_fin_comparison",
        icon=":material/download:",
    )

    st.markdown("---")
    st.markdown("### Detailed Budget by Week")
    for w in weeks:
        st.markdown(f'<div style="border-top:5px solid {w["color"]};border-radius:4px;"></div>',
                     unsafe_allow_html=True)
        st.markdown(f'#### {w["name"]}')
        used_currencies = sorted(set(b["moneda"] for b in w["budgetLines"] if b["moneda"] != "COP"))
        _, _, _, df_trm, _, _ = load_raw_sheets()
        trm_map = dict(zip(df_trm["Moneda"].astype(str).str.strip(), df_trm["TRM"]))
        if used_currencies:
            trm_note = "TRM used: " + "  ·  ".join(f"{c} = {_fmt_money(trm_map.get(c, 0))}" for c in used_currencies)
        else:
            trm_note = "All line items in COP."
        st.caption(trm_note)

        col_inc, col_exp = st.columns(2)
        income_lines = [b for b in w["budgetLines"] if b["tipo"] == "ingreso"]
        expense_lines = [b for b in w["budgetLines"] if b["tipo"] == "gasto"]

        def _lines_df(lines, is_expense):
            return pd.DataFrame([{
                "Concept": b["concepto"],
                "Unit Price": f'{b["precio"]:,.0f} {b["moneda"]}' if b["precio"] else "—",
                "Qty": b["cantidad"],
                "Total (COP)": _fmt_money(abs(b["presupuestado"]) if is_expense else b["presupuestado"]),
                "Executed": _fmt_money(abs(b["ejecutado"])) if b["ejecutado"] is not None else "pending",
            } for b in lines])

        with col_inc:
            st.markdown(f'<span style="color:{INCOME};font-weight:600;">Income</span>', unsafe_allow_html=True)
            st.dataframe(_lines_df(income_lines, False), use_container_width=True, hide_index=True)
            st.markdown(f'**Total Income: <span style="color:{INCOME};">{_fmt_money(w["ingresos"])}</span>**',
                        unsafe_allow_html=True)
        with col_exp:
            st.markdown(f'<span style="color:{EXPENSE};font-weight:600;">Expenses</span>', unsafe_allow_html=True)
            st.dataframe(_lines_df(expense_lines, True), use_container_width=True, hide_index=True)
            st.markdown(f'**Total Expenses: <span style="color:{EXPENSE};">{_fmt_money(w["egresos"])}</span>**',
                        unsafe_allow_html=True)

        st.markdown(f'**Balance ({w["marginPct"]}% margin): {_fmt_money(w["balance"])}**')

        budget_full_df = pd.DataFrame([{
            "Concept": b["concepto"], "Type": b["tipo"], "Unit Price": b["precio"], "Currency": b["moneda"],
            "Qty": b["cantidad"], "Budgeted_COP": b["presupuestado"], "Executed_COP": b["ejecutado"],
        } for b in w["budgetLines"]])
        st.download_button(
            "Download as Excel", data=_xlsx_bytes(budget_full_df, "Budget"),
            file_name=f'{w["key"]}_Budget.xlsx', key=f'dl_budget_{w["key"]}',
            icon=":material/download:",
        )
        st.markdown("---")


# ── 11) PDF — Evaluation Report per week ─────────────────────────────────
# Helpers de dibujo — traducidos 1:1 de EIV_report.py (_pdf_stacked_bar,
# _pdf_legend, _pdf_nps_bar, _wrap_to_width, _pdf_file_bytes): todo se
# razona en mm (y_mm = distancia desde ARRIBA de la página), convirtiendo a
# puntos de reportlab solo al dibujar. Reutilizables tal cual porque no
# dependen de nada específico de EIV (reciben color y datos por parámetro).
def _pdf_stacked_bar(c, options: dict, order: list, color_map: dict, x_mm, y_mm, w_mm, h_mm, page_h_mm):
    from reportlab.lib import colors as rl_colors
    from reportlab.lib.units import mm
    total = sum(options.get(k, 0) for k in order)
    if not total:
        return y_mm + h_mm + 2
    top_pt = (page_h_mm - y_mm) * mm
    h_pt = h_mm * mm
    cx = x_mm * mm
    for k in order:
        n = options.get(k, 0)
        if not n:
            continue
        sw = (w_mm * mm) * (n / total)
        col = color_map.get(k, (180, 180, 180))
        c.setFillColor(rl_colors.Color(col[0] / 255, col[1] / 255, col[2] / 255))
        c.rect(cx, top_pt - h_pt, sw, h_pt, fill=1, stroke=0)
        if sw > 14:
            c.setFont("Helvetica-Bold", 6.5)
            c.setFillColor(rl_colors.white)
            c.drawCentredString(cx + sw / 2, top_pt - h_pt / 2 - 2.2, f"{round(n / total * 100)}%")
        cx += sw
    c.setStrokeColor(rl_colors.HexColor("#C8C8C8"))
    c.setLineWidth(0.4)
    c.rect(x_mm * mm, top_pt - h_pt, w_mm * mm, h_pt, fill=0, stroke=1)
    return y_mm + h_mm + 2


def _pdf_legend(c, options: dict, order: list, color_map: dict, x_mm, y_mm, w_mm, page_h_mm):
    from reportlab.lib import colors as rl_colors
    from reportlab.lib.units import mm
    lx_mm, ly_mm = x_mm, y_mm
    c.setFont("Helvetica", 6.5)
    for k in order:
        n = options.get(k, 0)
        if not n:
            continue
        label = f"{k} ({n})"
        item_w_mm = (c.stringWidth(label, "Helvetica", 6.5) / mm) + 5 + 3
        if lx_mm + item_w_mm > x_mm + w_mm and lx_mm > x_mm:
            lx_mm, ly_mm = x_mm, ly_mm + 5
        col = color_map.get(k, (180, 180, 180))
        top_pt = (page_h_mm - ly_mm) * mm
        c.setFillColor(rl_colors.Color(col[0] / 255, col[1] / 255, col[2] / 255))
        c.rect(lx_mm * mm, top_pt - 2.4 * mm, 3.5 * mm, 3.2 * mm, fill=1, stroke=0)
        c.setFillColor(rl_colors.HexColor(INK))
        c.drawString(lx_mm * mm + 5 * mm, top_pt - 2.4 * mm, label)
        lx_mm += item_w_mm + 5
    return ly_mm + 6


def _pdf_nps_bar(c, options: dict, x_mm, y_mm, w_mm, h_mm, page_h_mm):
    from reportlab.lib import colors as rl_colors
    from reportlab.lib.units import mm
    NPS_COL = {
        0: (210, 45, 45), 1: (220, 55, 55), 2: (225, 85, 55), 3: (240, 110, 40), 4: (255, 145, 0),
        5: (255, 190, 0), 6: (205, 215, 55), 7: (155, 200, 80), 8: (95, 175, 95), 9: (55, 155, 75), 10: (25, 135, 55),
    }
    sw_mm = w_mm / 11
    top_pt = (page_h_mm - y_mm) * mm
    h_pt = h_mm * mm
    c.setFont("Helvetica", 6)
    c.setFillColor(rl_colors.HexColor("#8C7F87"))
    for i in range(11):
        c.drawCentredString((x_mm + i * sw_mm + sw_mm / 2) * mm, top_pt + 2 * mm, str(i))
    for i in range(11):
        n = 0
        for k, v in options.items():
            if str(k).strip() == str(i):
                n += v
        col = NPS_COL[i]
        cx = (x_mm + i * sw_mm) * mm
        c.setFillColor(rl_colors.Color(col[0] / 255, col[1] / 255, col[2] / 255))
        c.rect(cx, top_pt - h_pt, sw_mm * mm - 0.4 * mm, h_pt, fill=1, stroke=0)
        if n > 0:
            c.setFont("Helvetica-Bold", 7)
            c.setFillColor(rl_colors.white)
            c.drawCentredString(cx + (sw_mm * mm) / 2, top_pt - h_pt / 2 - 2.5, str(n))
    c.setStrokeColor(rl_colors.HexColor("#C8C8C8"))
    c.setLineWidth(0.4)
    c.rect(x_mm * mm, top_pt - h_pt, w_mm * mm, h_pt, fill=0, stroke=1)
    return y_mm + h_mm + 3


# Los datos de BD_encuesta_curso.xlsx ya vienen en español (LIKERT_SCORE
# arriba usa esas mismas claves), así que solo hace falta el juego de
# claves en español — se deja también el equivalente en inglés por si un
# futuro socio carga la encuesta en ese idioma.
LIKERT_ORDER_PDF = [
    "Totalmente en desacuerdo", "Strongly Disagree", "En desacuerdo", "Disagree",
    "Ni de acuerdo ni en desacuerdo", "Neither agree nor disagree",
    "De acuerdo", "Agree", "Totalmente de acuerdo", "Strongly Agree",
]
LIKERT_COLOR_PDF = {
    "Totalmente en desacuerdo": (210, 45, 45), "Strongly Disagree": (210, 45, 45),
    "En desacuerdo": (255, 112, 67), "Disagree": (255, 112, 67),
    "Ni de acuerdo ni en desacuerdo": (255, 193, 7), "Neither agree nor disagree": (255, 193, 7),
    "De acuerdo": (139, 195, 74), "Agree": (139, 195, 74),
    "Totalmente de acuerdo": (46, 125, 50), "Strongly Agree": (46, 125, 50),
}
WORKLOAD_ORDER_PDF = ["Nada", "Nothing", "Poco", "Low", "Mucho", "A lot", "Demasiado", "Too much"]
WORKLOAD_COLOR_PDF = {
    "Nada": (46, 125, 50), "Nothing": (46, 125, 50), "Poco": (139, 195, 74), "Low": (139, 195, 74),
    "Mucho": (255, 112, 67), "A lot": (255, 112, 67), "Demasiado": (210, 45, 45), "Too much": (210, 45, 45),
}
OBJ_ORDER_PDF = ["Ninguno", "None", "Algunos", "Some", "Todos", "All"]
OBJ_COLOR_PDF = {
    "Ninguno": (210, 45, 45), "None": (210, 45, 45), "Algunos": (255, 193, 7), "Some": (255, 193, 7),
    "Todos": (46, 125, 50), "All": (46, 125, 50),
}


def _wrap_to_width(c, text: str, font_name: str, font_size: float, max_width_mm: float) -> list:
    from reportlab.lib.units import mm
    max_w_pt = max_width_mm * mm * 0.98
    words = str(text).split()
    lines, cur = [], ""
    for word in words:
        trial = f"{cur} {word}".strip()
        if c.stringWidth(trial, font_name, font_size) <= max_w_pt or not cur:
            cur = trial
        else:
            lines.append(cur)
            cur = word
    if cur:
        lines.append(cur)
    return lines or [""]


def _pdf_file_bytes(path: Optional[str]) -> Optional[bytes]:
    if not path:
        return None
    try:
        with open(path, "rb") as f:
            return f.read()
    except Exception:
        return None


def _build_week_survey_pdf(w: dict) -> bytes:
    """Réplica en reportlab del PDF de evaluación de EIV_report.py
    (_build_professor_pdf): misma banda con logo+foto, mismas tarjetas KPI,
    mismas barras apiladas con leyenda por aspecto, misma barra NPS 0-10, y
    mismos comentarios paginados — mismo diseño y misma lógica de dibujo.
    Dos diferencias reales frente a EIV: (1) International Weeks no maneja
    periodos de matrícula (GR/UG) — cada semana es un solo periodo, así que
    se usa directo la rama de una sola columna; y (2) todo el texto del PDF
    va en español (idioma original de la encuesta y del reporte de Weeks),
    no en inglés."""
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.lib import colors as rl_colors
    from reportlab.pdfgen import canvas as rl_canvas
    from reportlab.lib.utils import ImageReader

    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=A4)
    PW, PH = 210.0, 297.0
    mg = 14.0
    cW = PW - 2 * mg
    accent = rl_colors.HexColor(w["color"])
    ink = rl_colors.HexColor(INK)
    muted = rl_colors.HexColor("#8C7F87")
    lgray = rl_colors.HexColor("#F8F6F8")
    soft = rl_colors.HexColor(w["colorSoft"])

    def top_pt(y_mm):
        return (PH - y_mm) * mm

    sv = w["survey"]
    logo_bytes = _pdf_file_bytes(_partner_logo_path(w["key"]))
    photo_bytes = _pdf_file_bytes(_photo_path(w["professor"]))

    def draw_stats_header():
        band_h = 34
        c.setFillColor(accent)
        c.rect(0, top_pt(band_h), PW * mm, band_h * mm, fill=1, stroke=0)
        if logo_bytes:
            try:
                c.setFillColor(rl_colors.white)
                c.roundRect(mg * mm, top_pt(23), 42 * mm, 20 * mm, 3 * mm, fill=1, stroke=0)
                img = ImageReader(io.BytesIO(logo_bytes))
                c.drawImage(img, (mg + 2) * mm, top_pt(21), width=38 * mm, height=16 * mm,
                            preserveAspectRatio=True, mask="auto")
            except Exception:
                pass
        name_right_edge = PW - mg
        if photo_bytes:
            try:
                c.setFillColor(rl_colors.white)
                c.roundRect((PW - mg - 24) * mm, top_pt(31), 24 * mm, 28 * mm, 2.5 * mm, fill=1, stroke=0)
                img = ImageReader(io.BytesIO(photo_bytes))
                c.drawImage(img, (PW - mg - 23) * mm, top_pt(30), width=22 * mm, height=26 * mm,
                            preserveAspectRatio=True, mask="auto")
                name_right_edge = PW - mg - 28
            except Exception:
                pass
        c.setFillColor(rl_colors.white)
        c.setFont("Helvetica-Bold", 13)
        c.drawRightString(name_right_edge * mm, top_pt(11), w["professor"])
        c.setFont("Helvetica", 8.5)
        c.drawRightString(name_right_edge * mm, top_pt(17), f'{w["location"]}, {w["country"]}')
        c.setFont("Helvetica", 8)
        c.drawRightString(name_right_edge * mm, top_pt(22.5), w["dates"])
        c.setFont("Helvetica-Bold", 12)
        c.drawString(mg * mm, top_pt(29.5), "EVALUACIÓN DEL CURSO")

    def draw_continuation_header():
        c.setFillColor(accent)
        c.rect(0, top_pt(14), PW * mm, 14 * mm, fill=1, stroke=0)
        c.setFillColor(rl_colors.white)
        c.setFont("Helvetica-Bold", 10)
        c.drawRightString((PW - mg) * mm, top_pt(9), w["professor"])

    def draw_footer(page_num):
        c.setFillColor(lgray)
        c.rect(0, 0, PW * mm, 10 * mm, fill=1, stroke=0)
        c.setFont("Helvetica", 7)
        c.setFillColor(muted)
        c.drawString(12 * mm, 4 * mm, "International Weeks 2026 — Evaluación del Curso")
        c.drawRightString((PW - 12) * mm, 4 * mm, str(page_num))

    page_num = 1
    draw_stats_header()

    def _estimate_page1_height():
        h = 0.0
        course_lines_est = _wrap_to_width(c, w["course"] or "", "Helvetica-Bold", 17, cW)
        h += len(course_lines_est) * 7 + 3
        h += 9  # subtítulo ubicación/fechas
        h += 23 + 8  # tarjeta de matrícula
        h += 32  # tarjetas KPI
        return h

    header_bottom_y = 34.0
    footer_top_y = PH - 10.0
    available_h = footer_top_y - header_bottom_y
    total_h = _estimate_page1_height()
    y = header_bottom_y + max(2.0, (available_h - total_h) / 2)

    c.setFillColor(accent)
    c.setFont("Helvetica-Bold", 17)
    course_lines = _wrap_to_width(c, w["course"] or "", "Helvetica-Bold", 17, cW)
    for i, line in enumerate(course_lines):
        c.drawCentredString((mg + cW / 2) * mm, top_pt(y + i * 7), line)
    y += len(course_lines) * 7 + 3
    c.setFillColor(muted)
    c.setFont("Helvetica", 8.5)
    c.drawCentredString((mg + cW / 2) * mm, top_pt(y), f'{w["location"]}, {w["country"]}  ·  {w["dates"]}')
    y += 9

    # ---- Tarjeta de matrícula: un solo periodo (todas las semanas caen en
    # el mismo periodo), a diferencia de EIV que muestra GR/UG lado a lado.
    c.setFillColor(soft)
    c.roundRect((mg + cW / 2 - 40) * mm, top_pt(y + 19), 80 * mm, 19 * mm, 3 * mm, fill=1, stroke=0)
    c.setFillColor(accent)
    c.setFont("Helvetica-Bold", 13)
    c.drawCentredString((mg + cW / 2) * mm, top_pt(y + 8.5), f'{w["abroadCount"]} viajaron')
    c.setFillColor(ink)
    c.setFont("Helvetica", 6.5)
    c.drawCentredString((mg + cW / 2) * mm, top_pt(y + 14), w["name"])
    c.setFillColor(muted)
    c.setFont("Helvetica-Oblique", 6)
    c.drawCentredString((mg + cW / 2) * mm, top_pt(y + 18.3), f'{sv["respondents"]} respondieron la encuesta')
    y += 23 + 8

    # ---- KPIs de respuesta, en esta misma página 1 ----
    if y + 32 > PH - 16:
        draw_footer(page_num)
        c.showPage()
        page_num += 1
        draw_continuation_header()
        y = 24.0
    stats = [
        (f'{sv["respondents"]} / {w["abroadCount"]}' if w["abroadCount"] else str(sv["respondents"]),
         "Estudiantes que Respondieron"),
        (f'{sv["responseRate"]}%', "Tasa de Respuesta"),
        (f'{sv["avg"]:.2f} / 5.0' if sv["avg"] is not None else "N/A", "Satisfacción Promedio"),
        (f'{sv["nps"]:.1f} / 10' if sv["npsN"] else "N/A", "Recomendación Promedio"),
        (f'{sv["objectivesAllPct"]:.0f}%' if sv["objectivesN"] else "N/A", "Objetivos Cumplidos (Todos)"),
    ]
    bw = cW / 5
    for i, (val, label) in enumerate(stats):
        bx = mg + i * bw
        c.setFillColor(lgray)
        c.roundRect((bx + 1) * mm, top_pt(y + 22), (bw - 3) * mm, 22 * mm, 3 * mm, fill=1, stroke=0)
        c.setFillColor(accent)
        c.setFont("Helvetica-Bold", 12)
        c.drawCentredString((bx + bw / 2 - 1) * mm, top_pt(y + 11), val)
        c.setFillColor(muted)
        c.setFont("Helvetica", 6)
        c.drawCentredString((bx + bw / 2 - 1) * mm, top_pt(y + 18), label)
    y += 32
    draw_footer(page_num)

    # ---- Evaluación cuantitativa (barras apiladas por aspecto) ----
    c.showPage()
    page_num += 1
    draw_continuation_header()
    draw_footer(page_num)
    y = 24.0

    def check(needed):
        nonlocal y, page_num
        if y + needed > PH - 16:
            c.showPage()
            page_num += 1
            draw_continuation_header()
            draw_footer(page_num)
            y = 24.0

    def section_bar(title):
        nonlocal y
        check(12)
        c.setFillColor(accent)
        c.rect(mg * mm, top_pt(y + 7.5), cW * mm, 7.5 * mm, fill=1, stroke=0)
        c.setFillColor(rl_colors.white)
        c.setFont("Helvetica-Bold", 8.5)
        c.drawString((mg + 5) * mm, top_pt(y + 5.3), title)
        y += 11

    def draw_likert_q(q):
        nonlocal y
        text = q.get("text") or ""
        q_lines = _wrap_to_width(c, text, "Helvetica", 9, cW)
        n_opts = sum(1 for k in LIKERT_ORDER_PDF if q["options"].get(k))
        legend_rows = 2 if n_opts >= 4 else 1
        needed = len(q_lines) * 4.6 + 7 + 14 + (legend_rows - 1) * 5
        check(needed)
        c.setFillColor(ink)
        c.setFont("Helvetica", 9)
        for i, line in enumerate(q_lines):
            c.drawString(mg * mm, top_pt(y + i * 4.6), line)
        total = sum(q["options"].values())
        if total:
            c.setFillColor(lgray)
            c.roundRect((mg + cW - 22) * mm, top_pt(y - 2), 22 * mm, 6 * mm, 2 * mm, fill=1, stroke=0)
            c.setFont("Helvetica", 7)
            c.setFillColor(muted)
            c.drawCentredString((mg + cW - 11) * mm, top_pt(y + 0.5), f"n={total}")
        y += len(q_lines) * 4.6 + 3
        y = _pdf_stacked_bar(c, q["options"], LIKERT_ORDER_PDF, LIKERT_COLOR_PDF, mg, y, cW, 7, PH)
        y = _pdf_legend(c, q["options"], LIKERT_ORDER_PDF, LIKERT_COLOR_PDF, mg, y, cW, PH)
        c.setStrokeColor(rl_colors.HexColor("#E6E6E6"))
        c.setLineWidth(0.3)
        c.line(mg * mm, top_pt(y), (mg + cW) * mm, top_pt(y))
        y += 5

    by_aspect: Dict[str, list] = {}
    nps_q = obj_q = wl_q = None
    for q in sv["questions"]:
        qtype = q.get("type")
        if qtype == "acuerdo":
            by_aspect.setdefault(q.get("aspect") or "General", []).append(q)
        elif qtype == "escala010" and nps_q is None:
            nps_q = q
        elif qtype == "cantidad" and obj_q is None:
            obj_q = q
        elif qtype == "carga" and wl_q is None:
            wl_q = q

    section_bar("EVALUACIÓN CUANTITATIVA")
    for asp, qs in by_aspect.items():
        check(16)
        c.setFillColor(soft)
        c.rect(mg * mm, top_pt(y + 6.5), cW * mm, 6.5 * mm, fill=1, stroke=0)
        c.setFillColor(ink)
        c.setFont("Helvetica-Bold", 7.5)
        c.drawString((mg + 4) * mm, top_pt(y + 4.6), str(asp).upper())
        y += 9.5
        for q in qs:
            draw_likert_q(q)
        y += 2

    if nps_q:
        section_bar("RECOMENDACIÓN (0-10)")
        y = _pdf_nps_bar(c, nps_q["options"], mg, y + 5, cW, 10, PH) + 8
    if obj_q:
        section_bar("OBJETIVOS DEL CURSO")
        y = _pdf_stacked_bar(c, obj_q["options"], OBJ_ORDER_PDF, OBJ_COLOR_PDF, mg, y, cW, 7, PH)
        y = _pdf_legend(c, obj_q["options"], OBJ_ORDER_PDF, OBJ_COLOR_PDF, mg, y, cW, PH) + 8
    if wl_q:
        section_bar("CARGA ACADÉMICA")
        y = _pdf_stacked_bar(c, wl_q["options"], WORKLOAD_ORDER_PDF, WORKLOAD_COLOR_PDF, mg, y, cW, 7, PH)
        y = _pdf_legend(c, wl_q["options"], WORKLOAD_ORDER_PDF, WORKLOAD_COLOR_PDF, mg, y, cW, PH)
    draw_footer(page_num)

    # ---- Comentarios de los estudiantes ----
    def ensure_space(needed):
        nonlocal y, page_num
        if y + needed > 283:
            draw_footer(page_num)
            c.showPage()
            page_num += 1
            draw_continuation_header()
            y = 24.0

    ensure_space(10)
    c.setFillColor(ink)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(mg * mm, top_pt(y), "COMENTARIOS DE LOS ESTUDIANTES")
    y += 7

    comment_groups = sv["commentGroups"]
    with_items = [g for g in comment_groups if g["items"]]
    if not with_items:
        c.setFont("Helvetica", 8)
        c.setFillColor(rl_colors.HexColor("#50484E"))
        c.drawString(mg * mm, top_pt(y), "No se recogieron comentarios para esta semana.")
        draw_footer(page_num)
    else:
        for g in comment_groups:
            group_needed = 7.0
            if not g["items"]:
                group_needed += 6
            else:
                for item in g["items"]:
                    lines_est = _wrap_to_width(c, f'1. "{item}"', "Helvetica", 8, cW)
                    group_needed += len(lines_est) * 4 + 2
                group_needed += 4
            if y + min(group_needed, 260) > 283 and y > 24.0:
                draw_footer(page_num)
                c.showPage()
                page_num += 1
                draw_continuation_header()
                y = 24.0
            c.setFont("Helvetica-Bold", 8.5)
            c.setFillColor(accent)
            c.drawString(mg * mm, top_pt(y), f"{g['label'].upper()}")
            y += 5.5
            if not g["items"]:
                c.setFont("Helvetica", 8)
                c.setFillColor(muted)
                c.drawString(mg * mm, top_pt(y), "No hay comentarios para esta pregunta.")
                y += 6
                continue
            for i, item in enumerate(g["items"]):
                lines = _wrap_to_width(c, f'{i + 1}. "{item}"', "Helvetica", 8, cW)
                needed = len(lines) * 4 + 2
                ensure_space(needed)
                c.setFont("Helvetica", 8)
                c.setFillColor(rl_colors.HexColor("#50484E"))
                for j, line in enumerate(lines):
                    c.drawString(mg * mm, top_pt(y + j * 4), line)
                y += needed
            y += 4
        draw_footer(page_num)

    c.save()
    buf.seek(0)
    return buf.getvalue()


# ── 12) NAVEGACIÓN ────────────────────────────────────────────────────────
# Mismo esqueleto que EIV_report.py: portada oculta el sidebar, nav superior
# con page_link + popover "More...", y flechas laterales fijas para pasar
# de sección. Solo las semanas llevan icono en el menú (la bandera del país
# del socio); Data Center / Overview / Financial no llevan emoji.
_weeks_for_nav, _ = build_weeks()

week_page_by_key: Dict[str, st.Page] = {
    w["key"]: st.Page(functools.partial(page_week, key=w["key"]), title=w["key"],
                       url_path=w["key"].lower(), icon=FLAGS.get(w["country"], ""))
    for w in _weeks_for_nav
}

pages = (
    [st.Page(page_cover, title="Data Center", icon=":material/dashboard:", url_path="cover", default=True),
     st.Page(page_overview, title="Overview", icon="🌍", url_path="overview")]
    + list(week_page_by_key.values())
    + [st.Page(page_financial, title="Financial Detail", icon="💰", url_path="financial")]
)
pg = st.navigation(pages, position="hidden")
IS_COVER = pg is pages[0]
nav_pages = pages[1:]
_NAV_VISIBLE = 6

if IS_COVER:
    st.markdown(
        "<style>section[data-testid='stSidebar']{display:none !important;}"
        "div[data-testid='stSidebarCollapsedControl']{display:none !important;}</style>",
        unsafe_allow_html=True,
    )

if not IS_COVER:
    with st.sidebar:
        try:
            _show_weeks_banner(use_container_width=True)
        except Exception:
            pass
        st.markdown(
            f'<div style="color:{INK};font-size:22px;font-weight:800;line-height:1.1;">'
            'International Weeks</div>',
            unsafe_allow_html=True,
        )
        st.caption("Facultad de Administración · Universidad de los Andes")
        with st.container(key="go_to_datacenter_btn"):
            st.page_link(pages[0], label="Go to Data Center", icon=":material/home:")
        st.markdown("---")

if not IS_COVER:
    with st.container(key="nav_toggle"):
        visible_pages = nav_pages[:_NAV_VISIBLE]
        overflow_pages = nav_pages[_NAV_VISIBLE:]
        n_cols = len(visible_pages) + (1 if overflow_pages else 0)
        nav_cols = st.columns(n_cols)
        for col, page_obj in zip(nav_cols, visible_pages):
            with col:
                st.page_link(page_obj)
        if overflow_pages:
            with nav_cols[-1]:
                with st.popover("More...", use_container_width=False):
                    for page_obj in overflow_pages:
                        st.page_link(page_obj)

    # ---- Flechas laterales para pasar de sección ----
    _idx = nav_pages.index(pg)
    _prev_pg = nav_pages[_idx - 1]
    _next_pg = nav_pages[(_idx + 1) % len(nav_pages)]
    with st.container(key="side_arrows"):
        arrow_l, arrow_r = st.columns(2)
        with arrow_l:
            st.page_link(_prev_pg, label="‹")
        with arrow_r:
            st.page_link(_next_pg, label="›")

pg.run()

st.markdown(
    f'<div style="text-align:center;padding:40px 0 10px;font-size:11.5px;color:{MUTED};'
    'font-family:monospace;">International Weeks · Facultad de Administración · '
    'Universidad de los Andes</div>',
    unsafe_allow_html=True,
)
