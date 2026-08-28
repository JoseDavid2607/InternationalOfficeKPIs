# ===========================================================================
#  UASM · International Summer School (EIV) Analytics · App multipágina
#  Traducción del reporte HTML "EIV_REPORT" a Streamlit — mismas secciones
#  (Data Center, Overview, Summary, Dashboard, Financial, Visiting Faculty,
#  School Enrollment/Gap Analysis, Conclusions), mismos cálculos, misma
#  identidad visual (paleta rosa/tinta del HTML original).
#
#  Simplificaciones honestas frente al HTML (documentadas donde aplican):
#  - El PDF por profesor es un resumen de reportlab, no una réplica
#    pixel-a-pixel del PDF con jsPDF (barras degradadas, radar dibujado a
#    mano, etc.) — mismo contenido numérico y comentarios.
#  - El árbol de navegación "Block > Course" del sidebar original se
#    reemplaza por selectboxes en el sidebar de Streamlit.
#  - SCHOOL_ENROLLMENT / MOBILITY_DATA / WEEKS_DATA eran, ya en el HTML
#    original, constantes congeladas (no recalculadas en vivo desde
#    BD_matriculados ni desde BD_movilidad, que ni siquiera vive en este
#    reporte) — se reproducen aquí como las mismas constantes.
# ===========================================================================
from __future__ import annotations

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import io
import time
import re
import unicodedata
import base64
from typing import Optional, Dict, List, Tuple

try:
    from google.oauth2.service_account import Credentials
    _GSPREAD_OK = True
    _GSPREAD_IMPORT_ERR = None
except ImportError as _e:
    _GSPREAD_OK = False
    _GSPREAD_IMPORT_ERR = str(_e)

import requests

# ── 1) CONFIGURACIÓN GLOBAL ────────────────────────────────────────────────
st.set_page_config(
    page_title="EIV Analytics — International Summer School",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="expanded",
)

PINK = "#E61166"; INK = "#241420"; INK_SOFT = "#6E5C68"; PAPER = "#FAF8FA"
LINE = "#E9E2E7"; MUTED = "#8C7F87"
BLUE = "#12BDFB"; PURPLE = "#B540D7"; GREEN = "#BFD466"; ORANGE = "#F2994A"
ELEGANT_PURPLE = "#5B3A8E"

st.markdown(
    "<style>"
    f".suite-header{{display:flex !important;flex-direction:column !important;align-items:center !important;"
    "padding:16px 24px 12px !important;"
    f"background:linear-gradient(135deg,#7A0A3C 0%,#B00D50 55%,{PINK} 100%) !important;"
    "border-radius:12px !important;box-shadow:0 2px 8px rgba(230,17,102,.18) !important;margin-bottom:14px !important;}}"
    f"div.suite-header div.sh-super{{font-size:11px !important;font-weight:700 !important;letter-spacing:2px !important;"
    "color:#FBD6E4 !important;text-transform:uppercase !important;margin-bottom:2px !important;}}"
    "div.suite-header div.sh-title{font-size:26px !important;font-weight:800 !important;color:#fff !important;"
    "text-align:center !important;line-height:1.2 !important;}"
    "div.suite-header div.sh-sub{font-size:13px !important;color:rgba(255,255,255,.80) !important;"
    "margin-top:4px !important;text-align:center !important;font-weight:400 !important;}"
    f".kv{{font-size:24px;font-weight:700;line-height:1.1;font-family:monospace;color:{INK};}}"
    f".kv.accent{{color:{PINK};}}"
    f".kl{{font-size:11px;font-weight:600;color:{MUTED};"
    "text-transform:uppercase;letter-spacing:.5px;margin-top:3px;}}"
    f"section[data-testid='stSidebar']{{background:{PAPER} !important;}}"
    "div[data-testid='stButton'] button{background:#FFFFFF !important;"
    f"border:1px solid {LINE} !important;border-radius:10px !important;"
    "color:#374151 !important;font-size:14px !important;"
    "font-weight:600 !important;height:48px !important;"
    "box-shadow:0 1px 3px rgba(0,0,0,.04) !important;}"
    f"div[data-testid='stButton'] button:hover{{background:{PAPER} !important;border-color:{PINK} !important;}}"
    "div.stDownloadButton>button{background:transparent !important;"
    "border:none !important;box-shadow:none !important;"
    f"color:{PINK} !important;font-size:13px !important;"
    "padding:0 !important;text-decoration:underline !important;}"
    f"thead th{{background:#FCE9F1 !important;color:{INK} !important;"
    "font-weight:700 !important;}}"
    ".pending-card{background:#FAFAFA;border:1px dashed #DCD3D8;border-radius:12px;"
    "padding:22px 24px;margin-top:14px;color:#6B7280;font-size:13.5px;}"
    ".pending-card .tag{display:inline-block;font-family:monospace;font-size:10px;"
    "letter-spacing:.06em;text-transform:uppercase;color:#8C7F87;"
    "background:#F1EAEE;padding:3px 9px;border-radius:5px;margin-bottom:8px;}"
    ".prof-quote{background:#FAF8FA;border-left:3px solid var(--pink,#E61166);"
    "border-radius:6px;padding:8px 12px;margin-bottom:10px;font-size:12.5px;"
    "font-style:italic;color:#3B2C34;}"
    ".prof-quote strong{font-style:normal;font-family:monospace;font-size:10.5px;"
    "color:#8C7F87;display:block;margin-bottom:3px;text-transform:uppercase;}"
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
    f"color:{PINK} !important;font-size:13px !important;font-weight:400 !important;"
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
    f"font-size:0 !important;font-weight:400;color:{PINK} !important;opacity:.55;text-decoration:none;"
    "transition:opacity .15s ease;}"
    ".st-key-side_arrows a p{font-size:26px !important;}"
    ".st-key-side_arrows a span:first-child{display:none !important;}"
    ".st-key-side_arrows a:hover{opacity:1;}"
    ".st-key-cover_enter_btn div[data-testid='stButton'] button{"
    f"background:{PINK} !important;border:none !important;color:#fff !important;"
    "font-size:16px !important;font-weight:700 !important;height:52px !important;"
    "box-shadow:0 4px 14px rgba(230,17,102,.28) !important;}"
    ".st-key-cover_enter_btn div[data-testid='stButton'] button:hover{background:#B00D50 !important;}"
    ".st-key-cover_banner{display:flex !important;justify-content:center !important;width:100% !important;}"
    ".st-key-cover_banner div[data-testid='stImage']{margin:0 auto !important;width:auto !important;"
    "display:flex !important;justify-content:center !important;}"
    ".st-key-cover_banner img{margin:0 auto !important;display:block !important;}"
    ".block-container{padding-top:3.2rem !important;}"
    ".st-key-go_to_datacenter_btn a{"
    f"display:flex !important;align-items:center;justify-content:center;gap:6px;"
    f"background:{PINK} !important;border:none !important;border-radius:10px !important;"
    "color:#fff !important;font-weight:700 !important;height:44px !important;text-decoration:none !important;}"
    ".st-key-go_to_datacenter_btn a span{color:#fff !important;}"
    ".st-key-go_to_datacenter_btn a:hover{background:#B00D50 !important;}"
    "div[class*='st-key-pdf_btn_'] div[data-testid='stDownloadButton'] button, "
    "div[class*='st-key-bulk_pdf_'] div[data-testid='stButton'] button, "
    "div[class*='st-key-bulk_pdf_'] div[data-testid='stDownloadButton'] button{"
    f"background:{ELEGANT_PURPLE} !important;color:#fff !important;border:none !important;"
    "text-decoration:none !important;font-weight:700 !important;height:48px !important;"
    "box-shadow:0 3px 10px rgba(91,58,142,.28) !important;}"
    "div[class*='st-key-pdf_btn_'] div[data-testid='stDownloadButton'] button:hover, "
    "div[class*='st-key-bulk_pdf_'] div[data-testid='stButton'] button:hover, "
    "div[class*='st-key-bulk_pdf_'] div[data-testid='stDownloadButton'] button:hover{background:#4A2F73 !important;}"
    "</style>",
    unsafe_allow_html=True,
)


# ── 2) HELPERS COMPARTIDOS ──────────────────────────────────────────────
def _render_header(title: str, subtitle: str = ""):
    # Estilos inline (no clases CSS): máxima especificidad posible, no puede
    # perder contra ninguna hoja de estilos de Streamlit -- si el título
    # sigue sin verse distinto con esto, el problema ya no es de CSS sino
    # de que el despliegue no está sirviendo este archivo actualizado
    # (revisar: ¿se reemplazó el archivo en GitHub? ¿se reinició la app en
    # Streamlit Cloud / "Manage app" → "Reboot"? ¿caché del navegador?).
    banner_style = (
        "display:flex;flex-direction:column;align-items:center;padding:16px 24px 12px;"
        f"background:linear-gradient(135deg,#7A0A3C 0%,#B00D50 55%,{PINK} 100%);"
        "border-radius:12px;box-shadow:0 2px 8px rgba(230,17,102,.18);margin-bottom:14px;"
    )
    super_style = (
        "font-size:11px;font-weight:700;letter-spacing:2px;color:#EAF7DE;"
        "text-transform:uppercase;margin-bottom:2px;"
    )
    title_style = "font-size:26px;font-weight:800;color:#ffffff;text-align:center;line-height:1.2;"
    sub_style = "font-size:13px;color:#EAF7DE;margin-top:4px;text-align:center;"

    sub = f'<div style="{sub_style}">{subtitle}</div>' if subtitle else ""
    st.markdown(
        f'<div style="{banner_style}">'
        f'<div style="{super_style}">EIV · International Summer School</div>'
        f'<div style="{title_style}">{title}</div>{sub}</div>',
        unsafe_allow_html=True,
    )


def _kpi(label: str, value, cls: str = ""):
    st.markdown(f'<div class="kv {cls}">{value}</div><div class="kl">{label}</div>', unsafe_allow_html=True)


def _pending_card(label: str, note: str = ""):
    note_html = f"<br>{note}" if note else ""
    st.markdown(
        f'<div class="pending-card"><span class="tag">Coming soon</span>'
        f'<p style="margin-top:8px;">{label}{note_html}</p></div>',
        unsafe_allow_html=True,
    )


def _quote(label: str, text):
    if text is None or (isinstance(text, float) and pd.isna(text)) or str(text).strip() == "":
        return ""
    return f'<div class="prof-quote"><strong>{label}</strong>“{text}”</div>'


def _fmt_cop(n) -> str:
    if n is None or (isinstance(n, float) and pd.isna(n)):
        return "—"
    return f"${round(n):,.0f}"


def _xlsx_bytes(df: pd.DataFrame, sheet_name: str = "Data") -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf) as w:
        df.to_excel(w, index=False, sheet_name=sheet_name[:31])
    buf.seek(0)
    return buf.getvalue()


import os as _os

_LOGO_CANDIDATES = [
    "PICS/LOGO/2026/logo2026.png", "PICS/LOGO/2026/Logo2026.png",
    "pics/LOGO/2026/logo2026.png", "PICS/logo/2026/logo2026.png",
    "PICS/LOGO/2026/logo2026.PNG", "./PICS/LOGO/2026/logo2026.png",
]
_BANNER_CANDIDATES = [
    "PICS/LOGO/2026/Banner_EIV.png", "PICS/LOGO/2026/banner_eiv.png",
    "pics/LOGO/2026/Banner_EIV.png", "PICS/LOGO/2026/Banner_EIV.PNG",
    "./PICS/LOGO/2026/Banner_EIV.png",
]


def _show_eiv_logo(width: int):
    """Prueba varias rutas/variantes de mayúsculas del logo — el filesystem
    de despliegue es sensible a mayúsculas y no sabemos con certeza cuál usa."""
    base = _os.path.dirname(_os.path.abspath(__file__))
    for rel in _LOGO_CANDIDATES:
        full = _os.path.join(base, rel)
        if _os.path.exists(full):
            st.image(full, width=width)
            return
    # último intento: buscar cualquier .png dentro de PICS/LOGO (nombre exacto desconocido)
    logo_dir = _os.path.join(base, "PICS", "LOGO", "2026")
    if _os.path.isdir(logo_dir):
        for fname in sorted(_os.listdir(logo_dir)):
            if fname.lower().endswith(".png") and "logo" in fname.lower():
                st.image(_os.path.join(logo_dir, fname), width=width)
                return


def _show_eiv_banner(width: Optional[int] = None, use_container_width: bool = False):
    """El banner de portada (Banner_EIV.png), centrado con HTML+base64 en vez
    de st.image -- st.image no se dejaba centrar de forma confiable con CSS
    (el wrapper que genera Streamlit no respondía a flex/margin:auto)."""
    base = _os.path.dirname(_os.path.abspath(__file__))

    def _find_banner_path():
        for rel in _BANNER_CANDIDATES:
            full = _os.path.join(base, rel)
            if _os.path.exists(full):
                return full
        logo_dir = _os.path.join(base, "PICS", "LOGO", "2026")
        if _os.path.isdir(logo_dir):
            for fname in sorted(_os.listdir(logo_dir)):
                if fname.lower().endswith(".png") and "banner" in fname.lower():
                    return _os.path.join(logo_dir, fname)
        return None

    path = _find_banner_path()
    if not path:
        return
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
    width_css = "100%" if use_container_width else f"{width}px"
    st.markdown(
        f'<div style="width:100%;text-align:center;">'
        f'<img src="data:image/png;base64,{b64}" style="width:{width_css};max-width:100%;">'
        f'</div>',
        unsafe_allow_html=True,
    )


def color_scale(pct: float) -> str:
    """Rojo -> ámbar -> verde, 0-100 (idéntico a EIV.Colors.scale)."""
    stops = [(209, 70, 69), (232, 178, 70), (76, 138, 63)]
    p = max(0.0, min(100.0, pct))
    seg = 0 if p < 50 else 1
    t = (p / 50) if p < 50 else ((p - 50) / 50)
    a, b = stops[seg], stops[seg + 1]
    rgb = tuple(round(a[i] + (b[i] - a[i]) * t) for i in range(3))
    return f"rgb({rgb[0]},{rgb[1]},{rgb[2]})"


def occ_color(pct: float) -> str:
    """Rojo desde 20% (o menos) hasta verde en 100%, para la barra de ocupación."""
    return color_scale(max(0.0, (pct - 20) / 80 * 100))


def _style_courses_table(disp: pd.DataFrame):
    """Barra de ocupación (rojo→verde) con ⚠️ en las 3 con menor ocupación,
    negrilla en Modality solo si es Online, fondo gris tenue para los
    bloques 1 y 3, y borde grueso entre bloques — todo vía pandas Styler."""
    low_threshold = sorted(disp["Occupancy"])[2] if len(disp) >= 3 else disp["Occupancy"].max()

    def _style(d: pd.DataFrame) -> pd.DataFrame:
        s = pd.DataFrame("", index=d.index, columns=d.columns)
        for i in d.index:
            pct = d.loc[i, "Occupancy"]
            color = occ_color(pct)
            s.loc[i, "Occupancy"] = f"background:linear-gradient(90deg,{color} {pct:.0f}%,#F3F1F2 {pct:.0f}%);"
            if d.loc[i, "Modality"] == "Online":
                s.loc[i, "Modality"] += "font-weight:700;"
            if str(d.loc[i, "Block"]) in ("1", "3"):
                for c in d.columns:
                    if c != "Occupancy":
                        s.loc[i, c] += "background-color:#F7F5F6;"
            is_last_in_block = (i == d.index[-1]) or (d.loc[i, "Block"] != d.loc[d.index[d.index.get_loc(i) + 1], "Block"])
            if is_last_in_block and i != d.index[-1]:
                for c in d.columns:
                    s.loc[i, c] += f"border-bottom:3px solid {INK};"
        return s

    fmt = {"Occupancy": lambda v: f"{v:.1f}%" + (" ⚠️" if v <= low_threshold else "")}
    return disp.style.apply(_style, axis=None).format(fmt)


# ── 3) FILE IDs (Google Drive) — carpeta Reportes/EIV ─────────────────────
CURSOS_FILE_ID = "1qZin81h9oQ4SfxadcNpdjzSO6w3ADKZr"       # BD_cursos.xlsx
LISTAS_FILE_ID = "1BpREXyP3KHIARxik6ah_av9R9wV_ElL3"       # BD_listas.xlsx (listas, programas)
EVALUACION_FILE_ID = "1M-0ucmSmgk2xAbG1Ll2Pvkz5Ir3oeAaH"   # BD_evaluacion_curso.xlsx (Frecuencias, Comentarios)
SATISFACCION_FILE_ID = "1moO_-D_btpTCXWvnKI7p0RotRvdrY5fM" # BD_satisfaccionEIV.xlsx
GASTOS_FILE_ID = "1bCKmesZ_50PhoXh6XTgRlBbFwJWrrVJt"       # BD_gastos.xlsx
ELECTIVAS_FILE_ID = "1wU3rjqr4ECGRsSJf8CYcF7zWROTql5dA"    # BD_electivas.xlsx (BD_ANÁLISIS)

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
            st.error(f"📦 Falta instalar `google-auth` en el entorno (import falló: `{_GSPREAD_IMPORT_ERR}`).")
        elif "gcp_service_account" not in st.secrets:
            st.error("🔑 No encuentro `st.secrets['gcp_service_account']`. Revisa Settings → Secrets.")
        else:
            st.error("🔑 Las credenciales de `gcp_service_account` no se pudieron usar para autenticar.")
        st.stop()

    url = f"https://www.googleapis.com/drive/v3/files/{file_id}?alt=media"
    headers = {"Authorization": f"Bearer {token}"}
    resp, last_err = None, None
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
        f"🌐 No se pudo descargar el archivo de Drive ({file_id}) tras varios intentos: {last_err}\n\n"
        "Verifica que el archivo esté compartido con el correo de la service account, con permiso de Editor."
    )
    st.stop()


# ── 4) CONSTANTES HISTÓRICAS (idénticas a EIV.HIST / SCHOOL_ENROLLMENT /
#      MOBILITY_DATA / WEEKS_DATA del HTML — no derivables de los archivos
#      de este reporte; el HTML original ya las trae congeladas) ──────────
HIST_PARTICIPANTS = [
    {"year": 2012, "val": 345}, {"year": 2013, "val": 363}, {"year": 2014, "val": 352},
    {"year": 2015, "val": 350}, {"year": 2016, "val": 362}, {"year": 2017, "val": 411},
    {"year": 2018, "val": 548}, {"year": 2019, "val": 528}, {"year": 2020, "val": 507},
    {"year": 2021, "val": 559}, {"year": 2022, "val": 527}, {"year": 2023, "val": 525},
    {"year": 2024, "val": 589, "seats": 660}, {"year": 2025, "val": 554, "seats": 630},
]
HIST_CATEGORY = [
    {"year": 2023, "EPOS": 27, "PRE": 53, "OTHER": 15, "EXTERNAL": 5},
    {"year": 2024, "EPOS": 42, "PRE": 44, "OTHER": 13, "EXTERNAL": 1},
    {"year": 2025, "EPOS": 23.8, "PRE": 59.0, "OTHER": 17.0, "EXTERNAL": 0.2},
]

SCHOOL_ENROLLMENT = {
    "currentPeriod": "2026", "currentPre": 2019, "currentPos": 1441,
    "preIntlExperience": {"movilidad": 533, "convocatoria": 99, "escuelaVerano": 438},
    "newStudents": {"pre": 277, "pos": 320},
    "historicalEnrollment": [
        {"year": 2022, "PRE": 1567, "POS": 1419}, {"year": 2023, "PRE": 1649, "POS": 1332},
        {"year": 2024, "PRE": 1811, "POS": 1369}, {"year": 2025, "PRE": 1923, "POS": 1532},
        {"year": 2026, "PRE": 2019, "POS": 1441},
    ],
    "posProgramDistribution": [
        {"program": "Maestría en Finanzas", "count": 217},
        {"program": "Especialización en Administración Financiera", "count": 172},
        {"program": "Maestría en Administración", "count": 167},
        {"program": "Maestría en Administración (MBA)", "count": 158},
        {"program": "Maestría en Regeneración y Desarrollo Sostenible", "count": 146},
        {"program": "Maestría en Gerencia para la Práctica del Desarrollo", "count": 80},
        {"program": "Maestría en Analítica y Gestión Financiera", "count": 71},
        {"program": "Maestría en Mercadeo", "count": 68},
        {"program": "Maestría en Administración Ejecutiva (EMBA)", "count": 67},
        {"program": "Maestría en Gerencia Estratégica", "count": 59},
        {"program": "Maestría en Gestión de la Cadena de Suministro", "count": 58},
        {"program": "Maestría en Gerencia Ambiental", "count": 55},
        {"program": "Especialización en Negociación", "count": 46},
        {"program": "Especialización en Inteligencia de Mercados", "count": 34},
        {"program": "Doctorado en Administración", "count": 31},
        {"program": "Maestría en Gerencia Internacional", "count": 10},
        {"program": "Maestría Internacional en Finanzas", "count": 1},
        {"program": "Maestría en Investigación en Administración", "count": 1},
    ],
}
MOBILITY_DATA = {
    "preUniqueLast5": 533, "posUniqueLast5": 34,
    "byYear": [
        {"year": 2021, "types": {"Intercambio Internacional": 49, "Doble Titulación": 5, "Pasantía de Investigación": 8}},
        {"year": 2022, "types": {"Intercambio Internacional": 96, "Doble Titulación": 12, "Pasantía de Investigación": 3}},
        {"year": 2023, "types": {"Intercambio Internacional": 143, "Doble Titulación": 12, "Pasantía de Investigación": 9}},
        {"year": 2024, "types": {"Intercambio Internacional": 94, "Doble Titulación": 8, "Pasantía de Investigación": 3}},
        {"year": 2025, "types": {"Intercambio Internacional": 125, "Doble Titulación": 8, "Pasantía de Investigación": 6}},
    ],
}
WEEKS_DATA = {"byWeek": [
    {"week": "KLU", "MBA": 12, "EPOS": 18}, {"week": "NOVA", "MBA": 32, "EPOS": 8},
    {"week": "FGV", "MBA": 22, "EPOS": 3}, {"week": "BABSON", "MBA": 26, "EPOS": 8},
]}

COUNTRY_GEO = {
    "Brazil": (-14.24, -51.93), "China": (35.86, 104.20), "Denmark": (56.26, 9.50),
    "France": (46.23, 2.21), "Poland": (51.92, 19.15), "UK": (55.38, -3.44),
    "Spain": (40.46, -3.75), "The Netherlands": (52.13, 5.29), "United States": (37.09, -95.71),
}
COUNTRY_COLORS = {
    "Brazil": BLUE, "China": PINK, "Denmark": PURPLE, "France": GREEN, "Poland": ORANGE,
    "UK": "#3FA34D", "Spain": "#F2B84B", "The Netherlands": "#56B3B4", "United States": "#E0568B",
}
BOGOTA = (4.71, -74.07)

LIKERT_WEIGHTS = {
    # Español
    "Totalmente en desacuerdo": 1, "En desacuerdo": 2,
    "Ni de acuerdo ni en desacuerdo": 3, "De acuerdo": 4, "Totalmente de acuerdo": 5,
    # English (BD_evaluacion_curso.xlsx trae ambos idiomas por fila; el
    # profesor puede tener "answer" en inglés según su Idioma en BD_cursos)
    "Strongly Disagree": 1, "Disagree": 2,
    "Neither agree nor disagree": 3, "Agree": 4, "Strongly Agree": 5,
}
WORKLOAD_WEIGHTS = {
    "Nada": 1, "Poco": 2, "Mucho": 3, "Demasiado": 4,
    "Nothing": 1, "Low": 2, "A lot": 3, "Too much": 4,
}
OBJ_ALL_VALUES = {"todos", "all"}
NPS_QID, WORKLOAD_QID, OBJ_QID = "E008V01", "P458V20", "P1003"
POSITIVE_QIDS = {"P216": "Course", "P216V01": "Faculty"}
IMPROVE_QIDS = {"P217": "Course", "P217V01": "Faculty"}
NAME_ALIASES = {"EDGAR VIRGUEZ": "EDGAR ANDRES VIRGUEZ RODRIGUEZ"}
PHOTO_FILENAME_OVERRIDES = {
    "ARASH AZADEGAN": "Arash Azadegan", "GLYN ATWAL": "Glyn Atwal",
    "SERVAAS VAN BILSEN": "Servaas Van Bilsen", "CHRISTIAM OLIVEIRA": "Cristiam Oliveira",
    "FERNANDO GOMEZ BAQUERO": "Fernando Gomez", "MIRKO ANTINO": "Mirko Antino",
    "EDGAR ANDRES VIRGUEZ RODRIGUEZ": "Edgar Virguez", "JAMIE SMITH": "Jamie Smith",
    "MARK SAMUEL": "Mark Samuel", "ANETA HRYCKIEWICZ GONTARCZYK": "Aneta Hryckiewicz",
    "BIN SHEN": "Bin Shen", "ALFRED VERNIS": "Alfred Vernis", "INTEKHAB ALAM": "Intekhab Alam",
    "CHRISTINA LUBINSKI": "Christina Lubinski",
}


# ── 5) CARGA DE DATOS ────────────────────────────────────────────────────
@st.cache_data(ttl=300)
def load_cursos() -> pd.DataFrame:
    raw = io.BytesIO(_download_drive_file_bytes(CURSOS_FILE_ID))
    df = pd.read_excel(raw, sheet_name="cursos")
    df.columns = df.columns.str.strip()
    return df


@st.cache_data(ttl=300)
def professor_language_map() -> Dict[str, str]:
    """Profesor -> 'es'/'en', desde la columna 'Idioma' de BD_cursos.xlsx
    (si existe). Si la columna todavía no existe en el archivo real, o el
    profesor no tiene un valor reconocible, se asume inglés por defecto —
    la mayoría de profesores visitantes no leen español y el resto del
    reporte ya está en inglés."""
    df = load_cursos()
    idioma_col = next((c for c in df.columns if c.strip().casefold() == "idioma"), None)
    prof_col = next((c for c in df.columns if c.strip().casefold() == "profesor"), None)
    out: Dict[str, str] = {}
    if not idioma_col or not prof_col:
        return out
    for _, r in df.iterrows():
        prof = r.get(prof_col)
        if pd.isna(prof):
            continue
        val = str(r.get(idioma_col) or "").strip().casefold()
        out[_norm_name(prof)] = "es" if val.startswith(("esp", "spa")) else "en"
    return out


def _prof_lang(prof_name: str) -> str:
    return professor_language_map().get(str(prof_name).strip(), "en")


@st.cache_data(ttl=300)
def load_listas() -> Tuple[pd.DataFrame, pd.DataFrame]:
    raw1 = io.BytesIO(_download_drive_file_bytes(LISTAS_FILE_ID))
    listas = pd.read_excel(raw1, sheet_name="listas")
    listas.columns = listas.columns.str.strip()
    raw2 = io.BytesIO(_download_drive_file_bytes(LISTAS_FILE_ID))
    programas = pd.read_excel(raw2, sheet_name="programas")
    programas.columns = programas.columns.str.strip()
    return listas, programas


@st.cache_data(ttl=300)
def load_evaluacion() -> Tuple[pd.DataFrame, pd.DataFrame]:
    raw1 = io.BytesIO(_download_drive_file_bytes(EVALUACION_FILE_ID))
    frec = pd.read_excel(raw1, sheet_name="frecuencias")
    frec.columns = frec.columns.str.strip()
    raw2 = io.BytesIO(_download_drive_file_bytes(EVALUACION_FILE_ID))
    com = pd.read_excel(raw2, sheet_name="comentarios")
    com.columns = com.columns.str.strip()
    return frec, com


@st.cache_data(ttl=300)
def load_satisfaccion() -> pd.DataFrame:
    raw = io.BytesIO(_download_drive_file_bytes(SATISFACCION_FILE_ID))
    df = pd.read_excel(raw, sheet_name="satisfaccionEIV")
    df.columns = df.columns.str.strip()
    return df


@st.cache_data(ttl=300)
def load_gastos() -> pd.DataFrame:
    raw = io.BytesIO(_download_drive_file_bytes(GASTOS_FILE_ID))
    df = pd.read_excel(raw, sheet_name="gastos")
    df.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in df.columns]
    return df


@st.cache_data(ttl=300)
def load_electivas() -> pd.DataFrame:
    raw = io.BytesIO(_download_drive_file_bytes(ELECTIVAS_FILE_ID))
    df = pd.read_excel(raw, sheet_name="BD_electivas")
    df.columns = df.columns.str.strip()
    return df


# ── 6) EIV.Compute equivalente ───────────────────────────────────────────
@st.cache_data(ttl=300)
def profesor_to_curso_map(df_cursos: pd.DataFrame) -> Dict[str, str]:
    m = {}
    for _, r in df_cursos.iterrows():
        if pd.notna(r.get("Profesor")):
            m[str(r["Profesor"]).strip().upper()] = r["Curso"]
    return m


def curso_de_fila(row, prof_curso_map: Dict[str, str]) -> Optional[str]:
    prof_field = row.get("Profesor(es)") or ""
    for p in str(prof_field).split("|"):
        hit = prof_curso_map.get(p.strip().upper())
        if hit:
            return hit
    return None


@st.cache_data(ttl=300)
def cursos_unicos_map(df_cursos: pd.DataFrame) -> Dict[str, dict]:
    m = {}
    for _, r in df_cursos.iterrows():
        nombre = r.get("Curso")
        if not nombre or pd.isna(nombre):
            continue
        if nombre not in m:
            m[nombre] = {
                "curso": nombre, "ciclo": r.get("Ciclo"), "fechas": r.get("Fechas curso"),
                "modalidad": r.get("Modalidad"), "capacidad": r.get("Capacidad") or 0,
                "area": str(r.get("Area curso") or "").strip(),
                "profesores": [r.get("Profesor")],
            }
        else:
            m[nombre]["profesores"].append(r.get("Profesor"))
    return m


def tipo_programa(row) -> Optional[str]:
    t = row.get("Tipo de programa")
    if t is None or (isinstance(t, float) and pd.isna(t)):
        return None
    t = str(t).strip()
    return "External" if t == "Extensión" else t


CATEGORY_LABELS = {
    "EPOS UASM": "UASM GR", "PRE UASM": "UASM UG", "Otros Pre y Posgrados": "Other UG/GR",
    "Extensión": "External", "External": "External",
}
PROGRAM_TYPE_LABELS = {
    "Pregrado": "Undergraduate", "Otros Pre y Posgrados": "Other Undergraduate/Graduate",
    "Especializaciones": "Specializations",
}


@st.cache_data(ttl=300)
def compute_all():
    """Agrega todo lo que EIV.Compute calculaba en JS, en un solo objeto
    reutilizable por todas las páginas (cacheado, así que las páginas no
    repiten el trabajo)."""
    df_cursos = load_cursos()
    df_listas, df_programas = load_listas()

    p2c = profesor_to_curso_map(df_cursos)
    df_listas = df_listas.copy()
    df_listas["_curso"] = df_listas.apply(lambda r: curso_de_fila(r, p2c), axis=1)
    df_listas["_tipo"] = df_listas.apply(tipo_programa, axis=1)

    cmap = cursos_unicos_map(df_cursos)
    c2b = {c: meta["ciclo"] for c, meta in cmap.items()}
    df_listas["_bloque"] = df_listas["_curso"].map(c2b)

    inscritos_por_curso = df_listas.dropna(subset=["_curso"])["_curso"].value_counts().to_dict()

    total_capacidad = sum(meta["capacidad"] or 0 for meta in cmap.values())
    total_inscritos = len(df_listas)
    participantes_unicos = df_listas["Código est"].dropna().nunique()

    tabla_cursos_rows = []
    for curso, meta in cmap.items():
        ins = inscritos_por_curso.get(curso, 0)
        cap = meta["capacidad"] or 0
        tabla_cursos_rows.append({
            "bloque": meta["ciclo"], "curso": curso,
            "profesores": " & ".join([p for p in meta["profesores"] if p]),
            "modalidad": meta["modalidad"], "inscritos": ins, "cupos": cap,
            "ocupacion": (ins / cap * 100) if cap else 0,
        })

    tabla_por_profesor_rows = []
    for _, r in df_cursos.iterrows():
        curso = r.get("Curso")
        meta = cmap.get(curso, {})
        ins = inscritos_por_curso.get(curso, 0)
        cap = meta.get("capacidad", 0)
        tabla_por_profesor_rows.append({
            "bloque": r.get("Ciclo"), "curso": curso, "profesor": r.get("Profesor"),
            "genero": str(r.get("Género") or "").strip(), "area": str(r.get("Area curso") or "").strip(),
            "universidad": str(r.get("Universidad") or "").strip(),
            "pais": str(r.get("País Universidad") or "").strip(),
            "modalidad": r.get("Modalidad"), "inscritos": ins, "cupos": cap or 0,
            "ocupacion": (ins / cap * 100) if cap else 0,
        })

    modalidad_counts: Dict[str, int] = {}
    for meta in cmap.values():
        m = meta["modalidad"] or "Unspecified"
        modalidad_counts[m] = modalidad_counts.get(m, 0) + 1

    bloques = sorted(set(meta["ciclo"] for meta in cmap.values()))

    # Notas (grades) — joined via professor -> course
    notas = df_listas.dropna(subset=["_curso"]).copy()
    notas["_nota"] = pd.to_numeric(
        notas["Nota final"].astype(str).str.replace(",", ".").str.extract(r"([\d.]+)")[0], errors="coerce"
    )
    notas = notas.dropna(subset=["_nota"])

    return {
        "df_cursos": df_cursos, "df_listas": df_listas, "df_programas": df_programas,
        "cmap": cmap, "c2b": c2b, "bloques": bloques,
        "inscritos_por_curso": inscritos_por_curso,
        "total_capacidad": total_capacidad, "total_inscritos": total_inscritos,
        "participantes_unicos": participantes_unicos,
        "ocupacion_pct": (total_inscritos / total_capacidad * 100) if total_capacidad else 0,
        "profesores_unicos": df_cursos["Profesor"].dropna().nunique(),
        "modalidad_counts": modalidad_counts,
        "promedio_estudiantes_curso": (total_inscritos / len(cmap)) if cmap else 0,
        "tabla_cursos": pd.DataFrame(tabla_cursos_rows),
        "tabla_por_profesor": pd.DataFrame(tabla_por_profesor_rows),
        "notas": notas,
    }


def filtered_listas(bloque: str, curso: str) -> pd.DataFrame:
    d = compute_all()
    df = d["df_listas"]
    if curso != "all":
        df = df[df["_curso"] == curso]
    elif bloque != "all":
        df = df[df["_bloque"].astype(str) == str(bloque)]
    return df


def filtered_notas(bloque: str, curso: str) -> pd.DataFrame:
    d = compute_all()
    df = d["notas"]
    if curso != "all":
        df = df[df["_curso"] == curso]
    elif bloque != "all":
        df = df[df["_bloque"].astype(str) == str(bloque)]
    return df


# ── 7) EIV.Evaluation equivalente ────────────────────────────────────────
def _norm_name(name) -> str:
    """Normaliza nombres de profesor (espacios/mayúsculas) antes de cruzar
    entre BD_cursos.xlsx y BD_evaluacion_curso.xlsx / BD_satisfaccionEIV.xlsx
    -- si alguno de los 3 archivos trae un espacio doble o una diferencia de
    mayúsculas, una comparación exacta pierde silenciosamente TODOS los
    datos de ese profesor (parecía que 'no leía la información')."""
    return re.sub(r"\s+", " ", str(name)).strip().upper()


def _resolve_alias(name: str) -> str:
    n = _norm_name(name)
    return NAME_ALIASES.get(n, n)


@st.cache_data(ttl=300)
def evaluation_by_professor() -> Dict[str, dict]:
    frec, _ = load_evaluacion()
    lang_map = professor_language_map()
    periodo_eiv_col = next((c for c in frec.columns if re.sub(r"\s+", "", c).strip().casefold() == "periodo_eiv"), None)
    profs: Dict[str, dict] = {}
    for _, r in frec.iterrows():
        prof = r.get("nombre_profesor")
        if pd.isna(prof):
            continue
        prof = _resolve_alias(prof)
        qid = r.get("id_pregunta")
        if pd.isna(qid):
            continue
        lang = lang_map.get(prof, "en")
        # BD_evaluacion_curso.xlsx trae pregunta/aspecto/respuesta en
        # español Y en inglés en la misma fila (columnas H/I, E/F, L/M) —
        # se elige la columna según el idioma real del profesor (Idioma en
        # BD_cursos), no un idioma fijo para todos.
        if lang == "es":
            pregunta_col, aspecto_col, respuesta_col = "pregunta", "aspecto_evaluado", "respuesta"
        else:
            pregunta_col, aspecto_col, respuesta_col = "question", "evaluated_aspect", "answer"
        pregunta = r.get(pregunta_col) if pregunta_col in r.index and pd.notna(r.get(pregunta_col)) else r.get("pregunta")
        aspecto = r.get(aspecto_col) if aspecto_col in r.index and pd.notna(r.get(aspecto_col)) else r.get("aspecto_evaluado")
        respuesta = r.get(respuesta_col) if respuesta_col in r.index and pd.notna(r.get(respuesta_col)) else r.get("respuesta")

        d = profs.setdefault(prof, {"questions": {}, "courses": set(), "lang": lang})
        if pd.notna(r.get("nombre_curso")):
            d["courses"].add(r["nombre_curso"])
        q = d["questions"].setdefault(qid, {"text": pregunta, "aspect": aspecto, "options": {}, "by_cursec": {}, "by_period": {}})
        n = int(r.get("respuestas_por_opcion") or 0)
        q["options"][respuesta] = q["options"].get(respuesta, 0) + n
        cursec_val = r.get("curSec")
        cursec_key = str(cursec_val).strip() if pd.notna(cursec_val) else "?"
        cs = q["by_cursec"].setdefault(cursec_key, {})
        cs[respuesta] = cs.get(respuesta, 0) + n
        # periodo_EIV: columna nueva que identifica el periodo REAL de
        # matrícula del estudiante que respondió (202613 EPOS / 202618
        # Pregrado y otros) -- permite partir resultados por periodo de
        # verdad, en vez de la heurística por prefijo de curSec.
        periodo_val = r.get(periodo_eiv_col) if periodo_eiv_col else None
        periodo_key = str(periodo_val).strip() if pd.notna(periodo_val) else "?"
        pe = q["by_period"].setdefault(periodo_key, {})
        pe[respuesta] = pe.get(respuesta, 0) + n
    return profs


@st.cache_data(ttl=300)
def comments_by_professor() -> Dict[str, dict]:
    _, com = load_evaluacion()
    lang_map = professor_language_map()
    respuesta_candidates_es = ["respuesta", "Respuesta"]
    respuesta_candidates_en = ["answer", "Answer", "respuesta en inglés", "respuesta_ingles"]
    periodo_eiv_col = next((c for c in com.columns if re.sub(r"\s+", "", c).strip().casefold() == "periodo_eiv"), None)
    profs: Dict[str, dict] = {}
    for _, r in com.iterrows():
        prof = r.get("nombre_profesor")
        if pd.isna(prof):
            continue
        prof = _resolve_alias(prof)
        qid = r.get("id_pregunta")
        if pd.isna(qid):
            continue
        lang = lang_map.get(prof, "en")
        candidates = respuesta_candidates_en if lang == "en" else respuesta_candidates_es
        ans = None
        for c in candidates:
            if c in r.index and pd.notna(r.get(c)):
                ans = r.get(c)
                break
        if ans is None:
            ans = r.get("respuesta")  # respaldo: solo hay una columna de comentarios (no bilingüe)
        clean = "" if (ans is None or (isinstance(ans, float) and pd.isna(ans))) else str(ans).strip()
        if not clean or clean.lower() == "n/a" or clean == "0":
            continue
        pregunta_col = "question" if (lang == "en" and "question" in r.index and pd.notna(r.get("question"))) else "pregunta"
        d = profs.setdefault(prof, {})
        q = d.setdefault(qid, {"text": r.get(pregunta_col), "items": []})
        periodo_val = r.get(periodo_eiv_col) if periodo_eiv_col else None
        periodo = str(periodo_val).strip() if pd.notna(periodo_val) else None
        q["items"].append({"text": clean, "periodo": periodo})
    return profs


def weighted_avg(option_counts: Dict[str, int]) -> Optional[float]:
    s, n = 0.0, 0
    for ans, c in option_counts.items():
        w = LIKERT_WEIGHTS.get(ans)
        if w is None:
            continue
        s += w * c
        n += c
    return (s / n) if n else None


def build_professor_eval_data(prof_name: str) -> Optional[dict]:
    by_prof = evaluation_by_professor()
    prof_name = _resolve_alias(prof_name)
    q = by_prof.get(prof_name)
    if not q:
        return None
    questions = q["questions"]

    aspects: Dict[str, dict] = {}
    nps_counts: Dict[str, int] = {}
    workload_counts: Dict[str, int] = {}
    obj_counts: Dict[str, int] = {}
    total_sum, total_n = 0.0, 0

    for qid, qd in questions.items():
        for ans, c in qd["options"].items():
            w = LIKERT_WEIGHTS.get(ans)
            if w is not None:
                total_sum += w * c
                total_n += c
                asp = qd["aspect"] or "Unknown"
                a = aspects.setdefault(asp, {"sum": 0.0, "n": 0})
                a["sum"] += w * c
                a["n"] += c
        if qid == NPS_QID:
            for ans, c in qd["options"].items():
                nps_counts[ans] = nps_counts.get(ans, 0) + c
        if qid == WORKLOAD_QID:
            for ans, c in qd["options"].items():
                workload_counts[ans] = workload_counts.get(ans, 0) + c
        if qid == OBJ_QID:
            for ans, c in qd["options"].items():
                obj_counts[ans] = obj_counts.get(ans, 0) + c

    avg = (total_sum / total_n) if total_n else None
    nps_n = sum(nps_counts.values())
    nps = None
    if nps_n:
        s = 0
        for ans, c in nps_counts.items():
            m = re.search(r"\d+", str(ans))
            if m:
                s += int(m.group()) * c
        nps = s / nps_n
    obj_n = sum(obj_counts.values())
    obj_pct = None
    if obj_n:
        todos = sum(c for ans, c in obj_counts.items() if str(ans).strip().casefold() in OBJ_ALL_VALUES)
        obj_pct = todos / obj_n * 100

    aspects_list = sorted(
        [{"aspect": a, "avg": v["sum"] / v["n"], "n": v["n"]} for a, v in aspects.items()],
        key=lambda x: -x["avg"],
    )

    d = compute_all()
    # OJO: 'nombre_curso' de BD_evaluacion_curso.xlsx viene TRUNCADO (ej.
    # "LEAD. DISRUPT: RESIL. & INNOV"), mientras que df_listas/_curso usa el
    # nombre COMPLETO de BD_cursos.xlsx (ej. "LEADING THROUGH DISRUPTION:
    # BUILDING RESILIENCE AND INNOVATION") -- comparar esos dos strings
    # directamente casi nunca calza, dejando 'inscritos' en 0 para casi
    # todos los profesores (y por eso Response Rate salía sobre 100%: el
    # numerador sumaba bien pero el denominador casi no acumulaba). Se
    # resuelve el curso real por el PROFESOR (consistente entre archivos),
    # no por el nombre truncado.
    df_cursos_all = d["df_cursos"]
    prof_norm_key = _norm_name(prof_name)
    match_row = df_cursos_all[df_cursos_all["Profesor"].astype(str).apply(_norm_name) == prof_norm_key]
    course = match_row["Curso"].iloc[0] if not match_row.empty else next(iter(q["courses"]), None)
    row = d["df_listas"][d["df_listas"]["_curso"] == course] if course else pd.DataFrame()
    inscritos = len(row)

    # ---- División por periodo real de matrícula (columna periodo_EIV en
    # BD_evaluacion_curso.xlsx: 202613 = EPOS/GR, 202618 = Pregrado y otros
    # /UG) -- calcula el mismo set de estadísticas (avg, NPS, objetivos,
    # aspectos) para CADA periodo que de verdad aparezca en las respuestas
    # de este profesor, sin inventar una división si no hay datos.
    PERIOD_GROUP_LABEL = {"202613": "GR", "202618": "UG"}
    all_periods = set()
    for qd in questions.values():
        all_periods.update(qd.get("by_period", {}).keys())
    all_periods.discard("?")

    periods: Dict[str, dict] = {}
    for periodo in all_periods:
        p_aspects: Dict[str, dict] = {}
        p_nps_counts: Dict[str, int] = {}
        p_obj_counts: Dict[str, int] = {}
        p_workload_counts: Dict[str, int] = {}
        p_questions: Dict[str, dict] = {}
        for qid, qd in questions.items():
            opts = qd.get("by_period", {}).get(periodo, {})
            if not opts:
                continue
            p_questions[qid] = {"text": qd["text"], "aspect": qd["aspect"], "options": opts}
            for ans, c in opts.items():
                w = LIKERT_WEIGHTS.get(ans)
                if w is not None:
                    asp = qd["aspect"] or "Unknown"
                    a = p_aspects.setdefault(asp, {"sum": 0.0, "n": 0})
                    a["sum"] += w * c
                    a["n"] += c
            if qid == NPS_QID:
                for ans, c in opts.items():
                    p_nps_counts[ans] = p_nps_counts.get(ans, 0) + c
            if qid == OBJ_QID:
                for ans, c in opts.items():
                    p_obj_counts[ans] = p_obj_counts.get(ans, 0) + c
            if qid == WORKLOAD_QID:
                for ans, c in opts.items():
                    p_workload_counts[ans] = p_workload_counts.get(ans, 0) + c
        p_aspects_list = sorted(
            [{"aspect": a, "avg": v["sum"] / v["n"], "n": v["n"]} for a, v in p_aspects.items()],
            key=lambda x: -x["avg"],
        )
        p_n_total = sum(v["n"] for v in p_aspects.values())
        p_score_total = sum(v["sum"] for v in p_aspects.values())
        p_avg = (p_score_total / p_n_total) if p_n_total else None
        p_nps_n = sum(p_nps_counts.values())
        p_nps = None
        if p_nps_n:
            s = sum(int(m.group()) * c for ans, c in p_nps_counts.items() for m in [re.search(r"\d+", str(ans))] if m)
            p_nps = s / p_nps_n
        p_obj_n = sum(p_obj_counts.values())
        p_obj_pct = None
        if p_obj_n:
            todos = sum(c for ans, c in p_obj_counts.items() if str(ans).strip().casefold() in OBJ_ALL_VALUES)
            p_obj_pct = todos / p_obj_n * 100
        periods[periodo] = {
            "group": PERIOD_GROUP_LABEL.get(periodo, periodo), "avg": p_avg, "n": p_n_total,
            "nps": p_nps, "npsN": p_nps_n, "objectivesAllPct": p_obj_pct, "objN": p_obj_n,
            "respondents": p_nps_n, "aspects": p_aspects_list, "questions": p_questions,
            "workloadCounts": p_workload_counts,
        }

    return {
        "avg": avg, "n": total_n, "nps": nps, "npsN": nps_n, "objectivesAllPct": obj_pct, "objN": obj_n,
        "workloadCounts": workload_counts, "aspects": aspects_list, "respondents": nps_n,
        "inscritos": inscritos, "course": course, "questions": questions, "lang": q.get("lang", "en"),
        "periods": periods,
    }


# ── 8) PÁGINA — Data Center ──────────────────────────────────────────────
# ── 8) PÁGINA — Cover (portada, sin sidebar ni nav) ──────────────────────
# ── 8) PÁGINA — Cover (portada, funciona como Data Center — sin sidebar) ──
def page_cover():
    st.markdown(f"<div style='height:0.2vh;'></div>", unsafe_allow_html=True)
    with st.container(key="cover_banner"):
        try:
            _show_eiv_banner(width=777)
        except Exception:
            pass
    col_l2, col_mid2, col_r2 = st.columns([1, 2, 1])
    with col_mid2:
        st.markdown(
            f'<div style="text-align:center;font-size:34px;font-weight:800;color:{INK};margin-top:18px;">'
            'International Summer School</div>'
            f'<div style="text-align:center;font-size:16px;color:{MUTED};margin:10px auto 0;line-height:1.5;">'
            "Analytics for the 2026 edition of EIV — enrollment, course evaluation, faculty "
            "satisfaction, and financial performance across every visiting-faculty course.</div>",
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
            ("BD_cursos.xlsx", "Courses & 2026 capacity", CURSOS_FILE_ID),
            ("BD_listas.xlsx", "Enrollment, program catalog & final grades", LISTAS_FILE_ID),
            ("BD_evaluacion_curso.xlsx", "Course evaluation — frequencies & comments", EVALUACION_FILE_ID),
            ("BD_satisfaccionEIV.xlsx", "Faculty satisfaction survey", SATISFACCION_FILE_ID),
            ("BD_gastos.xlsx", "Expenses & actual spend", GASTOS_FILE_ID),
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
                    "⇩", data=_download_drive_file_bytes(fid), file_name=fname,
                    key=f"cover_dl_{fname}", use_container_width=True,
                )


# ── 9) PÁGINA — Overview (Historical Evolution) ──────────────────────────
def page_overview():
    d = compute_all()
    _render_header(
        "Overview & Historical Evolution",
        "Participants and program-category composition, 2012–2026 (2026 is live from the current data).",
    )

    serie = [dict(p) for p in HIST_PARTICIPANTS]
    serie.append({"year": 2026, "val": d["total_inscritos"], "seats": d["total_capacidad"], "live": True})

    st.markdown("### Participants & Available Capacity")
    fig = go.Figure()
    with_seats = [p for p in serie if p.get("seats")]
    fig.add_trace(go.Bar(
        x=[p["year"] for p in with_seats], y=[p["seats"] for p in with_seats], name="Available capacity",
        marker=dict(color="rgba(18,189,251,0.10)", line=dict(color=BLUE, width=1.6)), width=0.6,
    ))
    point_text = []
    for p in serie:
        if p.get("seats"):
            occ = p["val"] / p["seats"] * 100
            point_text.append(f'{p["val"]} ({occ:.0f}% occ.)')
        else:
            point_text.append(str(p["val"]))
    fig.add_trace(go.Scatter(
        x=[p["year"] for p in serie], y=[p["val"] for p in serie], mode="lines+markers+text",
        name="Participants (seats)", line=dict(color=INK, width=3, shape="spline"),
        marker=dict(size=7, color=INK), text=point_text, textposition="top center",
        textfont=dict(size=10.5, color=INK),
    ))
    fig.update_layout(
        margin=dict(t=20, r=24, b=40, l=50), barmode="overlay",
        legend=dict(orientation="h", x=0, y=1.15),
        yaxis=dict(range=[0, 800], gridcolor="#EFEBEE"), xaxis=dict(dtick=1),
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", height=380,
    )
    st.plotly_chart(fig, use_container_width=True)

    hist_df = pd.DataFrame([{
        "Year": p["year"], "Participants (seats)": p["val"], "Available capacity": p.get("seats") or "",
        "Occupancy %": round(p["val"] / p["seats"] * 100, 1) if p.get("seats") else "",
    } for p in serie])
    st.download_button("Download as Excel", data=_xlsx_bytes(hist_df, "Historical_Participants"),
                        file_name="Historical_Participants.xlsx", key="dl_hist_part")

    st.markdown("### Program-Category Composition")
    listas = d["df_listas"]
    total = len(listas) or 1
    def _pct(cat):
        return listas["Categoría académica"].eq(cat).sum() / total * 100
    cat26 = {
        "year": 2026, "EPOS": _pct("EPOS UASM"), "PRE": _pct("PRE UASM"), "OTHER": _pct("Otros Pre y Posgrados"),
        "EXTERNAL": _pct("Extensión") + _pct("External"), "live": True,
    }
    cat_all = HIST_CATEGORY + [cat26]

    fig2 = go.Figure()
    series = [("EPOS", "UASM GR", BLUE), ("PRE", "UASM UG", GREEN), ("OTHER", "Other UG/GR", PINK), ("EXTERNAL", "External", PURPLE)]
    for key, name, color in series:
        fig2.add_trace(go.Bar(
            x=[str(c["year"]) for c in cat_all], y=[c[key] for c in cat_all], name=name,
            marker=dict(color=color), text=[f"{c[key]:.1f}%" for c in cat_all], textposition="outside",
            textfont=dict(size=11, color=INK),
        ))
    fig2.update_layout(
        margin=dict(t=20, r=16, b=36, l=46), barmode="group", legend=dict(orientation="h", x=0, y=1.2),
        yaxis=dict(ticksuffix="%", range=[0, 76], gridcolor="#EFEBEE"), xaxis=dict(type="category"),
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", height=360,
    )
    st.plotly_chart(fig2, use_container_width=True)

    # Total de participantes por año (para convertir % -> # de inscritos por
    # categoría); 2026 usa el conteo real, los años anteriores se estiman
    # sobre el total histórico de esa edición.
    year_totals = {p["year"]: p["val"] for p in HIST_PARTICIPANTS}
    year_totals[2026] = total
    cat_df_long = pd.DataFrame([{
        "Year": c["year"],
        "UASM GR": round(c["EPOS"] / 100 * year_totals.get(c["year"], total)),
        "UASM UG": round(c["PRE"] / 100 * year_totals.get(c["year"], total)),
        "Other UG/GR": round(c["OTHER"] / 100 * year_totals.get(c["year"], total)),
        "External": round(c["EXTERNAL"] / 100 * year_totals.get(c["year"], total)),
    } for c in cat_all])
    cat_df = cat_df_long.set_index("Year").T.rename_axis("Category").reset_index()
    cat_df.columns = [str(c) for c in cat_df.columns]
    st.dataframe(cat_df, use_container_width=True, hide_index=True)
    st.download_button("Download as Excel", data=_xlsx_bytes(cat_df, "Historical_Composition"),
                        file_name="Historical_Composition.xlsx", key="dl_hist_cat")


# ── 10) PÁGINA — Summary (2026 general) ──────────────────────────────────
def page_summary():
    d = compute_all()
    _render_header("2026 General Summary", "Courses, delivery modality, and where our visiting faculty came from.")

    kpis = [
        ("Total courses", len(d["cmap"])),
        ("On Campus", d["modalidad_counts"].get("On Campus", 0)),
        ("Online", d["modalidad_counts"].get("Online", 0)),
        ("Bootcamp", d["modalidad_counts"].get("Bootcamp", 0)),
        ("Avg. students / course", f"{d['promedio_estudiantes_curso']:.1f}"),
        ("Blocks", len(d["bloques"])),
    ]
    cols = st.columns(len(kpis))
    for col, (label, val) in zip(cols, kpis):
        with col:
            _kpi(label, val)

    tabla = d["tabla_por_profesor"]
    st.markdown("### Where Our Visiting Faculty Came From")
    by_country: Dict[str, dict] = {}
    for _, r in tabla.iterrows():
        for country in [c.strip() for c in str(r["pais"]).split(",") if c.strip()]:
            bc = by_country.setdefault(country, {"count": 0, "profs": set(), "unis": set()})
            bc["count"] += 1
            bc["profs"].add(r["profesor"])
            bc["unis"].add(r["universidad"])

    col_map, col_side = st.columns([3, 1])

    # Lee la selección del click anterior sobre el mapa (antes de construirlo).
    _sel = st.session_state.get("summary_map", {})
    _sel_points = _sel.get("selection", {}).get("points", []) if _sel else []
    focus = None
    if _sel_points:
        cd = _sel_points[0].get("customdata")
        if cd:
            focus = cd[0] if isinstance(cd, (list, tuple)) else cd

    with col_side:
        _kpi("Countries", len(by_country), "accent")
        all_unis = sorted({u for info in by_country.values() for u in info["unis"]})
        _kpi("Universities", len(all_unis), "accent")
        st.markdown("**Universities**" + (f" — {focus}" if focus else ""))
        unis_to_show = sorted(by_country[focus]["unis"]) if focus and focus in by_country else all_unis
        with st.container(height=260):
            for u in unis_to_show:
                st.markdown(f"- {u}")

    with col_map:
        fig_map = go.Figure()
        for country, info in by_country.items():
            coords = COUNTRY_GEO.get(country)
            if not coords:
                continue
            lat0, lon0 = coords
            highlight = not focus or focus == country
            fig_map.add_trace(go.Scattergeo(
                lat=[lat0, BOGOTA[0]], lon=[lon0, BOGOTA[1]], mode="lines",
                line=dict(color=COUNTRY_COLORS.get(country, PINK), width=1.6),
                opacity=0.7 if highlight else 0.12, showlegend=False, hoverinfo="skip",
            ))
            fig_map.add_trace(go.Scattergeo(
                lat=[lat0], lon=[lon0], mode="markers+text", text=[country], customdata=[country],
                textposition="top center", textfont=dict(size=11, color=INK if highlight else "#C9BEC5"),
                marker=dict(size=10 + len(info["profs"]) * 3, color=COUNTRY_COLORS.get(country, PINK),
                            opacity=1 if highlight else 0.25, line=dict(color="#fff", width=1.5)),
                hovertext=f"{country}: {len(info['profs'])} faculty, {len(info['unis'])} universities — click to filter",
                hoverinfo="text",
            ))
        fig_map.add_trace(go.Scattergeo(
            lat=[BOGOTA[0]], lon=[BOGOTA[1]], mode="markers+text", text=["Bogotá"],
            textposition="bottom center", textfont=dict(size=11, color=INK),
            marker=dict(size=11, color=INK, symbol="diamond"), showlegend=False, hoverinfo="skip",
        ))
        fig_map.update_geos(showland=True, landcolor="#F1EDEF", showcountries=True, countrycolor="#DDD4D9",
                            showocean=True, oceancolor="#FAF8FA", bgcolor="rgba(0,0,0,0)", projection_type="natural earth")
        fig_map.update_layout(margin=dict(t=6, r=6, b=6, l=6), height=420, showlegend=False, paper_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig_map, use_container_width=True, on_select="rerun", selection_mode="points", key="summary_map")
        if focus:
            st.caption("Double click the map to view all countries again.")

    st.markdown("### Courses by Block")
    tabla_f = tabla[tabla["pais"].str.contains(focus, case=False, na=False)] if focus else tabla
    total_profs = tabla_f["profesor"].nunique()
    gender_counts = tabla_f.drop_duplicates(subset=["profesor"])["genero"].value_counts()
    gtotal = gender_counts.sum() or 1
    male_pct = gender_counts.get("Male", 0) / gtotal * 100
    female_pct = gender_counts.get("Female", 0) / gtotal * 100

    col_tbl, col_donut = st.columns([3, 1])
    with col_tbl:
        st.markdown(
            f'<div style="display:flex;justify-content:flex-start;align-items:flex-start;gap:26px;margin-bottom:6px;">'
            f'<div style="text-align:left;"><div class="kv accent">{total_profs}</div>'
            f'<div class="kl">Visiting Faculty</div></div>'
            f'<div style="text-align:left;"><div style="font-size:14px;font-weight:700;color:{INK};">{male_pct:.0f}%</div>'
            f'<div class="kl">% Male</div></div>'
            f'<div style="text-align:left;"><div style="font-size:14px;font-weight:700;color:{INK};">{female_pct:.0f}%</div>'
            f'<div class="kl">% Female</div></div></div>',
            unsafe_allow_html=True,
        )
        if focus:
            st.caption(f"Filtered to courses taught by faculty from **{focus}**.")
        display_cols = ["bloque", "curso", "profesor", "area", "universidad", "pais", "modalidad", "inscritos", "cupos", "ocupacion"]
        disp = tabla_f[display_cols].copy().sort_values("bloque").reset_index(drop=True)
        disp.columns = ["Block", "Course", "Professor", "Area", "University", "Country", "Modality", "Enrolled", "Seats", "Occupancy"]
        styled = _style_courses_table(disp)
        st.dataframe(styled, use_container_width=True, hide_index=True, height=460)
        st.download_button("Download as Excel", data=_xlsx_bytes(disp, "Courses_by_Block"),
                            file_name="Courses_by_Block.xlsx", key="dl_courses_block")
    with col_donut:
        st.markdown("##### By Area")
        area_counts = tabla_f.drop_duplicates(subset=["curso"])["area"].value_counts()
        area_labels = [f"<b>{str(a).replace(' & ', ' &<br>')}</b>" for a in area_counts.index]
        fig_area = go.Figure(go.Pie(
            labels=area_labels, values=area_counts.values, hole=0.55,
            marker=dict(colors=[BLUE, PURPLE, PINK, GREEN, "#F2994A", "#3FA34D", "#8C7F87"],
                        line=dict(color="#fff", width=1.5)),
            textinfo="label+value", textposition="inside", insidetextorientation="horizontal",
            textfont=dict(size=15, color=INK),
        ))
        fig_area.update_layout(margin=dict(t=10, r=10, b=10, l=10), showlegend=False,
                                paper_bgcolor="rgba(0,0,0,0)", height=460, uniformtext_minsize=10,
                                uniformtext_mode="show")
        st.plotly_chart(fig_area, use_container_width=True)


# ── 11) PÁGINA — Dashboard (Enrollment Composition) ──────────────────────
GRADE_BUCKETS = [
    ("< 3.0", 0, 3.0), ("3.1 – 3.5", 3.0, 3.6), ("3.6 – 4.0", 3.6, 4.1),
    ("4.1 – 4.5", 4.1, 4.6), ("4.6 – 5.0", 4.6, 5.01),
]


def page_dashboard():
    d = compute_all()
    cmap = d["cmap"]

    with st.sidebar:
        st.markdown("#### Scope")
        bloque_opts = ["All"] + [str(b) for b in d["bloques"]]
        bloque_sel = st.selectbox("Block", bloque_opts, key="dash_bloque")
        bloque = "all" if bloque_sel == "All" else bloque_sel
        cursos_in_scope = sorted(
            [c for c, m in cmap.items() if bloque == "all" or str(m["ciclo"]) == str(bloque)]
        )
        curso_sel = st.selectbox("Course", ["All"] + cursos_in_scope, key="dash_curso")
        curso = "all" if curso_sel == "All" else curso_sel

    _render_header("Enrollment Composition", "Program-type mix, courses taken per student, and grade dispersion.")

    f_listas = filtered_listas(bloque, curso)
    f_notas = filtered_notas(bloque, curso)

    note = []
    if bloque != "all":
        note.append(f"Block {bloque}")
    if curso != "all":
        note.append(curso)
    st.caption("Filtered by: " + " · ".join(note) if note else "Showing all 2026 enrollments")

    if curso != "all":
        meta = cmap.get(curso, {})
        profs = " & ".join([p for p in meta.get("profesores", []) if p])
        if profs:
            st.markdown(f"**{profs}** — {curso}")

    col_chart, col_kpi = st.columns([2, 1], vertical_alignment="center")
    with col_kpi:
        st.markdown(
            f'<div style="display:flex;flex-direction:column;justify-content:center;height:100%;gap:22px;">'
            f'<div><div class="kv accent" style="font-size:32px;">{len(f_listas):,}</div>'
            f'<div class="kl">Enrolled seats</div></div>'
            f'<div><div class="kv accent" style="font-size:32px;">{f_listas["Código est"].dropna().nunique():,}</div>'
            f'<div class="kl">Unique participants</div></div></div>',
            unsafe_allow_html=True,
        )
    with col_chart:
        st.markdown("##### Program Type Distribution")
        total = len(f_listas) or 1
        counts = f_listas["_tipo"].fillna("Unspecified").value_counts()
        labels = [PROGRAM_TYPE_LABELS.get(l, l) for l in counts.index]
        values = (counts / total * 100)
        fig = go.Figure(go.Bar(
            x=values.values, y=labels, orientation="h",
            marker=dict(color=[{"External": PURPLE, "Otros Pre y Posgrados": PINK, "Pregrado": GREEN}.get(l, BLUE)
                                for l in counts.index]),
            text=[f"{v:.1f}%" for v in values.values], textposition="outside",
        ))
        fig.update_layout(margin=dict(t=6, r=36, b=36, l=160), xaxis=dict(ticksuffix="%", gridcolor="#EFEBEE"),
                           yaxis=dict(autorange="reversed"), plot_bgcolor="rgba(0,0,0,0)",
                           paper_bgcolor="rgba(0,0,0,0)", height=340)
        st.plotly_chart(fig, use_container_width=True)

    st.markdown(f"##### Grade Dispersion — {'All Courses' if curso=='all' and bloque=='all' else (curso if curso!='all' else f'Block {bloque}')}")
    total_n = len(f_notas) or 1
    counts_g = []
    for label, lo, hi in GRADE_BUCKETS:
        n = f_notas[(f_notas["_nota"] >= lo) & (f_notas["_nota"] < hi)].shape[0]
        counts_g.append(n)
    fig3 = go.Figure(go.Bar(
        x=[b[0] for b in GRADE_BUCKETS], y=[c / total_n * 100 for c in counts_g],
        marker=dict(color=[color_scale(i / (len(GRADE_BUCKETS) - 1) * 100) for i in range(len(GRADE_BUCKETS))]),
        text=[f"{c} students" for c in counts_g], textposition="outside",
    ))
    fig3.update_layout(margin=dict(t=6, r=16, b=30, l=46), yaxis=dict(ticksuffix="%", gridcolor="#EFEBEE"),
                        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", height=280)
    st.plotly_chart(fig3, use_container_width=True)

    st.markdown("### Enrollment Composition Table")
    rows = d["tabla_cursos"]
    rows = rows[
        ((bloque == "all") | (rows["bloque"].astype(str) == str(bloque)))
        & ((curso == "all") | (rows["curso"] == curso))
    ]
    tipos = sorted(d["df_listas"]["_tipo"].dropna().unique().tolist())
    grid: Dict[str, Dict[str, int]] = {}
    for _, r in d["df_listas"].dropna(subset=["_curso"]).iterrows():
        gc = grid.setdefault(r["_curso"], {})
        t = r["_tipo"] or "Unspecified"
        gc[t] = gc.get(t, 0) + 1

    table_rows = []
    for _, r in rows.iterrows():
        rowgrid = grid.get(r["curso"], {})
        row_total = sum(rowgrid.get(t, 0) for t in tipos)
        table_rows.append({"Course / Faculty": f"{r['curso']} ({r['profesores']})",
                            **{t: rowgrid.get(t, "") for t in tipos}, "Total": row_total})
    comp_df = pd.DataFrame(table_rows)

    col_totals = {t: comp_df[t].apply(lambda v: v if isinstance(v, (int, float)) else 0).sum() for t in tipos}
    grand_total = sum(col_totals.values()) or 1
    total_row = {"Course / Faculty": "Total", **col_totals, "Total": grand_total}
    pct_row = {"Course / Faculty": "% of Total",
               **{t: f"{col_totals[t] / grand_total * 100:.1f}%" for t in tipos}, "Total": "100.0%"}
    comp_df = pd.concat([comp_df, pd.DataFrame([total_row, pct_row])], ignore_index=True)

    pct_vals = {t: col_totals[t] / grand_total * 100 for t in tipos}
    lo, hi = min(pct_vals.values()), max(pct_vals.values())
    span = (hi - lo) or 1
    pct_idx = comp_df.index[-1]

    def _style_comp(row: pd.Series):
        if row.name != pct_idx:
            return [""] * len(row)
        out = []
        for c in row.index:
            if c in tipos:
                scaled = (pct_vals[c] - lo) / span * 100
                # background-color (no shorthand): el renderer de st.dataframe
                # no siempre respeta 'background' de forma fiable con colores sólidos.
                out.append(f"background-color:{color_scale(scaled)};color:#fff;font-weight:700;")
            else:
                out.append("font-weight:700;")
        return out

    comp_height = 38 * (len(comp_df) + 1) + 3  # sin scroll: alto = todas las filas + encabezado
    st.dataframe(comp_df.style.apply(_style_comp, axis=1), use_container_width=True, hide_index=True, height=comp_height)
    st.download_button("Download as Excel", data=_xlsx_bytes(comp_df, "Enrollment_Composition"),
                        file_name="Enrollment_Composition.xlsx", key="dl_enrollment_comp")

    st.markdown("##### Courses Taken per Student")
    per_student = d["df_listas"].dropna(subset=["Código est"]).groupby("Código est").agg(
        count=("Código est", "size"), tipo=("_tipo", "first")
    )
    per_student["tipo"] = per_student["tipo"].fillna("Unspecified")
    per_student["bucket"] = per_student["count"].clip(upper=3)
    pivot = per_student.groupby(["tipo", "bucket"]).size().unstack(fill_value=0)
    for b in [1, 2, 3]:
        if b not in pivot.columns:
            pivot[b] = 0
    pivot = pivot[[1, 2, 3]].sort_index()
    fig2 = go.Figure()
    for b, color, name in [(1, BLUE, "1 course"), (2, GREEN, "2 courses"), (3, PINK, "3 courses")]:
        fig2.add_trace(go.Bar(x=pivot.index, y=pivot[b], name=name, marker=dict(color=color)))
    fig2.update_layout(margin=dict(t=6, r=16, b=70, l=46), barmode="stack",
                        legend=dict(orientation="h", x=0, y=1.14), plot_bgcolor="rgba(0,0,0,0)",
                        paper_bgcolor="rgba(0,0,0,0)", height=280)
    st.plotly_chart(fig2, use_container_width=True)


# ── 12) PÁGINA — Financial ───────────────────────────────────────────────
def _gastos_col(row, name):
    norm = lambda s: re.sub(r"\s+", " ", str(s)).strip()
    target = norm(name)
    for k in row.index:
        if norm(k) == target:
            return row[k]
    return None


def page_financial():
    df = load_gastos()
    _render_header("Financial Balance", "Budget vs. actual spend, from BD_gastos.xlsx.")

    budget_col = next((c for c in df.columns if "Valor COP" in c), None)
    actual_col = next((c for c in df.columns if "Total Ra" in c), None)

    df["_budget"] = pd.to_numeric(df[budget_col], errors="coerce")
    df["_actual"] = pd.to_numeric(df[actual_col], errors="coerce")
    df["_concepto"] = df["Concepto"].astype(str).str.replace(r"\r?\n", " ", regex=True).str.strip()

    total_budget = df["_budget"].fillna(0).sum()
    reconciled = df.dropna(subset=["_actual"])
    total_executed = reconciled["_actual"].sum()
    diff = total_executed - total_budget

    c1, c2, c3 = st.columns(3)
    with c1:
        _kpi("Total Budget", _fmt_cop(total_budget), "accent")
    with c2:
        _kpi("Total Executed", _fmt_cop(total_executed), "accent")
    with c3:
        _kpi("Difference", ("+" if diff >= 0 else "") + _fmt_cop(diff), "accent")

    st.markdown("##### Budget by Concept")
    by_concept = df.groupby("_concepto")["_budget"].sum().sort_values(ascending=False)
    by_concept = by_concept[by_concept > 0]
    fig = go.Figure(go.Bar(
        x=by_concept.values, y=by_concept.index, orientation="h", marker=dict(color=BLUE),
        text=[_fmt_cop(v) for v in by_concept.values], textposition="outside",
    ))
    fig.update_layout(margin=dict(t=6, r=100, b=36, l=230), xaxis=dict(type="log", gridcolor="#EFEBEE"),
                       yaxis=dict(autorange="reversed"), plot_bgcolor="rgba(0,0,0,0)",
                       paper_bgcolor="rgba(0,0,0,0)", height=max(300, 26 * len(by_concept)))
    st.plotly_chart(fig, use_container_width=True)

    if not reconciled.empty:
        st.markdown(f"##### Budget vs. Actual — by Concept ({len(reconciled)} Executed Items)")
        by_c2 = reconciled.groupby("_concepto").agg(budget=("_budget", "sum"), actual=("_actual", "sum"))
        by_c2 = by_c2.sort_values("budget", ascending=False)
        fig2 = go.Figure()
        fig2.add_trace(go.Bar(x=by_c2.index, y=by_c2["budget"], name="Budget", marker=dict(color=PURPLE)))
        fig2.add_trace(go.Bar(x=by_c2.index, y=by_c2["actual"], name="Actual", marker=dict(color=BLUE)))
        fig2.update_layout(margin=dict(t=20, r=16, b=100, l=60), barmode="group",
                            legend=dict(orientation="h", x=0, y=1.14),
                            yaxis=dict(type="log", gridcolor="#EFEBEE"),
                            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", height=340)
        st.plotly_chart(fig2, use_container_width=True)
    else:
        st.info("Budget vs. Actual — no items executed yet.")

    st.markdown("##### Detail")
    desc_col = next((c for c in df.columns if "Profesor" in c and "Descrip" in c), None)
    disp = pd.DataFrame({
        "Concept": df["_concepto"], "Description": df[desc_col] if desc_col else "",
        "Provider": df.get("Proveedor", ""), "Budget": df["_budget"].apply(_fmt_cop),
        "Actual": df["_actual"].apply(_fmt_cop),
        "Status": np.where(df["_actual"].notna(), "executed", "pending"),
    }).sort_values("Concept")
    st.dataframe(disp, use_container_width=True, hide_index=True, height=420)
    st.download_button("Download as Excel", data=_xlsx_bytes(disp, "Financial_Detail"),
                        file_name="Financial_Detail.xlsx", key="dl_financial")


# ── 13) PÁGINA — Visiting Faculty ─────────────────────────────────────────
def page_visiting():
    d = compute_all()
    cmap = d["cmap"]

    with st.sidebar:
        st.markdown("#### Scope")
        bloque_opts = ["All"] + [str(b) for b in d["bloques"]]
        bloque_sel = st.selectbox("Block", bloque_opts, key="vf_bloque")
        bloque = "all" if bloque_sel == "All" else bloque_sel
        cursos_in_scope = sorted(
            [c for c, m in cmap.items() if bloque == "all" or str(m["ciclo"]) == str(bloque)]
        )
        curso_sel = st.selectbox("Course (selects one faculty member)", ["All"] + cursos_in_scope, key="vf_curso")
        curso = "all" if curso_sel == "All" else curso_sel

    _render_header("Evaluation & Feedback", "Course evaluation and faculty satisfaction survey — pick a course for a full profile.")

    df_cursos = d["df_cursos"]
    entries = df_cursos.to_dict("records")

    if curso != "all":
        matches = [e for e in entries if e.get("Curso") == curso]
        if matches:
            idx_key = f"vf_prof_idx__{curso}"
            idx = st.session_state.get(idx_key, 0) % len(matches)
            _render_professor_detail(matches[idx], entries, siblings=matches, sib_idx=idx, sib_key=idx_key)
            return

    # ---- Aggregate view ----
    scope_entries = entries if bloque == "all" else [e for e in entries if str(e.get("Ciclo")) == str(bloque)]
    # Deduplicar por nombre NORMALIZADO (no crudo): si 'Profesor' en BD_cursos
    # trae el mismo profesor con distinta mayúscula/espacio en dos filas (p.
    # ej. por un curso co-dictado), un set() sobre el valor crudo los trataría
    # como 2 personas distintas y sus datos se sumarían dos veces en el
    # agregado (Students Responded, Avg. Satisfaction, etc.).
    seen_prof_norm: Dict[str, str] = {}
    for e in scope_entries:
        p = e.get("Profesor")
        if not p:
            continue
        seen_prof_norm.setdefault(_norm_name(p), p)
    prof_names = list(seen_prof_norm.values())

    st.markdown("### Course Evaluation")
    resp_total = insc_total = 0
    avg_w = avg_wt = nps_w = nps_wt = obj_w = obj_wt = 0.0
    eval_by_course = []
    for name in prof_names:
        ed = build_professor_eval_data(name)
        if not ed:
            continue
        resp_total += ed["respondents"] or 0
        insc_total += ed["inscritos"] or 0
        if ed["avg"] is not None:
            avg_w += ed["avg"] * (ed["respondents"] or 1)
            avg_wt += (ed["respondents"] or 1)
        if ed["npsN"]:
            nps_w += ed["nps"] * ed["npsN"]
            nps_wt += ed["npsN"]
        if ed["objN"]:
            obj_w += ed["objectivesAllPct"] * ed["objN"]
            obj_wt += ed["objN"]
        entry = next((e for e in scope_entries if e.get("Profesor") == name), None)
        curso_label = entry["Curso"] if entry else name
        if ed["avg"] is not None or ed["npsN"]:
            eval_by_course.append({"curso": curso_label, "profesor": name, "avg": ed["avg"],
                                    "nps": ed["nps"] if ed["npsN"] else None})

    combined_avg = (avg_w / avg_wt) if avg_wt else None
    combined_nps = (nps_w / nps_wt) if nps_wt else None
    combined_obj = (obj_w / obj_wt) if obj_wt else None
    df_sat = load_satisfaccion()

    c1, c2, c_btn = st.columns([1, 1, 1])
    with c1:
        _kpi("Students Responded", f"{resp_total} / {insc_total}" if insc_total else str(resp_total), "accent")
    with c2:
        _kpi("Response Rate", f"{resp_total/insc_total*100:.1f}%" if insc_total else "—", "accent")
    with c_btn:
        st.caption(f"One PDF per faculty member ({len(prof_names)}), as a ZIP.")
        # Se genera de una sola vez (con caché por alcance) y se muestra
        # directo el botón de descarga -- antes había que darle click a
        # "Generate" y LUEGO a "Download" (dos pasos); ahora es un solo click.
        scope_key = f"{bloque}|{'|'.join(sorted(prof_names))}"
        cache = st.session_state.get("_eiv_bulk_pdf_cache", {})
        if cache.get("key") != scope_key:
            try:
                import zipfile
                zbuf = io.BytesIO()
                with st.spinner(f"Preparando {len(prof_names)} reportes…"), zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as zf:
                    for name in prof_names:
                        ed_name = build_professor_eval_data(name)
                        if not ed_name:
                            continue
                        entry_name = next((e for e in scope_entries if e.get("Profesor") == name), None)
                        curso_name = entry_name["Curso"] if entry_name else ""
                        row_name = _find_satisfaction_row(df_sat, curso_name, name) if entry_name else None
                        pdf_bytes = _build_professor_pdf(name, curso_name, ed_name, row_name or {})
                        fname = f"ISS_{name.replace(' ', '_')}_Evaluation_Report.pdf"
                        zf.writestr(fname, pdf_bytes)
                zbuf.seek(0)
                st.session_state["_eiv_bulk_pdf_cache"] = {"key": scope_key, "data": zbuf.getvalue()}
            except ModuleNotFoundError:
                st.session_state["_eiv_bulk_pdf_cache"] = {"key": scope_key, "data": None}
        cache = st.session_state.get("_eiv_bulk_pdf_cache", {})
        if cache.get("data"):
            with st.container(key="bulk_pdf_download"):
                st.download_button(
                    "Download PDF Reports (ZIP)", data=cache["data"],
                    file_name="ISS_2026_Evaluation_Reports.zip", mime="application/zip", key="dl_bulk_pdf_zip",
                    icon=":material/download:", use_container_width=True,
                )
    c3, c4, c5 = st.columns(3)
    with c3:
        _kpi("Avg. Satisfaction", f"{combined_avg:.2f} / 5" if combined_avg is not None else "—", "accent")
    with c4:
        _kpi("Avg. Recommendation", f"{combined_nps:.1f} / 10" if combined_nps is not None else "—", "accent")
    with c5:
        _kpi("Objectives Met (All)", f"{combined_obj:.1f}%" if combined_obj is not None else "—", "accent")

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("##### Ranked by Satisfaction")
        ranked = sorted([c for c in eval_by_course if c["avg"] is not None], key=lambda c: -c["avg"])
        if ranked:
            fig = go.Figure(go.Bar(
                x=[c["avg"] for c in ranked],
                y=[f'{c["profesor"]} — {c["curso"][:28]}' for c in ranked], orientation="h",
                marker=dict(color=[color_scale((c["avg"] - 4) / 1 * 100) for c in ranked]),
                text=[f'{c["avg"]:.2f}' for c in ranked], textposition="outside",
            ))
            fig.update_layout(margin=dict(t=6, r=36, b=26, l=240), xaxis=dict(range=[4, 5]),
                               yaxis=dict(autorange="reversed"), plot_bgcolor="rgba(0,0,0,0)",
                               paper_bgcolor="rgba(0,0,0,0)", height=max(200, 28 * len(ranked)))
            st.plotly_chart(fig, use_container_width=True)
    with col_b:
        st.markdown("##### Ranked by Recommendation (NPS)")
        ranked_n = sorted([c for c in eval_by_course if c["nps"] is not None], key=lambda c: -c["nps"])
        if ranked_n:
            fig = go.Figure(go.Bar(
                x=[c["nps"] for c in ranked_n],
                y=[f'{c["profesor"]} — {c["curso"][:28]}' for c in ranked_n], orientation="h",
                marker=dict(color=[color_scale((c["nps"] - 7) / 3 * 100) for c in ranked_n]),
                text=[f'{c["nps"]:.1f}' for c in ranked_n], textposition="outside",
            ))
            fig.update_layout(margin=dict(t=6, r=36, b=26, l=240), xaxis=dict(range=[7, 10]),
                               yaxis=dict(autorange="reversed"), plot_bgcolor="rgba(0,0,0,0)",
                               paper_bgcolor="rgba(0,0,0,0)", height=max(200, 28 * len(ranked_n)))
            st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")
    st.markdown("### Faculty Satisfaction")
    df_sat = load_satisfaccion()
    sat_rows = []
    for e in scope_entries:
        row = _find_satisfaction_row(df_sat, e["Curso"], e["Profesor"])
        if row is not None:
            sat_rows.append(row)

    overall_col = next((c for c in df_sat.columns if "overall experience" in c.lower()), None)
    matched_col = next((c for c in df_sat.columns if "match your expectations" in c.lower()), None)
    ov_vals = [r[overall_col] for r in sat_rows if overall_col and isinstance(r.get(overall_col), (int, float))]
    mt_vals = [r[matched_col] for r in sat_rows if matched_col and isinstance(r.get(matched_col), (int, float))]

    c1, c2, c3 = st.columns(3)
    with c1:
        _kpi("Faculty Satisfaction Responses", f"{len(sat_rows)} / {len(scope_entries)}", "accent")
    with c2:
        _kpi("Avg. Overall Experience", f"{np.mean(ov_vals):.2f} / 5" if ov_vals else "—", "accent")
    with c3:
        _kpi("Avg. Matched Expectations", f"{np.mean(mt_vals):.2f} / 5" if mt_vals else "—", "accent")

    if sat_rows:
        col_radar, col_bars = st.columns(2)
        with col_radar:
            st.markdown("##### Overall Ratings (Radar)")
            axes = [(a, v) for a, v in _radar_axes_multi(sat_rows) if v is not None]
            if len(axes) >= 3:
                r_vals = [v for _, v in axes] + [axes[0][1]]
                theta_vals = [a for a, _ in axes] + [axes[0][0]]
                fig = go.Figure(go.Scatterpolar(
                    r=r_vals, theta=theta_vals,
                    fill="toself", fillcolor="rgba(230,17,102,0.18)", line=dict(color=PINK, width=2),
                    mode="lines+markers+text", marker=dict(size=6, color=PINK),
                    text=[f"{v:.2f}" for v in r_vals], textposition="top center",
                    textfont=dict(size=11, color=INK),
                ))
                fig.update_layout(margin=dict(t=20, r=50, b=20, l=50),
                                   polar=dict(radialaxis=dict(visible=True, range=[0, 5], tickfont=dict(size=9))),
                                   showlegend=False, paper_bgcolor="rgba(0,0,0,0)", height=340)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Not enough satisfaction categories with data to draw the radar for this scope.")
        with col_bars:
            st.markdown("##### Student Performance")
            perf = [(label, _avg_multi(sat_rows, [q])) for label, q in STUDENT_PERF_ITEMS]
            perf = [(l, v) for l, v in perf if v is not None]
            if perf:
                fig_b = go.Figure(go.Bar(
                    x=[v for _, v in perf], y=[l for l, _ in perf], orientation="h",
                    marker=dict(color=[color_scale((v - 1) / 4 * 100) for _, v in perf]),
                    text=[f"{v:.2f}" for _, v in perf], textposition="outside",
                ))
                fig_b.update_layout(margin=dict(t=10, r=36, b=26, l=140), xaxis=dict(range=[1, 5.5]),
                                     yaxis=dict(autorange="reversed"), plot_bgcolor="rgba(0,0,0,0)",
                                     paper_bgcolor="rgba(0,0,0,0)", height=320)
                st.plotly_chart(fig_b, use_container_width=True)

        st.markdown("##### Important Insights")
        cards = [
            _insight_card("Enrollment Size", sat_rows, "Regarding the number of enrolled students, do you consider that it was"),
            _insight_card("Office Hours Use", sat_rows, "How frequently did students used your office hours? (Frequently/Occasionally/Rarely)"),
            _insight_card("UG vs. Grad Differences", sat_rows, "Did you perceive any differences between undergraduate and graduate students performance  in the different  course activities?"),
        ]
        _render_insight_cards(cards)

        st.markdown("##### Comments")
        prof_col_for_sat = next((c for c in df_sat.columns if re.sub(r"\s+", " ", str(c)).strip().casefold() == "full name"), None)
        title_col_for_sat = next((c for c in df_sat.columns if re.sub(r"\s+", " ", str(c)).strip().casefold() == "course title"), None)
        comment_fields = [
            ("What They Valued Most", "Please mention two aspects that you valued the most about your course delivery"),
            ("What They Would Change", "Please mention two aspects that you would change if delivering this course again"),
            ("Student Academic Impression", "Please provide the general impressions on the academic level and performance of the students in your course"),
            ("General Comments", "General comments"),
            ("Changes for Next Year", "What changes would you introduce to make the International Summer School even better?"),
        ]
        ccols = st.columns(2)
        for i, (label, question) in enumerate(comment_fields):
            with ccols[i % 2]:
                items = []
                for row in sat_rows:
                    val = _fcol(row, question)
                    if val is None or (isinstance(val, float) and pd.isna(val)) or not str(val).strip():
                        continue
                    who = row.get(prof_col_for_sat) if prof_col_for_sat else None
                    what = row.get(title_col_for_sat) if title_col_for_sat else None
                    items.append((who, what, str(val).strip()))
                st.markdown(f"**{label} ({len(items)})**")
                if items:
                    with st.container(height=200):
                        for who, what, txt in items:
                            attribution = f"{who} — {what}" if who else (what or "")
                            st.markdown(
                                f'<div class="prof-quote"><strong>{attribution}</strong>“{txt}”</div>',
                                unsafe_allow_html=True,
                            )
                else:
                    st.caption("No comments for this question in this scope.")


def _fcol(row: dict, target: str, default=None):
    """Búsqueda de columna normalizada (colapsa espacios/NBSP a un espacio
    simple, y normaliza Unicode a NFC) — BD_satisfaccionEIV trae algunas
    cabeceras con \\xa0 (espacio duro) en lugares distintos, y algunas
    tildes ('Bogotá', 'Neón') pueden llegar en forma NFD (á = 'a' + acento
    combinante) en vez de NFC (á precompuesta); ambos casos hacen que una
    búsqueda exacta por clave falle en silencio, aunque el texto se vea
    idéntico -- esto era lo que dejaba 3 de los 6 ejes del radar en None."""
    norm = lambda s: unicodedata.normalize("NFC", re.sub(r"\s+", " ", str(s)).strip())
    t = norm(target)
    for k in row.keys():
        if norm(k) == t:
            v = row[k]
            return default if (v is None or (isinstance(v, float) and pd.isna(v))) else v
    return default


def _find_satisfaction_row(df_sat: pd.DataFrame, curso: str, profesor: str) -> Optional[dict]:
    def _resolve_header(df, target):
        t = unicodedata.normalize("NFC", re.sub(r"\s+", " ", target).strip().casefold())
        for k in df.columns:
            if unicodedata.normalize("NFC", re.sub(r"\s+", " ", str(k)).strip().casefold()) == t:
                return k
        return None

    title_col = _resolve_header(df_sat, "Course title")
    if not title_col:
        return None
    curso_norm = unicodedata.normalize("NFC", re.sub(r"\s+", " ", str(curso)).strip().casefold())
    title_norm = df_sat[title_col].astype(str).str.replace(r"\s+", " ", regex=True).str.strip().str.casefold().apply(
        lambda s: unicodedata.normalize("NFC", s))
    matches = df_sat[title_norm == curso_norm]
    if len(matches) == 0:
        return None
    if len(matches) == 1:
        return matches.iloc[0].to_dict()
    name_col = _resolve_header(df_sat, "Full name")
    prof_words = [w for w in profesor.lower().split() if len(w) > 2]
    best, best_score = matches.iloc[0].to_dict(), -1
    for _, m in matches.iterrows():
        full = str(m.get(name_col, "")).lower() if name_col else ""
        score = sum(1 for w in prof_words if w in full)
        if score > best_score:
            best_score, best = score, m.to_dict()
    return best


def _radar_axes(row: dict) -> List[Tuple[str, Optional[float]]]:
    def avg(keys):
        vals = [_fcol(row, k) for k in keys if isinstance(_fcol(row, k), (int, float))]
        return (sum(vals) / len(vals)) if vals else None

    return [
        ("IO Staff Support", avg([
            "Please rate the support provided by the International Office staff in the following items (1: Poor – 5: Excellent): General assistance provided by the International Office staff",
            "Effectiveness of the hiring process (documentation and payment procedures)",
            "Agility in scheduling your inbound/outbound flights",
            "Agility in arranging your accommodation in Bogotá",
            "Clarity of the virtual learning platform instructions and access credentials"])),
        ("Student Performance", avg([
            "Please rate the following aspects (1: Poor – 5: Excellent): Students English proficiency",
            "Student punctuality", "Students analytical skills"])),
        ("Tech & Platform Support", avg([
            "Please rate the following aspects (1: Poor – 5: Excellent): Virtual learning platform (Bloque Neón)",
            "Support provided by DSIT (Technology Office)", "Support provided by CSL (Logistics Service Center)"])),
        ("Teaching Assistant", avg([
            "Please rate the following logistical aspects during your visit  (1: Poor – 5: Excellent): English proficiency of the teaching assistant",
            "Proactivity of the teaching assistant", "Punctuality of the teaching assistant",
            "Professionalism and attitude of the teaching assistant", "Responsiveness to students’ questions and needs"])),
        ("Visit Logistics", avg([
            "Please rate the following logistical aspects during your visit  (1: Poor – 5: Excellent)2: Organization of the welcome lunch",
            "Punctuality and reliability of the daily transportation provided", "Quality of accommodation at BH La Quinta"])),
        ("Overall Experience", _fcol(row, "Please rate your overall experience during this year’s International Summer School  (1: Not at all – 5: Exceeded expectations)")),
    ]


STUDENT_PERF_ITEMS = [
    ("English Proficiency", "Please rate the following aspects (1: Poor – 5: Excellent): Students English proficiency"),
    ("Punctuality", "Student punctuality"),
    ("Analytical Skills", "Students analytical skills"),
]


def _avg_multi(rows: List[dict], keys: List[str]) -> Optional[float]:
    vals = [_fcol(r, k) for r in rows for k in keys if isinstance(_fcol(r, k), (int, float))]
    return (sum(vals) / len(vals)) if vals else None


def _radar_axes_multi(rows: List[dict]) -> List[Tuple[str, Optional[float]]]:
    """Igual que _radar_axes(), pero promediando sobre varias filas (vista
    agregada de Faculty Satisfaction, en vez de un solo profesor)."""
    return [
        ("IO Staff Support", _avg_multi(rows, [
            "Please rate the support provided by the International Office staff in the following items (1: Poor – 5: Excellent): General assistance provided by the International Office staff",
            "Effectiveness of the hiring process (documentation and payment procedures)",
            "Agility in scheduling your inbound/outbound flights",
            "Agility in arranging your accommodation in Bogotá",
            "Clarity of the virtual learning platform instructions and access credentials"])),
        ("Student Performance", _avg_multi(rows, [
            "Please rate the following aspects (1: Poor – 5: Excellent): Students English proficiency",
            "Student punctuality", "Students analytical skills"])),
        ("Tech & Platform Support", _avg_multi(rows, [
            "Please rate the following aspects (1: Poor – 5: Excellent): Virtual learning platform (Bloque Neón)",
            "Support provided by DSIT (Technology Office)", "Support provided by CSL (Logistics Service Center)"])),
        ("Teaching Assistant", _avg_multi(rows, [
            "Please rate the following logistical aspects during your visit  (1: Poor – 5: Excellent): English proficiency of the teaching assistant",
            "Proactivity of the teaching assistant", "Punctuality of the teaching assistant",
            "Professionalism and attitude of the teaching assistant", "Responsiveness to students’ questions and needs"])),
        ("Visit Logistics", _avg_multi(rows, [
            "Please rate the following logistical aspects during your visit  (1: Poor – 5: Excellent)2: Organization of the welcome lunch",
            "Punctuality and reliability of the daily transportation provided", "Quality of accommodation at BH La Quinta"])),
        ("Overall Experience", _avg_multi(rows, [
            "Please rate your overall experience during this year’s International Summer School  (1: Not at all – 5: Exceeded expectations)"])),
    ]


def _insight_card(label: str, rows: List[dict], question: str) -> Tuple[str, str, str]:
    """Respuesta más común a una pregunta categórica + % y conteo, para las
    tarjetas 'Important Insights'."""
    vals = [str(_fcol(r, question)).strip() for r in rows if _fcol(r, question) not in (None, "")]
    vals = [v for v in vals if v and v.lower() != "nan"]
    if not vals:
        return label, "—", ""
    counts = pd.Series(vals).value_counts()
    top, n = counts.index[0], counts.iloc[0]
    pct = n / len(vals) * 100
    return label, top, f"{pct:.1f}% of respondents ({n}/{len(vals)})"


def _render_insight_cards(cards: List[Tuple[str, str, str]]):
    cols = st.columns(len(cards))
    for col, (label, value, sub) in zip(cols, cards):
        with col:
            st.markdown(
                f'<div style="background:#F6F8FB;border-radius:10px;padding:14px 16px;">'
                f'<div style="font-family:monospace;font-size:10.5px;letter-spacing:.06em;'
                f'text-transform:uppercase;color:{PINK};font-weight:700;">{label}</div>'
                f'<div style="font-size:19px;font-weight:800;color:{INK};margin:4px 0 2px;">{value}</div>'
                f'<div style="font-size:12px;color:#8C7F87;">{sub}</div></div>',
                unsafe_allow_html=True,
            )


def _photo_path(profesor: str) -> str:
    fname = PHOTO_FILENAME_OVERRIDES.get(profesor, profesor.title())
    base = _os.path.dirname(_os.path.abspath(__file__))
    for candidate in (f"PICS/PROFES/2026/{fname}.jpg", f"PICS/PROFES/{fname}.jpg"):
        full = _os.path.join(base, candidate)
        if _os.path.exists(full):
            return full
    return _os.path.join(base, f"PICS/PROFES/2026/{fname}.jpg")


def _render_professor_detail(entry: dict, entries: List[dict], siblings: Optional[list] = None,
                              sib_idx: int = 0, sib_key: str = ""):
    d = compute_all()
    curso = entry["Curso"]
    profesor = entry["Profesor"]
    inscritos = d["inscritos_por_curso"].get(curso, 0)
    capacidad = entry.get("Capacidad") or 0
    ocupacion = f"{inscritos/capacidad*100:.1f}%" if capacidad else "—"

    co_entries = [e for e in entries if e.get("Curso") == curso]

    if siblings and len(siblings) > 1:
        col_prev, col_mid_nav, col_next = st.columns([1, 8, 1])
        with col_prev:
            if st.button("‹", key=f"{sib_key}_prev", help="Previous faculty on this course", use_container_width=True):
                st.session_state[sib_key] = (sib_idx - 1) % len(siblings)
                st.rerun()
        with col_next:
            if st.button("›", key=f"{sib_key}_next", help="Next faculty on this course", use_container_width=True):
                st.session_state[sib_key] = (sib_idx + 1) % len(siblings)
                st.rerun()

    col_photo, col_info, col_pdf = st.columns([1, 2.4, 1])
    with col_photo:
        try:
            st.image(_photo_path(profesor), width=220)
        except Exception:
            pass
    with col_info:
        st.markdown(f"### {profesor}")
        st.markdown(f"**University:** {entry.get('Universidad','—')} · {entry.get('País Universidad','—')}")
        st.markdown(f"**Course:** {curso}")
        st.markdown(f"**Block / Dates:** Block {entry.get('Ciclo','—')} ({entry.get('Fechas curso','—')})")
        st.markdown(f"**Modality:** {entry.get('Modalidad','—')}")
        st.markdown(f"**Enrolled / Capacity:** {inscritos} / {capacidad} ({ocupacion} occupancy)")
        if len(co_entries) > 1:
            st.caption(f"Co-teaching faculty: {len(co_entries)} on this course.")

    ed = build_professor_eval_data(profesor)
    df_sat = load_satisfaccion()
    row = _find_satisfaction_row(df_sat, curso, profesor)

    with col_pdf:
        st.markdown("<div style='height:14px;'></div>", unsafe_allow_html=True)
        if ed:
            try:
                pdf_key = f"{profesor}|{curso}"
                cache = st.session_state.get("_eiv_prof_pdf_cache", {})
                if cache.get("key") != pdf_key:
                    pdf_bytes = _build_professor_pdf(profesor, curso, ed, row if row else {})
                    st.session_state["_eiv_prof_pdf_cache"] = {"key": pdf_key, "data": pdf_bytes}
                    cache = st.session_state["_eiv_prof_pdf_cache"]
                with st.container(key=f"pdf_btn_{profesor}"):
                    st.download_button(
                        "Evaluation Report (PDF)", data=cache["data"],
                        file_name=f"EIV_{profesor.replace(' ', '_')}_Evaluation_Report.pdf",
                        mime="application/pdf", key=f"dl_pdf_{profesor}", use_container_width=True,
                        icon=":material/picture_as_pdf:",
                    )
            except ModuleNotFoundError:
                pass

    st.markdown("---")
    st.markdown("### Course Evaluation")
    if not ed:
        st.info("No course evaluation data found for this professor.")
    else:
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1:
            _kpi("Students Responded", ed["respondents"], "accent")
        with c2:
            rr = (ed["respondents"] / ed["inscritos"] * 100) if ed["inscritos"] else None
            _kpi("Response Rate", f"{rr:.1f}%" if rr is not None else "—", "accent")
        with c3:
            _kpi("Avg. Satisfaction", f"{ed['avg']:.2f} / 5" if ed["avg"] is not None else "—", "accent")
        with c4:
            _kpi("Would Recommend", f"{ed['nps']:.1f} / 10" if ed["npsN"] else "—", "accent")
        with c5:
            _kpi("Objectives Met", f"{ed['objectivesAllPct']:.1f}%" if ed["objN"] else "—", "accent")

        if ed["aspects"]:
            st.markdown("##### By Aspect")
            fig = go.Figure(go.Bar(
                x=[a["avg"] for a in ed["aspects"]], y=[a["aspect"] for a in ed["aspects"]], orientation="h",
                marker=dict(color=PINK), text=[f'{a["avg"]:.2f}' for a in ed["aspects"]], textposition="outside",
            ))
            fig.update_layout(margin=dict(t=10, r=50, b=26, l=270), xaxis=dict(range=[1, 5.5]),
                               yaxis=dict(autorange="reversed"), plot_bgcolor="rgba(0,0,0,0)",
                               paper_bgcolor="rgba(0,0,0,0)", height=280)
            st.plotly_chart(fig, use_container_width=True)

        st.markdown("##### Comments")
        com_by_prof = comments_by_professor()
        prof_comments = com_by_prof.get(_resolve_alias(profesor), {})
        ccols = st.columns(2)
        all_groups = {**POSITIVE_QIDS, **IMPROVE_QIDS}
        for i, (qid, label) in enumerate(all_groups.items()):
            with ccols[i % 2]:
                group = prof_comments.get(qid, {"items": []})
                exact_q = group.get("text") or label
                sorted_items = sorted(group["items"], key=lambda it: {"202613": 0, "202618": 1}.get(it.get("periodo"), 2))
                st.markdown(f"**{exact_q} ({len(sorted_items)})**")
                if sorted_items:
                    with st.container(height=200):
                        for item in sorted_items:
                            tag = f'<strong style="color:{PINK};">[{item["periodo"]}]</strong> ' if item.get("periodo") else ""
                            st.markdown(f'<div class="prof-quote">{tag}“{item["text"]}”</div>', unsafe_allow_html=True)
                else:
                    st.caption("No comments for this question.")

    st.markdown("---")
    st.markdown("### Faculty Satisfaction Survey")
    if not row:
        st.info("No satisfaction survey response found for this course.")
    else:
        overall = _fcol(row, "Please rate your overall experience during this year’s International Summer School  (1: Not at all – 5: Exceeded expectations)")
        matched = _fcol(row, "Did your experience match your expectations?  (1: Not at all – 5: Exceeded expectations)")
        rec_colleagues = _fcol(row, "Would you recommend the International Summer School to colleagues?")
        rec_students = _fcol(row, "Would you recommend the International Summer School to your home students?")

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            _kpi("Overall Experience", f"{overall} / 5" if overall is not None else "—", "accent")
        with c2:
            _kpi("Matched Expectations", f"{matched} / 5" if matched is not None else "—", "accent")
        with c3:
            _kpi("Would Recommend to Colleagues", rec_colleagues or "—", "accent")
        with c4:
            _kpi("Would Recommend to Students", rec_students or "—", "accent")

        col_radar, col_fields = st.columns([1, 1])
        with col_radar:
            st.markdown("##### Overall Ratings (Radar)")
            axes = [(a, v) for a, v in _radar_axes(row) if v is not None]
            if len(axes) >= 3:
                r_vals = [v for _, v in axes] + [axes[0][1]]
                theta_vals = [a for a, _ in axes] + [axes[0][0]]
                fig = go.Figure(go.Scatterpolar(
                    r=r_vals, theta=theta_vals,
                    fill="toself", fillcolor="rgba(230,17,102,0.18)", line=dict(color=PINK, width=2),
                    mode="lines+markers+text", marker=dict(size=6, color=PINK),
                    text=[f"{v:.2f}" for v in r_vals], textposition="top center",
                    textfont=dict(size=11, color=INK),
                ))
                fig.update_layout(margin=dict(t=20, r=50, b=20, l=50),
                                   polar=dict(radialaxis=dict(visible=True, range=[0, 5], tickfont=dict(size=9))),
                                   showlegend=False, paper_bgcolor="rgba(0,0,0,0)", height=340)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Not enough satisfaction categories with data to draw the radar.")
        with col_fields:
            st.markdown(f"**Would choose the same modality again?** {_fcol(row, 'Would you prefer to deliver the course in the same modality again','—')}")
            st.markdown(f"**Enrollment size:** {_fcol(row, 'Regarding the number of enrolled students, do you consider that it was','—')}")
            st.markdown(f"**Office hours usage:** {_fcol(row, 'How frequently did students used your office hours? (Frequently/Occasionally/Rarely)','—')}")
            st.markdown(f"**Materials uploaded on time?** {_fcol(row, 'Was all course material uploaded in a timely manner?','—')}")
            st.markdown(f"**UG vs. Grad differences?** {_fcol(row, 'Did you perceive any differences between undergraduate and graduate students performance  in the different  course activities?','—')}")

        st.markdown("##### Comments & Open-Ended Responses")
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            st.markdown(_quote("What they valued most", _fcol(row, "Please mention two aspects that you valued the most about your course delivery")), unsafe_allow_html=True)
            st.markdown(_quote("What they would change", _fcol(row, "Please mention two aspects that you would change if delivering this course again")), unsafe_allow_html=True)
            st.markdown(_quote("Student academic impression", _fcol(row, "Please provide the general impressions on the academic level and performance of the students in your course")), unsafe_allow_html=True)
            st.markdown(_quote("General comments", _fcol(row, "General comments")), unsafe_allow_html=True)
            st.markdown(_quote("Changes suggested for next year", _fcol(row, "What changes would you introduce to make the International Summer School even better?")), unsafe_allow_html=True)
        with col_c2:
            st.markdown(_quote("International Office support", _fcol(row, "Comments or suggestions about the support provided by the International Office")), unsafe_allow_html=True)
            st.markdown(_quote("Tech / logistics support to improve", _fcol(row, "Are there any aspects that should be improved by the DSIT or CSL units?")), unsafe_allow_html=True)
            st.markdown(_quote("Teaching assistant support to improve", _fcol(row, "Are there any aspects that should be improved by the teaching assistants?")), unsafe_allow_html=True)
            st.markdown(_quote("Visit logistics to improve", _fcol(row, "Are there any logistical aspects that should be improved?")), unsafe_allow_html=True)
            st.markdown(_quote("General observations", _fcol(row, "General observations or suggestions")), unsafe_allow_html=True)


# ── 14) PDF — Evaluation Report per professor ────────────────────────────
def _pdf_check_page(y_mm, needed, page_h_mm, draw_header_fn):
    if y_mm + needed > page_h_mm - 16:
        draw_header_fn()
        return 24.0
    return y_mm


def _pdf_stacked_bar(c, options: dict, order: list, color_map: dict, x_mm, y_mm, w_mm, h_mm, page_h_mm):
    """Traduce EIV.Export.pdfStackedBar a reportlab -- todo se razona en mm
    (y_mm = distancia desde ARRIBA de la página, igual que el resto del
    layout), convirtiendo a puntos de reportlab solo al dibujar. Devuelve
    el nuevo y_mm (más abajo) para encadenar el siguiente elemento."""
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
        # Revisa ANTES de dibujar si el ítem cabe en lo que queda de la fila
        # -- antes se revisaba DESPUÉS de dibujar, así que una etiqueta larga
        # que arrancaba cerca del borde se salía del margen de la columna.
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


# Órdenes bilingües (español + inglés) — cada fila de BD_evaluacion_curso.xlsx
# ya trae ambos idiomas, y evaluation_by_professor() elige uno según el
# profesor, así que estos mapas cubren ambos casos; solo se dibujan las
# claves que de verdad tengan datos.
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
    """Envuelve texto según el ANCHO REAL renderizado (c.stringWidth), no
    una estimación de caracteres por línea -- esa estimación quedaba corta
    para columnas angostas y dejaba preguntas largas en una sola línea que
    se salía del margen. Se resta un pequeño margen de seguridad (2%) para
    absorber cualquier diferencia mínima de redondeo."""
    from reportlab.lib.units import mm
    max_w_pt = max_width_mm * mm * 0.98
    words = str(text).split()
    lines, cur = [], ""
    for w in words:
        trial = f"{cur} {w}".strip()
        if c.stringWidth(trial, font_name, font_size) <= max_w_pt or not cur:
            cur = trial
        else:
            lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines or [""]


def _pdf_file_bytes(path: str) -> Optional[bytes]:
    try:
        with open(path, "rb") as f:
            return f.read()
    except Exception:
        return None


def _build_professor_pdf(profesor: str, curso: str, ed: dict, sat_row: dict) -> bytes:
    """Réplica en reportlab del PDF de evaluación aprobado en el HTML
    (EIV.Export.generateSingleProfessorPDF): banda rosa con logo+foto,
    tarjetas KPI, barras apiladas con leyenda de conteos reales agrupadas
    por aspecto, barra NPS 0-10 con gradiente, barras de objetivos y carga
    académica, y comentarios paginados — mismo diseño y lógica aprobados.
    Todo el layout se razona en mm (como el jsPDF original), convirtiendo a
    puntos de reportlab solo en las llamadas de dibujo."""
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.lib import colors as rl_colors
    from reportlab.pdfgen import canvas as rl_canvas
    from reportlab.lib.utils import ImageReader
    import textwrap

    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=A4)
    PW, PH = 210.0, 297.0
    mg = 14.0
    cW = PW - 2 * mg
    pink = rl_colors.HexColor(PINK)
    ink = rl_colors.HexColor(INK)
    muted = rl_colors.HexColor("#8C7F87")
    lgray = rl_colors.HexColor("#F8F6F8")

    def top_pt(y_mm):
        return (PH - y_mm) * mm

    entry = next((e for e in load_cursos().to_dict("records") if e.get("Profesor") == profesor), None)
    info = None
    if entry:
        d0 = compute_all()
        info = {
            "universidad": entry.get("Universidad"), "pais": entry.get("País Universidad"),
            "curso": entry.get("Curso"), "modalidad": entry.get("Modalidad"), "bloque": entry.get("Ciclo"),
            "inscritos": d0["inscritos_por_curso"].get(entry.get("Curso"), 0),
        }

    # ---- Distribución de programas por periodo, para el curso de este
    # profesor -- 202613 (EPOS Admi) y 202618 (Pregrado Admin y otros
    # programas) son los dos periodos de matrícula reales de EIV. Se
    # calculan directo de BD_listas.xlsx (columnas Periodo / Nombre
    # Programa), filtrando por el mismo _curso ya resuelto (por profesor,
    # no por el nombre truncado de BD_evaluacion_curso.xlsx).
    period_program_dist: Dict[str, List[Tuple[str, int]]] = {}
    if info:
        d0 = compute_all()
        df_listas_all = d0["df_listas"]
        df_programas_all = d0["df_programas"]
        periodo_col = next((c for c in df_listas_all.columns if c.strip().casefold() == "periodo"), None)
        programa_col = next((c for c in df_listas_all.columns
                              if c.strip().casefold() in ("nombre programa", "programa principal", "programa")), None)
        # Traduce el nombre del programa (español) al nombre oficial en
        # inglés, usando la hoja 'programas' de BD_listas.xlsx (columna
        # PRORGAMA=español, PROGRAM=inglés) -- el PDF va siempre en inglés.
        prog_es_to_en: Dict[str, str] = {}
        prorgama_col = next((c for c in df_programas_all.columns if c.strip().casefold() == "prorgama"), None)
        program_col = next((c for c in df_programas_all.columns if c.strip().casefold() == "program"), None)
        if prorgama_col and program_col:
            for _, pr in df_programas_all.iterrows():
                es = pr.get(prorgama_col)
                en = pr.get(program_col)
                if pd.notna(es) and pd.notna(en):
                    prog_es_to_en[str(es).strip().casefold()] = str(en).strip()
        if periodo_col and programa_col:
            course_rows = df_listas_all[df_listas_all["_curso"] == info["curso"]]
            for periodo_val in ("202613", "202618"):
                sub = course_rows[course_rows[periodo_col].astype(str).str.strip() == periodo_val]
                if sub.empty:
                    continue
                counts = sub[programa_col].dropna().astype(str).str.strip().value_counts()
                translated = [
                    (prog_es_to_en.get(name.casefold(), name), n) for name, n in counts.items()
                ]
                period_program_dist[periodo_val] = translated

    logo_bytes = None
    try:
        _base = _os.path.dirname(_os.path.abspath(__file__))
        for rel in _LOGO_CANDIDATES:
            full = _os.path.join(_base, rel)
            if _os.path.exists(full):
                logo_bytes = _pdf_file_bytes(full)
                break
    except Exception:
        logo_bytes = None
    photo_bytes = _pdf_file_bytes(_photo_path(profesor))

    def draw_stats_header():
        band_h = 34
        c.setFillColor(pink)
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
        c.drawRightString(name_right_edge * mm, top_pt(11), profesor)
        c.setFont("Helvetica", 8.5)
        c.drawRightString(name_right_edge * mm, top_pt(17), info["universidad"] if info else "")
        c.setFont("Helvetica", 8)
        c.drawRightString(name_right_edge * mm, top_pt(22.5), info["pais"] if info else "")
        c.setFont("Helvetica-Bold", 12)
        c.drawString(mg * mm, top_pt(29.5), "COURSE EVALUATION")

    def draw_continuation_header():
        c.setFillColor(pink)
        c.rect(0, top_pt(14), PW * mm, 14 * mm, fill=1, stroke=0)
        c.setFillColor(rl_colors.white)
        c.setFont("Helvetica-Bold", 10)
        c.drawRightString((PW - mg) * mm, top_pt(9), profesor)

    def draw_footer(page_num):
        c.setFillColor(lgray)
        c.rect(0, 0, PW * mm, 10 * mm, fill=1, stroke=0)
        c.setFont("Helvetica", 7)
        c.setFillColor(muted)
        c.drawString(12 * mm, 4 * mm, "International Summer School 2026 — Course Evaluation")
        c.drawRightString((PW - 12) * mm, 4 * mm, str(page_num))

    page_num = 1
    PERIOD_LABELS = {"202613": "GR", "202618": "UG"}
    PERIOD_TOTAL_LABELS = {"202613": "GR — 202613", "202618": "UG — 202618"}

    # ---- Página 1: banda + título de curso centrado + matrícula por
    # periodo (GR/UG) con su gráfica de programas debajo -- los KPIs de
    # respuesta van DESPUÉS de esto, no antes ----
    draw_stats_header()

    def _estimate_page1_height():
        """'Ensayo en seco' del mismo bloque de la página 1 (sin dibujar
        nada), solo para saber cuánta altura total ocupa y poder centrarlo
        verticalmente entre la banda superior y el pie de página."""
        h = 0.0
        course_lines_est = _wrap_to_width(c, info["curso"] if info else (curso or ""), "Helvetica-Bold", 17, cW)
        h += len(course_lines_est) * 7 + 3
        if info:
            h += 9
        if period_program_dist:
            pkeys_est = list(period_program_dist.keys())
            bw2_est = cW / max(len(pkeys_est), 1)
            h += 23 + 8
            col_heights_est = []
            for pk in pkeys_est:
                pairs = period_program_dist[pk]
                yy = 0.0
                for prog_name, _n in pairs:
                    label_lines_est = _wrap_to_width(c, prog_name, "Helvetica", 6.5, bw2_est - 2)
                    yy += len(label_lines_est) * 3.2 + 1 + 6.5
                col_heights_est.append(yy)
            h += max(col_heights_est, default=0) + 6
        h += 32  # tarjetas KPI
        return h

    header_bottom_y = 34.0
    footer_top_y = PH - 10.0
    available_h = footer_top_y - header_bottom_y
    total_h = _estimate_page1_height()
    y = header_bottom_y + max(2.0, (available_h - total_h) / 2)
    c.setFillColor(pink)
    c.setFont("Helvetica-Bold", 17)
    course_lines = _wrap_to_width(c, info["curso"] if info else (curso or ""), "Helvetica-Bold", 17, cW)
    for i, line in enumerate(course_lines):
        c.drawCentredString((mg + cW / 2) * mm, top_pt(y + i * 7), line)
    y += len(course_lines) * 7 + 3
    if info:
        c.setFillColor(muted)
        c.setFont("Helvetica", 8.5)
        c.drawCentredString((mg + cW / 2) * mm, top_pt(y), f"Block {info['bloque']}  ·  {info['modalidad']}")
        y += 9

    if period_program_dist:
        pkeys = list(period_program_dist.keys())
        bw2 = cW / max(len(pkeys), 1)
        periods_for_resp = ed.get("periods", {})
        for i, pk in enumerate(pkeys):
            total_p = sum(n for _, n in period_program_dist[pk])
            responded_p = periods_for_resp.get(pk, {}).get("respondents")
            bx = mg + i * bw2
            c.setFillColor(rl_colors.HexColor("#FBE3ED"))
            c.roundRect((bx + 1) * mm, top_pt(y + 19), (bw2 - 3) * mm, 19 * mm, 3 * mm, fill=1, stroke=0)
            c.setFillColor(pink)
            c.setFont("Helvetica-Bold", 13)
            c.drawCentredString((bx + bw2 / 2 - 1) * mm, top_pt(y + 8.5), f"{total_p} enrolled")
            c.setFillColor(ink)
            c.setFont("Helvetica", 6.5)
            c.drawCentredString((bx + bw2 / 2 - 1) * mm, top_pt(y + 14),
                                 PERIOD_TOTAL_LABELS.get(pk, pk))
            if responded_p is not None:
                c.setFillColor(muted)
                c.setFont("Helvetica-Oblique", 6)
                c.drawCentredString((bx + bw2 / 2 - 1) * mm, top_pt(y + 18.3), f"{responded_p} responded the survey")
        y += 23
        c.setFillColor(muted)
        c.setFont("Helvetica-Oblique", 6.5)
        c.drawCentredString((mg + cW / 2) * mm, top_pt(y), "GR = Graduate (EPOS)   ·   UG = Undergraduate")
        y += 8

        # Gráficas de barras de programas, una por periodo, lado a lado --
        # la etiqueta va ARRIBA de la barra (no al lado) para tener todo el
        # ancho de la columna disponible y que el nombre nunca se corte.
        col_heights = []
        for i, pk in enumerate(pkeys):
            bx = mg + i * bw2
            pairs = period_program_dist[pk]
            max_n = max((n for _, n in pairs), default=1)
            yy = y
            for prog_name, n in pairs:
                label_lines = _wrap_to_width(c, prog_name, "Helvetica", 6.5, bw2 - 2)
                c.setFillColor(ink)
                c.setFont("Helvetica", 6.5)
                for li, line in enumerate(label_lines):
                    c.drawString(bx * mm, top_pt(yy + 3 + li * 3.2), line)
                yy += len(label_lines) * 3.2 + 1
                bar_w = (bw2 - 14) * mm
                fill_w = max(bar_w * (n / max_n), 1)
                c.setFillColor(rl_colors.HexColor("#F1D9E7"))
                c.rect(bx * mm, top_pt(yy + 3.6), bar_w, 3.6 * mm, fill=1, stroke=0)
                c.setFillColor(pink)
                c.rect(bx * mm, top_pt(yy + 3.6), fill_w, 3.6 * mm, fill=1, stroke=0)
                c.setFillColor(ink)
                c.setFont("Helvetica-Bold", 6.5)
                c.drawString(bx * mm + bar_w + 1.5, top_pt(yy + 2.6), str(n))
                yy += 6.5
            col_heights.append(yy - y)
        y += max(col_heights, default=0) + 6

    # ---- KPIs de respuesta, al final de esta MISMA página 1 (no en una
    # página aparte) -- solo salta de página si de verdad no cabe.
    if y + 32 > PH - 16:
        draw_footer(page_num)
        c.showPage()
        page_num += 1
        draw_continuation_header()
        y = 24.0
    stats = [
        (f"{ed['respondents']} / {ed['inscritos']}" if ed.get("inscritos") else str(ed["respondents"]), "Students Responded"),
        (f"{ed['respondents']/ed['inscritos']*100:.0f}%" if ed.get("inscritos") else "—", "Response Rate"),
        (f"{ed['avg']:.2f} / 5.0" if ed["avg"] is not None else "N/A", "Avg. Satisfaction"),
        (f"{ed['nps']:.1f} / 10" if ed["npsN"] else "N/A", "Avg. Recommendation"),
        (f"{ed['objectivesAllPct']:.0f}%" if ed["objN"] else "N/A", "Objectives Met (All)"),
    ]
    bw = cW / 5
    for i, (val, label) in enumerate(stats):
        bx = mg + i * bw
        c.setFillColor(lgray)
        c.roundRect((bx + 1) * mm, top_pt(y + 22), (bw - 3) * mm, 22 * mm, 3 * mm, fill=1, stroke=0)
        c.setFillColor(pink)
        c.setFont("Helvetica-Bold", 12)
        c.drawCentredString((bx + bw / 2 - 1) * mm, top_pt(y + 11), val)
        c.setFillColor(muted)
        c.setFont("Helvetica", 6)
        c.drawCentredString((bx + bw / 2 - 1) * mm, top_pt(y + 18), label)
    y += 32
    draw_footer(page_num)

    # ---- Páginas de evaluación cuantitativa (barras apiladas por aspecto) ----
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
        c.setFillColor(pink)
        c.rect(mg * mm, top_pt(y + 7.5), cW * mm, 7.5 * mm, fill=1, stroke=0)
        c.setFillColor(rl_colors.white)
        c.setFont("Helvetica-Bold", 8.5)
        c.drawString((mg + 5) * mm, top_pt(y + 5.3), title)
        y += 11

    def draw_likert_q(q):
        nonlocal y
        text = q.get("text") or ""
        q_lines = _wrap_to_width(c, text, "Helvetica", 8, cW)
        n_opts = sum(1 for k in LIKERT_ORDER_PDF if q["options"].get(k))
        legend_rows = 2 if n_opts >= 4 else 1
        needed = len(q_lines) * 4 + 7 + 14 + (legend_rows - 1) * 5
        check(needed)
        c.setFillColor(ink)
        c.setFont("Helvetica", 8)
        for i, line in enumerate(q_lines):
            c.drawString(mg * mm, top_pt(y + i * 4), line)
        total = sum(q["options"].values())
        if total:
            c.setFillColor(lgray)
            c.roundRect((mg + cW - 22) * mm, top_pt(y - 2), 22 * mm, 6 * mm, 2 * mm, fill=1, stroke=0)
            c.setFont("Helvetica", 7)
            c.setFillColor(muted)
            c.drawCentredString((mg + cW - 11) * mm, top_pt(y + 0.5), f"n={total}")
        y += len(q_lines) * 4 + 3
        y = _pdf_stacked_bar(c, q["options"], LIKERT_ORDER_PDF, LIKERT_COLOR_PDF, mg, y, cW, 7, PH)
        y = _pdf_legend(c, q["options"], LIKERT_ORDER_PDF, LIKERT_COLOR_PDF, mg, y, cW, PH)
        c.setStrokeColor(rl_colors.HexColor("#E6E6E6"))
        c.setLineWidth(0.3)
        c.line(mg * mm, top_pt(y), (mg + cW) * mm, top_pt(y))
        y += 5

    periods_dict = ed.get("periods", {})
    p13 = periods_dict.get("202613")
    p18 = periods_dict.get("202618")
    has_periods = bool(p13 or p18)

    if not has_periods:
        # Respaldo: si no hay periodo_EIV en los datos, usa el combinado de
        # siempre en una sola columna (comportamiento previo).
        by_aspect: Dict[str, list] = {}
        for q in ed["questions"].values():
            if not any(k in LIKERT_WEIGHTS for k in q["options"]):
                continue
            asp = q.get("aspect") or "General"
            by_aspect.setdefault(asp, []).append(q)
        section_bar("QUANTITATIVE EVALUATION")
        for asp, qs in by_aspect.items():
            check(16)
            c.setFillColor(rl_colors.HexColor("#FBE3ED"))
            c.rect(mg * mm, top_pt(y + 6.5), cW * mm, 6.5 * mm, fill=1, stroke=0)
            c.setFillColor(ink)
            c.setFont("Helvetica-Bold", 7.5)
            c.drawString((mg + 4) * mm, top_pt(y + 4.6), str(asp).upper())
            y += 9.5
            for q in qs:
                draw_likert_q(q)
            y += 2
        nps_q = ed["questions"].get(NPS_QID)
        obj_q = ed["questions"].get(OBJ_QID)
        wl_q = ed["questions"].get(WORKLOAD_QID)
        if nps_q:
            section_bar("RECOMMENDATION (0-10)")
            y = _pdf_nps_bar(c, nps_q["options"], mg, y + 5, cW, 10, PH) + 8
        if obj_q:
            section_bar("COURSE OBJECTIVES")
            y = _pdf_stacked_bar(c, obj_q["options"], OBJ_ORDER_PDF, OBJ_COLOR_PDF, mg, y, cW, 7, PH)
            y = _pdf_legend(c, obj_q["options"], OBJ_ORDER_PDF, OBJ_COLOR_PDF, mg, y, cW, PH) + 8
        if wl_q:
            section_bar("ACADEMIC WORKLOAD")
            y = _pdf_stacked_bar(c, wl_q["options"], WORKLOAD_ORDER_PDF, WORKLOAD_COLOR_PDF, mg, y, cW, 7, PH)
            _pdf_legend(c, wl_q["options"], WORKLOAD_ORDER_PDF, WORKLOAD_COLOR_PDF, mg, y, cW, PH)
        draw_footer(page_num)
    else:
        col_gap = 6.0
        col_w = (cW - col_gap) / 2
        col_x = [mg, mg + col_w + col_gap]
        col_title = ["Period 202613 (GR)", "Period 202618 (UG)"]
        col_data = [p13, p18]

        def _count_legend_rows(options, order, w):
            """Cuenta las filas EXACTAS que ocupará la leyenda en un ancho
            `w` (misma lógica de empaquetado que _pdf_legend), en vez de
            adivinar 1 o 2 filas -- una leyenda con opciones/etiquetas
            largas en columna angosta puede necesitar 3+ filas, y quedarse
            corto en la estimación era lo que cortaba preguntas a la mitad
            entre páginas."""
            lx = 0.0
            rows = 1
            for k in order:
                n = options.get(k, 0)
                if not n:
                    continue
                label = f"{k} ({n})"
                item_w = (c.stringWidth(label, "Helvetica", 6.5) / mm) + 5 + 3
                if lx + item_w > w and lx > 0:
                    rows += 1
                    lx = 0.0
                lx += item_w
            return rows

        def estimate_q_height(q, w):
            """Alto real necesario para una pregunta en la columna de ancho
            `w`: líneas del texto envuelto + barra + filas EXACTAS de
            leyenda (antes se usaba una heurística de 1-2 filas que se
            quedaba corta con etiquetas largas, cortando preguntas a la
            mitad entre páginas)."""
            text = q.get("text") or ""
            n_lines = len(_wrap_to_width(c, text, "Helvetica", 7, w))
            legend_rows = _count_legend_rows(q["options"], LIKERT_ORDER_PDF, w)
            return n_lines * 3.6 + 2 + 6 + legend_rows * 5 + 6

        def draw_likert_q_col(q, x_off, w, y_in):
            """Misma lógica que draw_likert_q pero en un ancho/columna
            propios, para el layout de 2 columnas lado a lado."""
            text = q.get("text") or ""
            q_lines = _wrap_to_width(c, text, "Helvetica", 7, w)
            c.setFillColor(ink)
            c.setFont("Helvetica", 7)
            for i, line in enumerate(q_lines):
                c.drawString(x_off * mm, top_pt(y_in + i * 3.6), line)
            y_in += len(q_lines) * 3.6 + 2
            y_in = _pdf_stacked_bar(c, q["options"], LIKERT_ORDER_PDF, LIKERT_COLOR_PDF, x_off, y_in, w, 6, PH)
            y_in = _pdf_legend(c, q["options"], LIKERT_ORDER_PDF, LIKERT_COLOR_PDF, x_off, y_in, w, PH)
            c.setStrokeColor(rl_colors.HexColor("#E6E6E6"))
            c.setLineWidth(0.3)
            c.line(x_off * mm, top_pt(y_in), (x_off + w) * mm, top_pt(y_in))
            return y_in + 4

        def col_aspects(pdata):
            by_asp: Dict[str, list] = {}
            if not pdata:
                return by_asp
            for q in pdata["questions"].values():
                if not any(k in LIKERT_WEIGHTS for k in q["options"]):
                    continue
                asp = q.get("aspect") or "General"
                by_asp.setdefault(asp, []).append(q)
            return by_asp

        by_aspect_cols = [col_aspects(p13), col_aspects(p18)]
        all_aspects, seen_asp = [], set()
        for bac in by_aspect_cols:
            for asp in bac:
                if asp not in seen_asp:
                    seen_asp.add(asp)
                    all_aspects.append(asp)

        section_bar("RESPONSES BY CATEGORY")
        for ci in range(2):
            c.setFillColor(pink)
            c.setFont("Helvetica-Bold", 8)
            c.drawString(col_x[ci] * mm, top_pt(y + 4), col_title[ci])
        y += 8

        for asp in all_aspects:
            qs_l = by_aspect_cols[0].get(asp, [])
            qs_r = by_aspect_cols[1].get(asp, [])

            def draw_aspect_header():
                nonlocal y
                for ci in range(2):
                    c.setFillColor(rl_colors.HexColor("#FBE3ED"))
                    c.rect(col_x[ci] * mm, top_pt(y + 6), col_w * mm, 6 * mm, fill=1, stroke=0)
                    c.setFillColor(ink)
                    c.setFont("Helvetica-Bold", 6.8)
                    c.drawString((col_x[ci] + 3) * mm, top_pt(y + 4.3), str(asp).upper())
                y += 8.5

            # Revisa el bloque COMPLETO del aspecto (encabezado + TODAS las
            # preguntas de AMBAS columnas) antes de dibujar nada, para
            # empezarlo en página nueva si de plano no cabe.
            total_l = sum(estimate_q_height(q, col_w) for q in qs_l)
            total_r = sum(estimate_q_height(q, col_w) for q in qs_r)
            block_needed = 8.5 + max(total_l, total_r)
            if y + min(block_needed, PH - 40) > PH - 16 and y > 24.0:
                draw_footer(page_num)
                c.showPage()
                page_num += 1
                draw_continuation_header()
                y = 24.0

            draw_aspect_header()
            y_l = y_r = y

            # Dibuja las columnas EMPAREJADAS pregunta por pregunta (no toda
            # la izquierda y luego toda la derecha) -- así, si de verdad no
            # alcanza el espacio para una pareja, ambas columnas saltan de
            # página JUNTAS y se repite el encabezado del aspecto en la
            # página nueva. Antes, una columna podía terminar de dibujar
            # sus preguntas cerca del fondo de la página mientras la otra
            # apenas empezaba, dejando el encabezado de un lado sin sus
            # preguntas y esas preguntas reapareciendo solas, sin
            # encabezado, en la siguiente página.
            n_pairs = max(len(qs_l), len(qs_r))
            for i in range(n_pairs):
                q_l = qs_l[i] if i < len(qs_l) else None
                q_r = qs_r[i] if i < len(qs_r) else None
                needed_l = estimate_q_height(q_l, col_w) if q_l else 0
                needed_r = estimate_q_height(q_r, col_w) if q_r else 0
                if max(y_l + needed_l, y_r + needed_r) > PH - 16:
                    draw_footer(page_num)
                    c.showPage()
                    page_num += 1
                    draw_continuation_header()
                    y = 24.0
                    draw_aspect_header()
                    y_l = y_r = y
                if q_l:
                    y_l = draw_likert_q_col(q_l, col_x[0], col_w, y_l)
                if q_r:
                    y_r = draw_likert_q_col(q_r, col_x[1], col_w, y_r)
            y = max(y_l, y_r) + 1

        # ---- Recommendation (NPS): periodo 202613 arriba, 202618 abajo ----
        nps_q13 = (p13 or {}).get("questions", {}).get(NPS_QID)
        nps_q18 = (p18 or {}).get("questions", {}).get(NPS_QID)
        if nps_q13 or nps_q18:
            draw_footer(page_num)
            c.showPage()
            page_num += 1
            draw_continuation_header()
            y = 24.0
            section_bar("RECOMMENDATION (0-10)")
            for periodo, nps_q, grp in (("202613", nps_q13, "GR"), ("202618", nps_q18, "UG")):
                if not nps_q:
                    continue
                check(20)
                c.setFillColor(pink)
                c.setFont("Helvetica-Bold", 7.5)
                c.drawString(mg * mm, top_pt(y + 3), f"Period {periodo} ({grp})")
                y += 6
                q_lines = _wrap_to_width(c, nps_q.get("text") or "", "Helvetica", 7.5, cW)
                c.setFillColor(ink)
                c.setFont("Helvetica", 7.5)
                for i, line in enumerate(q_lines):
                    c.drawString(mg * mm, top_pt(y + i * 3.6), line)
                y += len(q_lines) * 3.6 + 3
                y = _pdf_nps_bar(c, nps_q["options"], mg, y, cW, 9, PH) + 7

        # ---- Objectives y Academic Workload: 202613 / 202618 lado a lado ----
        obj_q13 = (p13 or {}).get("questions", {}).get(OBJ_QID)
        obj_q18 = (p18 or {}).get("questions", {}).get(OBJ_QID)
        if obj_q13 or obj_q18:
            check(24)
            section_bar("COURSE OBJECTIVES")
            for ci in range(2):
                c.setFillColor(pink)
                c.setFont("Helvetica-Bold", 7.5)
                c.drawString(col_x[ci] * mm, top_pt(y + 3), col_title[ci])
            y0 = y + 6
            y_l = y_r = y0
            if obj_q13:
                q_lines = _wrap_to_width(c, obj_q13.get("text") or "", "Helvetica", 6.8, col_w)
                c.setFillColor(ink); c.setFont("Helvetica", 6.8)
                for i, line in enumerate(q_lines):
                    c.drawString(col_x[0] * mm, top_pt(y_l + i * 3.4), line)
                y_l += len(q_lines) * 3.4 + 2
                y_l = _pdf_stacked_bar(c, obj_q13["options"], OBJ_ORDER_PDF, OBJ_COLOR_PDF, col_x[0], y_l, col_w, 6, PH)
                y_l = _pdf_legend(c, obj_q13["options"], OBJ_ORDER_PDF, OBJ_COLOR_PDF, col_x[0], y_l, col_w, PH)
            if obj_q18:
                q_lines = _wrap_to_width(c, obj_q18.get("text") or "", "Helvetica", 6.8, col_w)
                c.setFillColor(ink); c.setFont("Helvetica", 6.8)
                for i, line in enumerate(q_lines):
                    c.drawString(col_x[1] * mm, top_pt(y_r + i * 3.4), line)
                y_r += len(q_lines) * 3.4 + 2
                y_r = _pdf_stacked_bar(c, obj_q18["options"], OBJ_ORDER_PDF, OBJ_COLOR_PDF, col_x[1], y_r, col_w, 6, PH)
                y_r = _pdf_legend(c, obj_q18["options"], OBJ_ORDER_PDF, OBJ_COLOR_PDF, col_x[1], y_r, col_w, PH)
            y = max(y_l, y_r) + 6

        wl_q13 = (p13 or {}).get("questions", {}).get(WORKLOAD_QID)
        wl_q18 = (p18 or {}).get("questions", {}).get(WORKLOAD_QID)
        if wl_q13 or wl_q18:
            check(24)
            section_bar("ACADEMIC WORKLOAD")
            for ci in range(2):
                c.setFillColor(pink)
                c.setFont("Helvetica-Bold", 7.5)
                c.drawString(col_x[ci] * mm, top_pt(y + 3), col_title[ci])
            y0 = y + 6
            y_l = y_r = y0
            if wl_q13:
                q_lines = _wrap_to_width(c, wl_q13.get("text") or "", "Helvetica", 6.8, col_w)
                c.setFillColor(ink); c.setFont("Helvetica", 6.8)
                for i, line in enumerate(q_lines):
                    c.drawString(col_x[0] * mm, top_pt(y_l + i * 3.4), line)
                y_l += len(q_lines) * 3.4 + 2
                y_l = _pdf_stacked_bar(c, wl_q13["options"], WORKLOAD_ORDER_PDF, WORKLOAD_COLOR_PDF, col_x[0], y_l, col_w, 6, PH)
                y_l = _pdf_legend(c, wl_q13["options"], WORKLOAD_ORDER_PDF, WORKLOAD_COLOR_PDF, col_x[0], y_l, col_w, PH)
            if wl_q18:
                q_lines = _wrap_to_width(c, wl_q18.get("text") or "", "Helvetica", 6.8, col_w)
                c.setFillColor(ink); c.setFont("Helvetica", 6.8)
                for i, line in enumerate(q_lines):
                    c.drawString(col_x[1] * mm, top_pt(y_r + i * 3.4), line)
                y_r += len(q_lines) * 3.4 + 2
                y_r = _pdf_stacked_bar(c, wl_q18["options"], WORKLOAD_ORDER_PDF, WORKLOAD_COLOR_PDF, col_x[1], y_r, col_w, 6, PH)
                y_r = _pdf_legend(c, wl_q18["options"], WORKLOAD_ORDER_PDF, WORKLOAD_COLOR_PDF, col_x[1], y_r, col_w, PH)
            y = max(y_l, y_r) + 4

        draw_footer(page_num)

    # ---- Página de comentarios ----
    c.showPage()
    page_num += 1
    draw_continuation_header()
    y = 24.0

    com_by_prof = comments_by_professor()
    com_qs = com_by_prof.get(_resolve_alias(profesor), {})

    def _sort_by_period(items):
        # 202613 primero, luego 202618, luego cualquier otro/():sin periodo
        order = {"202613": 0, "202618": 1}
        return sorted(items, key=lambda it: order.get(it.get("periodo"), 2))

    comment_groups = []
    for qid, label in {**POSITIVE_QIDS, **IMPROVE_QIDS}.items():
        g = com_qs.get(qid)
        exact_q = (g.get("text") if g else None) or label
        comment_groups.append({"label": exact_q, "items": _sort_by_period(g["items"]) if g else []})

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
    c.drawString(mg * mm, top_pt(y), "STUDENT COMMENTS")
    y += 7

    with_items = [g for g in comment_groups if g["items"]]
    if not with_items:
        c.setFont("Helvetica", 8)
        c.setFillColor(rl_colors.HexColor("#50484E"))
        c.drawString(mg * mm, top_pt(y), "No comments collected for this faculty member.")
        draw_footer(page_num)
    else:
        for g in comment_groups:
            # Estima el alto TOTAL del grupo (encabezado + todos sus items)
            # ANTES de empezar a dibujarlo -- si no cabe completo en lo que
            # queda de la página actual, se salta la página entera para
            # este grupo, en vez de partirlo a la mitad.
            group_needed = 7.0  # encabezado
            if not g["items"]:
                group_needed += 6
            else:
                for item in g["items"]:
                    tag = f'[{item["periodo"]}] ' if item.get("periodo") else ""
                    lines_est = _wrap_to_width(c, f'1. {tag}"{item["text"]}"', "Helvetica", 8, cW)
                    group_needed += len(lines_est) * 4 + 2 + 0.001 * 0  # aproximación consistente con el dibujo real
                group_needed += 4
            if y + min(group_needed, 260) > 283 and y > 24.0:
                draw_footer(page_num)
                c.showPage()
                page_num += 1
                draw_continuation_header()
                y = 24.0
            c.setFont("Helvetica-Bold", 8.5)
            c.setFillColor(pink)
            n = len(g["items"])
            c.drawString(mg * mm, top_pt(y), f"{g['label'].upper()}")
            y += 5.5
            if not g["items"]:
                c.setFont("Helvetica", 8)
                c.setFillColor(muted)
                c.drawString(mg * mm, top_pt(y), "No comments for this question.")
                y += 6
                continue
            for i, item in enumerate(g["items"]):
                tag = f'[{item["periodo"]}] ' if item.get("periodo") else ""
                lines = _wrap_to_width(c, f'{i + 1}. {tag}"{item["text"]}"', "Helvetica", 8, cW)
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



# ── 15) PÁGINA — School Enrollment & Gap Analysis ────────────────────────
def page_faculty():
    d = compute_all()
    _render_header(
        "Other Analysis — School Enrollment & International Exposure",
        "Historical enrollment, graduate program mix, PRE mobility, and the international-exposure gap analysis.",
    )

    S = SCHOOL_ENROLLMENT
    years = [r["year"] for r in S["historicalEnrollment"]]
    st.markdown("### Historical Enrollment")
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=years, y=[r["PRE"] for r in S["historicalEnrollment"]], mode="lines+markers+text",
                              name="PRE (undergraduate)", line=dict(color=BLUE, width=3, shape="spline"),
                              text=[r["PRE"] for r in S["historicalEnrollment"]], textposition="top center"))
    fig.add_trace(go.Scatter(x=years, y=[r["POS"] for r in S["historicalEnrollment"]], mode="lines+markers+text",
                              name="EPOS (graduate)", line=dict(color=PURPLE, width=3, shape="spline"),
                              text=[r["POS"] for r in S["historicalEnrollment"]], textposition="bottom center"))
    fig.update_layout(margin=dict(t=20, r=20, b=30, l=46), legend=dict(orientation="h", x=0, y=1.15),
                       xaxis=dict(type="category"), plot_bgcolor="rgba(0,0,0,0)",
                       paper_bgcolor="rgba(0,0,0,0)", height=320)
    st.plotly_chart(fig, use_container_width=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        _kpi("Current PRE enrolled", f"{S['currentPre']:,}")
    with c2:
        _kpi("Current EPOS enrolled", f"{S['currentPos']:,}")
    with c3:
        _kpi("Reference period", S["currentPeriod"])

    st.markdown("### Graduate Program Distribution")
    prog = sorted(S["posProgramDistribution"], key=lambda p: -p["count"])
    fig2 = go.Figure(go.Bar(x=[p["count"] for p in prog], y=[p["program"] for p in prog], orientation="h",
                             marker=dict(color=PURPLE), text=[p["count"] for p in prog], textposition="outside"))
    fig2.update_layout(margin=dict(t=10, r=30, b=30, l=260), yaxis=dict(autorange="reversed"),
                        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", height=460)
    st.plotly_chart(fig2, use_container_width=True)

    st.markdown("---")
    st.markdown("### Undergraduate (PRE) Mobility")
    M = MOBILITY_DATA
    myears = [str(y["year"]) for y in M["byYear"]]
    type_colors = {"Intercambio Internacional": BLUE, "Doble Titulación": PURPLE, "Pasantía de Investigación": GREEN}
    fig3 = go.Figure()
    for t, color in type_colors.items():
        fig3.add_trace(go.Bar(x=myears, y=[y["types"].get(t, 0) for y in M["byYear"]], name=t, marker=dict(color=color)))
    totals = [sum(y["types"].values()) for y in M["byYear"]]
    fig3.add_trace(go.Scatter(x=myears, y=totals, mode="text", text=[f"Total: {t}" for t in totals],
                               textposition="top center", showlegend=False, hoverinfo="skip"))
    fig3.update_layout(margin=dict(t=26, r=16, b=30, l=40), barmode="stack",
                        legend=dict(orientation="h", x=0, y=1.18), plot_bgcolor="rgba(0,0,0,0)",
                        paper_bgcolor="rgba(0,0,0,0)", height=340)
    st.plotly_chart(fig3, use_container_width=True)

    st.markdown("### Graduate (EPOS) — International Weeks & Electives")
    W = WEEKS_DATA["byWeek"]
    week_colors = {"KLU": ORANGE, "NOVA": PURPLE, "FGV": BLUE, "BABSON": "#3FA34D"}
    fig4 = go.Figure()
    fig4.add_trace(go.Bar(x=[w["week"] for w in W], y=[w["MBA"] for w in W], name="MBA",
                           marker=dict(color=[week_colors[w["week"]] + "88" for w in W]),
                           text=[f'{w["MBA"]} MBA' for w in W], textposition="inside"))
    fig4.add_trace(go.Bar(x=[w["week"] for w in W], y=[w["EPOS"] for w in W], name="EPOS",
                           marker=dict(color=[week_colors[w["week"]] for w in W]),
                           text=[f'{w["EPOS"]} EPOS' for w in W], textposition="inside"))
    fig4.update_layout(margin=dict(t=10, r=16, b=40, l=40), barmode="stack", showlegend=False,
                        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", height=300)
    st.plotly_chart(fig4, use_container_width=True)

    df_electivas = load_electivas()
    areas: Dict[str, dict] = {}
    for _, r in df_electivas.iterrows():
        materia = str(r.get("Materia") or "")
        area = materia.split("-")[0] or "Other"
        a = areas.setdefault(area, {"courses": 0, "enrolled": 0})
        a["courses"] += 1
        a["enrolled"] += int(r.get("Inscritos") or 0)
    area_list = sorted(areas.items(), key=lambda kv: -kv[1]["enrolled"])
    fig5 = go.Figure(go.Bar(
        x=[v["enrolled"] for _, v in area_list],
        y=[f'{a} ({v["courses"]} course{"s" if v["courses"]>1 else ""})' for a, v in area_list],
        orientation="h", marker=dict(color=PINK),
        text=[f'{v["enrolled"]} enrolled' for _, v in area_list], textposition="outside",
    ))
    fig5.update_layout(margin=dict(t=10, r=30, b=30, l=170), yaxis=dict(autorange="reversed"),
                        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", height=max(240, 30 * len(area_list)))
    st.plotly_chart(fig5, use_container_width=True)

    st.markdown("---")
    st.markdown("### International-Exposure Gap Analysis")
    st.caption(
        "Share of currently enrolled students with at least one form of international "
        "exposure (EIV, mobility, convocatoria, international weeks)."
    )
    pre_exp = S["preIntlExperience"]
    pos_eiv = d["df_listas"][d["df_listas"]["Categoría académica"] == "EPOS UASM"]["Código est"].dropna().nunique()
    pos_semanas = sum(w["MBA"] + w["EPOS"] for w in W)

    parts = [
        {"label": "Escuela de Verano (EIV)", "pre": pre_exp["escuelaVerano"], "pos": pos_eiv, "color": PINK},
        {"label": "Movilidad", "pre": pre_exp["movilidad"], "pos": MOBILITY_DATA["posUniqueLast5"], "color": BLUE},
        {"label": "Convocatoria", "pre": pre_exp["convocatoria"], "pos": 0, "color": PURPLE},
        {"label": "Semanas Internacionales", "pre": 0, "pos": pos_semanas, "color": GREEN},
    ]
    pre_with = sum(p["pre"] for p in parts)
    pos_with = sum(p["pos"] for p in parts)
    cats = ["UG", "GR"]

    fig6 = go.Figure()
    for p in parts:
        pre_p = p["pre"] / S["currentPre"] * 100
        pos_p = p["pos"] / S["currentPos"] * 100
        fig6.add_trace(go.Bar(
            x=cats, y=[pre_p, pos_p], name=p["label"], marker=dict(color=p["color"]),
            text=[f"{pre_p:.1f}% ({p['pre']})" if p["pre"] else "", f"{pos_p:.1f}% ({p['pos']})" if p["pos"] else ""],
            textposition="inside",
        ))
    without_pre = S["currentPre"] - pre_with
    without_pos = S["currentPos"] - pos_with
    new_pre = min(S["newStudents"]["pre"], without_pre)
    new_pos = min(S["newStudents"]["pos"], without_pos)
    other_pre = without_pre - new_pre
    other_pos = without_pos - new_pos

    fig6.add_trace(go.Bar(x=cats, y=[new_pre / S["currentPre"] * 100, new_pos / S["currentPos"] * 100],
                           name="New students", marker=dict(color="#C7B9C1")))
    fig6.add_trace(go.Bar(x=cats, y=[other_pre / S["currentPre"] * 100, other_pos / S["currentPos"] * 100],
                           name="No International Experience", marker=dict(color="#EAE3E8")))
    fig6.update_layout(margin=dict(t=10, r=16, b=30, l=46), barmode="stack",
                        legend=dict(orientation="h", x=0, y=1.2), yaxis=dict(ticksuffix="%", range=[0, 100]),
                        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", height=380)
    st.plotly_chart(fig6, use_container_width=True)


# ── 16) PÁGINA — Conclusions ──────────────────────────────────────────────
def page_conclusions():
    _render_header("Achievements & Challenges", "Findings and recommendations for the 2026 edition.")

    GREEN_H = "#1FBE72"

    st.markdown(f'<div style="font-size:30px;font-weight:700;color:{GREEN_H};margin:8px 0 18px;">Key Achievements</div>', unsafe_allow_html=True)
    achievements = [
        ("Nivel muy top de profesores y universidades representantes.", ""),
        ("Programa muy conocido entre profesores internacionales.",
         "La convocatoria cerrada tuvo más de 100 propuestas para los 13 cursos finalmente escogidos."),
        ("Apoya la internacionalización de la facultad y de la universidad",
         "con actividades extra que se dan gracias a la escuela de verano: reuniones de área, reuniones con "
         "profesores, cursos de educación ejecutiva, seminarios, etc."),
        ("Principal vía para que los estudiantes cumplan su experiencia internacional.",
         "Ayuda a cumplir el componente de internacionalización sin salir del campus para quienes no puedan "
         "realizar un intercambio en el exterior. Es el programa top of mind de los estudiantes, y ayuda a "
         "quienes cuentan con pocos ingresos a vivir una experiencia internacional."),
    ]
    for title, sub in achievements:
        sub_html = f'<div style="font-size:13.5px;color:{MUTED};margin-top:4px;">{sub}</div>' if sub else ""
        st.markdown(
            f'<div style="margin-bottom:22px;"><div style="font-size:17px;font-weight:700;color:{INK};">{title}</div>{sub_html}</div>',
            unsafe_allow_html=True,
        )

    st.markdown("---")
    st.markdown(f'<div style="font-size:30px;font-weight:700;color:{GREEN_H};margin:8px 0 18px;">Challenges</div>', unsafe_allow_html=True)
    challenges = [
        ("Experiencia virtual vs. presencial.", ""),
        ("Atracción de estudiantes internacionales.", ""),
        ("Dispersión de notas.",
         "Al final la tendencia es superar 4.5 y no hay pérdidas — la mayoría de quienes pierden retiran el "
         "curso, o no pagaron y se les retira."),
        ("Claridad de reglas y expectativas del curso con los profesores.",
         "Reforzar con los profesores visitantes cómo se califica en Uniandes."),
        ("Baja participación histórica en cursos de sostenibilidad.", ""),
        ("Contar con un estudiante representante de pregrado",
         "para escoger cursos que sean pertinentes y de valor para los estudiantes."),
        ("El bootcamp se percibe como un curso totalmente diferente", "a los demás cursos de la escuela de verano."),
        ("% de participación estratégica.",
         "Importante para generar una buena participación en clase, sin llenar el curso con evaluaciones de "
         "participación. máximo un 10% de la nota final."),
    ]
    for title, sub in challenges:
        sub_html = f'<div style="font-size:13.5px;color:{MUTED};margin-top:4px;">{sub}</div>' if sub else ""
        st.markdown(
            f'<div style="margin-bottom:22px;"><div style="font-size:17px;font-weight:700;color:{INK};">{title}</div>{sub_html}</div>',
            unsafe_allow_html=True,
        )


# ── 17) NAVEGACIÓN ────────────────────────────────────────────────────────
pages = [
    st.Page(page_cover, title="Data Center", icon="🌐", url_path="cover", default=True),
    st.Page(page_overview, title="Overview", icon="📈", url_path="overview"),
    st.Page(page_summary, title="Summary", icon="🌍", url_path="summary"),
    st.Page(page_dashboard, title="Dashboard", icon="📊", url_path="dashboard"),
    st.Page(page_visiting, title="Evaluation & Feedback", icon="🧑‍🏫", url_path="visiting"),
    st.Page(page_conclusions, title="Achievements & Challenges", icon="📝", url_path="conclusions"),
    st.Page(page_financial, title="Financial", icon="💰", url_path="financial"),
    st.Page(page_faculty, title="Other Analysis", icon="🎓", url_path="enrollment"),
]
pg = st.navigation(pages, position="hidden")
IS_COVER = pg is pages[0]
nav_pages = pages[1:]
_NAV_VISIBLE = 5

if IS_COVER:
    st.markdown(
        "<style>section[data-testid='stSidebar']{display:none !important;}"
        "div[data-testid='stSidebarCollapsedControl']{display:none !important;}</style>",
        unsafe_allow_html=True,
    )

if not IS_COVER:
    with st.sidebar:
        try:
            _show_eiv_logo(140)
        except Exception:
            pass
        st.markdown(
            f'<div style="color:{INK};font-size:22px;font-weight:800;line-height:1.1;">EIV Analytics</div>',
            unsafe_allow_html=True,
        )
        st.caption("International Summer School · Facultad de Administración")
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

    # ---- Flechas laterales para pasar de sección (un solo contenedor,
    # mismo patrón flex que el nav de arriba, para evitar el problema de
    # posicionamiento con dos contenedores 'fixed' separados) ----
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
    'font-family:monospace;">International Summer School (EIV) · Facultad de Administración · '
    'Universidad de los Andes</div>',
    unsafe_allow_html=True,
)
