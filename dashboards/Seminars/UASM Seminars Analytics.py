# ===========================================================================
#  UASM · Seminars Analytics · App multipágina
#  Traducción 1:1 del reporte HTML "International Seminars" a Streamlit —
#  mismas 4 secciones (Data Center, Dashboard, Seminar Search, Tables), mismo
#  modelo de datos (Seminarios/Conferencistas/Participantes unidos por el ID
#  "Seminario"), misma paleta de color por dominio (Seminars=teal,
#  Speakers=blue, Participants=pink).
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
    "border-radius:12px;box-shadow:0 2px 8px rgba(14,165,160,.18);margin-bottom:14px;}}"
    f".sh-super{{font-size:11px;font-weight:700;letter-spacing:2px;"
    f"color:{SEM_SOFT};text-transform:uppercase;margin-bottom:2px;}}"
    ".sh-title{font-size:26px;font-weight:800;color:#fff;text-align:center;line-height:1.2;}"
    ".sh-sub{font-size:13px;color:rgba(255,255,255,.80);margin-top:4px;text-align:center;}"
    f".kv{{font-size:26px;font-weight:700;line-height:1.1;font-family:monospace;}}"
    f".kl{{font-size:11px;font-weight:600;color:{MUTED};"
    "text-transform:uppercase;letter-spacing:.5px;margin-top:3px;}}"
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
    "font-weight:700 !important;}}"
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
    ".st-key-nav_toggle{position:fixed;top:0.25rem;left:50%;transform:translateX(-50%);"
    "z-index:999999;width:70vw;max-width:900px;}"
    ".st-key-nav_toggle div[data-testid='stHorizontalBlock']{"
    "display:flex !important;flex-wrap:nowrap !important;width:100% !important;"
    "justify-content:space-between !important;gap:8px !important;}"
    ".st-key-nav_toggle div[data-testid='column']{width:auto !important;min-width:fit-content !important;flex:none !important;}"
    ".st-key-nav_toggle div[data-testid='stPageLink']{width:auto !important;min-width:fit-content !important;overflow:visible !important;}"
    ".st-key-nav_toggle div[data-testid='stPageLink'] a{white-space:nowrap !important;overflow:visible !important;text-overflow:unset !important;width:auto !important;min-width:fit-content !important;}"
    ".st-key-nav_toggle div[data-testid='stPageLink'] a p{white-space:nowrap !important;overflow:visible !important;}"
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


def _pending_card(label: str, note: str = ""):
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


# Nacionalidad (español, tal como aparece en BD_conferencistas) -> centroide,
# para el mapa de burbujas de Speakers. Idéntico a countryGeo del HTML.
COUNTRY_GEO = {
    "Alemania": (51.16, 10.45), "Argentina": (-38.42, -63.62), "Brazil": (-14.24, -51.93),
    "Chile": (-35.68, -71.54), "China": (35.86, 104.20), "Colombia": (4.57, -74.30),
    "Costa Rica": (9.75, -83.75), "Dinamarca": (56.26, 9.50), "Egipto": (26.82, 30.80),
    "Eslovenia": (46.15, 14.99), "España": (40.46, -3.75), "Estados Unidos": (37.09, -95.71),
    "Francia": (46.23, 2.21), "Grecia": (39.07, 21.82), "Holanda": (52.13, 5.29),
    "India": (20.59, 78.96), "Iran": (32.43, 53.69), "Irlanda": (53.14, -7.69),
    "Italia": (41.87, 12.57), "Mexico": (23.63, -102.55), "Peru": (-9.19, -75.02),
    "Polonia": (51.92, 19.15), "Portugal": (39.40, -8.22), "Turquía": (38.96, 35.24),
    "UK": (55.38, -3.44),
}

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

import os
import base64


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
    raw_s = io.BytesIO(_download_drive_file_bytes(SEMINARIOS_FILE_ID))
    sem_df = pd.read_excel(raw_s, sheet_name="seminarios")
    raw_c = io.BytesIO(_download_drive_file_bytes(CONFERENCISTAS_FILE_ID))
    conf_df = pd.read_excel(raw_c, sheet_name="conferencistas")
    raw_p = io.BytesIO(_download_drive_file_bytes(PARTICIPANTES_FILE_ID))
    part_df = pd.read_excel(raw_p, sheet_name="participantes")
    return sem_df, conf_df, part_df


@st.cache_data(ttl=300)
def build_dataset() -> List[dict]:
    """Replica exacta de buildDataset() del HTML: une seminarios/conferencistas/
    participantes por el ID 'Seminario', y desambigua automáticamente
    cualquier ID duplicado (dato repetido por error de captura), igual que
    el original — para cualquier cantidad de duplicados, no solo un caso
    conocido."""
    sem_df, conf_df, part_df = load_raw_sheets()

    participants_by_sem: Dict[str, List[dict]] = {}
    for _, r in part_df.iterrows():
        sem = r.get("Seminario al que asistió")
        nombre = r.get("Nombre")
        if pd.isna(sem) or pd.isna(nombre):
            continue
        participants_by_sem.setdefault(sem, []).append({
            "nombre": str(nombre).strip(), "correo": r.get("Correo"),
            "genero": norm_genero(r.get("Género")), "tipo": norm_tipo(r.get("Tipo Asistente")),
            "modalidad": norm_modalidad(r.get("Modalidad")),
        })

    def _clean_text_g(v, default="Unknown"):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return default
        txt = str(v).strip()
        return txt if txt else default

    speakers_by_sem: Dict[str, List[dict]] = {}
    for _, r in conf_df.iterrows():
        sem = r.get("Seminario que dirigió")
        if pd.isna(sem) or pd.isna(r.get("Nombre Completo")):
            continue
        speakers_by_sem.setdefault(sem, []).append({
            "nombre": _clean_text_g(r.get("Nombre Completo"), ""),
            "genero": _clean_text_g(r.get("Genero")),
            "nacionalidad": _clean_text_g(r.get("Nacionalidad")),
            "universidad": _clean_text_g(r.get("Nombre Universidad origen")),
            "paisUniversidad": _clean_text_g(r.get("País Universidad origen")),
            "region": _clean_text_g(r.get("Ubicación")),
        })

    seen_ids: Dict[str, int] = {}
    seminars: List[dict] = []
    for _, r in sem_df.iterrows():
        raw_id = r.get("Seminario")
        if pd.isna(raw_id):
            continue
        seen_ids[raw_id] = seen_ids.get(raw_id, 0) + 1
        sid = raw_id
        data_note = None
        if seen_ids[raw_id] > 1:
            sid = f"{raw_id} ({seen_ids[raw_id]})"
            data_note = (
                f'Source data reused ID "{raw_id}" for more than one seminar — participant '
                f"records could not be split automatically and are shown only under the "
                f"first occurrence."
            )
        def _clean_speaker(v):
            if v is None or (isinstance(v, float) and pd.isna(v)):
                return None
            txt = str(v).strip()
            if not txt or txt.upper() == "NA":
                return None
            return txt

        speakers = [_clean_speaker(r.get(f"Conferencista {i}")) for i in range(1, 6)]
        speakers = [s for s in speakers if s]

        date_val = r.get("Fecha")
        date_val = date_val if (date_val is not None and not pd.isna(date_val)) else None
        hora_raw = r.get("Hora")
        hora_str = None
        if hora_raw is not None and not pd.isna(hora_raw):
            try:
                t = pd.Timestamp(hora_raw)
                hora_str = f"{t.hour:02d}:{t.minute:02d}"
            except Exception:
                hora_str = str(hora_raw)

        plist = participants_by_sem.get(raw_id, []) if seen_ids[raw_id] == 1 else []

        def _clean_text(v, default=None):
            if v is None or (isinstance(v, float) and pd.isna(v)):
                return default
            txt = str(v).strip()
            return txt if txt else default

        seminars.append({
            "id": sid, "categoria": _clean_text(r.get("Categoría"), "Other"),
            "nombre": _clean_text(r.get("Nombre del seminario"), "Untitled seminar"),
            "speakers": speakers, "fecha": fmt_date(date_val), "hora": hora_str,
            "anio": r.get("Año"), "semestre": semester_of(date_val),
            "modalidad": norm_modalidad(r.get("Modalidad")), "lugar": _clean_text(r.get("Lugar"), "TBD"),
            "area": _clean_text(r.get("Área"), "Unknown"), "participantsCount": len(plist), "participants": plist,
            "speakerDetails": speakers_by_sem.get(raw_id, []), "dataNote": data_note,
        })
    return seminars


def all_speakers(seminars: List[dict]) -> List[dict]:
    seen = set()
    out = []
    for s in seminars:
        for sp in s["speakerDetails"]:
            if sp["nombre"] in seen:
                continue
            seen.add(sp["nombre"])
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
        st.info("No hay datos para este filtro.")
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


def _hbar(pairs: List[Tuple[str, int]], color: str, height: int = 220):
    if not pairs:
        st.info("No hay datos para este filtro.")
        return
    labels = [p[0] for p in pairs]
    values = [p[1] for p in pairs]
    fig = go.Figure(go.Bar(
        x=values, y=labels, orientation="h", marker=dict(color=color),
        text=[str(v) for v in values], textposition="outside", cliponaxis=False,
        textfont=dict(size=10.5, color=INK),
    ))
    fig.update_layout(
        margin=dict(t=6, r=30, b=24, l=150),
        xaxis=dict(gridcolor=LINE, tickfont=dict(color=INK_SOFT, size=9)),
        yaxis=dict(tickfont=dict(color=INK, size=10.5), automargin=True, autorange="reversed"),
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", height=height,
    )
    st.plotly_chart(fig, use_container_width=True)


def _world_map(speakers: List[dict], focus_country: Optional[str] = None):
    by_nat = count_by(speakers, lambda s: s.get("nacionalidad"))
    by_nat = [(n, c) for n, c in by_nat if n in COUNTRY_GEO]
    if not by_nat:
        st.info("No hay nacionalidades mapeables para este filtro.")
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


# ── 6) PÁGINA — Data Center ─────────────────────────────────────────────
def page_datacenter():
    seminars = build_dataset()
    _render_header(
        "Data Center",
        "Verification of the sources powering this report. Every seminar/speaker/participant "
        "is matched by the \"Seminario\" ID.",
    )
    files = [
        ("BD_seminarios.xlsx", "Seminar catalog — category, speakers, date, modality, area"),
        ("BD_conferencistas.xlsx", "Speaker profiles — gender, nationality, origin university"),
        ("BD_participantes.xlsx", "Attendance roster — name, email, gender, type, modality"),
    ]
    for fname, desc in files:
        st.markdown(f"**{fname}** — {desc}  \n:material/check_circle: Loaded")
    st.markdown(f"**{len(files)} sources loaded** · {len(seminars)} seminars")
    st.page_link(pages[1], label="Enter the Report →", icon=":material/arrow_forward:")


# ── 7) PÁGINA — Dashboard ────────────────────────────────────────────────
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
        _hbar(count_by(sems, lambda s: s["area"]), SEM, height=230)

    st.markdown("---")

    # ---------------- Speakers domain ----------------
    _domain_badge("🎤 Speakers", "spk")
    col1, col2 = st.columns(2)
    with col1:
        _kpi("Speakers", len(speakers), SPK)
    with col2:
        _kpi("Universities Represented", len(set(s["universidad"] for s in speakers)), SPK)

    col_g, col_r = st.columns(2)
    with col_g:
        st.markdown("##### By Gender")
        _donut(count_by(speakers, lambda s: s["genero"]), [SPK, PAR, MUTED], height=200)
    with col_r:
        st.markdown("##### By Origin Region")
        _hbar(count_by(speakers, lambda s: s["region"]), SPK, height=200)

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
        _hbar(count_by(parts, lambda p: p["tipo"]), PAR, height=200)

    st.markdown("---")
    _render_pdf_download(sems, speakers, parts, avg_part)


# ── 8) PÁGINA — Seminar Search ──────────────────────────────────────────
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
        part_df.columns = ["Nombre", "Correo", "Género", "Tipo Asistente", "Modalidad"]
    else:
        part_df = pd.DataFrame(columns=["Nombre", "Correo", "Género", "Tipo Asistente", "Modalidad"])
    st.dataframe(part_df, use_container_width=True, hide_index=True)
    st.download_button(
        "Download as Excel", data=_xlsx_bytes(part_df, "Participants"),
        file_name=f"{s['id']}_participants.xlsx".replace(" ", "_"), key=f"dl_sem_{s['id']}",
    )


# ── 9) PÁGINA — Tables ───────────────────────────────────────────────────
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
                "Participants": s["participantsCount"],
            })
        rows_disp.append({"Category": "", "Seminar": "", "Speaker(s)": "", "Date": "", "Modality": "",
                           "Area": f"Total participants — {cat}", "Participants": cat_total})
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
        c.drawString(mg, H - 21 * mm, "Facultad de Administración · Universidad de los Andes")

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
            len(set(s["universidad"] for s in speakers)), "Universities", SPK)
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
    c.setFillColor(rl_colors.HexColor("#241420"))
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
        c.setFillColor(rl_colors.HexColor("#241420"))
        c.setFont("Helvetica-Bold", 8.5)
        c.drawString(mg + 3, y - 0.5 * mm, f"{cat.upper()} — {len(in_cat)} seminar{'s' if len(in_cat) != 1 else ''}")
        y -= 10 * mm
        for s in in_cat:
            check_page(12 * mm)
            c.setFont("Helvetica", 7.5)
            c.setFillColor(rl_colors.HexColor("#241420"))
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


def _render_pdf_download(sems, speakers, parts, avg_part):
    try:
        pdf_bytes = _build_summary_pdf(sems, speakers, parts, avg_part)
        st.download_button(
            "📄 Generate PDF Report", data=pdf_bytes,
            file_name="International_Seminars_Report.pdf", mime="application/pdf",
            key="dl_summary_pdf", use_container_width=False,
        )
    except Exception as e:
        st.caption(f"PDF export unavailable: {e}")


# ── 11) NAVEGACIÓN ────────────────────────────────────────────────────────
pages = [
    st.Page(page_datacenter, title="Data Center", icon="🗄️", url_path="datacenter", default=True),
    st.Page(page_dashboard, title="Dashboard", icon="📊", url_path="dashboard"),
    st.Page(page_search, title="Seminar Search", icon="🔎", url_path="search"),
    st.Page(page_tables, title="Tables", icon="📋", url_path="tables"),
]
pg = st.navigation(pages, position="hidden")
IS_DATACENTER = pg is pages[0]
nav_pages = pages[1:]

with st.sidebar:
    st.markdown(
        f'<div style="color:{INK};font-size:22px;font-weight:800;line-height:1.1;cursor:pointer;">'
        'UASM Seminars</div>',
        unsafe_allow_html=True,
    )
    st.caption("Facultad de Administración · International Seminars")
    st.page_link(pages[0], label="Back to Data Center", icon=":material/home:")
    st.markdown("---")

if not IS_DATACENTER:
    with st.container(key="nav_toggle"):
        nav_cols = st.columns(len(nav_pages))
        for col, page_obj in zip(nav_cols, nav_pages):
            with col:
                st.page_link(page_obj)

pg.run()

st.markdown(
    f'<div style="text-align:center;padding:40px 0 10px;font-size:11.5px;color:{MUTED};'
    'font-family:monospace;">International Seminars · Facultad de Administración · '
    'Universidad de los Andes</div>',
    unsafe_allow_html=True,
)
