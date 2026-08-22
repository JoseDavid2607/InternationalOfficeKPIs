# ===========================================================================
#  UASM · Internationalization Analytics · App multipágina
#  Basada en la misma arquitectura de UASM_Faculty_Analytics.py: navegación
#  fija arriba, sidebar con logo, tarjetas KPI, Plotly para gráficos.
#  Fuente de datos: Google Drive (Service Account), varios archivos .xlsx.
#
#  Secciones completas por ahora: Faculty, Visiting Faculty, Intl. Weeks.
#  Las demás (Research, Home Campus, Graduates, Agreement Utilization,
#  Mobility by Program, PhD Mobility, Agreements) quedan como "Coming soon",
#  listas para completarse en próximas iteraciones.
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
from typing import Optional, Dict, List

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
    page_title="UASM Internationalization Analytics",
    page_icon="🌐",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Paleta morado metálico elegante (no muy oscuro)
_PURPLE_DEEP = "#4A2E7D"
_PURPLE_MID = "#7B5AAE"
_PURPLE_SOFT = "#A98FD2"
_PURPLE_ACCENT = "#8A63C9"
_PURPLE_BG_TINT = "#F4F0FA"

st.markdown(
    "<style>"
    ".suite-header{display:flex;flex-direction:column;align-items:center;"
    "padding:16px 24px 12px;"
    f"background:linear-gradient(135deg,{_PURPLE_DEEP} 0%,{_PURPLE_MID} 60%,{_PURPLE_SOFT} 100%);"
    "border-radius:12px;box-shadow:0 2px 8px rgba(74,46,125,.22);margin-bottom:14px;}"
    ".sh-super{font-size:11px;font-weight:700;letter-spacing:2px;"
    "color:#E4D9F7;text-transform:uppercase;margin-bottom:2px;}"
    ".sh-title{font-size:26px;font-weight:800;color:#fff;text-align:center;line-height:1.2;}"
    ".sh-sub{font-size:13px;color:rgba(255,255,255,.80);margin-top:4px;text-align:center;}"
    ".kv{font-size:28px;font-weight:800;color:#6941A8;line-height:1.1;}"
    ".kl{font-size:11px;font-weight:600;color:#6B7280;"
    "text-transform:uppercase;letter-spacing:.5px;margin-top:3px;}"
    "section[data-testid='stSidebar']{background:#F7F4FC !important;}"
    "div[data-testid='stButton'] button{background:#FFFFFF !important;"
    "border:1px solid #E1D5F2 !important;border-radius:10px !important;"
    "color:#374151 !important;font-size:14px !important;"
    "font-weight:600 !important;height:48px !important;"
    "box-shadow:0 1px 3px rgba(0,0,0,.04) !important;}"
    "div[data-testid='stButton'] button:hover{"
    "background:#FBF9FE !important;border-color:#C9B4E8 !important;}"
    "div.stDownloadButton>button{background:transparent !important;"
    "border:none !important;box-shadow:none !important;"
    "color:#6941A8 !important;font-size:13px !important;"
    "padding:0 !important;text-decoration:underline !important;}"
    "thead th{background:#EFE7FA !important;color:#4A2E7D !important;"
    "font-weight:700 !important;}"
    ".pending-card{background:#FAF8FD;border:1px dashed #D8C7EF;border-radius:12px;"
    "padding:22px 24px;margin-top:14px;color:#6B7280;font-size:13.5px;}"
    ".pending-card .tag{display:inline-block;font-family:monospace;font-size:10px;"
    "letter-spacing:.06em;text-transform:uppercase;color:#7B5AAE;"
    "background:#EDE3F8;padding:3px 9px;border-radius:5px;margin-bottom:8px;}"
    ".report-link-card{display:flex;align-items:center;justify-content:space-between;"
    "flex-wrap:wrap;gap:12px;background:#FAF8FD;border:1px solid #E7DBF6;"
    "border-radius:12px;padding:16px 20px;margin-top:6px;}"
    ".report-link-card h4{margin:0;color:#3B2560;font-size:15px;}"
    ".st-key-nav_toggle{position:fixed;top:0.25rem;left:50%;transform:translateX(-50%);"
    "z-index:999999;width:82vw;max-width:1050px;}"
    ".st-key-nav_toggle div[data-testid='stHorizontalBlock']{"
    "display:flex !important;flex-wrap:wrap !important;width:100% !important;"
    "justify-content:center !important;gap:4px 18px !important;}"
    ".st-key-nav_toggle div[data-testid='column']{width:auto !important;min-width:fit-content !important;flex:none !important;}"
    ".st-key-nav_toggle div[data-testid='stPageLink']{width:auto !important;min-width:fit-content !important;overflow:visible !important;}"
    ".st-key-nav_toggle div[data-testid='stPageLink'] a{white-space:nowrap !important;overflow:visible !important;text-overflow:unset !important;width:auto !important;min-width:fit-content !important;font-size:13px !important;padding:6px 4px !important;}"
    ".st-key-nav_toggle div[data-testid='stPageLink'] a p{white-space:nowrap !important;overflow:visible !important;}"
    ".st-key-nav_toggle div[data-testid='stButton']{width:auto !important;min-width:fit-content !important;}"
    ".st-key-nav_toggle div[data-testid='stButton'] button{"
    "background:transparent !important;border:none !important;box-shadow:none !important;"
    "color:#6941A8 !important;font-size:13px !important;font-weight:400 !important;"
    "height:auto !important;width:auto !important;padding:6px 4px !important;white-space:nowrap !important;"
    "min-width:fit-content !important;}"
    ".st-key-nav_toggle div[data-testid='stButton'] button:hover{"
    "background:transparent !important;color:#4A2E7D !important;text-decoration:underline;}"
    "</style>",
    unsafe_allow_html=True,
)


# ── 2) HELPERS COMPARTIDOS ──────────────────────────────────────────────
def _render_header(title: str, subtitle: str = ""):
    sub = f'<div class="sh-sub">{subtitle}</div>' if subtitle else ""
    st.markdown(
        f'<div class="suite-header"><div class="sh-super">UASM · Internationalization Analytics</div>'
        f'<div class="sh-title">{title}</div>{sub}</div>',
        unsafe_allow_html=True,
    )


def _kpi(label: str, value):
    st.markdown(
        f'<div class="kv">{value}</div><div class="kl">{label}</div>',
        unsafe_allow_html=True,
    )


def _pending_card(label: str, note: str = ""):
    note_html = f"<br>{note}" if note else ""
    st.markdown(
        f'<div class="pending-card"><span class="tag">Coming soon</span>'
        f'<p style="margin-top:8px;">{label}{note_html}</p></div>',
        unsafe_allow_html=True,
    )


def _section_title_with_link(title: str, url: str, button_label: str, level: str = "###"):
    """Título de sub-sección con el botón 'Go to report' justo al lado, en la misma fila."""
    col1, col2 = st.columns([4, 1])
    with col1:
        st.markdown(f"{level} {title}")
    with col2:
        if url:
            st.link_button(button_label, url, use_container_width=True)
        else:
            st.button(button_label, disabled=True, use_container_width=True,
                       help="Falta configurar la URL de este reporte en el código.")


def _xlsx_bytes(df: pd.DataFrame, sheet_name: str = "Data") -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf) as w:
        df.to_excel(w, index=False, sheet_name=sheet_name[:31])
    buf.seek(0)
    return buf.getvalue()


# ── 3) FILE IDs (Google Drive) ──────────────────────────────────────────
# Carpeta raíz: UASM School of Management > Internationalization
PROFESORES_FILE_ID = "1ncnUk_8VsDt1I0Hui9g0VyoTkA-8P376"   # BD_profesores.xlsx (compartido con Faculty Analytics) → hojas: planta, Info. Profesores, Faculty Distribution
CONVENIOS_FILE_ID = "1tuPLhr8ttvtFKebi1eoWyrL7E7NimNRn"    # BD_convenios.xlsx
MOVILIDAD_FILE_ID = "1zkxMmITgSfpjb1o2AKxKDyqhYdRcpwJw"    # BD_movilidad.xlsx

# EIV (International Summer School) — Reportes/EIV
EIV_CURSOS_FILE_ID = "1qZin81h9oQ4SfxadcNpdjzSO6w3ADKZr"   # BD_cursos.xlsx → Profesor, Universidad, País Universidad, Ciclo

# Seminars — Reportes/Seminars
SEMINARS_FILE_ID = "1YFSUmZ95Md9qHoHvc114Eed-oA_uqjST"     # BD_seminarios.xlsx

# International Weeks — Reportes/INT_WEEKS
WEEKS_LISTAS_FILE_ID = "1Gxev_FWI_mfav3dVWaZczC49F_68Qj9f"  # BD_listas.xlsx → estudiantes UASM asistentes (Producto=universidad, Programa)
WEEKS_SEMANAS_FILE_ID = "1UXmTsOp1X9DKA_OpFy7kmg4fz2_uQT-W" # BD_semanas.xlsx → semanas ofrecidas por la Facultad

# URLs de los reportes externos (HTML estáticos ya existentes). Configúralas
# aquí cuando tengas la URL pública de publicación (p.ej. GitHub Pages).
EIV_REPORT_URL = ""
SEMINARS_REPORT_URL = ""
WEEKS_REPORT_URL = ""

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
    """Descarga un archivo .xlsx de Drive (autenticado con la service account)."""
    token = _get_gspread_access_token()
    if not token:
        if not _GSPREAD_OK:
            st.error(
                "📦 Falta instalar las librerías `google-auth` en el entorno "
                f"(el import falló con: `{_GSPREAD_IMPORT_ERR}`). Agrégalas a "
                "tu `requirements.txt`."
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


def _norm_id(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    try:
        f = float(v)
        return str(int(f)) if f.is_integer() else str(f)
    except (ValueError, TypeError):
        s = str(v).strip().upper()
        return s if s else None


def _period_sort_key(p):
    s = str(p).strip()
    try:
        return (int(s[:4]), 30 if "Intersemestral" in s else int(s[-2:].replace("-", "")))
    except (ValueError, IndexError):
        return (-1, -1)


# Demonym (Country of Birth en BD_profesores) -> país, para el mapa de burbujas.
_NATIONALITY_TO_COUNTRY = {
    "American": "United States", "Argentinian": "Argentina", "Australian": "Australia",
    "Brazilian": "Brazil", "British": "United Kingdom", "Bulgarian": "Bulgaria",
    "Canadian": "Canada", "Chilean": "Chile", "Dominican": "Dominican Republic",
    "Egyptian": "Egypt", "French": "France", "German": "Germany", "Indian": "India",
    "Italian": "Italy", "Kenyan": "Kenya", "New Zealander": "New Zealand",
    "Peruvian": "Peru", "Philippine": "Philippines", "Portuguese": "Portugal",
    "Russian": "Russia", "South African": "South Africa", "Spanish": "Spain",
    "Turkish": "Turkey", "Venezuelan": "Venezuela", "Dutch": "Netherlands",
    "Belgian": "Belgium", "Finnish": "Finland", "Mexican": "Mexico",
}


# ── 4) CARGA DE DATOS ────────────────────────────────────────────────────
@st.cache_data(ttl=300)
def load_planta_fulltime() -> pd.DataFrame:
    """Roster de planta (full-time) leído directamente de la hoja 'planta' de
    BD_profesores.xlsx — trae ya 'Country of Birth'/'Nationality' y 'Double
    Nationality' propias, sin necesidad de merge (mismo patrón que
    demo_load_fulltime() en Faculty Analytics)."""
    raw = io.BytesIO(_download_drive_file_bytes(PROFESORES_FILE_ID))
    df_ = pd.read_excel(raw, sheet_name="planta")
    df_.columns = df_.columns.str.strip()

    sem = df_["Semestre"].astype(str).str.strip() if "Semestre" in df_.columns else df_.iloc[:, 0].astype(str).str.strip()
    is_inter = sem.str.contains("inter", case=False, na=False)
    df_.loc[~is_inter, "Periodo"] = sem.str[:4] + "-" + sem.str[-2:]
    df_.loc[is_inter, "Periodo"] = sem.str[:4] + " Intersemestral"
    df_["Año"] = sem.str[:4]

    if "ID Nr." in df_.columns and "ID" not in df_.columns:
        df_ = df_.rename(columns={"ID Nr.": "ID"})
    return df_


@st.cache_data(ttl=300)
def load_eiv_cursos() -> pd.DataFrame:
    raw = io.BytesIO(_download_drive_file_bytes(EIV_CURSOS_FILE_ID))
    df_ = pd.read_excel(raw, sheet_name="cursos")
    df_.columns = df_.columns.str.strip()
    return df_


@st.cache_data(ttl=300)
def load_seminarios() -> pd.DataFrame:
    raw = io.BytesIO(_download_drive_file_bytes(SEMINARS_FILE_ID))
    df_ = pd.read_excel(raw, sheet_name="seminarios")
    df_.columns = df_.columns.str.strip()
    return df_


@st.cache_data(ttl=300)
def load_weeks_listas() -> pd.DataFrame:
    raw = io.BytesIO(_download_drive_file_bytes(WEEKS_LISTAS_FILE_ID))
    df_ = pd.read_excel(raw, sheet_name="listas")
    df_.columns = df_.columns.str.strip()
    return df_


@st.cache_data(ttl=300)
def load_weeks_semanas() -> pd.DataFrame:
    raw = io.BytesIO(_download_drive_file_bytes(WEEKS_SEMANAS_FILE_ID))
    df_ = pd.read_excel(raw, sheet_name="Semanas Internacionales")
    df_.columns = df_.columns.str.strip()
    return df_


# ── 5) PÁGINA — Faculty (International Full-Time Faculty) ──────────────
def page_faculty():
    df = load_planta_fulltime()

    nat_col = next((c for c in ["Country of Birth", "Nationality"] if c in df.columns), None)
    dual_col = next((c for c in ["Double Nationality", "Doble Nacionalidad"] if c in df.columns), None)
    id_col = "ID" if "ID" in df.columns else None

    _render_header(
        "International Full-Time Faculty",
        "Country of birth and dual nationality of the School of Management's full-time faculty, from the planta roster in BD_profesores.",
    )

    if nat_col is None or id_col is None or df.empty:
        st.info("No pude encontrar las columnas necesarias (Country of Birth / ID) en la hoja 'planta' de BD_profesores.xlsx.")
        return

    with st.sidebar:
        st.markdown("#### Timeframe")
        mode = st.radio("View by", ["Year", "Period"], key="fac_mode")

    if mode == "Period":
        keys = sorted(df["Periodo"].dropna().unique().tolist(), key=_period_sort_key)
        cols_def = [(k, k, df[df["Periodo"] == k]) for k in keys]
    else:
        years = sorted(df["Año"].dropna().unique().tolist())
        cols_def = []
        for y in years:
            periods_in_year = sorted(
                df[df["Año"] == y]["Periodo"].dropna().unique().tolist(), key=_period_sort_key
            )
            if not periods_in_year:
                continue
            latest = periods_in_year[-1]
            cols_def.append((y, y, df[df["Periodo"] == latest]))

    if not cols_def:
        st.info("No hay datos suficientes para mostrar esta vista.")
        return

    labels = [c[1] for c in cols_def]
    with st.sidebar:
        snap_label = st.selectbox("Periodo (snapshot for map & highlight)", labels[::-1], index=0, key="fac_snapshot")

    total_counts, intl_counts, nat_lists, dual_counts_list, dual_nat_lists = [], [], [], [], []
    for key, label, rows in cols_def:
        total_counts.append(rows[id_col].nunique())
        intl_rows = rows[rows[nat_col].notna() & ~rows[nat_col].astype(str).str.strip().str.lower().isin(
            ["colombian", "colombia", ""])]
        intl_counts.append(intl_rows[id_col].nunique())
        nat_lists.append(
            intl_rows.groupby(nat_col)[id_col].nunique().sort_values(ascending=False)
        )
        if dual_col and dual_col in rows.columns:
            dual_rows = rows[rows[dual_col].notna() & ~rows[dual_col].astype(str).str.strip().str.lower().isin(
                ["no", ""])]
            dual_counts_list.append(dual_rows[id_col].nunique())
            dual_nat_lists.append(dual_rows.groupby(dual_col)[id_col].nunique().sort_values(ascending=False))
        else:
            dual_counts_list.append(0)
            dual_nat_lists.append(pd.Series(dtype=int))

    pct_intl = [round((intl_counts[i] / total_counts[i] * 100), 1) if total_counts[i] else 0.0
                for i in range(len(labels))]

    # ---- Tabla 1: Total / Intl / % / # nacionalidades / lista ----
    st.markdown("### International Full-Time Faculty")
    max_nat = max((len(n) for n in nat_lists), default=0)
    table1 = {"": ["Total Faculty", "International Full-time Faculty", "% International",
                   "Number of nationalities"]}
    for i, label in enumerate(labels):
        table1[label] = [total_counts[i], intl_counts[i], f"{pct_intl[i]}%", len(nat_lists[i])]
    df_table1 = pd.DataFrame(table1)
    st.dataframe(df_table1, use_container_width=True, hide_index=True)

    if max_nat:
        nat_rows = []
        for i in range(max_nat):
            row = {"": "Nationalities" if i == 0 else ""}
            for j, label in enumerate(labels):
                nl = nat_lists[j]
                row[label] = f"{nl.index[i]} ({int(nl.iloc[i])})" if i < len(nl) else ""
            nat_rows.append(row)
        st.dataframe(pd.DataFrame(nat_rows), use_container_width=True, hide_index=True)

    st.download_button(
        "Download as Excel", data=_xlsx_bytes(df_table1, "International_Faculty"),
        file_name="International_Faculty.xlsx", key="dl_fac_table1",
    )

    # ---- Trend chart: % international over time ----
    st.markdown("### International Faculty Over Time")
    col_trend, col_donut = st.columns([2, 1])
    with col_trend:
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=labels, y=intl_counts, mode="lines+markers+text", name="International Faculty",
            line=dict(color=_PURPLE_MID, width=3, shape="spline"),
            marker=dict(
                size=[13 if l == snap_label else 7 for l in labels],
                color=[_PURPLE_ACCENT if l == snap_label else _PURPLE_MID for l in labels],
            ),
            text=intl_counts, textposition="top center", textfont=dict(size=10),
        ))
        fig.update_layout(
            margin=dict(t=10, r=16, b=36, l=36), xaxis=dict(type="category"),
            yaxis=dict(gridcolor="#EFEBEE", title="International Faculty"),
            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", height=300,
        )
        st.plotly_chart(fig, use_container_width=True)
    with col_donut:
        snap_idx = labels.index(snap_label)
        donut_df = pd.DataFrame({
            "Group": ["International", "Domestic"],
            "Count": [intl_counts[snap_idx], total_counts[snap_idx] - intl_counts[snap_idx]],
        })
        fig_donut = px.pie(
            donut_df, names="Group", values="Count", hole=0.55,
            color="Group", color_discrete_map={"International": _PURPLE_ACCENT, "Domestic": "#E5DDF2"},
            title=f"{snap_label} snapshot",
        )
        fig_donut.update_traces(textinfo="label+percent", textfont=dict(size=11))
        fig_donut.update_layout(height=300, margin=dict(t=40, r=10, b=10, l=10), showlegend=False)
        st.plotly_chart(fig_donut, use_container_width=True)

    # ---- Nationality bubble map + bar chart for snapshot ----
    st.markdown(f"### Nationalities Represented — {snap_label}")
    snap_nat = nat_lists[snap_idx]
    map_df = snap_nat.rename("Count").reset_index().rename(columns={nat_col: "Nationality"})
    map_df["Country"] = map_df["Nationality"].map(_NATIONALITY_TO_COUNTRY)
    map_df_geo = map_df.dropna(subset=["Country"])

    col_map, col_bar = st.columns([2, 1])
    with col_map:
        if not map_df_geo.empty:
            fig_nat = px.scatter_geo(
                map_df_geo, locations="Country", locationmode="country names", size="Count",
                text="Nationality", hover_name="Nationality", hover_data={"Count": True, "Country": False},
                projection="natural earth",
            )
            fig_nat.update_traces(
                marker=dict(color=_PURPLE_ACCENT, opacity=0.8, line=dict(width=1, color="#FFFFFF")),
                mode="markers+text", textposition="top center", textfont=dict(size=10, color="#3B2560"),
            )
            fig_nat.update_geos(
                showcountries=True, countrycolor="#E5DDF2", showland=True, landcolor=_PURPLE_BG_TINT,
                showocean=True, oceancolor="#EDE7F7", bgcolor="rgba(0,0,0,0)",
            )
            fig_nat.update_layout(height=360, margin=dict(l=10, r=10, t=10, b=6))
            st.plotly_chart(fig_nat, use_container_width=True)
        else:
            st.info("No hay nacionalidades internacionales mapeables para este snapshot.")
    with col_bar:
        if not map_df.empty:
            fig_bar = px.bar(
                map_df.sort_values("Count"), x="Count", y="Nationality", orientation="h",
                color_discrete_sequence=[_PURPLE_MID],
            )
            fig_bar.update_layout(
                margin=dict(t=10, r=16, b=26, l=10), plot_bgcolor="rgba(0,0,0,0)",
                paper_bgcolor="rgba(0,0,0,0)", height=360, yaxis=dict(title=""), xaxis=dict(title="Faculty"),
            )
            st.plotly_chart(fig_bar, use_container_width=True)

    # ---- Tabla 2: Dual nationality ----
    st.markdown("### Full-Time Faculty with Dual Nationality")
    max_dual = max((len(n) for n in dual_nat_lists), default=0)
    table2 = {"": ["Faculty with dual nationality"]}
    for i, label in enumerate(labels):
        table2[label] = [dual_counts_list[i]]
    df_table2 = pd.DataFrame(table2)
    st.dataframe(df_table2, use_container_width=True, hide_index=True)

    if max_dual:
        dual_rows_disp = []
        for i in range(max_dual):
            row = {"": "Nationalities" if i == 0 else ""}
            for j, label in enumerate(labels):
                dn = dual_nat_lists[j]
                row[label] = f"{dn.index[i]} ({int(dn.iloc[i])})" if i < len(dn) else ""
            dual_rows_disp.append(row)
        st.dataframe(pd.DataFrame(dual_rows_disp), use_container_width=True, hide_index=True)

    st.download_button(
        "Download as Excel", data=_xlsx_bytes(df_table2, "Dual_Nationality_Faculty"),
        file_name="Dual_Nationality_Faculty.xlsx", key="dl_fac_table2",
    )


# ── 6) PÁGINA — Visiting Faculty & International Seminars ──────────────
def page_visiting():
    _render_header(
        "Visiting Faculty & International Seminars",
        "Summary of faculty who visited through the International Summer School (EIV) and the International Seminars series — full detail lives in their own reports.",
    )

    # ---- EIV ----
    df_eiv = load_eiv_cursos()
    prof_col = next((c for c in ["Profesor"] if c in df_eiv.columns), None)
    country_col = next((c for c in ["País Universidad", "Pais Universidad"] if c in df_eiv.columns), None)
    ciclo_col = "Ciclo" if "Ciclo" in df_eiv.columns else None

    _section_title_with_link(
        "International Summer School (EIV) — Visiting Faculty", EIV_REPORT_URL, "Go to EIV report →",
    )

    if prof_col and country_col:
        col1, col2 = st.columns(2)
        with col1:
            _kpi("Visiting Faculty (this edition)", df_eiv[prof_col].nunique())
        with col2:
            _kpi("Countries of Origin", df_eiv[country_col].nunique())

        col_chart, col_table = st.columns([3, 2])
        by_country = df_eiv.groupby(country_col)[prof_col].nunique().sort_values(ascending=False).reset_index(
            name="Faculty")
        with col_chart:
            st.markdown("##### Faculty by Country of Origin")
            fig_c = px.bar(
                by_country.sort_values("Faculty"), x="Faculty", y=country_col, orientation="h",
                color_discrete_sequence=[_PURPLE_MID],
            )
            fig_c.update_layout(
                margin=dict(t=6, r=16, b=26, l=140), plot_bgcolor="rgba(0,0,0,0)",
                paper_bgcolor="rgba(0,0,0,0)", height=max(220, 32 * len(by_country)), yaxis=dict(title=""),
            )
            st.plotly_chart(fig_c, use_container_width=True)
        with col_table:
            st.markdown("##### Detail")
            st.dataframe(by_country, use_container_width=True, hide_index=True)

        if ciclo_col:
            by_ciclo = df_eiv.groupby(ciclo_col)[prof_col].nunique().reset_index(name="Visiting Faculty")
            st.markdown("##### Visiting Faculty by Cycle")
            col_ch2, col_tb2 = st.columns([3, 2])
            with col_ch2:
                fig_cy = px.bar(by_ciclo, x=ciclo_col, y="Visiting Faculty",
                                 color_discrete_sequence=[_PURPLE_ACCENT])
                fig_cy.update_layout(
                    margin=dict(t=6, r=16, b=26, l=36), plot_bgcolor="rgba(0,0,0,0)",
                    paper_bgcolor="rgba(0,0,0,0)", height=260, xaxis=dict(type="category"),
                )
                st.plotly_chart(fig_cy, use_container_width=True)
            with col_tb2:
                st.dataframe(by_ciclo, use_container_width=True, hide_index=True)
    else:
        st.info("No pude encontrar las columnas necesarias (Profesor / País Universidad) en BD_cursos.xlsx (EIV).")

    st.markdown("---")

    # ---- Seminars ----
    _section_title_with_link(
        "International Seminars", SEMINARS_REPORT_URL, "Go to Seminars report →",
    )

    df_sem = load_seminarios()
    year_col = "Año" if "Año" in df_sem.columns else None
    speaker_cols = [c for c in df_sem.columns if c.strip().lower().startswith("conferencista")]

    if speaker_cols:
        speakers = set()
        for c in speaker_cols:
            vals = df_sem[c].dropna().astype(str).str.strip()
            speakers.update(v for v in vals if v and v.upper() != "NA")

        col1, col2 = st.columns(2)
        with col1:
            _kpi("Total Seminars (all years)", len(df_sem))
        with col2:
            _kpi("Speakers", len(speakers))

        if year_col:
            by_year = df_sem[year_col].dropna().astype(str).value_counts().sort_index()
            by_year = by_year[~by_year.index.str.contains("REF", case=False, na=False)]
            if not by_year.empty:
                by_year_df = by_year.rename("Seminars").reset_index().rename(columns={"index": "Year"})
                st.markdown("##### Seminars by Year")
                col_cy, col_ty = st.columns([3, 2])
                with col_cy:
                    fig_y = px.bar(by_year_df, x="Year", y="Seminars", color_discrete_sequence=[_PURPLE_MID])
                    fig_y.update_layout(
                        margin=dict(t=6, r=16, b=26, l=36), plot_bgcolor="rgba(0,0,0,0)",
                        paper_bgcolor="rgba(0,0,0,0)", height=260, xaxis=dict(type="category"),
                    )
                    st.plotly_chart(fig_y, use_container_width=True)
                with col_ty:
                    st.dataframe(by_year_df, use_container_width=True, hide_index=True)

        area_col = "Área" if "Área" in df_sem.columns else None
        if area_col:
            by_area = df_sem[area_col].dropna().value_counts().reset_index()
            by_area.columns = ["Area", "Seminars"]
            st.markdown("##### Seminars by Area")
            col_ca, col_ta = st.columns([3, 2])
            with col_ca:
                fig_a = px.pie(by_area, names="Area", values="Seminars", hole=0.5)
                fig_a.update_traces(textinfo="percent+label", textfont=dict(size=10))
                fig_a.update_layout(height=320, margin=dict(t=10, r=10, b=10, l=10), showlegend=False)
                st.plotly_chart(fig_a, use_container_width=True)
            with col_ta:
                st.dataframe(by_area, use_container_width=True, hide_index=True)
    else:
        st.info("No pude encontrar las columnas de conferencistas en BD_seminarios.xlsx.")


# ── 7) PÁGINA — International Weeks ──────────────────────────────────────
def page_weeks():
    _render_header(
        "International Weeks",
        "UASM graduate-program students attending International Weeks abroad, and International Weeks or activities offered by the School.",
    )

    # ---- Students attending int'l weeks abroad ----
    _section_title_with_link(
        "Students Attending International Weeks Abroad", WEEKS_REPORT_URL, "Go to Weeks report →",
    )
    df_listas = load_weeks_listas()
    univ_col = "Producto" if "Producto" in df_listas.columns else None
    prog_col = "Programa" if "Programa" in df_listas.columns else None

    if univ_col and prog_col:
        col1, col2, col3 = st.columns(3)
        with col1:
            _kpi("Total Students", len(df_listas))
        with col2:
            _kpi("Programs Represented", df_listas[prog_col].nunique())
        with col3:
            _kpi("Universities Visited", df_listas[univ_col].nunique())

        rows = []
        for prog, sub in df_listas.groupby(prog_col):
            univ_counts = sub[univ_col].value_counts()
            rows.append({
                "Program": prog,
                "Students": len(sub),
                "Universities": univ_counts.shape[0],
                "University Breakdown": ", ".join(f"{u} ({n})" for u, n in univ_counts.items()),
            })
        table_students = pd.DataFrame(rows).sort_values("Students", ascending=False)

        col_chart, col_table = st.columns([3, 2])
        with col_chart:
            st.markdown("##### Students by Program")
            fig_p = px.bar(
                table_students.sort_values("Students"), x="Students", y="Program", orientation="h",
                color_discrete_sequence=[_PURPLE_MID],
            )
            fig_p.update_layout(
                margin=dict(t=6, r=16, b=26, l=180), plot_bgcolor="rgba(0,0,0,0)",
                paper_bgcolor="rgba(0,0,0,0)", height=max(220, 34 * len(table_students)), yaxis=dict(title=""),
            )
            st.plotly_chart(fig_p, use_container_width=True)
        with col_table:
            st.markdown("##### Detail")
            st.dataframe(table_students, use_container_width=True, hide_index=True)

        univ_counts_all = df_listas[univ_col].value_counts().reset_index()
        univ_counts_all.columns = ["University", "Students"]
        st.markdown("##### Students by Host University")
        fig_u = px.pie(univ_counts_all, names="University", values="Students", hole=0.5)
        fig_u.update_traces(textinfo="percent+label", textfont=dict(size=10))
        fig_u.update_layout(height=340, margin=dict(t=10, r=10, b=10, l=10), showlegend=False)
        st.plotly_chart(fig_u, use_container_width=True)

        st.download_button(
            "Download as Excel", data=_xlsx_bytes(table_students, "International_Weeks_by_Program"),
            file_name="International_Weeks_by_Program.xlsx", key="dl_weeks_students",
        )
    else:
        st.info("No pude encontrar las columnas necesarias (Producto / Programa) en BD_listas.xlsx (Int. Weeks).")

    st.markdown("---")

    # ---- International Weeks offered by the School ----
    st.markdown("### International Weeks or Activities Offered by the School")
    df_semanas = load_weeks_semanas()
    id_col = df_semanas.columns[0] if len(df_semanas.columns) else None
    name_col = "Nombre" if "Nombre" in df_semanas.columns else None
    prof_col = next((c for c in ["Profesor/s", "Profesor"] if c in df_semanas.columns), None)
    fechas_col = "fechas" if "fechas" in df_semanas.columns else None
    asist_col = "Asistentes" if "Asistentes" in df_semanas.columns else None
    country_col = "País" if "País" in df_semanas.columns else None
    city_col = "Ubicación" if "Ubicación" in df_semanas.columns else None

    if name_col:
        col1, col2, col3 = st.columns(3)
        with col1:
            _kpi("Weeks Offered", len(df_semanas))
        with col2:
            _kpi("Countries", df_semanas[country_col].nunique() if country_col else "—")
        with col3:
            _kpi("Total Attendees", int(df_semanas[asist_col].sum()) if asist_col else "—")

        display_cols = [c for c in [id_col, name_col, prof_col, fechas_col, asist_col, city_col, country_col]
                         if c and c in df_semanas.columns]
        st.dataframe(df_semanas[display_cols], use_container_width=True, hide_index=True)

        col_ca, col_cb = st.columns(2)
        if country_col:
            with col_ca:
                by_country = df_semanas[country_col].value_counts().reset_index()
                by_country.columns = ["Country", "Weeks"]
                st.markdown("##### Weeks by Country")
                fig = px.bar(by_country, x="Weeks", y="Country", orientation="h",
                             color_discrete_sequence=[_PURPLE_MID])
                fig.update_layout(
                    margin=dict(t=6, r=16, b=26, l=120), plot_bgcolor="rgba(0,0,0,0)",
                    paper_bgcolor="rgba(0,0,0,0)", height=max(220, 40 * len(by_country)), yaxis=dict(title=""),
                )
                st.plotly_chart(fig, use_container_width=True)
        if asist_col and name_col:
            with col_cb:
                st.markdown("##### Attendees by Week")
                fig_at = px.bar(
                    df_semanas.sort_values(asist_col), x=asist_col, y=name_col, orientation="h",
                    color_discrete_sequence=[_PURPLE_ACCENT],
                )
                fig_at.update_layout(
                    margin=dict(t=6, r=16, b=26, l=160), plot_bgcolor="rgba(0,0,0,0)",
                    paper_bgcolor="rgba(0,0,0,0)", height=max(220, 40 * len(df_semanas)), yaxis=dict(title=""),
                )
                st.plotly_chart(fig_at, use_container_width=True)

        st.download_button(
            "Download as Excel", data=_xlsx_bytes(df_semanas[display_cols], "International_Weeks_Offered"),
            file_name="International_Weeks_Offered.xlsx", key="dl_weeks_offered",
        )
    else:
        st.info("No pude encontrar las columnas necesarias en BD_semanas.xlsx (Int. Weeks).")


# ── 8) PÁGINAS PENDIENTES (placeholders) ────────────────────────────────
def _make_pending_page(title: str, subtitle: str, note: str):
    def _page():
        _render_header(title, subtitle)
        _pending_card("This section is pending.", note)
    return _page


page_research = _make_pending_page(
    "Research & Faculty Activities",
    "International academic journal articles, conference presentations, faculty activities, credit-granting and non-credit courses taught abroad.",
    "Needs a source database with international publications, conference presentations, and faculty visiting-teaching records.",
)
page_homecampus = _make_pending_page(
    "Home Campus",
    "Incoming international students hosted at UASM's home campus.",
    "Needs BD_movilidad.xlsx (incoming) mapped and confirmed for this view.",
)
page_graduates = _make_pending_page(
    "Graduates",
    "International mobility experience among UASM graduates.",
    "Needs BD_grados joined against BD_movilidad / BD_experiencia.",
)
page_agreements_mobility = _make_pending_page(
    "Agreement Utilization",
    "Outgoing & incoming mobility mapped against each international agreement.",
    "Needs BD_movilidad.xlsx cross-referenced with BD_convenios.xlsx.",
)
page_program_mobility = _make_pending_page(
    "Mobility by Program",
    "Outgoing exchange mobility broken down by academic program.",
    "Needs BD_movilidad.xlsx segmented by program and level.",
)
page_phd = _make_pending_page(
    "PhD Mobility",
    "International mobility and research stays among PhD students.",
    "Needs a dedicated PhD mobility source database.",
)
page_agreements = _make_pending_page(
    "Agreements",
    "Full catalog of international agreements — type, status, dates, accreditation.",
    "BD_convenios.xlsx is already loaded; full-table view coming next.",
)


# ── 9) NAVEGACIÓN ─────────────────────────────────────────────────────────
pages = [
    st.Page(page_faculty, title="Faculty", icon="🌎", url_path="faculty", default=True),
    st.Page(page_visiting, title="Visiting Faculty", icon="🧑‍🏫", url_path="visiting"),
    st.Page(page_research, title="Research", icon="🔬", url_path="research"),
    st.Page(page_homecampus, title="Home Campus", icon="🏫", url_path="homecampus"),
    st.Page(page_graduates, title="Graduates", icon="🎓", url_path="graduates"),
    st.Page(page_weeks, title="Intl. Weeks", icon="📅", url_path="weeks"),
    st.Page(page_agreements_mobility, title="Agreement Utilization", icon="🤝", url_path="agreements-mobility"),
    st.Page(page_program_mobility, title="Mobility by Program", icon="🧭", url_path="program-mobility"),
    st.Page(page_phd, title="PhD Mobility", icon="🎯", url_path="phd"),
    st.Page(page_agreements, title="Agreements", icon="📄", url_path="agreements"),
]
pg = st.navigation(pages, position="hidden")

with st.sidebar:
    col_logo, col_title = st.columns([1, 3])
    with col_logo:
        try:
            st.image("imagenes/logo.png", width=65)
        except Exception:
            pass
    with col_title:
        st.markdown(
            '<div style="padding-top:10px;color:#4A2E7D;font-size:22px;'
            'font-weight:800;line-height:1.1;">UASM Intl. KPIs</div>',
            unsafe_allow_html=True,
        )
        st.caption("Internationalization Analytics")
    st.markdown("---")

_NAV_VISIBLE = 6  # cuántas secciones caben directamente antes de agrupar el resto bajo "More"
with st.container(key="nav_toggle"):
    visible_pages = pages[:_NAV_VISIBLE]
    overflow_pages = pages[_NAV_VISIBLE:]
    show_more = st.session_state.get("nav_show_more", False)

    shown_pages = visible_pages + (overflow_pages if show_more else [])
    n_cols = len(shown_pages) + (1 if overflow_pages else 0)
    nav_cols = st.columns(n_cols)

    for col, page_obj in zip(nav_cols, shown_pages):
        with col:
            st.page_link(page_obj)

    if overflow_pages:
        with nav_cols[-1]:
            if st.button("Less" if show_more else "More", key="nav_more_toggle"):
                st.session_state.nav_show_more = not show_more
                st.rerun()

pg.run()
