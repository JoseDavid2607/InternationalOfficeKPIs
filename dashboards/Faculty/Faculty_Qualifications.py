import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import webbrowser
import os
import re
from io import BytesIO

# ================== GENERAL CONFIG ======================
st.set_page_config(
    page_title="Faculty Qualifications",
    page_icon=":bar_chart:",
    layout="wide",
    initial_sidebar_state="expanded",
)

with st.container():
    st.markdown(
        """
        <style>
        .header-title {
            color:#21877D; font-weight:700; text-align:center; font-size:32px;
        }
        .header-btn {
            background-color:#21877D; padding:8px 16px; border:none;
            border-radius:8px; cursor:pointer; font-size:14px;
            display:inline-block;
        }
        a.header-btn, a.header-btn:link, a.header-btn:visited, a.header-btn:hover, a.header-btn:active {
            color:#ffffff !important; text-decoration:none !important;
        }
        .scroll-wrap-600 { max-height:600px; overflow-y:auto; }
        .scroll-wrap-400 { max-height:400px; overflow-y:auto; }
        .scroll-wrap-program { max-height:520px; overflow-y:auto; }
        </style>
        """,
        unsafe_allow_html=True
    )

# ====== DOWNLOAD BUTTONS — MINIMAL STYLE ======
st.markdown("""
<style>
div.stDownloadButton > button {
  background: transparent !important;
  border: none !important;
  box-shadow: none !important;
  color: #21877D !important;
  font-size: 13px !important;
  padding: 0 !important;
  text-decoration: underline !important;
}
div.stDownloadButton { margin: 2px 0 8px 0; }
div.stDownloadButton > button:hover { opacity: 0.9; }
</style>
""", unsafe_allow_html=True)

# ================== HEADER (solo título, sin prev/next) ==================
st.markdown('<div class="header-title">Full-time Faculty Qualifications</div>', unsafe_allow_html=True)

# ================== DATA LOAD ======================
@st.cache_data(ttl=0)
def load_faculty_distribution():
    xls = pd.ExcelFile("data/Faculty/BD_Faculty.xlsx")
    df = pd.read_excel(xls, sheet_name="Faculty Distribution")
    df.columns = df.columns.str.strip()
    return df

@st.cache_data(ttl=0)
def load_cartelera():
    xls = pd.ExcelFile("data/Faculty/BD_Faculty.xlsx")
    df = pd.read_excel(xls, sheet_name="BD Cartelera 2020-2025")
    df.columns = df.columns.str.strip()
    return df

df_fd  = load_faculty_distribution()
df_car = load_cartelera()

# ================== HELPERS =============================
MINT = "#1FA89B"        # Participating
SUPPORTING = "#7FD3FF"  # Supporting (light turquoise)
TOTAL_SERIES_COLOR = "#D09E33"

def _resolve(df: pd.DataFrame, target: str):
    t = target.strip().casefold()
    for c in df.columns:
        if c.strip().casefold() == t: return c
    return None

def _norm_str(s: pd.Series) -> pd.Series:
    return s.astype(str).str.strip().str.lower()

def normalize_ps(val: str) -> str:
    v = str(val).strip().lower()
    if v in {"p","participating","participante","participating faculty"}: return "P"
    if v in {"s","supporting","soporte","supporting faculty"}:            return "S"
    return ""

def normalize_tipo(val: str) -> str:
    v = str(val).strip().lower()
    if v in {"sa","scholarly academics","scholarly academic"}:                 return "SA"
    if v in {"pa","practice academics","practice academic"}:                    return "PA"
    if v in {"sp","scholarly practitioners","scholarly practitioner"}:         return "SP"
    if v in {"ip","instructional practitioners","instructional practitioner"}: return "IP"
    if v in {"o","other","others","otro","otros"}:                             return "OTHER"
    m = re.search(r"\b(sa|pa|sp|ip|o|other)\b", v)
    if m:
        code = m.group(1).upper()
        return "OTHER" if code in {"O","OTHER"} else code
    return "OTHER"

def _get_any(df: pd.DataFrame, *cands) -> str | None:
    for c in cands:
        got = _resolve(df, c)
        if got: return got
    return None

def extract_year_from_period(p: str) -> int | None:
    if p is None: return None
    m = re.search(r"(19|20)\d{2}", str(p))
    return int(m.group(0)) if m else None

def period_suffix(p: str) -> str | None:
    m = re.search(r"(?:19|20)\d{2}[-_/ ]?(\d+)", str(p))
    return m.group(1) if m else None

def is_regular_period(p) -> bool:
    s = str(p).strip().lower()
    if "inter" in s:  # intersemestral
        return False
    suf = period_suffix(s)
    return (suf in {"10", "20"}) or (suf is None)

def list_periods_semestral():
    sem_col = _get_any(df_car, "Semestre", "Periodo", "Periodo Académico", "Periodo academico")
    vals = []
    if sem_col:
        vals = df_car[sem_col].dropna().astype(str).str.strip().tolist()
    regs = [v for v in vals if is_regular_period(v) and period_suffix(v) in {"10","20"}]
    def sort_key(p):
        y = extract_year_from_period(p) or -1
        suf = int(period_suffix(p) or 0)
        return (y, suf)
    return sorted(sorted(set(regs)), key=sort_key)

def list_years_from_sem():
    sem_col = _get_any(df_car, "Semestre", "Periodo", "Periodo Académico", "Periodo academico")
    years = set()
    if sem_col:
        for s in df_car[sem_col].dropna().astype(str):
            y = extract_year_from_period(s)
            if y: years.add(y)
    ycol_fd = _get_any(df_fd, "Year", "Año")
    if ycol_fd:
        for y in pd.to_numeric(df_fd[ycol_fd], errors="coerce").dropna().astype(int):
            years.add(int(y))
    return sorted(years)

def years_with_inter():
    sem_col = _get_any(df_car, "Semestre", "Periodo", "Periodo Académico", "Periodo academico")
    inter = set()
    if sem_col:
        for s in df_car[sem_col].dropna().astype(str):
            if "inter" in s.lower():
                y = extract_year_from_period(s)
                if y: inter.add(y)
    return sorted(inter)

def _slugify(s: str) -> str:
    return re.sub(r'[^A-Za-z0-9]+', '_', str(s)).strip('_')

# —— utilidades de descarga (Excel en memoria) ——
def _sanitize_for_export(df: pd.DataFrame) -> pd.DataFrame:
    return df[[c for c in df.columns if not str(c).startswith("_")]].copy()

def _xlsx_bytes(df: pd.DataFrame, sheet_name: str = "Data") -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf) as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    buf.seek(0)
    return buf.getvalue()

def _download_xlsx_button(df: pd.DataFrame, fname: str, key: str, label: str = "Download Excel"):
    safe = _sanitize_for_export(df)
    clean = re.sub(r"[^\w\sÁÉÍÓÚÜÑáéíóúüñ().%/-]+", "", label).strip()
    st.download_button(
        clean,
        data=_xlsx_bytes(safe),
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=key,
        use_container_width=False
    )

# ================== SENSITIVITY HELPERS (multi-ops) ==================
def build_member_list_for_view(df_period: pd.DataFrame, view_mode: str, col_areaCourse, col_field, program_col) -> list[str]:
    if view_mode == "By Academic Area" and col_areaCourse:
        items = sorted(df_period[col_areaCourse].astype(str).str.strip().dropna().unique().tolist())
    elif view_mode == "By Field" and col_field:
        items = sorted(df_period[col_field].astype(str).str.strip().dropna().unique().tolist())
    elif view_mode == "By Program" and program_col:
        items = sorted(df_period[program_col].astype(str).str.strip().dropna().unique().tolist())
    else:
        items = []
    return ["All"] + items

def apply_ops_to_aggs(agg_ps: pd.DataFrame, agg_tipo: pd.DataFrame, ops: list, member_all_label="All", index_name="member") -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Aplica una lista de operaciones sobre agregados por miembro (fila = área/field/program).
    - scope 'PS': afecta columnas 'P'/'S' sumando créditos (no redistribuye).
    - scope 'QUAL': afecta una de ['SA','SP','IP','PA','OTHER'] sumando créditos.
    """
    mod_ps = agg_ps.copy()
    mod_tipo = agg_tipo.copy()
    for op in ops or []:
        scope = op.get("scope")
        cat = op.get("cat")
        member = op.get("member", member_all_label)
        delta = float(op.get("credits", 0.0)) * int(op.get("count", 0))
        if delta == 0:
            continue

        if scope == "PS":
            if cat not in ["P", "S"]:
                continue
            if cat not in mod_ps.columns:
                mod_ps[cat] = 0.0
            if member == member_all_label:
                mod_ps[cat] = (mod_ps[cat] + delta).clip(lower=0.0)
            else:
                if member in mod_ps.index:
                    mod_ps.at[member, cat] = max(0.0, float(mod_ps.at[member, cat]) + delta)

        elif scope == "QUAL":
            cats = ["SA","SP","IP","PA","OTHER"]
            if cat not in cats:
                continue
            if cat not in mod_tipo.columns:
                mod_tipo[cat] = 0.0
            if member == member_all_label:
                mod_tipo[cat] = (mod_tipo[cat] + delta).clip(lower=0.0)
            else:
                if member in mod_tipo.index:
                    mod_tipo.at[member, cat] = max(0.0, float(mod_tipo.at[member, cat]) + delta)
    return mod_ps, mod_tipo

def impact_column_generic(base_df: pd.DataFrame, mod_df: pd.DataFrame, target: str, mode: str) -> pd.Series:
    """
    Impacto como Δ%target por fila.
    - mode="PS": target in {"P","S"}; % sobre (P+S).
    - mode="QUAL": target in {"SA","SP","IP","PA","OTHER"}; % sobre suma QUAL.
    """
    if mode == "PS":
        den0 = (base_df.get("P",0) + base_df.get("S",0)).replace(0, pd.NA)
        den1 = (mod_df.get("P",0) + mod_df.get("S",0)).replace(0, pd.NA)
        pct0 = (base_df[target] / den0 * 100).fillna(0.0)
        pct1 = (mod_df[target]  / den1 * 100).fillna(0.0)
        return (pct1 - pct0).round(2)
    else:
        cats = ["SA","SP","IP","PA","OTHER"]
        for c in cats:
            if c not in base_df.columns: base_df[c] = 0.0
            if c not in mod_df.columns:  mod_df[c]  = 0.0
        den0 = base_df[cats].sum(axis=1).replace(0, pd.NA)
        den1 = mod_df[cats].sum(axis=1).replace(0, pd.NA)
        pct0 = (base_df[target] / den0 * 100).fillna(0.0)
        pct1 = (mod_df[target]  / den1 * 100).fillna(0.0)
        return (pct1 - pct0).round(2)

# ================== SIDEBAR ==================
SEMESTRAL_PERIODS = list_periods_semestral()
YEARS_ALL = list_years_from_sem()
INTER_YEARS = years_with_inter()

with st.sidebar:
    # TOP: KPI selector (en inglés)
    st.markdown("### 📊 Go to KPI")
    options = {
        "1 Full-time Composition": "https://facultycompositiondashboardpy-dtacyzfa3otmpbewqc5axu.streamlit.app/",
        "2 Full-time Staffing Levels": "https://facultystaffinglevelsdashboardpy-phv4t8jzbyyz5rrepqttuf.streamlit.app/",
        "3 Distribution by Academic Area": "https://facultydistributionareadashboardpy-yzwpiqdlukfdp6qcygxjhj.streamlit.app/",
        "4 Faculty Demographics": "https://facultydemographicsdashboardpy-kmsnpswxs35psbqtdtvb6y.streamlit.app/",
        "5 Full-time Faculty Questionnaire": "https://full-timefacultyactivitiespy-bbe7fmmyrxvssadnygm4fx.streamlit.app/",
        "6 Faculty Qualifications": "https://facultyqualificationspy-drvj3wpyrxvm2lrnafdwx5.streamlit.app/",
    }
    choices = [k for k, u in options.items() if isinstance(u, str) and (u.startswith("http://") or u.startswith("https://"))]
    default_label = "6 Faculty Qualifications"
    default_idx = choices.index(default_label) if default_label in choices else 0
    choice = st.selectbox("Select…", choices, index=default_idx, key="kpi_nav_top")
    st.link_button("Open", options[choice], use_container_width=True)

    # Timeframe & View (al medio)
    st.markdown('---')
    st.markdown("#### Timeframe")
    st.session_state.setdefault("time_mode", "Semestral")
    time_mode = st.radio(
        "Timeframe",
        ["Semestral", "Anual", "Intersemestral"],
        key="time_mode",
        label_visibility="collapsed",
        horizontal=False
    )

    if time_mode == "Semestral":
        default_sem = SEMESTRAL_PERIODS[-1] if SEMESTRAL_PERIODS else "202510"
        st.session_state.setdefault("sel_sem", default_sem)
        sel_sem = st.selectbox("Semester", SEMESTRAL_PERIODS or [default_sem], key="sel_sem")
        sel_year = extract_year_from_period(sel_sem) or (YEARS_ALL[-1] if YEARS_ALL else 2025)
        sel_label = str(sel_sem)
    elif time_mode == "Anual":
        default_year = YEARS_ALL[-1] if YEARS_ALL else 2025
        st.session_state.setdefault("sel_year", default_year)
        sel_year = st.selectbox("Year", YEARS_ALL or [default_year], key="sel_year")
        sel_sem = None
        sel_label = f"{sel_year} (Annual)"
    else:
        default_i = INTER_YEARS[-1] if INTER_YEARS else (YEARS_ALL[-1] if YEARS_ALL else 2025)
        st.session_state.setdefault("sel_inter_year", default_i)
        sel_year = st.selectbox("Year (Intersemestral)", INTER_YEARS or YEARS_ALL or [default_i], key="sel_inter_year")
        sel_sem = None
        sel_label = f"{sel_year} Intersemestral"

    st.session_state.setdefault("view_mode", "By Academic Area")
    view_mode = st.selectbox("View", ["By Program", "By Academic Area", "By Field"], key="view_mode")

    # Export placeholder (no mover)
    dl_bd_placeholder = st.empty()

    # BOTTOM: Sensitivity (toggle + setup)
    st.markdown('---')
    st.markdown("#### Sensitivity analysis")
    sens_mode = st.toggle("Enable sensitivity mode", value=False, help="Switch dashboard to sensitivity view")

    if sens_mode:
        st.caption("Add hypothetical professors to see the impact by Area/Field/Program.")

        # Init contenedor de operaciones acumuladas
        if "sens_ops" not in st.session_state:
            st.session_state.sens_ops = []   # cada op: {"scope":"PS"|"QUAL", "cat":"P"/"S"/"SA"/"SP"/"IP"/"PA"/"OTHER", "member":"...", "credits":float, "count":int}

        target_group = st.radio("Target group", ["P/S", "Qualifications"], horizontal=True, key="sens_target_group")

        if target_group == "P/S":
            cat = st.selectbox("Category", ["P", "S"], key="sens_cat_ps")
        else:
            cat = st.selectbox("Category", ["SA", "SP", "IP", "PA", "OTHER"], key="sens_cat_qual")

        # El selector "Apply to" se pobla luego del filtrado del período (abajo)
        st.session_state.setdefault("sens_member", "All")
        sens_member_placeholder = st.empty()

        count = st.number_input("Professors to add", min_value=1, step=1, value=1, key="sens_count")
        credits = st.number_input("Credits per professor", min_value=0.0, step=0.5, value=8.0, key="sens_credits")

        col_add, col_reset = st.columns(2)
        with col_add:
            if st.button("Add", type="primary", use_container_width=True):
                scope = "PS" if target_group == "P/S" else "QUAL"
                member = st.session_state.get("sens_member", "All")
                st.session_state.sens_ops.append({
                    "scope": scope,
                    "cat": cat,
                    "member": member,
                    "credits": float(credits),
                    "count": int(count),
                })
                st.experimental_rerun()
        with col_reset:
            if st.button("Reset to original", use_container_width=True):
                st.session_state.sens_ops = []
                st.experimental_rerun()

# ------------- Normalizadores de columnas básicas -------------
col_sem  = _get_any(df_car, "Semestre","Periodo","Periodo Académico","Periodo academico")
if "_SEM" not in df_car.columns and col_sem:
    df_car["_SEM"] = df_car[col_sem].astype(str).str.strip()
else:
    df_car["_SEM"] = df_car.get("_SEM", pd.Series(dtype=str))
df_car["_YEAR"] = df_car["_SEM"].map(extract_year_from_period)
df_car["_IS_INTER"] = df_car["_SEM"].str.lower().str.contains("inter", na=False)

# ================== FUNCIONES DE FILTRO ==================
def mask_timeframe(series_sem: pd.Series, mode: str, selected_year: int | None, selected_sem: str | None) -> pd.Series:
    s = series_sem.astype(str)
    if mode == "Semestral" and selected_sem:
        return s.str.strip().eq(str(selected_sem))
    if mode == "Anual" and selected_year is not None:
        return s.str.startswith(str(selected_year))
    if mode == "Intersemestral" and selected_year is not None:
        return s.str.startswith(str(selected_year)) & s.str.lower().str.contains("inter")
    return pd.Series([True]*len(s), index=series_sem.index)

def filter_df_car(df: pd.DataFrame, mode: str, selected_year: int | None, selected_sem: str | None) -> pd.DataFrame:
    if "_SEM" not in df.columns:
        sc = _get_any(df, "Semestre","Periodo","Periodo Académico","Periodo academico")
        if sc: df = df.assign(_SEM=df[sc].astype(str).str.strip())
        else: return df
    m = mask_timeframe(df["_SEM"], mode, selected_year, selected_sem)
    return df[m].copy()

def filter_df_fd(df: pd.DataFrame, mode: str, selected_year: int | None, selected_sem: str | None) -> pd.DataFrame:
    semc = _get_any(df, "Semestre","Periodo","Periodo Académico","Periodo academico")
    ycol = _get_any(df, "Year","Año")
    out = df.copy()
    if semc:
        sem_series = out[semc].astype(str).str.strip()
        m = mask_timeframe(sem_series, mode, selected_year, selected_sem)
        out = out[m].copy()
    elif ycol and selected_year is not None:
        out = out[pd.to_numeric(out[ycol], errors="coerce").astype("Int64") == int(selected_year)].copy()
    return out

# ================== PRE-FILTROS ==================
df_car_base = df_car.copy()
df_fd_base  = df_fd.copy()

df_car_filt_all = filter_df_car(df_car_base, time_mode, sel_year, sel_sem)
df_fd_f          = filter_df_fd(df_fd_base, time_mode, sel_year, sel_sem)

# Sidebar: botón de descarga de la BD filtrada
if 'dl_bd_placeholder' in locals():
    safe = _sanitize_for_export(df_car_filt_all)
    dl_bd_placeholder.download_button(
        "Download DB (Excel)",
        data=_xlsx_bytes(safe),
        file_name=f"BD_Cartelera_{_slugify(sel_label)}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"dl_bd_sidebar_main_{_slugify(sel_label)}_{st.session_state.get('time_mode','NA')}_{'sens' if sens_mode else 'hist'}"
    )

# ================== PRE-FILTROS ==================
df_car_base = df_car.copy()
df_fd_base  = df_fd.copy()

df_car_filt_all = filter_df_car(df_car_base, time_mode, sel_year, sel_sem)
df_fd_f          = filter_df_fd(df_fd_base, time_mode, sel_year, sel_sem)

# Sidebar: botón de descarga de la BD filtrada según el período seleccionado
if 'dl_bd_placeholder' in locals():
    safe = _sanitize_for_export(df_car_filt_all)
    dl_bd_placeholder.download_button(
        "Download DB (Excel)",
        data=_xlsx_bytes(safe),
        file_name=f"BD_Cartelera_{_slugify(sel_label)}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"dl_bd_{_slugify(sel_label)}"
    )

# ====== SENSITIVITY: preparar lista de miembros y variables específicas ======
if 'sens_mode' in locals() and sens_mode:
    col_areaCourse = _get_any(df_car_filt_all, "Area del curso","Área del curso","Area del Curso","AREA DEL CURSO")
    col_field = _get_any(df_car_filt_all, "Field","FIELD","Campo","Área de conocimiento")
    program_col = _get_any(df_car_filt_all, "Program","PROGRAM","program","Materia")

    # Poblar selector "Apply to" según View (EN INGLÉS) — SIN asignar a session_state manualmente
    members = build_member_list_for_view(df_car_filt_all.copy(), view_mode, col_areaCourse, col_field, program_col)
    with st.sidebar:
        st.selectbox("Apply to", members, key="sens_member")

    # Marcador de que hay sensibilidad activa y lista de operaciones
    SENS = {"on": True, "ops": st.session_state.get("sens_ops", [])}
else:
    SENS = {"on": False, "ops": []}



# ================== RELEVANT COLUMNS ==================
col_ps_fd   = _get_any(df_fd_f, "P/S", "P - S", "Participating/Supporting")
col_area_fd = _get_any(df_fd_f, "AREA_PROFESOR", "Area_Profesor", "Area Profesor", "Área", "Area")
col_tipo_fd = _get_any(df_fd_f, "TIPO", "Tipo", "Ranking", "Tipo Ranking")

col_cred  = _get_any(df_car, "Créditos", "Creditos", "Credits")
col_tipoC = _get_any(df_car, "TIPO", "Tipo", "Tipo Ranking")
col_areaC = _get_any(df_car, "AREA_PROFESOR", "Area_Profesor", "Area Profesor", "Área", "Area")
col_ps_C  = _get_any(df_car, "P/S", "P - S", "Participating/Supporting")
col_areaCourse = _get_any(df_car, "Area del curso","Área del curso","Area del Curso","AREA DEL CURSO")
col_prof = _get_any(df_car, "Profesor","PROFESOR","Docente")
col_code = _get_any(df_car, "Código Materia","Codigo Materia","CODIGO MATERIA","Código","Codigo","Course Code")
col_name = _get_any(df_car, "Nombre largo curso","Nombre Curso","Nombre del curso","Course Name")
col_field = _get_any(df_car, "Field","FIELD","Campo","Área de conocimiento")
col_prog = _get_any(df_car, "Program","PROGRAM","program","Materia")

# ========== NORMALIZATION ==========
if col_ps_fd:
    df_fd_f["_PS"] = _norm_str(df_fd_f[col_ps_fd]).map(normalize_ps)
if col_area_fd:
    df_fd_f["_AREA"] = df_fd_f[col_area_fd].astype(str).str.strip()
if col_tipo_fd:
    df_fd_f["_TIPO"] = _norm_str(df_fd_f[col_tipo_fd]).map(normalize_tipo)

# ---------- Utilities for stylers (rules) ----------
def style_percent_tables(df_, id_col):
    sty = pd.DataFrame('', index=df_.index, columns=df_.columns)
    colP = "%P"; colSA = "%SA"; colOTHER = "%OTHER"

    p_vals     = pd.to_numeric(df_[colP], errors="coerce")
    sa_vals    = pd.to_numeric(df_[colSA], errors="coerce")
    other_vals = pd.to_numeric(df_[colOTHER], errors="coerce")
    is_total   = df_[id_col].astype(str).str.upper().eq("TOTAL")

    sty.loc[(~is_total) & (p_vals < 60), colP] = 'background-color:#FDE2E2;'
    sty.loc[is_total & (p_vals < 75), colP]    = 'background-color:#FDE2E2; font-weight:700;'
    sty.loc[sa_vals < 40, colSA]               = 'background-color:#FDE2E2;'
    sty.loc[other_vals > 10, colOTHER]         = 'background-color:#FDE2E2;'

    for c in sty.columns:
        sty.loc[is_total, c] = (sty.loc[is_total, c].astype(str) + 'font-weight:700;').str.replace(';;',';', regex=False)
    return sty

# ================== SECCIÓN PRINCIPAL ==================
st.markdown("---")
st.subheader(f"Faculty Sufficiency and Qualifications — {sel_label}")

if not all([col_cred, col_tipoC, col_areaCourse]):
    st.error("Missing columns in 'BD Cartelera 2020-2025': 'Credits', 'TIPO', and/or 'Academic Area (course)'.")
else:
    # ---------- Base normalization ----------
    df_car_n = df_car.copy()
    df_car_n["_CRED"]  = pd.to_numeric(df_car_n[col_cred], errors="coerce").fillna(0.0)
    df_car_n["_TIPO"]  = _norm_str(df_car_n[col_tipoC]).map(normalize_tipo)
    if "_SEM" not in df_car_n.columns:
        sc = _get_any(df_car_n, "Semestre","Periodo","Periodo Académico","Periodo academico")
        df_car_n["_SEM"] = df_car_n[sc].astype(str).str.strip() if sc else ""
    df_car_n["_YEAR"] = df_car_n["_SEM"].map(extract_year_from_period)
    df_car_n["_IS_INTER"] = df_car_n["_SEM"].str.lower().str.contains("inter", na=False)
    df_car_n["_AREA"]  = df_car_n[col_areaCourse].astype(str).str.strip()
    col_ps_C_local     = _get_any(df_car_n, "P/S","P - S","Participating/Supporting")
    df_car_n["_PS"]    = _norm_str(df_car_n[col_ps_C_local]).map(normalize_ps) if col_ps_C_local else ""

    # Excluir materias específicas (global)
    program_col = _get_any(df_car_n, "program")
    EXCLUDE_SUBJ = {"CONT", "E-IMER", "E-ENEG", "E-AFIN"}
    if program_col:
        mask_ok = ~df_car_n[program_col].astype(str).str.strip().str.upper().isin(EXCLUDE_SUBJ)
        df_car_global = df_car_n[mask_ok].copy()
    else:
        df_car_global = df_car_n.copy()

    # ---------- Filtro por timeframe seleccionado ----------
    fil = filter_df_car(df_car_global, time_mode, sel_year, sel_sem)

    if fil.empty:
        st.info(f"No records for the selected timeframe: {sel_label}.")
    else:
        # ============================================================
        #      VISTA SEGÚN SELECCIÓN
        # ============================================================
        def build_percent_table(base_idx_name, agg_tipo, agg_ps):
            den_ps   = (agg_ps["P"] + agg_ps["S"]).replace(0, pd.NA)
            p_share  = (agg_ps["P"] / den_ps) * 100
            s_share  = 100 - p_share
            denom_q  = (agg_tipo.sum(axis=1)).replace(0, pd.NA)

            dfm = pd.DataFrame({
                base_idx_name: agg_tipo.index,
                "%P":  p_share,
                "%S":  s_share,
                "%SA": (agg_tipo["SA"] / denom_q) * 100,
                "%OTHER": (agg_tipo["OTHER"] / denom_q) * 100,
            }).fillna(0.0)

            tot_P, tot_S = agg_ps["P"].sum(), agg_ps["S"].sum()
            tot_den_ps   = tot_P + tot_S
            p_tot = (tot_P / tot_den_ps * 100) if tot_den_ps else 0.0
            s_tot = 100 - p_tot
            tipo_sums = agg_tipo[["SA","PA","SP","IP","OTHER"]].sum(axis=0)
            denom_q_tot = float(tipo_sums.sum())

            total_row = {
                base_idx_name: "TOTAL",
                "%P":  round(p_tot, 1),
                "%S":  round(s_tot, 1),
                "%SA": round((tipo_sums["SA"] / denom_q_tot * 100) if denom_q_tot else 0.0, 1),
                "%OTHER": round((tipo_sums["OTHER"] / denom_q_tot * 100) if denom_q_tot else 0.0, 1),
            }
            dfm[["%P","%S","%SA","%OTHER"]] = dfm[["%P","%S","%SA","%OTHER"]].round(1)
            dfm = pd.concat([dfm, pd.DataFrame([total_row])], ignore_index=True)
            return dfm[[base_idx_name, "%P", "%S", "%SA", "%OTHER"]]

        # --------- helpers para eje temporal según modo ---------
        def build_time_axis_for_history(df_hist: pd.DataFrame):
            if time_mode == "Semestral":
                reg = sorted(
                    {s for s in df_hist["_SEM"].dropna().unique() if period_suffix(s) in {"10","20"}},
                    key=lambda s: (extract_year_from_period(s) or -1, int(period_suffix(s) or 0))
                )
                x_map = {s:i for i, s in enumerate(reg)}
                return "_SEM", reg, x_map
            if time_mode == "Anual":
                years = sorted({extract_year_from_period(s) for s in df_hist["_SEM"] if extract_year_from_period(s)}, key=int)
                x_map = {y:i for i, y in enumerate(years)}
                return "_YEAR", years, x_map
            inter = sorted({f"{extract_year_from_period(s)} Intersemestral"
                            for s in df_hist["_SEM"] if "inter" in str(s).lower()
                            and extract_year_from_period(s)}, key=lambda x: int(x.split()[0]))
            x_map = {lab:i for i, lab in enumerate(inter)}
            return "_INTER_LABEL", inter, x_map

        def transform_for_time_mode_ps(df_ps: pd.DataFrame):
            base = df_ps.copy()
            base["_YEAR"] = base["_SEM"].map(extract_year_from_period)
            base["_INTER_LABEL"] = base["_SEM"].map(lambda s: f"{extract_year_from_period(s)} Intersemestral" if "inter" in str(s).lower() else None)
            if time_mode == "Semestral":
                return base
            if time_mode == "Anual":
                need_cols = [c for c in base.columns if c not in {"P_share"}]
                g = base[need_cols].groupby(["_YEAR"] + [c for c in base.columns if c.startswith("_") and c not in {"_SEM","_YEAR","_INTER_LABEL"}], dropna=False).sum(numeric_only=True).reset_index()
                if "P" in g and "S" in g:
                    g["P_share"] = (g["P"] / (g["P"] + g["S"]).replace(0, pd.NA)) * 100
                return g.rename(columns={"_YEAR":"_SEM"})
            base = base[~base["_INTER_LABEL"].isna()].copy()
            g = base.groupby(["_INTER_LABEL"] + [c for c in base.columns if c.startswith("_") and c not in {"_SEM","_YEAR","_INTER_LABEL"}], dropna=False).sum(numeric_only=True).reset_index()
            if "P" in g and "S" in g:
                g["P_share"] = (g["P"] / (g["P"] + g["S"]).replace(0, pd.NA)) * 100
            g = g.rename(columns={"_INTER_LABEL":"_SEM"})
            return g

        def transform_for_time_mode_tipo(df_tipo: pd.DataFrame, share_col_name: str):
            base = df_tipo.copy()
            base["_YEAR"] = base["_SEM"].map(extract_year_from_period)
            base["_INTER_LABEL"] = base["_SEM"].map(lambda s: f"{extract_year_from_period(s)} Intersemestral" if "inter" in str(s).lower() else None)
            cats = ["SA","PA","SP","IP","OTHER"]
            if time_mode == "Semestral":
                return base
            if time_mode == "Anual":
                keys = ["_YEAR"] + [c for c in base.columns if c.startswith("_") and c not in {"_SEM","_YEAR","_INTER_LABEL"}]
                g = base.groupby(keys, dropna=False)[cats].sum().reset_index()
                den = (g[cats].sum(axis=1)).replace(0, pd.NA)
                if share_col_name == "SA_share":
                    g["SA_share"] = (g["SA"] / den) * 100
                else:
                    g["OTHER_share"] = (g["OTHER"] / den) * 100
                return g.rename(columns={"_YEAR":"_SEM"})
            base = base[~base["_INTER_LABEL"].isna()].copy()
            keys = ["_INTER_LABEL"] + [c for c in base.columns if c.startswith("_") and c not in {"_SEM","_YEAR","_INTER_LABEL"}]
            g = base.groupby(keys, dropna=False)[cats].sum().reset_index()
            den = (g[cats].sum(axis=1)).replace(0, pd.NA)
            if share_col_name == "SA_share":
                g["SA_share"] = (g["SA"] / den) * 100
            else:
                g["OTHER_share"] = (g["OTHER"] / den) * 100
            return g.rename(columns={"_INTER_LABEL":"_SEM"})

        def draw_history(fig_title, level_name, level_values, metric_kind, total_series_builders,
                         agg_ps_all, agg_tipo_all, x_labels, x_map, sel_x):
            palette = px.colors.qualitative.Safe + px.colors.qualitative.Bold + px.colors.qualitative.Pastel
            color_map = {a: palette[i % len(palette)] for i, a in enumerate(level_values)}

            st.markdown(f"<h4 style='margin:0 0 6px 0; font-weight:500;'>{fig_title}</h4>", unsafe_allow_html=True)
            sel_col, radio_col = st.columns([6,4])
            options = ["(All)", "(TOTAL)"] + level_values
            with sel_col:
                opt = st.selectbox("", options, index=0, key=f"{level_name}_filter", label_visibility="collapsed")
            with radio_col:
                metric_choice = st.radio("", ["%P", "%SA", "%OTHER"],
                                         index={ "%P":0, "%SA":1, "%OTHER":2 }[metric_kind],
                                         horizontal=True, key=f"metric_{level_name}",
                                         label_visibility="collapsed")

            fig = go.Figure()

            if metric_choice == "%P":
                thr = 75 if opt == "(TOTAL)" else 60
                if opt == "(All)":
                    for a in level_values:
                        sub = agg_ps_all[(agg_ps_all[level_name] == a)].copy()
                        sub["x"] = sub["_SEM"].map(x_map)
                        sub = sub.sort_values("x")
                        if sub.empty: continue
                        fig.add_trace(go.Scatter(
                            x=sub["x"], y=sub["P_share"], mode="lines+markers", name=a,
                            marker=dict(size=6, color=color_map[a]),
                            line=dict(width=2, color=color_map[a]),
                            hovertemplate=a + "<br>%{y:.1f}%<extra></extra>"
                        ))
                elif opt == "(TOTAL)":
                    sub = total_series_builders["P"].copy()
                    sub["x"] = sub["_SEM"].map(x_map)
                    sub = sub.sort_values("x")
                    fig.add_trace(go.Scatter(
                        x=sub["x"], y=sub["P_share"], mode="lines+markers", name="TOTAL",
                        marker=dict(size=6, color=TOTAL_SERIES_COLOR),
                        line=dict(width=2, color=TOTAL_SERIES_COLOR),
                        hovertemplate="TOTAL<br>%{y:.1f}%<extra></extra>"
                    ))
                else:
                    sub = agg_ps_all[(agg_ps_all[level_name] == opt)].copy()
                    sub["x"] = sub["_SEM"].map(x_map)
                    sub = sub.sort_values("x")
                    fig.add_trace(go.Scatter(
                        x=sub["x"], y=sub["P_share"], mode="lines+markers", name=opt,
                        marker=dict(size=6, color=MINT), line=dict(width=2, color=MINT),
                        hovertemplate=opt + "<br>%{y:.1f}%<extra></extra>"
                    ))
                y_min = 40
                bad_high = False

            elif metric_choice == "%SA":
                thr = 40
                share_col = "SA_share"
                if opt == "(All)":
                    for a in level_values:
                        sub = agg_tipo_all[(agg_tipo_all[level_name] == a)].copy()
                        sub["x"] = sub["_SEM"].map(x_map)
                        sub = sub.sort_values("x")
                        if sub.empty: continue
                        fig.add_trace(go.Scatter(
                            x=sub["x"], y=sub[share_col], mode="lines+markers", name=a,
                            marker=dict(size=6, color=color_map[a]),
                            line=dict(width=2, color=color_map[a]),
                            hovertemplate=a + "<br>%{y:.1f}%<extra></extra>"
                        ))
                elif opt == "(TOTAL)":
                    sub = total_series_builders["SA"].copy()
                    sub["x"] = sub["_SEM"].map(x_map)
                    sub = sub.sort_values("x")
                    fig.add_trace(go.Scatter(
                        x=sub["x"], y=sub[share_col], mode="lines+markers", name="TOTAL",
                        marker=dict(size=6, color=TOTAL_SERIES_COLOR),
                        line=dict(width=2, color=TOTAL_SERIES_COLOR),
                        hovertemplate="TOTAL<br>%{y:.1f}%<extra></extra>"
                    ))
                else:
                    sub = agg_tipo_all[(agg_tipo_all[level_name] == opt)].copy()
                    sub["x"] = sub["_SEM"].map(x_map)
                    sub = sub.sort_values("x")
                    fig.add_trace(go.Scatter(
                        x=sub["x"], y=sub[share_col], mode="lines+markers", name=opt,
                        marker=dict(size=6, color=MINT), line=dict(width=2, color=MINT),
                        hovertemplate=opt + "<br>%{y:.1f}%<extra></extra>"
                    ))
                y_min = 20
                bad_high = False

            else:  # "%OTHER"
                thr = 10
                share_col = "OTHER_share"
                if opt == "(All)":
                    for a in level_values:
                        sub = agg_tipo_all[(agg_tipo_all[level_name] == a)].copy()
                        sub["x"] = sub["_SEM"].map(x_map)
                        sub = sub.sort_values("x")
                        if sub.empty: continue
                        fig.add_trace(go.Scatter(
                            x=sub["x"], y=sub[share_col], mode="lines+markers", name=a,
                            marker=dict(size=6, color=color_map[a]),
                            line=dict(width=2, color=color_map[a]),
                            hovertemplate=a + "<br>%{y:.1f}%<extra></extra>"
                        ))
                elif opt == "(TOTAL)":
                    sub = total_series_builders["OTHER"].copy()
                    sub["x"] = sub["_SEM"].map(x_map)
                    sub = sub.sort_values("x")
                    fig.add_trace(go.Scatter(
                        x=sub["x"], y=sub[share_col], mode="lines+markers", name="TOTAL",
                        marker=dict(size=6, color=TOTAL_SERIES_COLOR),
                        line=dict(width=2, color=TOTAL_SERIES_COLOR),
                        hovertemplate="TOTAL<br>%{y:.1f}%<extra></extra>"
                    ))
                else:
                    sub = agg_tipo_all[(agg_tipo_all[level_name] == opt)].copy()
                    sub["x"] = sub["_SEM"].map(x_map)
                    sub = sub.sort_values("x")
                    fig.add_trace(go.Scatter(
                        x=sub["x"], y=sub[share_col], mode="lines+markers", name=opt,
                        marker=dict(size=6, color=MINT), line=dict(width=2, color=MINT),
                        hovertemplate=opt + "<br>%{y:.1f}%<extra></extra>"
                    ))
                y_min = 0
                y_max = 40
                bad_high = True

            if bad_high:
                fig.update_layout(shapes=[dict(type="rect", xref="paper", yref="y", x0=0, x1=1, y0=thr, y1=100,
                                               fillcolor="#FDE2E2", opacity=0.35, layer="below", line_width=0)])
                fig.add_hline(y=thr, line_color="#F5A3A3", line_dash="dash")
            else:
                fig.update_layout(shapes=[dict(type="rect", xref="paper", yref="y", x0=0, x1=1, y0=0, y1=thr,
                                               fillcolor="#FDE2E2", opacity=0.35, layer="below", line_width=0)])
                fig.add_hline(y=thr, line_color="red", line_dash="dash")

            if sel_x is not None:
                fig.add_vrect(x0=sel_x-0.5, x1=sel_x+0.5,
                              fillcolor="#E8FAF7", opacity=0.5, layer="below", line_width=0)

            if metric_choice == "%OTHER":
                fig.update_layout(
                    xaxis=dict(tickmode="array", tickvals=list(range(len(x_labels))), ticktext=[str(x) for x in x_labels]),
                    yaxis=dict(range=[y_min, y_max]),
                )
            else:
                fig.update_layout(
                    xaxis=dict(tickmode="array", tickvals=list(range(len(x_labels))), ticktext=[str(x) for x in x_labels]),
                    yaxis=dict(range=[y_min, 100]),
                )
            fig.update_xaxes(title=None)
            fig.update_yaxes(title=None)

            # Render gráfico
            st.plotly_chart(fig, use_container_width=True)

            # ===== Descarga de datos del gráfico (abajo) =====
            # Construimos un DF con las series visibles según selección
            def _series_for(level_val: str, ycol: str):
                if ycol == "P_share":
                    sub = agg_ps_all[(agg_ps_all[level_name] == level_val)]
                else:
                    sub = agg_tipo_all[(agg_tipo_all[level_name] == level_val)]
                m = sub.set_index("_SEM")[ycol].to_dict()
                return [m.get(x, None) for x in x_labels]

            if metric_choice == "%P":
                ycol = "P_share"
                base_cols = {}
                if opt == "(All)":
                    for a in level_values:
                        base_cols[a] = _series_for(a, ycol)
                elif opt == "(TOTAL)":
                    sub = total_series_builders["P"].set_index("_SEM")["P_share"].to_dict()
                    base_cols["TOTAL"] = [sub.get(x, None) for x in x_labels]
                else:
                    base_cols[opt] = _series_for(opt, ycol)
            elif metric_choice == "%SA":
                ycol = "SA_share"
                base_cols = {}
                if opt == "(All)":
                    for a in level_values:
                        base_cols[a] = _series_for(a, ycol)
                elif opt == "(TOTAL)":
                    sub = total_series_builders["SA"].set_index("_SEM")[ycol].to_dict()
                    base_cols["TOTAL"] = [sub.get(x, None) for x in x_labels]
                else:
                    base_cols[opt] = _series_for(opt, ycol)
            else:
                ycol = "OTHER_share"
                base_cols = {}
                if opt == "(All)":
                    for a in level_values:
                        base_cols[a] = _series_for(a, ycol)
                elif opt == "(TOTAL)":
                    sub = total_series_builders["OTHER"].set_index("_SEM")[ycol].to_dict()
                    base_cols["TOTAL"] = [sub.get(x, None) for x in x_labels]
                else:
                    base_cols[opt] = _series_for(opt, ycol)

            export_df = pd.DataFrame({"Period": x_labels, **base_cols})
            fname = f"chart_{_slugify(fig_title)}_{_slugify(metric_choice)}_{_slugify(opt)}_{_slugify(sel_label)}.xlsx"
            _download_xlsx_button(export_df, fname, key=f"dl_hist_{_slugify(fig_title)}_{metric_choice}_{_slugify(opt)}_{_slugify(sel_label)}", label="⬇️ Datos de la gráfica (Excel)")

        # ======================== BY ACADEMIC AREA ========================
        if view_mode == "By Academic Area":
            colT, colG = st.columns([6,6], gap="large")

            # --- Agregaciones por Área (para tabla del timeframe seleccionado) ---
            agg_tipo = (fil.groupby(["_AREA","_TIPO"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["SA","PA","SP","IP","OTHER"]:
                if k not in agg_tipo.columns: agg_tipo[k] = 0.0
            agg_tipo = agg_tipo[["SA","PA","SP","IP","OTHER"]]
            agg_ps = (fil.groupby(["_AREA","_PS"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["P","S"]:
                if k not in agg_ps.columns: agg_ps[k] = 0.0
            agg_ps = agg_ps[["P","S"]]

            with colT:
                # ---------- SENSIBILIDAD / HISTÓRICO (Academic Area) ----------
                base_agg_ps = agg_ps.copy()
                base_agg_tipo = agg_tipo.copy()

                if SENS["on"]:
                    # Aplica TODAS las operaciones acumuladas (Add) sobre P/S o SA/SP/IP/PA/OTHER
                    mod_agg_ps, mod_agg_tipo = apply_ops_to_aggs(
                        base_agg_ps, base_agg_tipo, SENS.get("ops", []),
                        member_all_label="All", index_name="Academic Area"
                    )
                    last_op = SENS["ops"][-1] if SENS.get("ops") else None

                    if last_op and last_op["scope"] == "PS":
                        # ------ Modo PS (P o S) ------
                        target_cat = last_op["cat"]  # "P" o "S"
                        impact = impact_column_generic(base_agg_ps, mod_agg_ps, target_cat, mode="PS")

                        den = (mod_agg_ps["P"] + mod_agg_ps["S"]).replace(0, pd.NA)
                        tbl = pd.DataFrame({
                            "Academic Area": mod_agg_ps.index,
                            f"%{target_cat}": (mod_agg_ps[target_cat] / den * 100).round(1).fillna(0.0),
                            f"Impact (Δ%{target_cat})": impact.values
                        })

                        # TOTAL
                        totP, totS = float(mod_agg_ps["P"].sum()), float(mod_agg_ps["S"].sum())
                        tot_den = totP + totS
                        pct_tot = ((totP if target_cat == "P" else totS) / tot_den * 100) if tot_den else 0.0

                        base_den = float(base_agg_ps["P"].sum() + base_agg_ps["S"].sum()) or 1.0
                        base_pct_tot = float(base_agg_ps[target_cat].sum()) / base_den * 100

                        total_row = {
                            "Academic Area": "TOTAL",
                            f"%{target_cat}": round(pct_tot, 1),
                            f"Impact (Δ%{target_cat})": round(pct_tot - base_pct_tot, 2)
                        }
                        metrics_tbl = pd.concat([tbl, pd.DataFrame([total_row])], ignore_index=True)

                        _download_xlsx_button(
                            metrics_tbl,
                            f"table_ByArea_PS_{_slugify(sel_label)}.xlsx",
                            key=f"dl_tbl_area_ps_{_slugify(sel_label)}",
                            label="⬇️ Download table (Excel)"
                        )

                        styled_tbl = (
                            metrics_tbl.style
                            .format({f"%{target_cat}": "{:.1f}%", f"Impact (Δ%{target_cat})": "{:+.2f}"})
                            .background_gradient(subset=[f"Impact (Δ%{target_cat})"], cmap="RdYlGn")
                            .hide(axis="index")
                        )
                        st.markdown(f"<div class='scroll-wrap-400'>{styled_tbl.to_html(escape=False)}</div>", unsafe_allow_html=True)

                    else:
                        # ------ Modo Qualifications (SA/SP/IP/PA/OTHER) ------
                        target_cat = (last_op["cat"] if last_op else "SA")
                        impact = impact_column_generic(base_agg_tipo, mod_agg_tipo, target_cat, mode="QUAL")

                        cats = ["SA", "SP", "IP", "PA", "OTHER"]
                        for c in cats:
                            if c not in mod_agg_tipo.columns:
                                mod_agg_tipo[c] = 0.0

                        den = (mod_agg_tipo[cats].sum(axis=1)).replace(0, pd.NA)
                        tbl = pd.DataFrame({
                            "Academic Area": mod_agg_tipo.index,
                            f"%{target_cat}": (mod_agg_tipo[target_cat] / den * 100).round(1).fillna(0.0),
                            f"Impact (Δ%{target_cat})": impact.values
                        })

                        sums = mod_agg_tipo[cats].sum(axis=0)
                        den_tot = float(sums.sum()) or 1.0
                        pct_tot = float(sums[target_cat]) / den_tot * 100

                        sums_base = base_agg_tipo[cats].sum(axis=0)
                        den_tot_base = float(sums_base.sum()) or 1.0
                        base_pct_tot = float(sums_base[target_cat]) / den_tot_base * 100

                        total_row = {
                            "Academic Area": "TOTAL",
                            f"%{target_cat}": round(pct_tot, 1),
                            f"Impact (Δ%{target_cat})": round(pct_tot - base_pct_tot, 2)
                        }
                        metrics_tbl = pd.concat([tbl, pd.DataFrame([total_row])], ignore_index=True)

                        _download_xlsx_button(
                            metrics_tbl,
                            f"table_ByArea_QUAL_{_slugify(sel_label)}.xlsx",
                            key=f"dl_tbl_area_qual_{_slugify(sel_label)}",
                            label="⬇️ Download table (Excel)"
                        )

                        styled_tbl = (
                            metrics_tbl.style
                            .format({f"%{target_cat}": "{:.1f}%", f"Impact (Δ%{target_cat})": "{:+.2f}"})
                            .background_gradient(subset=[f"Impact (Δ%{target_cat})"], cmap="RdYlGn")
                            .hide(axis="index")
                        )
                        st.markdown(f"<div class='scroll-wrap-400'>{styled_tbl.to_html(escape=False)}</div>", unsafe_allow_html=True)

                else:
                    # ======= MODO HISTÓRICO (sin sensibilidad) =======
                    metrics_tbl = build_percent_table("Academic Area", base_agg_tipo, base_agg_ps)
                    _download_xlsx_button(
                        metrics_tbl,
                        f"table_ByArea_{_slugify(sel_label)}.xlsx",
                        key=f"dl_tbl_area_{_slugify(sel_label)}",
                        label="⬇️ Download table (Excel)"
                    )
                    styled_tbl = (
                        metrics_tbl.style
                        .format({"%P": "{:.1f}%", "%S": "{:.1f}%", "%SA": "{:.1f}%", "%OTHER": "{:.1f}%"})
                        .apply(style_percent_tables, id_col="Academic Area", axis=None)
                        .hide(axis="index")
                    )
                    st.markdown(f"<div class='scroll-wrap-400'>{styled_tbl.to_html(escape=False)}</div>", unsafe_allow_html=True)



            # históricos por Área
            df_hist = df_car_global.copy()

            # PS shares
            agg_ps_all = (df_hist.groupby(["_SEM","_AREA","_PS"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["P","S"]:
                if k not in agg_ps_all.columns: agg_ps_all[k] = 0.0
            agg_ps_all["P_share"] = (agg_ps_all["P"] / (agg_ps_all["P"] + agg_ps_all["S"]).replace(0, pd.NA)) * 100
            agg_ps_all = agg_ps_all.reset_index()

            # Tipo shares
            agg_tipo_all = (df_hist.groupby(["_SEM","_AREA","_TIPO"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["SA","PA","SP","IP","OTHER"]:
                if k not in agg_tipo_all.columns: agg_tipo_all[k] = 0.0
            den_all = (agg_tipo_all[["SA","PA","SP","IP","OTHER"]].sum(axis=1)).replace(0, pd.NA)
            agg_tipo_all["SA_share"]    = (agg_tipo_all["SA"]    / den_all) * 100
            agg_tipo_all["OTHER_share"] = (agg_tipo_all["OTHER"] / den_all) * 100
            agg_tipo_all = agg_tipo_all.reset_index()

            # TOTAL series
            tot_by_sem_P = (df_hist.groupby(["_SEM","_PS"])["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["P","S"]:
                if k not in tot_by_sem_P.columns: tot_by_sem_P[k] = 0.0
            tot_by_sem_P["P_share"] = (tot_by_sem_P["P"] / (tot_by_sem_P["P"] + tot_by_sem_P["S"]).replace(0, pd.NA)) * 100
            tot_by_sem_P = tot_by_sem_P.reset_index()

            tot_by_sem_tipo = (df_hist.groupby(["_SEM","_TIPO"])["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["SA","PA","SP","IP","OTHER"]:
                if k not in tot_by_sem_tipo.columns: tot_by_sem_tipo[k] = 0.0
            den_tot = (tot_by_sem_tipo[["SA","PA","SP","IP","OTHER"]].sum(axis=1)).replace(0, pd.NA)
            tot_by_sem_tipo["SA_share"]    = (tot_by_sem_tipo["SA"]    / den_tot) * 100
            tot_by_sem_tipo["OTHER_share"] = (tot_by_sem_tipo["OTHER"] / den_tot) * 100
            tot_by_sem_tipo = tot_by_sem_tipo.reset_index()

            # --- Adaptar a modo temporal ---
            agg_ps_all_tm   = transform_for_time_mode_ps(agg_ps_all.rename(columns={"_AREA":"__LEVEL__"})).rename(columns={"__LEVEL__":"_AREA"})
            agg_tipo_all_tm = transform_for_time_mode_tipo(agg_tipo_all.rename(columns={"_AREA":"__LEVEL__"}), "SA_share").rename(columns={"__LEVEL__":"_AREA"})
            agg_tipo_all_tm_other = transform_for_time_mode_tipo(agg_tipo_all.rename(columns={"_AREA":"__LEVEL__"}), "OTHER_share").rename(columns={"__LEVEL__":"_AREA"})
            agg_tipo_all_tm = agg_tipo_all_tm.drop(columns=[c for c in ["OTHER_share"] if c in agg_tipo_all_tm], errors="ignore")\
                                             .merge(agg_tipo_all_tm_other[["_SEM","_AREA","OTHER","SA","PA","SP","IP","OTHER_share"]], on=["_SEM","_AREA","SA","PA","SP","IP","OTHER"], how="outer")

            tot_by_sem_P_tm = transform_for_time_mode_ps(tot_by_sem_P.copy())
            tot_by_sem_tipo_sa_tm = transform_for_time_mode_tipo(tot_by_sem_tipo.copy(), "SA_share")
            tot_by_sem_tipo_ot_tm = transform_for_time_mode_tipo(tot_by_sem_tipo.copy(), "OTHER_share")
            tot_by_sem_tipo_tm = tot_by_sem_tipo_sa_tm.drop(columns=[c for c in ["OTHER_share"] if c in tot_by_sem_tipo_sa_tm], errors="ignore")\
                                                      .merge(tot_by_sem_tipo_ot_tm[["_SEM","SA","PA","SP","IP","OTHER","OTHER_share"]], on=["_SEM","SA","PA","SP","IP","OTHER"], how="outer")

            # Eje X y selección
            key_col, x_labels, x_map = build_time_axis_for_history(df_hist)
            sel_x = None
            if time_mode == "Semestral" and sel_sem and str(sel_sem) in x_map: sel_x = x_map[str(sel_sem)]
            if time_mode == "Anual" and sel_year in x_map: sel_x = x_map[sel_year]
            if time_mode == "Intersemestral":
                lab = f"{sel_year} Intersemestral"
                if lab in x_map: sel_x = x_map[lab]

            areas_all = sorted(set(agg_ps_all_tm["_AREA"].unique()) | set(agg_tipo_all_tm["_AREA"].unique()))

            with colG:
                draw_history("Evolution by Academic Area",
                             level_name="_AREA",
                             level_values=areas_all,
                             metric_kind="%P",
                             total_series_builders={"P": tot_by_sem_P_tm, "SA": tot_by_sem_tipo_tm, "OTHER": tot_by_sem_tipo_tm},
                             agg_ps_all=agg_ps_all_tm, agg_tipo_all=agg_tipo_all_tm,
                             x_labels=x_labels, x_map=x_map, sel_x=sel_x)

        # ============================== BY FIELD ==============================
        elif view_mode == "By Field":
            if not col_field:
                st.info("Column 'Field' was not found.")
            else:
                fil_field = fil.copy()
                fil_field["_FIELD"] = fil_field[col_field].astype(str).str.strip()

                colF_L, colF_R = st.columns([6,6], gap="large")

                # Tabla timeframe
                agg_tipo_f = (fil_field.groupby(["_FIELD","_TIPO"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
                for k in ["SA","PA","SP","IP","OTHER"]:
                    if k not in agg_tipo_f.columns: agg_tipo_f[k] = 0.0
                agg_tipo_f = agg_tipo_f[["SA","PA","SP","IP","OTHER"]]

                agg_ps_f = (fil_field.groupby(["_FIELD","_PS"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
                for k in ["P","S"]:
                    if k not in agg_ps_f.columns: agg_ps_f[k] = 0.0
                agg_ps_f = agg_ps_f[["P","S"]]

                with colF_L:
                    base_agg_ps = agg_ps_f.copy()
                    base_agg_tipo = agg_tipo_f.copy()

                    if SENS["on"]:
                        mod_agg_ps, mod_agg_tipo = apply_ops_to_aggs(
                            base_agg_ps, base_agg_tipo, SENS.get("ops", []),
                            member_all_label="All", index_name="Field"
                        )
                        # Determinar la última op para decidir qué mostrar
                        last_op = SENS["ops"][-1] if SENS.get("ops") else None

                        if last_op and last_op["scope"] == "PS":
                            target_cat = last_op["cat"]  # "P" o "S"
                            impact = impact_column_generic(base_agg_ps, mod_agg_ps, target_cat, mode="PS")
                            den = (mod_agg_ps["P"] + mod_agg_ps["S"]).replace(0, pd.NA)
                            tbl = pd.DataFrame({
                                "Field": mod_agg_ps.index,
                                f"%{target_cat}": (mod_agg_ps[target_cat] / den * 100).round(1).fillna(0.0),
                                f"Impact (Δ%{target_cat})": impact.values
                            })
                            # TOTAL
                            totP, totS = float(mod_agg_ps["P"].sum()), float(mod_agg_ps["S"].sum())
                            tot_den = totP + totS
                            pct_tot = ((totP if target_cat=="P" else totS) / tot_den * 100) if tot_den else 0.0
                            base_den = float(base_agg_ps["P"].sum() + base_agg_ps["S"].sum()) or 1.0
                            base_pct_tot = float(base_agg_ps[target_cat].sum()) / base_den * 100
                            total_row = {"Field":"TOTAL", f"%{target_cat}": round(pct_tot,1), f"Impact (Δ%{target_cat})": round(pct_tot - base_pct_tot, 2)}
                            metrics_tbl_f = pd.concat([tbl, pd.DataFrame([total_row])], ignore_index=True)

                            _download_xlsx_button(metrics_tbl_f, f"table_ByField_PS_{_slugify(sel_label)}.xlsx",
                                                key=f"dl_tbl_field_ps_{_slugify(sel_label)}", label="⬇️ Download table (Excel)")
                            styled_tbl_f = (metrics_tbl_f.style
                                            .format({f"%{target_cat}":"{:.1f}%", f"Impact (Δ%{target_cat})":"{:+.2f}"})
                                            .background_gradient(subset=[f"Impact (Δ%{target_cat})"], cmap="RdYlGn")
                                            .hide(axis="index"))
                            st.markdown(f"<div class='scroll-wrap-400'>{styled_tbl_f.to_html(escape=False)}</div>", unsafe_allow_html=True)

                        else:
                            # Qualifications (SA/SP/IP/PA/OTHER)
                            target_cat = last_op["cat"] if last_op else "SA"
                            impact = impact_column_generic(base_agg_tipo, mod_agg_tipo, target_cat, mode="QUAL")
                            cats = ["SA","SP","IP","PA","OTHER"]
                            for c in cats:
                                if c not in mod_agg_tipo.columns: mod_agg_tipo[c] = 0.0
                            den = (mod_agg_tipo[cats].sum(axis=1)).replace(0, pd.NA)
                            tbl = pd.DataFrame({
                                "Field": mod_agg_tipo.index,
                                f"%{target_cat}": (mod_agg_tipo[target_cat] / den * 100).round(1).fillna(0.0),
                                f"Impact (Δ%{target_cat})": impact.values
                            })
                            sums = mod_agg_tipo[cats].sum(axis=0)
                            den_tot = float(sums.sum()) or 1.0
                            pct_tot = float(sums[target_cat]) / den_tot * 100

                            sums_base = base_agg_tipo[cats].sum(axis=0)
                            den_tot_base = float(sums_base.sum()) or 1.0
                            base_pct_tot = float(sums_base[target_cat]) / den_tot_base * 100

                            total_row = {"Field":"TOTAL", f"%{target_cat}": round(pct_tot,1), f"Impact (Δ%{target_cat})": round(pct_tot - base_pct_tot, 2)}
                            metrics_tbl_f = pd.concat([tbl, pd.DataFrame([total_row])], ignore_index=True)

                            _download_xlsx_button(metrics_tbl_f, f"table_ByField_QUAL_{_slugify(sel_label)}.xlsx",
                                                key=f"dl_tbl_field_qual_{_slugify(sel_label)}", label="⬇️ Download table (Excel)")
                            styled_tbl_f = (metrics_tbl_f.style
                                            .format({f"%{target_cat}":"{:.1f}%", f"Impact (Δ%{target_cat})":"{:+.2f}"})
                                            .background_gradient(subset=[f"Impact (Δ%{target_cat})"], cmap="RdYlGn")
                                            .hide(axis="index"))
                            st.markdown(f"<div class='scroll-wrap-400'>{styled_tbl_f.to_html(escape=False)}</div>", unsafe_allow_html=True)

                    else:
                        # ===== HISTÓRICO (como lo tenías) =====
                        metrics_tbl_f = build_percent_table("Field", base_agg_tipo, base_agg_ps)
                        _download_xlsx_button(metrics_tbl_f, f"table_ByField_{_slugify(sel_label)}.xlsx",
                                            key=f"dl_tbl_field_{_slugify(sel_label)}", label="⬇️ Download table (Excel)")
                        styled_tbl_f = (metrics_tbl_f.style
                                        .format({"%P":"{:.1f}%","%S":"{:.1f}%","%SA":"{:.1f}%","%OTHER":"{:.1f}%"})
                                        .apply(style_percent_tables, id_col="Field", axis=None)
                                        .hide(axis="index"))
                        st.markdown(f"<div class='scroll-wrap-400'>{styled_tbl_f.to_html(escape=False)}</div>", unsafe_allow_html=True)



                # Históricos Field
                df_hist_f = df_car_global.copy()
                df_hist_f["_FIELD"] = df_hist_f[col_field].astype(str).str.strip()

                agg_ps_all_f = (df_hist_f.groupby(["_SEM","_FIELD","_PS"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
                for k in ["P","S"]:
                    if k not in agg_ps_all_f.columns: agg_ps_all_f[k] = 0.0
                agg_ps_all_f["P_share"] = (agg_ps_all_f["P"] / (agg_ps_all_f["P"] + agg_ps_all_f["S"]).replace(0, pd.NA)) * 100
                agg_ps_all_f = agg_ps_all_f.reset_index()

                agg_tipo_all_f = (df_hist_f.groupby(["_SEM","_FIELD","_TIPO"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
                for k in ["SA","PA","SP","IP","OTHER"]:
                    if k not in agg_tipo_all_f.columns: agg_tipo_all_f[k] = 0.0
                den_all_f = (agg_tipo_all_f[["SA","PA","SP","IP","OTHER"]].sum(axis=1)).replace(0, pd.NA)
                agg_tipo_all_f["SA_share"]    = (agg_tipo_all_f["SA"]    / den_all_f) * 100
                agg_tipo_all_f["OTHER_share"] = (agg_tipo_all_f["OTHER"] / den_all_f) * 100
                agg_tipo_all_f = agg_tipo_all_f.reset_index()

                # TOTAL series
                tot_by_sem_f = (df_hist_f.groupby(["_SEM","_PS"])["_CRED"].sum().unstack(fill_value=0.0))
                for k in ["P","S"]:
                    if k not in tot_by_sem_f.columns: tot_by_sem_f[k] = 0.0
                tot_by_sem_f["P_share"] = (tot_by_sem_f["P"] / (tot_by_sem_f["P"] + tot_by_sem_f["S"]).replace(0, pd.NA)) * 100
                tot_by_sem_f = tot_by_sem_f.reset_index()

                tot_by_sem_tipo_f = (df_hist_f.groupby(["_SEM","_TIPO"])["_CRED"].sum().unstack(fill_value=0.0))
                for k in ["SA","PA","SP","IP","OTHER"]:
                    if k not in tot_by_sem_tipo_f.columns: tot_by_sem_tipo_f[k] = 0.0
                den_f = (tot_by_sem_tipo_f[["SA","PA","SP","IP","OTHER"]].sum(axis=1)).replace(0, pd.NA)
                tot_by_sem_tipo_f["SA_share"]    = (tot_by_sem_tipo_f["SA"]    / den_f) * 100
                tot_by_sem_tipo_f["OTHER_share"] = (tot_by_sem_tipo_f["OTHER"] / den_f) * 100
                tot_by_sem_tipo_f = tot_by_sem_tipo_f.reset_index()

                # Adaptar a modo temporal
                agg_ps_all_tm   = transform_for_time_mode_ps(agg_ps_all_f.rename(columns={"_FIELD":"__LEVEL__"})).rename(columns={"__LEVEL__":"_FIELD"})
                agg_tipo_all_tm = transform_for_time_mode_tipo(agg_tipo_all_f.rename(columns={"_FIELD":"__LEVEL__"}), "SA_share").rename(columns={"__LEVEL__":"_FIELD"})
                agg_tipo_all_tm_other = transform_for_time_mode_tipo(agg_tipo_all_f.rename(columns={"_FIELD":"__LEVEL__"}), "OTHER_share").rename(columns={"__LEVEL__":"_FIELD"})
                agg_tipo_all_tm = agg_tipo_all_tm.drop(columns=[c for c in ["OTHER_share"] if c in agg_tipo_all_tm], errors="ignore")\
                                                 .merge(agg_tipo_all_tm_other[["_SEM","_FIELD","OTHER","SA","PA","SP","IP","OTHER_share"]], on=["_SEM","_FIELD","SA","PA","SP","IP","OTHER"], how="outer")

                tot_by_sem_P_tm = transform_for_time_mode_ps(tot_by_sem_f.copy())
                tot_by_sem_tipo_sa_tm = transform_for_time_mode_tipo(tot_by_sem_tipo_f.copy(), "SA_share")
                tot_by_sem_tipo_ot_tm = transform_for_time_mode_tipo(tot_by_sem_tipo_f.copy(), "OTHER_share")
                tot_by_sem_tipo_tm = tot_by_sem_tipo_sa_tm.drop(columns=[c for c in ["OTHER_share"] if c in tot_by_sem_tipo_sa_tm], errors="ignore")\
                                                          .merge(tot_by_sem_tipo_ot_tm[["_SEM","SA","PA","SP","IP","OTHER","OTHER_share"]], on=["_SEM","SA","PA","SP","IP","OTHER"], how="outer")

                # Eje X y selección
                key_col, x_labels, x_map = build_time_axis_for_history(df_hist_f)
                sel_x = None
                if time_mode == "Semestral" and sel_sem and str(sel_sem) in x_map: sel_x = x_map[str(sel_sem)]
                if time_mode == "Anual" and sel_year in x_map: sel_x = x_map[sel_year]
                if time_mode == "Intersemestral":
                    lab = f"{sel_year} Intersemestral"
                    if lab in x_map: sel_x = x_map[lab]

                fields_all = sorted(set(agg_ps_all_tm["_FIELD"].astype(str).unique()) | set(agg_tipo_all_tm["_FIELD"].astype(str).unique()))

                with colF_R:
                    draw_history("Evolution by Academic Field",
                                 level_name="_FIELD",
                                 level_values=fields_all,
                                 metric_kind="%P",
                                 total_series_builders={"P": tot_by_sem_P_tm, "SA": tot_by_sem_tipo_tm, "OTHER": tot_by_sem_tipo_tm},
                                 agg_ps_all=agg_ps_all_tm, agg_tipo_all=agg_tipo_all_tm,
                                 x_labels=x_labels, x_map=x_map, sel_x=sel_x)

        # ============================= BY PROGRAM =============================
        else:  # "By Program"
            if not program_col:
                st.info("Column 'program' was not found.")
            else:
                fil_mat = fil.copy()
                fil_mat["_MAT"] = fil_mat[program_col].astype(str).str.strip()

                colM_L, colM_R = st.columns([6,6], gap="large")

                # Tabla timeframe
                agg_tipo_m = (fil_mat.groupby(["_MAT","_TIPO"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
                for k in ["SA","PA","SP","IP","OTHER"]:
                    if k not in agg_tipo_m.columns: agg_tipo_m[k] = 0.0
                agg_tipo_m = agg_tipo_m[["SA","PA","SP","IP","OTHER"]]

                agg_ps_m = (fil_mat.groupby(["_MAT","_PS"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
                for k in ["P","S"]:
                    if k not in agg_ps_m.columns: agg_ps_m[k] = 0.0
                agg_ps_m = agg_ps_m[["P","S"]]

                with colM_L:
                    base_agg_ps = agg_ps_m.copy()
                    base_agg_tipo = agg_tipo_m.copy()

                    if SENS["on"]:
                        mod_agg_ps, mod_agg_tipo = apply_ops_to_aggs(
                            base_agg_ps, base_agg_tipo, SENS.get("ops", []),
                            member_all_label="All", index_name="Program"
                        )
                        last_op = SENS["ops"][-1] if SENS.get("ops") else None

                        if last_op and last_op["scope"] == "PS":
                            target_cat = last_op["cat"]
                            impact = impact_column_generic(base_agg_ps, mod_agg_ps, target_cat, mode="PS")
                            den = (mod_agg_ps["P"] + mod_agg_ps["S"]).replace(0, pd.NA)
                            tbl = pd.DataFrame({
                                "Program": mod_agg_ps.index,
                                f"%{target_cat}": (mod_agg_ps[target_cat] / den * 100).round(1).fillna(0.0),
                                f"Impact (Δ%{target_cat})": impact.values
                            })
                            totP, totS = float(mod_agg_ps["P"].sum()), float(mod_agg_ps["S"].sum())
                            tot_den = totP + totS
                            pct_tot = ((totP if target_cat=="P" else totS) / tot_den * 100) if tot_den else 0.0
                            base_den = float(base_agg_ps["P"].sum() + base_agg_ps["S"].sum()) or 1.0
                            base_pct_tot = float(base_agg_ps[target_cat].sum()) / base_den * 100
                            total_row = {"Program":"TOTAL", f"%{target_cat}": round(pct_tot,1), f"Impact (Δ%{target_cat})": round(pct_tot - base_pct_tot, 2)}
                            metrics_tbl_m = pd.concat([tbl, pd.DataFrame([total_row])], ignore_index=True)

                            _download_xlsx_button(metrics_tbl_m, f"table_ByProgram_PS_{_slugify(sel_label)}.xlsx",
                                                key=f"dl_tbl_prog_ps_{_slugify(sel_label)}", label="⬇️ Download table (Excel)")
                            styled_tbl_m = (metrics_tbl_m.style
                                            .format({f"%{target_cat}":"{:.1f}%", f"Impact (Δ%{target_cat})":"{:+.2f}"})
                                            .background_gradient(subset=[f"Impact (Δ%{target_cat})"], cmap="RdYlGn")
                                            .hide(axis="index"))
                            st.markdown(f"<div class='scroll-wrap-program'>{styled_tbl_m.to_html(escape=False)}</div>", unsafe_allow_html=True)

                        else:
                            target_cat = last_op["cat"] if last_op else "SA"
                            impact = impact_column_generic(base_agg_tipo, mod_agg_tipo, target_cat, mode="QUAL")
                            cats = ["SA","SP","IP","PA","OTHER"]
                            for c in cats:
                                if c not in mod_agg_tipo.columns: mod_agg_tipo[c] = 0.0
                            den = (mod_agg_tipo[cats].sum(axis=1)).replace(0, pd.NA)
                            tbl = pd.DataFrame({
                                "Program": mod_agg_tipo.index,
                                f"%{target_cat}": (mod_agg_tipo[target_cat] / den * 100).round(1).fillna(0.0),
                                f"Impact (Δ%{target_cat})": impact.values
                            })
                            sums = mod_agg_tipo[cats].sum(axis=0)
                            den_tot = float(sums.sum()) or 1.0
                            pct_tot = float(sums[target_cat]) / den_tot * 100

                            sums_base = base_agg_tipo[cats].sum(axis=0)
                            den_tot_base = float(sums_base.sum()) or 1.0
                            base_pct_tot = float(sums_base[target_cat]) / den_tot_base * 100

                            total_row = {"Program":"TOTAL", f"%{target_cat}": round(pct_tot,1), f"Impact (Δ%{target_cat})": round(pct_tot - base_pct_tot, 2)}
                            metrics_tbl_m = pd.concat([tbl, pd.DataFrame([total_row])], ignore_index=True)

                            _download_xlsx_button(metrics_tbl_m, f"table_ByProgram_QUAL_{_slugify(sel_label)}.xlsx",
                                                key=f"dl_tbl_prog_qual_{_slugify(sel_label)}", label="⬇️ Download table (Excel)")
                            styled_tbl_m = (metrics_tbl_m.style
                                            .format({f"%{target_cat}":"{:.1f}%", f"Impact (Δ%{target_cat})":"{:+.2f}"})
                                            .background_gradient(subset=[f"Impact (Δ%{target_cat})"], cmap="RdYlGn")
                                            .hide(axis="index"))
                            st.markdown(f"<div class='scroll-wrap-program'>{styled_tbl_m.to_html(escape=False)}</div>", unsafe_allow_html=True)

                    else:
                        # ===== HISTÓRICO (como lo tenías) =====
                        metrics_tbl_m = build_percent_table("Program", base_agg_tipo, base_agg_ps)
                        _download_xlsx_button(metrics_tbl_m, f"table_ByProgram_{_slugify(sel_label)}.xlsx",
                                            key=f"dl_tbl_prog_{_slugify(sel_label)}", label="⬇️ Download table (Excel)")
                        styled_tbl_m = (metrics_tbl_m.style
                                        .format({"%P":"{:.1f}%","%S":"{:.1f}%","%SA":"{:.1f}%","%OTHER":"{:.1f}%"})
                                        .apply(style_percent_tables, id_col="Program", axis=None)
                                        .hide(axis="index"))
                        st.markdown(f"<div class='scroll-wrap-program'>{styled_tbl_m.to_html(escape=False)}</div>", unsafe_allow_html=True)



                # Históricos Program
                df_hist_m = df_car_global.copy()
                df_hist_m["_MAT"] = df_hist_m[program_col].astype(str).str.strip()

                agg_ps_all_m = (df_hist_m.groupby(["_SEM","_MAT","_PS"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
                for k in ["P","S"]:
                    if k not in agg_ps_all_m.columns: agg_ps_all_m[k] = 0.0
                agg_ps_all_m["P_share"] = (agg_ps_all_m["P"] / (agg_ps_all_m["P"] + agg_ps_all_m["S"]).replace(0, pd.NA)) * 100
                agg_ps_all_m = agg_ps_all_m.reset_index()

                agg_tipo_all_m = (df_hist_m.groupby(["_SEM","_MAT","_TIPO"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
                for k in ["SA","PA","SP","IP","OTHER"]:
                    if k not in agg_tipo_all_m.columns: agg_tipo_all_m[k] = 0.0
                den_all_m = (agg_tipo_all_m[["SA","PA","SP","IP","OTHER"]].sum(axis=1)).replace(0, pd.NA)
                agg_tipo_all_m["SA_share"]    = (agg_tipo_all_m["SA"]    / den_all_m) * 100
                agg_tipo_all_m["OTHER_share"] = (agg_tipo_all_m["OTHER"] / den_all_m) * 100
                agg_tipo_all_m = agg_tipo_all_m.reset_index()

                # TOTAL series
                tot_by_sem_m = (df_hist_m.groupby(["_SEM","_PS"])["_CRED"].sum().unstack(fill_value=0.0))
                for k in ["P","S"]:
                    if k not in tot_by_sem_m.columns: tot_by_sem_m[k] = 0.0
                tot_by_sem_m["P_share"] = (tot_by_sem_m["P"] / (tot_by_sem_m["P"] + tot_by_sem_m["S"]).replace(0, pd.NA)) * 100
                tot_by_sem_m = tot_by_sem_m.reset_index()

                tot_by_sem_tipo_m = (df_hist_m.groupby(["_SEM","_TIPO"])["_CRED"].sum().unstack(fill_value=0.0))
                for k in ["SA","PA","SP","IP","OTHER"]:
                    if k not in tot_by_sem_tipo_m.columns: tot_by_sem_tipo_m[k] = 0.0
                den_m = (tot_by_sem_tipo_m[["SA","PA","SP","IP","OTHER"]].sum(axis=1)).replace(0, pd.NA)
                tot_by_sem_tipo_m["SA_share"]    = (tot_by_sem_tipo_m["SA"]    / den_m) * 100
                tot_by_sem_tipo_m["OTHER_share"] = (tot_by_sem_tipo_m["OTHER"] / den_m) * 100
                tot_by_sem_tipo_m = tot_by_sem_tipo_m.reset_index()

                # Adaptar a modo temporal
                agg_ps_all_tm   = transform_for_time_mode_ps(agg_ps_all_m.rename(columns={"_MAT":"__LEVEL__"})).rename(columns={"__LEVEL__":"_MAT"})
                agg_tipo_all_tm = transform_for_time_mode_tipo(agg_tipo_all_m.rename(columns={"_MAT":"__LEVEL__"}), "SA_share").rename(columns={"__LEVEL__":"_MAT"})
                agg_tipo_all_tm_other = transform_for_time_mode_tipo(agg_tipo_all_m.rename(columns={"_MAT":"__LEVEL__"}), "OTHER_share").rename(columns={"__LEVEL__":"_MAT"})
                agg_tipo_all_tm = agg_tipo_all_tm.drop(columns=[c for c in ["OTHER_share"] if c in agg_tipo_all_tm], errors="ignore")\
                                                 .merge(agg_tipo_all_tm_other[["_SEM","_MAT","OTHER","SA","PA","SP","IP","OTHER_share"]], on=["_SEM","_MAT","SA","PA","SP","IP","OTHER"], how="outer")

                tot_by_sem_P_tm = transform_for_time_mode_ps(tot_by_sem_m.copy())
                tot_by_sem_tipo_sa_tm = transform_for_time_mode_tipo(tot_by_sem_tipo_m.copy(), "SA_share")
                tot_by_sem_tipo_ot_tm = transform_for_time_mode_tipo(tot_by_sem_tipo_m.copy(), "OTHER_share")
                tot_by_sem_tipo_tm = tot_by_sem_tipo_sa_tm.drop(columns=[c for c in ["OTHER_share"] if c in tot_by_sem_tipo_sa_tm], errors="ignore")\
                                                          .merge(tot_by_sem_tipo_ot_tm[["_SEM","SA","PA","SP","IP","OTHER","OTHER_share"]], on=["_SEM","SA","PA","SP","IP","OTHER"], how="outer")

                # Eje X y selección
                key_col, x_labels, x_map = build_time_axis_for_history(df_hist_m)
                sel_x = None
                if time_mode == "Semestral" and sel_sem and str(sel_sem) in x_map: sel_x = x_map[str(sel_sem)]
                if time_mode == "Anual" and sel_year in x_map: sel_x = x_map[sel_year]
                if time_mode == "Intersemestral":
                    lab = f"{sel_year} Intersemestral"
                    if lab in x_map: sel_x = x_map[lab]

                programs_all = sorted(set(agg_ps_all_tm["_MAT"].astype(str).unique()) | set(agg_tipo_all_tm["_MAT"].astype(str).unique()))

                with colM_R:
                    draw_history("Evolution by Academic Program",
                                 level_name="_MAT",
                                 level_values=programs_all,
                                 metric_kind="%P",
                                 total_series_builders={"P": tot_by_sem_P_tm, "SA": tot_by_sem_tipo_tm, "OTHER": tot_by_sem_tipo_tm},
                                 agg_ps_all=agg_ps_all_tm, agg_tipo_all=agg_tipo_all_tm,
                                 x_labels=x_labels, x_map=x_map, sel_x=sel_x)

# ====== (FULL-WIDTH) CREDIT SUMS BY DIMENSION (EXPANDER) ======
try:
    # Base del período actual
    sum_df = (df_car_filt_all.copy() if 'df_car_filt_all' in locals()
              else (fil.copy() if 'fil' in locals() else df_car.copy()))

    # EXCLUSIÓN LOCAL POR CÓDIGO DE MATERIA (solo para esta tabla)
    col_code_local = _get_any(sum_df, "Program")
    if col_code_local:
        code_up = sum_df[col_code_local].astype(str).str.strip().str.upper()
        pattern = r"^(?:E[-_ ]?)?(?:CONT|AFIN|ENEG|IMER)\b"
        sum_df = sum_df[~code_up.str.match(pattern, na=False)].copy()

    # Normalizaciones mínimas
    if "_CRED" not in sum_df.columns:
        cred_src = _get_any(sum_df, "Créditos","Creditos","Credits")
        sum_df["_CRED"] = pd.to_numeric(sum_df[cred_src], errors="coerce").fillna(0.0) if cred_src else 0.0
    if "_PS" not in sum_df.columns and col_ps_C:
        sum_df["_PS"] = _norm_str(sum_df[col_ps_C]).map(normalize_ps)
    if "_TIPO" not in sum_df.columns and col_tipoC:
        sum_df["_TIPO"] = _norm_str(sum_df[col_tipoC]).map(normalize_tipo)

    # Dimensión según vista
    view = st.session_state.view_mode if "view_mode" in st.session_state else "By Academic Area"
    if view == "By Academic Area":
        dim_col, dim_label = "_AREA", "Academic Area"
        if dim_col not in sum_df.columns and col_areaCourse:
            sum_df[dim_col] = sum_df[col_areaCourse].astype(str).str.strip()
    elif view == "By Field":
        dim_col, dim_label = "_FIELD", "Field"
        if dim_col not in sum_df.columns and 'col_field' in locals() and col_field:
            sum_df[dim_col] = sum_df[col_field].astype(str).str.strip()
    else:
        dim_col, dim_label = "_MAT", "Program"
        if dim_col not in sum_df.columns and 'program_col' in locals() and program_col:
            sum_df[dim_col] = sum_df[program_col].astype(str).str.strip()

    # Construcción de tabla
    if dim_col not in sum_df.columns:
        st.info("Cannot build credit sums table: missing dimension column.")
    else:
        base_index = sum_df.groupby(dim_col)["_CRED"].sum().sort_values(ascending=False)
        idx = base_index.index

        sum_total = base_index.rename("Credit Sum")
        sum_P  = (sum_df[sum_df["_PS"]   == "P"     ].groupby(dim_col)["_CRED"].sum().reindex(idx, fill_value=0.0)).rename("P Sum")
        sum_S  = (sum_df[sum_df["_PS"]   == "S"     ].groupby(dim_col)["_CRED"].sum().reindex(idx, fill_value=0.0)).rename("S Sum")
        sum_SA = (sum_df[sum_df["_TIPO"] == "SA"    ].groupby(dim_col)["_CRED"].sum().reindex(idx, fill_value=0.0)).rename("SA Sum")
        sum_PA = (sum_df[sum_df["_TIPO"] == "PA"    ].groupby(dim_col)["_CRED"].sum().reindex(idx, fill_value=0.0)).rename("PA Sum")
        sum_SP = (sum_df[sum_df["_TIPO"] == "SP"    ].groupby(dim_col)["_CRED"].sum().reindex(idx, fill_value=0.0)).rename("SP Sum")
        sum_IP = (sum_df[sum_df["_TIPO"] == "IP"    ].groupby(dim_col)["_CRED"].sum().reindex(idx, fill_value=0.0)).rename("IP Sum")
        sum_OT = (sum_df[sum_df["_TIPO"] == "OTHER" ].groupby(dim_col)["_CRED"].sum().reindex(idx, fill_value=0.0)).rename("OTHER Sum")

        tbl = pd.concat(
            [sum_total, sum_P, sum_S, sum_SA, sum_PA, sum_SP, sum_IP, sum_OT],
            axis=1
        ).fillna(0.0)

        # Fila TOTAL al final
        total_row = pd.DataFrame(tbl.sum(axis=0)).T
        total_row.index = ["TOTAL"]
        tbl_out = pd.concat([tbl, total_row], axis=0)

        display_label = sel_label if 'sel_label' in locals() else "Selected Period"
        with st.expander(f"Credit sums by {dim_label} — {display_label}", expanded=False):
            # Botón de descarga (arriba de la tabla)
            export_tbl = tbl_out.reset_index().rename(columns={"index": dim_label})
            _download_xlsx_button(
                export_tbl,
                f"credit_sums_{_slugify(dim_label)}_{_slugify(display_label)}.xlsx",
                key=f"dl_credit_sums_{_slugify(dim_label)}_{_slugify(display_label)}",
                label="⬇️ Descargar tabla (Excel)"
            )
            # Estilo y render
            def _bold_total(df_):
                sty = pd.DataFrame('', index=df_.index, columns=df_.columns)
                if "TOTAL" in df_.index:
                    sty.loc["TOTAL", :] = 'font-weight:700;'
                return sty
            styled_tbl = (tbl_out.style
                          .format({c: "{:,.0f}" for c in ["Credit Sum","P Sum","S Sum","SA Sum","PA Sum","SP Sum","IP Sum","OTHER Sum"]})
                          .apply(_bold_total, axis=None))
            st.dataframe(styled_tbl, use_container_width=True)

except Exception:
    pass

# ================== TABLA + DONA + TOP PROFES Y BUSCADOR ==================
try:
    cfg = {
        "By Academic Area": {"key": "_AREA_filter",  "col": "_AREA", "label": "area",    "metric_key": "metric__AREA"},
        "By Field":         {"key": "_FIELD_filter", "col": "_FIELD","label": "campo",   "metric_key": "metric__FIELD"},
        "By Program":       {"key": "_MAT_filter",   "col": "_MAT",  "label": "programa","metric_key": "metric__MAT"},
    }
    view = st.session_state.view_mode

    if view in cfg:
        key = cfg[view]["key"]
        col_tag = cfg[view]["col"]
        label = cfg[view]["label"]
        metric_key = cfg[view]["metric_key"]

        opt_val = st.session_state.get(key, "(All)")
        metric_choice = st.session_state.get(metric_key, "%P")  # "%P" | "%SA" | "%OTHER"

        if opt_val != "(All)":
            base = df_car_filt_all.copy()

            # Normalizaciones mínimas
            if "_AREA" not in base.columns and col_areaCourse:
                base["_AREA"] = base[col_areaCourse].astype(str).str.strip()
            if "_FIELD" not in base.columns and col_field:
                base["_FIELD"] = base[col_field].astype(str).str.strip()
            if "_MAT" not in base.columns and col_prog:
                base["_MAT"] = base[col_prog].astype(str).str.strip()
            if "_TIPO" not in base.columns and col_tipoC:
                base["_TIPO"] = _norm_str(base[col_tipoC]).map(normalize_tipo)
            if "_PS" not in base.columns and col_ps_C:
                base["_PS"] = _norm_str(base[col_ps_C]).map(normalize_ps)
            if "_CRED" not in base.columns and col_cred:
                base["_CRED"] = pd.to_numeric(base[col_cred], errors="coerce").fillna(0.0)

            # Filtro por dimensión (excepto TOTAL)
            if opt_val != "(TOTAL)" and col_tag in base.columns:
                base = base[base[col_tag] == opt_val].copy()

            # Columnas solicitadas (para tabla y buscador)
            wanted_map = {
                "Semestre": col_sem,
                "Código Materia": col_code,
                "Créditos": col_cred,
                "Nombre largo curso": col_name,
                "Program": col_prog,
                "Profesor": col_prof,
                "Area del curso": col_areaCourse,
                "Field": col_field,
                "TIPO": col_tipoC,
                "P/S": col_ps_C,
            }
            present = {k: v for k, v in wanted_map.items() if v in base.columns}
            out = base[list(present.values())].rename(columns={v: k for k, v in present.items()})

            # ====== Layout: Tabla (izq) + Dona (der) ======
            cL, cR = st.columns([7,5], gap="large")

            # ---- Título + filtro de tabla según métrica + tabla
            with cL:
                if metric_choice == "%P":
                    table_filter = st.radio(
                        "",
                        ["All", "Only P", "Only S"],
                        index=0, horizontal=True, key=f"table_filt_ps_{view}_{opt_val}"
                    )
                    base_tbl = base.copy()
                    if table_filter == "Only P":
                        base_tbl = base_tbl[base_tbl["_PS"] == "P"]
                    elif table_filter == "Only S":
                        base_tbl = base_tbl[base_tbl["_PS"] == "S"]
                else:
                    table_filter = st.radio(
                        "",
                        ["All", "Only SA", "Only OTHER"],
                        index=0, horizontal=True, key=f"table_filt_tipo_{view}_{opt_val}"
                    )
                    base_tbl = base.copy()
                    if table_filter == "Only SA":
                        base_tbl = base_tbl[base_tbl["_TIPO"] == "SA"]
                    elif table_filter == "Only OTHER":
                        base_tbl = base_tbl[base_tbl["_TIPO"] == "OTHER"]

                present_tbl = {k: v for k, v in wanted_map.items() if v in base_tbl.columns}
                out = base_tbl[list(present_tbl.values())].rename(columns={v: k for k, v in present_tbl.items()})

                display_label = sel_label if 'sel_label' in locals() else "Selected Period"
                n_courses = len(out)

                if metric_choice == "%P":
                    if table_filter == "Only P":
                        title = f"{n_courses} courses were taught in {display_label} by Participating Faculty"
                    elif table_filter == "Only S":
                        title = f"{n_courses} courses were taught in {display_label} by Supporting Faculty"
                    else:
                        title = (f"{n_courses} courses were taught in {display_label}"
                                 if opt_val == "(TOTAL)"
                                 else f"{n_courses} courses of {opt_val} were taught in {display_label}")
                else:
                    if table_filter == "Only SA":
                        title = f"{n_courses} courses were taught in {display_label} by Scholarly Academics"
                    elif table_filter == "Only OTHER":
                        title = f"{n_courses} courses were taught in {display_label} by Others"
                    else:
                        title = (f"{n_courses} courses were taught in {display_label}"
                                 if opt_val == "(TOTAL)"
                                 else f"{n_courses} courses of {opt_val} were taught in {display_label}")

                st.markdown(f"### {title}")

                # —— botón de descarga sobre la tabla ——
                _download_xlsx_button(
                    out,
                    f"table_detail_{_slugify(opt_val)}_{_slugify(display_label)}.xlsx",
                    key=f"dl_tbl_detail_{_slugify(opt_val)}_{_slugify(display_label)}",
                    label="⬇️ Descargar tabla (Excel)"
                )

                st.dataframe(out, use_container_width=True, hide_index=True)

            # ---- Dona (NO afectada por el filtro de tabla)
            with cR:
                MINT      = "#1FA89B"
                GRAY_2    = "#B0B0B0"
                GRAY_DARK = "#6B7280"
                RED_LIGHT = "#F5A3A3"

                donut_h   = 360
                header_h  = 56
                row_h     = 28
                title_h   = 28
                filter_h  = 40
                visible_rows = max(6, min(len(out), 12))
                table_h = title_h + filter_h + header_h + visible_rows * row_h + 16
                avail     = max(0, table_h - donut_h)
                shift_down = 40
                pad_top   = max(0, avail // 2 + shift_down)
                pad_bottom= max(0, avail - (pad_top - shift_down))
                st.markdown(f"<div style='height:{pad_top}px'></div>", unsafe_allow_html=True)

                if metric_choice == "%P":
                    agg = base.groupby("_PS")["_CRED"].sum()
                    p_val = float(agg.get("P", 0.0))
                    s_val = float(agg.get("S", 0.0))
                    den = p_val + s_val
                    p_share = (p_val/den*100) if den else 0.0
                    thr = 75.0 if opt_val == "(TOTAL)" else 60.0
                    alert = (p_share < thr)
                    color_map = {"P": (RED_LIGHT if alert else MINT), "S": GRAY_2}

                    fig = px.pie(
                        names=["P","S"], values=[p_val, s_val],
                        color=["P","S"],
                        color_discrete_map=color_map,
                        hole=0.55
                    )
                    fig.update_traces(textinfo="percent+label", hovertemplate="%{label}: %{percent:.1%}<extra></extra>")
                    fig.update_layout(
                        title=f"% Participating Distribution — {('TOTAL' if opt_val=='(TOTAL)' else opt_val)}",
                        height=donut_h, margin=dict(l=10, r=10, t=40, b=10),
                        legend=dict(orientation="v", yanchor="bottom", y=0.4, xanchor="center", x=0.9),
                    )
                    st.plotly_chart(fig, use_container_width=True)

                    # —— descarga de datos de la dona (abajo) ——
                    donut_df = pd.DataFrame({
                        "Group": ["P","S"],
                        "Credits": [p_val, s_val]
                    })
                    donut_df["Percent"] = (donut_df["Credits"] / max(1e-9, donut_df["Credits"].sum()))*100
                    _download_xlsx_button(
                        donut_df,
                        f"chart_donut_PS_{_slugify(opt_val)}_{_slugify(display_label)}.xlsx",
                        key=f"dl_donut_ps_{_slugify(opt_val)}_{_slugify(display_label)}",
                        label="⬇️ Datos de la gráfica (Excel)"
                    )

                else:
                    agg = base.groupby("_TIPO")["_CRED"].sum()
                    sa = float(agg.get("SA", 0.0))
                    pa = float(agg.get("PA", 0.0))
                    sp = float(agg.get("SP", 0.0))
                    ip = float(agg.get("IP", 0.0))
                    other = float(agg.get("OTHER", 0.0))
                    den = sa + pa + sp + ip + other
                    sa_share    = (sa/den*100)    if den else 0.0
                    other_share = (other/den*100) if den else 0.0

                    thr = 40.0
                    alert_sa     = (sa_share < thr)
                    alert_other  = (other_share > 10.0)

                    labels_all  = ["SA", "PA", "SP", "IP", "OTHER"]
                    values_all  = [sa,    pa,   sp,   ip,   other]
                    filtered    = [(l, v) for l, v in zip(labels_all, values_all) if v > 0]

                    if filtered:
                        labels = [l for l, _ in filtered]
                        values = [v for _, v in filtered]

                        color_map = {}
                        for l in labels:
                            if l == "SA":
                                color_map[l] = (RED_LIGHT if alert_sa else MINT)
                            elif l == "OTHER":
                                color_map[l] = (RED_LIGHT if alert_other else GRAY_DARK)
                            else:
                                color_map[l] = GRAY_2

                        fig = px.pie(
                            names=labels,
                            values=values,
                            color=labels,
                            color_discrete_map=color_map,
                            hole=0.55
                        )
                        fig.update_traces(
                            textinfo="percent+label",
                            sort=False,
                            hovertemplate="%{label}: %{percent:.1%}<extra></extra>"
                        )
                        title_txt = "%SA Distribution" if metric_choice == "%SA" else "%OTHER Distribution"
                        fig.update_layout(
                            title=f"{title_txt} — {('TOTAL' if opt_val=='(TOTAL)' else opt_val)}",
                            height=donut_h, margin=dict(l=10, r=10, t=40, b=10),
                            legend=dict(orientation="v", yanchor="bottom", y=0.4, xanchor="center", x=0.9),
                        )
                        st.plotly_chart(fig, use_container_width=True)

                        # —— descarga datos dona (abajo) ——
                        donut_df = pd.DataFrame({"Type": labels_all, "Credits": values_all})
                        donut_df["Percent"] = (donut_df["Credits"] / max(1e-9, donut_df["Credits"].sum()))*100
                        _download_xlsx_button(
                            donut_df,
                            f"chart_donut_TIPO_{_slugify(opt_val)}_{_slugify(display_label)}.xlsx",
                            key=f"dl_donut_tipo_{_slugify(opt_val)}_{_slugify(display_label)}",
                            label="⬇️ Datos de la gráfica (Excel)"
                        )
                    else:
                        st.caption("No hay registros de TIPO para esta métrica en este período.")
                # padding inferior para alinear con la tabla
                st.markdown(f"<div style='height:{pad_bottom}px'></div>", unsafe_allow_html=True)

            # ====== NUEVO: Top 5 Profes + Buscador (independiente del filtro de arriba) ======
            st.markdown("---")
            st.markdown("### Detail — Top Professors & Search")

            # Dataset SOLO del período seleccionado (ya filtrado globalmente)
            period_df = df_car_filt_all.copy()

            # Resolver columnas desde el dataset del período
            col_periodo = _get_any(period_df, "periodo", "Periodo", "PERIODO", "Semestre", "SEMESTRE")

            colTop, colInputs = st.columns([6, 6], gap="large")

            # --- Top 5 Profes (izquierda) — independiente de la selección de arriba
            with colTop:
                if col_prof and col_prof in period_df.columns:
                    tmp = period_df.copy()

                    # Asegurar créditos numéricos
                    if "_CRED" not in tmp.columns:
                        tmp["_CRED"] = (pd.to_numeric(tmp[col_cred], errors="coerce").fillna(0.0)
                                        if col_cred in tmp.columns else 1.0)
                    tmp["_PROF"] = tmp[col_prof].astype(str).str.strip()

                    # Contar cursos por profesor (si hay código de materia contamos apariciones)
                    if col_code and col_code in tmp.columns:
                        agg_courses = (col_code, "count")
                    else:
                        agg_courses = ("_PROF", "size")

                    top = (
                        tmp.groupby("_PROF", dropna=False)
                        .agg(Courses=agg_courses, Credits=("_CRED", "sum"))
                        .reset_index()
                        .rename(columns={"_PROF": "Profesor"})
                        .sort_values(["Courses", "Credits"], ascending=[False, False])
                        .head(5)
                    )

                    # Ranking 1..5
                    top.insert(0, "#", range(1, len(top) + 1))

                    # Botón de descarga arriba de la tabla
                    _download_xlsx_button(
                        top,
                        f"top_professors_{_slugify(sel_label)}.xlsx",
                        key=f"dl_top_prof_{_slugify(sel_label)}",
                        label="⬇️ Descargar tabla (Excel)"
                    )

                    st.dataframe(
                        top.style.format({"Credits": "{:,.1f}"}),
                        use_container_width=True, hide_index=True
                    )
                else:
                    st.info("No 'Profesor' column to build Top 5.")

            # --- Entradas de búsqueda (derecha)
            with colInputs:
                st.write("**Search**")
                q_prof   = st.text_input("Professor name contains:", value="", key="q_prof")
                q_course = st.text_input("Course name / code contains:", value="", key="q_course")

            # --- Resultados de la búsqueda a todo el ancho (debajo)
            if q_prof.strip() or q_course.strip():
                st.markdown("#### Search results")
                search_df = period_df.copy()

                # Preparar columnas visibles; incluir Periodo primero (si existe)
                show_cols_map = {
                    "Periodo": col_periodo,
                    "Profesor": col_prof,
                    "Código Materia": col_code,
                    "Nombre largo curso": col_name,
                    "Area del curso": col_areaCourse,
                    "Field": col_field,
                    "Program": col_prog,
                }
                have = {k: v for k, v in show_cols_map.items() if v in search_df.columns}

                # Filtros
                mask = pd.Series(True, index=search_df.index)
                if q_prof.strip() and col_prof in search_df.columns:
                    patp = q_prof.strip().lower()
                    mask = mask & search_df[col_prof].astype(str).str.lower().str.contains(patp, na=False)

                if q_course.strip():
                    patc = q_course.strip().lower()
                    m_name = (search_df[col_name].astype(str).str.lower().str.contains(patc, na=False)
                              if col_name in search_df.columns else False)
                    m_code = (search_df[col_code].astype(str).str.lower().str.contains(patc, na=False)
                              if col_code in search_df.columns else False)
                    m_any = (m_name | m_code) if isinstance(m_name, pd.Series) or isinstance(m_code, pd.Series) else False
                    mask = mask & (m_any if isinstance(m_any, pd.Series) else True)

                search_df = search_df[mask].copy()

                # Renombrar y ordenar columnas
                out_search = search_df[list(have.values())].rename(columns={v: k for k, v in have.items()})
                if out_search.empty:
                    st.caption("No results found for the given query.")
                else:
                    col_order = [c for c in [
                        "Periodo", "Profesor", "Código Materia", "Nombre largo curso",
                        "Area del curso", "Field", "Program"
                    ] if c in out_search.columns]
                    out_search = out_search[col_order]

                    # Botón de descarga sobre la tabla
                    _download_xlsx_button(
                        out_search,
                        f"search_results_{_slugify(sel_label)}.xlsx",
                        key=f"dl_search_{_slugify(sel_label)}",
                        label="⬇️ Descargar tabla (Excel)"
                    )

                    st.dataframe(out_search, use_container_width=True, hide_index=True)
            # Si no hay texto en ninguno, no mostrar nada abajo.

except Exception:
    pass

# ==================== (BOTTOM) COUNTS SECTION — PIVOT (interactive) ====================
st.markdown("---")
show_counts = st.checkbox("Show P/S counts", value=False)
if show_counts:
    st.subheader(f"Participating vs Supporting — {sel_label} (Counts & %)")

    col_ps_fd   = _get_any(df_fd_f, "P/S", "P - S", "Participating/Supporting")
    col_area_fd = _get_any(df_fd_f, "AREA_PROFESOR", "Area_Profesor", "Area Profesor", "Área", "Area")
    col_tipo_fd = _get_any(df_fd_f, "TIPO", "Tipo", "Ranking", "Tipo Ranking")

    if col_ps_fd:
        df_fd_f["_PS"] = _norm_str(df_fd_f[col_ps_fd]).map(normalize_ps)
    if col_area_fd:
        df_fd_f["_AREA"] = df_fd_f[col_area_fd].astype(str).str.strip()
    if col_tipo_fd:
        df_fd_f["_TIPO"] = _norm_str(df_fd_f[col_tipo_fd]).map(normalize_tipo)

    pivot_rows = st.radio("Pivot by", ["AREA", "Qualification Type"], index=0, horizontal=True)

    if pivot_rows == "AREA":
        row_name = "AREA"
        row_series = df_fd_f["_AREA"].astype(str).str.strip().replace({"": "N/A"})
        desired_order = None
    else:
        row_name = "Type"
        row_series = df_fd_f["_TIPO"].map(lambda v: str(v).upper())
        desired_order = ["SA", "PA", "SP", "IP", "OTHER"]

    base = pd.DataFrame({row_name: row_series, "_PS": df_fd_f["_PS"]})

    table = (base.groupby([row_name, "_PS"], dropna=False)
                  .size().unstack(fill_value=0)
                  .rename(columns={"P": "Participating", "S": "Supporting"}))
    for k in ["Participating", "Supporting"]:
        if k not in table.columns:
            table[k] = 0

    table["__Total__"] = table["Participating"] + table["Supporting"]
    if desired_order:
        for code in desired_order:
            if code not in table.index:
                table.loc[code] = [0, 0, 0]
        table = table.reindex(desired_order)
    else:
        table = table.sort_values("__Total__", ascending=False)

    df_counts = (
        table[["Participating", "Supporting"]]
        .astype(int)
        .reset_index()
    )
    total_row = pd.DataFrame([{
        row_name: "TOTAL",
        "Participating": int(df_counts["Participating"].sum()),
        "Supporting":    int(df_counts["Supporting"].sum())
    }])
    df_counts_out = pd.concat([df_counts, total_row], ignore_index=True)

    def _bold_total(df_):
        sty = pd.DataFrame('', index=df_.index, columns=df_.columns)
        mask = df_[row_name].astype(str).str.upper().eq("TOTAL")
        for c in df_.columns:
            sty.loc[mask, c] = 'font-weight:700;'
        return sty

    # ===== LAYOUT: TABLA (izq) + GRÁFICA (der) =====
    left, right = st.columns([6,6], gap="large")

    # --- Preparar datos de % antes de pintar ---
    denom = table["__Total__"].replace(0, pd.NA)
    perc_df = pd.DataFrame({
        row_name: table.index,
        "%Participating": (table["Participating"] / denom * 100).round(1).fillna(0.0),
        "%Supporting":    (table["Supporting"]    / denom * 100).round(1).fillna(0.0),
    })
    cat_order = (desired_order if desired_order
                 else df_counts_out.loc[df_counts_out[row_name] != "TOTAL", row_name].tolist())
    chart_export = perc_df.melt(id_vars=row_name, value_vars=["%Participating", "%Supporting"],
                                var_name="Group", value_name="Percent")

    # --- Columna izquierda: tabla + descarga
    with left:
        _download_xlsx_button(
            df_counts_out,
            f"ps_counts_{_slugify(row_name)}_{_slugify(sel_label)}.xlsx",
            key=f"dl_ps_counts_{_slugify(row_name)}_{_slugify(sel_label)}",
            label="Descargar tabla (Excel)"
        )
        styled_counts = (df_counts_out.style
                         .format({"Participating": "{:,.0f}", "Supporting": "{:,.0f}"})
                         .apply(_bold_total, axis=None))
        st.dataframe(styled_counts, use_container_width=True, hide_index=True)

    # --- Columna derecha: gráfica + descarga
    with right:
        fig = px.bar(
            chart_export, x=row_name, y="Percent", color="Group",
            barmode="group", text="Percent",
            color_discrete_map={"%Participating": MINT, "%Supporting": SUPPORTING},
            category_orders={row_name: cat_order}
        )
        fig.update_traces(texttemplate="%{text:.1f}%")
        fig.update_layout(
            xaxis_title=None, yaxis_title=None,
            height=340,
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
            legend_title_text=None,
            margin=dict(l=20, r=10, t=10, b=40)
        )
        st.plotly_chart(fig, use_container_width=True)

        _download_xlsx_button(
            chart_export,
            f"chart_ps_perc_{_slugify(row_name)}_{_slugify(sel_label)}.xlsx",
            key=f"dl_chart_ps_perc_{_slugify(row_name)}_{_slugify(sel_label)}",
            label="Descargar datos (Excel)"
        )



