import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import webbrowser
import os
import io, base64, re

# ====== utils descarga (dejar así) ======
def _xlsx_bytes(df: pd.DataFrame, sheet_name: str = "Data") -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf) as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    buf.seek(0)
    return buf.getvalue()

def _download_link(label: str, df: pd.DataFrame, filename: str):
    data = _xlsx_bytes(df)
    b64 = base64.b64encode(data).decode()
    href = f"data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}"
    st.markdown(
        f'<a class="dl-min" download="{filename}" href="{href}">{label}</a>',
        unsafe_allow_html=True
    )

# CSS del link minimalista (dejar así)
st.markdown("""
<style>
a.dl-min, a.dl-min:link, a.dl-min:visited {
  color:#1FA89B !important; text-decoration:underline !important; 
  font-size:13px; display:inline-block; margin-top:6px;
}
a.dl-min:hover { opacity:.85; }
</style>
""", unsafe_allow_html=True)

#================= GENERAL CONFIG (dejar así) ===============================
st.set_page_config(
    page_title="Dashboard",
    page_icon=":bar_chart:",
    layout="wide",
    initial_sidebar_state="expanded",
)

#================= DATA LOAD (cached) ===========================
@st.cache_data(ttl=0)
def load_fulltime():
    df = pd.read_excel("data/Faculty/BD_Faculty.xlsx", sheet_name="BD PLANTA 2020-2025")

    # Periodo soportando Intersemestral: YYYY10/YYYY20 o "YYYY Intersemestral"
    if "Semestre" in df.columns:
        sem = df["Semestre"].astype(str).str.strip()
        is_inter = sem.str.contains("inter", case=False, na=False)
        df["Periodo"] = np.where(is_inter, sem.str[:4] + " Intersemestral", sem.str[:4] + sem.str[-2:])
    else:
        raw = df.iloc[:, 0].astype(str).str.strip()
        is_inter = raw.str.contains("inter", case=False, na=False)
        df["Periodo"] = np.where(is_inter, raw.str[:4] + " Intersemestral", raw.str.slice(0, 4) + raw.str.slice(4, 6))

    if "Academic Area" in df.columns and "AREA_PROFESOR" not in df.columns:
        df["AREA_PROFESOR"] = df["Academic Area"]
    if "ID Nr." in df.columns and "ID" not in df.columns:
        df = df.rename(columns={"ID Nr.": "ID"})
    if "Full Name" not in df.columns:
        fn = df.get("First Name", "")
        ln = df.get("Last Name", "")
        df["Full Name"] = (pd.Series(fn).astype(str).fillna("") + " " + pd.Series(ln).astype(str).fillna("")).str.strip()
    return df

@st.cache_data(ttl=0)
def load_parttime():
    df = pd.read_excel("data/Faculty/BD_Faculty.xlsx", sheet_name="Faculty Distribution")

    # Filtra CÁTEDRA (normaliza acento)
    if "PLANTA_CATEDRA" in df.columns:
        col = df["PLANTA_CATEDRA"].astype(str).str.strip()
        col = col.str.normalize("NFKD").str.encode("ascii", errors="ignore").str.decode("ascii")
        df = df[col.str.upper().eq("CATEDRA")].copy()

    sem = df["Semestre"].astype(str).str.strip()
    is_inter = sem.str.contains("inter", case=False, na=False)
    df.loc[~is_inter, "Periodo"] = sem.str[:4] + sem.str[-2:]            # YYYY10/20 (sin guion)
    df.loc[is_inter,  "Periodo"] = sem.str[:4] + " Intersemestral"

    if "ID Nr." not in df.columns and "ID" in df.columns:
        df = df.rename(columns={"ID": "ID Nr."})
    if "AREA_PROFESOR" not in df.columns and "Academic Area" in df.columns:
        df["AREA_PROFESOR"] = df["Academic Area"]
    return df

df_full = load_fulltime()
df_part = load_parttime()

# ====== helpers timeframe ======
def _years_all() -> list[int]:
    y = []
    if "Periodo" in df_full.columns:
        y.append(df_full["Periodo"].dropna().astype(str).str[:4])
    if "Semestre" in df_full.columns:
        y.append(df_full["Semestre"].dropna().astype(str).str[:4])
    if "Semestre" in df_part.columns:
        y.append(df_part["Semestre"].dropna().astype(str).str[:4])
    if "Periodo" in df_part.columns:
        y.append(df_part["Periodo"].dropna().astype(str).str[:4])
    if not y:
        return []
    ys = pd.concat(y, ignore_index=True)
    ys = ys[ys.str.match(r"^\d{4}$", na=False)]
    return sorted(ys.astype(int).unique().tolist())

def _col_id(df_: pd.DataFrame):
    return "ID Nr." if "ID Nr." in df_.columns else ("ID" if "ID" in df_.columns else None)

def _filter_for_timeframe(df_in: pd.DataFrame, time_mode: str,
                          sel_sem: str | None = None,
                          sel_year: int | None = None,
                          sel_inter_label: str | None = None) -> pd.DataFrame:
    """Filtro puntual (para vistas que requieren un período/año concreto)."""
    dfb = df_in.copy()
    pcol = "Periodo" if "Periodo" in dfb.columns else None
    scol = "Semestre" if "Semestre" in dfb.columns else None
    idc  = _col_id(dfb)

    if time_mode == "Semestral" and sel_sem:
        # objetivo: YYYY10/20 (sin guion)
        target = str(sel_sem)
        mask = pd.Series(False, index=dfb.index)
        if pcol: mask |= dfb[pcol].astype(str).eq(target)
        if scol: mask |= dfb[scol].astype(str).str.replace("-", "", regex=False).str.fullmatch(target, na=False)
        dfb = dfb[mask].copy()

    elif time_mode == "Anual" and sel_year is not None:
        y = str(sel_year)
        mask = pd.Series(False, index=dfb.index)
        if pcol: mask |= dfb[pcol].astype(str).str.startswith(y)  # incluye inter
        if scol: mask |= dfb[scol].astype(str).str.startswith(y)
        dfb = dfb[mask].copy()
        if idc:
            sort_key = pcol if pcol else (scol if scol else None)
            if sort_key: dfb = dfb.sort_values(by=[sort_key])
            dfb = dfb.drop_duplicates(subset=[idc,], keep="last")

    elif time_mode == "Intersemestral" and sel_inter_label:
        y = sel_inter_label.split()[0]
        mask = pd.Series(False, index=dfb.index)
        if scol:
            scol_n = dfb[scol].astype(str)
            mask |= scol_n.str.contains("inter", case=False, na=False) & scol_n.str.contains(y, na=False)
        if pcol:
            mask |= dfb[pcol].astype(str).eq(sel_inter_label)
        dfb = dfb[mask].copy()

    return dfb

def _is_semester_label(p: str) -> bool:
    return bool(re.fullmatch(r"\d{4}(10|20)", str(p)))

def _is_inter_label(p: str) -> bool:
    return bool(re.fullmatch(r"\d{4}\s+Intersemestral", str(p)))

def _options_for_timeframe(df_src: pd.DataFrame, time_mode: str):
    per = df_src["Periodo"].dropna().astype(str)
    if time_mode == "Semestral":
        return sorted([p for p in per.unique() if _is_semester_label(p)])
    elif time_mode == "Intersemestral":
        return sorted([p for p in per.unique() if _is_inter_label(p)])
    else:
        return sorted(per.str[:4].unique().tolist())  # años

# ==== Time series builders (para las líneas de abajo) ====
def build_time_series(df_src: pd.DataFrame, time_mode: str, idcol: str,
                      degree_col: str | None, nat_col: str | None):
    """Devuelve (labels, phd%, intl%) según el time_mode."""
    labels, phd_pct, intl_pct = [], [], []

    def _phd_pct_df(sub):
        if sub.empty or degree_col is None: return 0.0
        sub = sub.copy()
        s = sub[degree_col].astype(str).str.strip()
        is_tbd = s.str.upper().eq("TBD") | s.eq("") | s.str.lower().eq("na") | s.str.lower().eq("none")
        s_norm = s.str.lower().str.replace(".", "", regex=False)
        s_norm = s_norm.str.normalize("NFKD").str.encode("ascii", errors="ignore").str.decode("ascii")
        is_phd = s_norm.str.contains(r"\bphd\b") | s_norm.str.contains("doctor")
        sub["__deg"] = np.where(is_tbd, "TBD", np.where(is_phd, "PhD", "Other"))
        tot = sub[idcol].nunique()
        return 0.0 if tot == 0 else round(100 * sub.loc[sub["__deg"]=="PhD", idcol].nunique() / tot, 1)

    def _intl_pct_df(sub):
        if sub.empty or nat_col is None: return 0.0
        nat = sub[nat_col].astype(str).str.strip()
        is_col, is_tbd, is_empty = nat.eq("Colombian"), nat.str.upper().eq("TBD"), nat.eq("")
        tot = sub[idcol].nunique()
        return 0.0 if tot == 0 else round(100 * sub.loc[(~is_col) & (~is_tbd) & (~is_empty), idcol].nunique() / tot, 1)

    if time_mode == "Semestral":
        periods = sorted([p for p in df_src["Periodo"].dropna().astype(str).unique() if _is_semester_label(p)])
        for p in periods:
            sub = df_src[df_src["Periodo"].astype(str).eq(p)]
            labels.append(p); phd_pct.append(_phd_pct_df(sub)); intl_pct.append(_intl_pct_df(sub))
    elif time_mode == "Intersemestral":
        periods = sorted([p for p in df_src["Periodo"].dropna().astype(str).unique() if _is_inter_label(p)])
        for p in periods:
            sub = df_src[df_src["Periodo"].astype(str).eq(p)]
            labels.append(p); phd_pct.append(_phd_pct_df(sub)); intl_pct.append(_intl_pct_df(sub))
    else:  # Anual
        years = sorted(df_src["Periodo"].dropna().astype(str).str[:4].unique())
        for y in years:
            sub = _filter_for_timeframe(df_src, "Anual", sel_year=int(y))
            labels.append(y); phd_pct.append(_phd_pct_df(sub)); intl_pct.append(_intl_pct_df(sub))

    return labels, phd_pct, intl_pct

#================= INITIAL STATE ================================
if "modo_faculty" not in st.session_state:
    st.session_state.modo_faculty = "Full-time"
if "time_mode_side" not in st.session_state:
    st.session_state.time_mode_side = "Semestral"
if "sel_tf_label" not in st.session_state:
    st.session_state.sel_tf_label = None  # periodo elegido en el selectbox (string)

#================= GLOBAL CSS (dejar así) =======================
st.markdown("""
<style>
#mode-pill [role="radiogroup"]{ display:flex; gap:8px; margin-top:0; }
#mode-pill [role="radio"]{
  flex:1; justify-content:center; border:1px solid #d0d4d9;
  border-radius:999px; padding:8px 12px; background:#f0f2f6; color:#666;
  font-weight:600; cursor:pointer; text-align:center;
}
#mode-pill [role="radio"][aria-checked="true"]{
  background:#dff7f2; color:#0b6b63; border-color:#8fd7cc;
}
#mode-pill [data-baseweb="radio"] input{ display:none !important; }
.period-label{ text-align:center; font-weight:700; font-size:1.05rem; }
.header-title { color:#21877D; font-weight:bold; text-align:center; font-size:32px; }
.header-btn, .header-btn:link, .header-btn:visited, .header-btn:hover, .header-btn:active {
  background-color:#21877D !important; color:#ffffff !important; padding:8px 16px !important;
  border:none !important; border-radius:8px !important; cursor:pointer !important;
  font-size:14px !important; text-decoration:none !important;
}
</style>
""", unsafe_allow_html=True)

#================= HEADER ===================
with st.container():
    cols = st.columns([1,3,1], gap="small")
    with cols[0]:
        st.markdown('<a href="http://157.253.69.67:8503" class="header-btn" target="_self">⬅ Previous KPI</a>', unsafe_allow_html=True)
    with cols[1]:
        st.markdown('<div class="header-title">UASM Faculty Demographics</div>', unsafe_allow_html=True)
    with cols[2]:
        st.markdown('<a href="http://157.253.69.67:8505" class="header-btn" target="_self">➡ Next KPI</a>', unsafe_allow_html=True)

#================= SIDEBAR: MODE + NAV ==========================
options = {
    "Select...": None,
    "1 Full-time Composition": "http://157.253.69.67:8501",
    "2 Full-time Staffing Levels": "http://157.253.69.67:8502",
    "3 Distribution by Academic Area": "http://157.253.69.67:8503",
    "4 Faculty Demographics": "http://157.253.69.67:8504",
    "5 Full-time Faculty Questionnaire": "http://157.253.69.67:8505",
    "6 Faculty Qualifications": "http://157.253.69.67:8506",
    "Open main HTML menu": "web/KPIs/Faculty/Web KPIs - Faculty.html"
}
with st.sidebar:
    st.markdown('</div>', unsafe_allow_html=True)
    choice = st.selectbox("📊 Go to KPI:", list(options.keys()))

    st.markdown('---')
    st.markdown("#### Faculty Type")
    st.markdown('<div id="mode-pill">', unsafe_allow_html=True)
    mode_sidebar = st.radio(
        "Mode", ["Full-time", "Part-time"],
        index=0 if st.session_state.modo_faculty == "Full-time" else 1,
        horizontal=True, label_visibility="collapsed", key="mode_pill_radio"
    )

if mode_sidebar != st.session_state.modo_faculty:
    st.session_state.modo_faculty = mode_sidebar
    st.session_state.sel_tf_label = None  # reset selección al cambiar modo
    st.rerun()

if options[choice]:
    target = options[choice]
    if target.endswith(".html"):
        abs_path = os.path.abspath(target)
        webbrowser.open(f"file:///{abs_path}")
        st.success("The Faculty menu was opened in a new browser tab.")
    else:
        st.markdown(f'<meta http-equiv="refresh" content="0; url={target}" />', unsafe_allow_html=True)

#================= SIDEBAR: TIMEFRAME + SELECTOR DE PERIODO ===================
with st.sidebar:
    st.markdown("#### Timeframe")
    time_mode_side = st.radio(
        "Timeframe", ["Semestral", "Anual", "Intersemestral"],
        key="time_mode_side", label_visibility="collapsed"
    )

    # Dataset según modo
    df_base = df_full if st.session_state.get("modo_faculty", "Full-time") == "Full-time" else df_part

    # Opciones del selector según timeframe
    options_tf = _options_for_timeframe(df_base, time_mode_side)
    # Default: el último disponible
    default_opt = options_tf[-1] if options_tf else None

    sel_label = st.selectbox(
        "Periodo",
        options_tf,
        index=(options_tf.index(st.session_state.sel_tf_label) if st.session_state.sel_tf_label in options_tf else (len(options_tf)-1 if options_tf else 0)),
        help="Selecciona el periodo para filtrar las gráficas y tablas.",
    ) if options_tf else None

    # Guarda selección
    if sel_label != st.session_state.get("sel_tf_label"):
        st.session_state.sel_tf_label = sel_label

    # --- Export según selección ---
    if sel_label:
        if time_mode_side == "Semestral":
            export_df = _filter_for_timeframe(df_base, "Semestral", sel_sem=sel_label)
            label_time = sel_label
        elif time_mode_side == "Anual":
            export_df = _filter_for_timeframe(df_base, "Anual", sel_year=int(sel_label))
            label_time = sel_label
        else:
            export_df = _filter_for_timeframe(df_base, "Intersemestral", sel_inter_label=sel_label)
            label_time = sel_label
    else:
        export_df = df_base.iloc[0:0].copy()
        label_time = ""

    fname = f"{'FT' if st.session_state.get('modo_faculty')=='Full-time' else 'PT'}_{time_mode_side}_{str(label_time).replace(' ','_')}.xlsx"
    _download_link(f"Descargar base (Excel) — {st.session_state.get('modo_faculty','Full-time')} — {label_time}", export_df, fname)

#================= MODE MERGE ===========================
mode_now = st.session_state.get("modo_faculty", "Full-time")
df = (df_full if mode_now == "Full-time" else df_part).copy()
if "ID Nr." not in df.columns and "ID" in df.columns:
    df["ID Nr."] = df["ID"]

#================= COLUMN HELPERS =======================
def col_id(df_):
    return "ID Nr." if "ID Nr." in df_.columns else ("ID" if "ID" in df_.columns else None)
def col_degree(df_):
    for c in ["Highest Degree", "TÍTULO"]:
        if c in df_.columns: return c
    return None
def col_gender(df_):
    for c in ["Gender", "GÉNERO"]:
        if c in df_.columns: return c
    return None
def col_nationality(df_):
    for c in ["Country of Birth", "Nationality"]:
        if c in df_.columns: return c
    return None
def normalize_degree(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip()
    is_tbd = s.str.upper().eq("TBD") | s.eq("") | s.str.lower().eq("na") | s.str.lower().eq("none")
    s_norm = s.str.lower().str.replace(".", "", regex=False)
    s_norm = s_norm.str.normalize("NFKD").str.encode("ascii", errors="ignore").str.decode("ascii")
    is_phd = s_norm.str.contains(r"\bphd\b") | s_norm.str.contains("doctor")
    is_master = s_norm.str.contains("master") | s_norm.str.contains(r"\bmsc\b") | s_norm.str.contains(r"\bms\b")
    is_bachelor = (s_norm.str.contains("bachelor") | s_norm.str.contains(r"\bbsc\b") | s_norm.str.contains(r"\bbs\b") |
                   s_norm.str.contains(r"\bba\b") | s_norm.str.contains("licen"))
    out = pd.Series("Other", index=s.index, dtype=object)
    out[is_tbd] = "TBD"
    out[~is_tbd & is_phd] = "PhD"
    out[~is_tbd & ~is_phd & is_master] = "Master"
    out[~is_tbd & ~is_phd & ~is_master & is_bachelor] = "Bachelor"
    return out

#================= PERIOD SELECTOR HEADER (texto del periodo seleccionado) ====
sel_period_text = st.session_state.get("sel_tf_label") or ""
st.subheader("Full-time demographics by Faculty ranking" if mode_now == "Full-time" else "Part-time demographic table")
# Mostrar el periodo SOLO en Full-time (en Part-time la tabla ya muestra todos los periodos)
if sel_period_text and mode_now == "Full-time":
    st.markdown(f"<div class='period-label'>{sel_period_text}</div>", unsafe_allow_html=True)


#================= LAYOUT: TABLE (LEFT) + SMALL CHARTS (RIGHT)
col_table, col_side = st.columns([3, 1.2])

#================= TABLE =========================================
with col_table:
    IDCOL = col_id(df)
    if not IDCOL:
        st.error("ID column not found."); st.stop()

    tmode = st.session_state.get("time_mode_side", "Semestral")
    sel_lbl = st.session_state.get("sel_tf_label")

    if not sel_lbl and mode_now == "Full-time":
        st.info("No data for the selected mode/period.")
    else:
        # --- Full-time: tabla del periodo puntual (según timeframe) ---
        if mode_now == "Full-time":
            if tmode == "Semestral":
                active = _filter_for_timeframe(df, "Semestral", sel_sem=sel_lbl)
            elif tmode == "Anual":
                active = _filter_for_timeframe(df, "Anual", sel_year=int(sel_lbl))
            else:
                active = _filter_for_timeframe(df, "Intersemestral", sel_inter_label=sel_lbl)

            # Ranking order
            if "Faculty Ranking" in active.columns:
                base_order = ["Full Professor", "Associate Professor", "Assistant Professor", "Instructor"]
                uniq = active["Faculty Ranking"].dropna().unique().tolist()
                ranking_order = [x for x in base_order if x in uniq] + [x for x in uniq if x not in base_order]
            else:
                ranking_order = []

            cols_out = ["Category"] + ranking_order + ["Total"]

            def counts_by_ranking(df_sub: pd.DataFrame) -> pd.Series:
                if ranking_order:
                    s = (df_sub.groupby("Faculty Ranking")[IDCOL].nunique()
                         .reindex(ranking_order, fill_value=0).astype(int))
                else:
                    s = pd.Series(dtype=int)
                s.loc["Total"] = int(s.sum()) if not s.empty else int(df_sub[IDCOL].nunique())
                return s

            rows = []
            groups = {"Highest Degree","Nationality","Gender","Age"}

            # 1) Highest Degree
            dcol = col_degree(active)
            if dcol and not active.empty:
                active["Degree_norm"] = normalize_degree(active[dcol])
                rows.append(pd.Series({"Category": "Highest Degree", **counts_by_ranking(active[active["Degree_norm"].isin(["PhD","Master","Bachelor"]) ]).to_dict()}))
                for d in ["PhD", "Master", "Bachelor"]:
                    rows.append(pd.Series({"Category": d, **counts_by_ranking(active[active["Degree_norm"] == d]).to_dict()}))

            # 2) Nationality
            ncol = col_nationality(active)
            if ncol and not active.empty:
                rows.append(pd.Series({"Category":"Nationality", **counts_by_ranking(active).to_dict()}))
                rows.append(pd.Series({"Category":"Colombian", **counts_by_ranking(active[active[ncol].astype(str).eq("Colombian")]).to_dict()}))
                rows.append(pd.Series({"Category":"International", **counts_by_ranking(active[~active[ncol].astype(str).eq("Colombian")]).to_dict()}))

            # 3) Gender
            gcol = col_gender(active)
            if gcol and not active.empty:
                rows.append(pd.Series({"Category":"Gender", **counts_by_ranking(active[active[gcol].astype(str).isin(["Male","Female"])]).to_dict()}))
                for g in ["Male","Female"]:
                    rows.append(pd.Series({"Category": g, **counts_by_ranking(active[active[gcol].astype(str)==g]).to_dict()}))

            # 4) Age (buckets)
            if not active.empty and "Age" in active.columns:
                age = pd.to_numeric(active["Age"], errors="coerce")
                active["Age_bucket"] = pd.cut(age, bins=[-np.inf,29,40,50,60,np.inf],
                                              labels=["Under 30","31-40","41-50","51-60","over 61"])
                rows.append(pd.Series({"Category":"Age", **counts_by_ranking(active[active["Age_bucket"].notna()]).to_dict()}))
                for b in ["Under 30","31-40","41-50","51-60","over 61"]:
                    rows.append(pd.Series({"Category": b, **counts_by_ranking(active[active["Age_bucket"]==b]).to_dict()}))

            table_df = pd.DataFrame(rows).reindex(columns=cols_out).fillna(0)

            if not table_df.empty:
                numeric_cols = [c for c in (ranking_order + ["Total"]) if c in table_df.columns]
                is_group = table_df["Category"].isin(groups)
                all_zero = (table_df[numeric_cols].sum(axis=1) == 0) if numeric_cols else pd.Series(False, index=table_df.index)
                table_df = table_df.loc[~(all_zero & ~is_group)].copy()
                for c in numeric_cols:
                    table_df[c] = pd.to_numeric(table_df[c], errors="coerce").fillna(0).astype(int)

            if table_df.empty:
                st.info("No rows to display for this selection.")
            else:
                mint_dark  = "#004d47"; mint_light = "#dff7f2"
                def _style_group_rows(df_):
                    styles = pd.DataFrame('', index=df_.index, columns=df_.columns)
                    mask = df_["Category"].isin(groups)
                    styles.loc[mask, ["Category"] + ranking_order + ["Total"]] = 'background-color:#f2f2f2;'
                    styles.loc[mask, "Category"] = styles.loc[mask, "Category"] + 'font-weight:700;'
                    return styles
                def _style_total_col_bold(df_):
                    styles = pd.DataFrame('', index=df_.index, columns=df_.columns)
                    styles.loc[:, "Total"] = 'font-weight:700;'
                    return styles
                def _style_group_totals_mint(df_):
                    styles = pd.DataFrame('', index=df_.index, columns=df_.columns)
                    mask = df_["Category"].isin(groups)
                    styles.loc[mask, "Total"] = f'background-color:{mint_light}; color:{mint_dark}; font-weight:800;'
                    return styles

                styled_table = (
                    table_df.style
                    .apply(_style_group_rows, axis=None)
                    .apply(_style_total_col_bold, axis=None)
                    .apply(_style_group_totals_mint, axis=None)
                    .format(precision=0, na_rep="")
                    .hide(axis="index")
                )
                st.dataframe(styled_table, use_container_width=True, height=48 + 33*(len(table_df)+1))

        # --- Part-time: tabla por universo (todos del tipo seleccionado) ---
        else:
            tmode = st.session_state.get("time_mode_side", "Semestral")
            active = df.copy()
            if tmode == "Semestral":
                active = active[active["Periodo"].astype(str).apply(_is_semester_label)].copy()
                keys = sorted(active["Periodo"].dropna().astype(str).unique().tolist())
            elif tmode == "Intersemestral":
                active = active[active["Periodo"].astype(str).apply(_is_inter_label)].copy()
                keys = sorted(active["Periodo"].dropna().astype(str).unique().tolist())
            else:  # Anual
                active["__Year"] = active["Periodo"].astype(str).str[:4]
                active = active.sort_values(by=["Periodo"]).drop_duplicates(subset=[IDCOL, "__Year"], keep="last")
                keys = sorted(active["__Year"].dropna().astype(str).unique().tolist())

            def counts_by_key(df_sub: pd.DataFrame) -> pd.Series:
                if tmode == "Anual":
                    s = df_sub.groupby("__Year")[IDCOL].nunique().reindex(keys, fill_value=0).astype(int)
                else:
                    s = df_sub.groupby("Periodo")[IDCOL].nunique().reindex(keys, fill_value=0).astype(int)
                return s

            deg_col = "TÍTULO" if "TÍTULO" in active.columns else ("Highest Degree" if "Highest Degree" in active.columns else None)
            nat_col = "Nationality" if "Nationality" in active.columns else ("Country of Birth" if "Country of Birth" in active.columns else None)
            gen_col = "GÉNERO" if "GÉNERO" in active.columns else ("Gender" if "Gender" in active.columns else None)

            rows = []
            groups = {"Highest Degree", "Nationality", "Gender", "Age"}

            if deg_col and not active.empty:
                active["Degree_norm"] = normalize_degree(active[deg_col])
                rows.append(pd.Series({"Category": "Highest Degree", **counts_by_key(active).to_dict()}))
                for d in ["PhD", "Master", "Bachelor", "TBD"]:
                    rows.append(pd.Series({"Category": d, **counts_by_key(active[active["Degree_norm"] == d]).to_dict()}))

            if nat_col and not active.empty:
                nat = active[nat_col].astype(str).str.strip()
                is_tbd_nat = nat.str.upper().eq("TBD")
                is_col = nat.eq("Colombian")
                rows.append(pd.Series({"Category": "Nationality", **counts_by_key(active).to_dict()}))
                rows.append(pd.Series({"Category": "Colombian", **counts_by_key(active[is_col]).to_dict()}))
                is_int = (~is_col) & (~is_tbd_nat) & nat.ne("")
                rows.append(pd.Series({"Category": "International", **counts_by_key(active[is_int]).to_dict()}))
                rows.append(pd.Series({"Category": "TBD (Nationality)", **counts_by_key(active[is_tbd_nat]).to_dict()}))

            if gen_col and not active.empty:
                g = active[gen_col].astype(str).str.strip()
                is_tbd_g = g.str.upper().eq("TBD") | g.eq("")
                rows.append(pd.Series({"Category": "Gender", **counts_by_key(active).to_dict()}))
                for gval in ["Male", "Female"]:
                    rows.append(pd.Series({"Category": gval, **counts_by_key(active[g.eq(gval)]).to_dict()}))
                rows.append(pd.Series({"Category": "TBD (Gender)", **counts_by_key(active[is_tbd_g]).to_dict()}))

            if "Age" in active.columns and not active.empty:
                age = pd.to_numeric(active["Age"], errors="coerce")
                active = active.assign(
                    Age_bucket=pd.cut(age, bins=[-np.inf,29,40,50,60,np.inf],
                                      labels=["Under 30","31-40","41-50","51-60","over 61"])
                )
                is_tbd_age = active["Age_bucket"].isna()
                rows.append(pd.Series({"Category": "Age", **counts_by_key(active).to_dict()}))
                for b in ["Under 30","31-40","41-50","51-60","over 61"]:
                    rows.append(pd.Series({"Category": b, **counts_by_key(active[active["Age_bucket"] == b]).to_dict()}))
                rows.append(pd.Series({"Category": "TBD (Age)", **counts_by_key(active[is_tbd_age]).to_dict()}))

            ycol = "Years Industry experience"
            if ycol in active.columns:
                yseries = pd.to_numeric(active[ycol], errors="coerce")
                if tmode == "Anual":
                    avg_series = active.assign(**{ycol: yseries}).groupby("__Year")[ycol].mean().reindex(keys).round(1)
                else:
                    avg_series = active.assign(**{ycol: yseries}).groupby("Periodo")[ycol].mean().reindex(keys).round(1)
                rows.append(pd.Series({"Category": "Avg years of Work Exp.", **avg_series.to_dict()}))
                is_tbd_exp = active[ycol].isna()
                rows.append(pd.Series({"Category": "TBD (Work Exp.)", **counts_by_key(active[is_tbd_exp]).to_dict()}))

            cols_out = ["Category"] + keys
            table_df = pd.DataFrame(rows).reindex(columns=cols_out).fillna(0)

            if not table_df.empty:
                numeric_cols = keys
                is_group_or_avg = table_df["Category"].isin(groups) | table_df["Category"].eq("Avg years of Work Exp.")
                all_zero = (table_df[numeric_cols].sum(axis=1) == 0)
                table_df = table_df.loc[~(all_zero & ~is_group_or_avg)].copy()

            if table_df.empty:
                st.info("No rows to display."); st.stop()

            mask_avg = table_df["Category"].eq("Avg years of Work Exp.")
            display_df = table_df.copy().astype(object)
            for c in keys:
                display_df.loc[mask_avg, c] = pd.to_numeric(table_df.loc[mask_avg, c], errors="coerce").map(lambda x: ("" if pd.isna(x) else f"{x:.1f}"))
                display_df.loc[~mask_avg, c] = pd.to_numeric(table_df.loc[~mask_avg, c], errors="coerce").fillna(0).astype(int).map(lambda x: f"{x}")

            blue_light = "#e6f0fb"; blue_dark = "#184a90"; red_light  = "#f8d7da"; red_dark  = "#721c24"

            def _style_gray_groups_and_avg(df_):
                styles = pd.DataFrame('', index=df_.index, columns=df_.columns)
                gray_rows = df_["Category"].isin(groups) | df_["Category"].eq("Avg years of Work Exp.")
                styles.loc[gray_rows, ["Category"] + keys] = 'background-color:#f2f2f2; font-weight:700;'
                return styles
            def _style_last_col_blue(df_):
                styles = pd.DataFrame('', index=df_.index, columns=df_.columns)
                if keys:
                    last_col = keys[-1]
                    mask = df_["Category"].isin(groups)
                    styles.loc[mask, last_col] = f'background-color:{blue_light}; color:{blue_dark}; font-weight:800;'
                return styles
            def _style_tbd_red(df_):
                styles = pd.DataFrame('', index=df_.index, columns=df_.columns)
                for i, row in df_.iterrows():
                    if "TBD" in str(row["Category"]):
                        for c in keys:
                            try: v = float(table_df.loc[i, c])
                            except Exception: v = np.nan
                            if pd.notna(v) and v > 0:
                                styles.at[i, c] = f'background-color:{red_light}; color:{red_dark}; font-weight:800;'
                return styles

            styled_table = (
                display_df.style
                .apply(_style_gray_groups_and_avg, axis=None)
                .apply(_style_last_col_blue, axis=None)
                .apply(_style_tbd_red, axis=None)
                .hide(axis="index")
            )
            st.dataframe(styled_table, use_container_width=True, height=48 + 33 * (len(display_df) + 1))

#================= SIDE CHARTS ===================================
with col_side:
    IDCOL = col_id(df)
    if not IDCOL:
        st.stop()

    sel_lbl = st.session_state.get("sel_tf_label")
    tmode = st.session_state.get("time_mode_side", "Semestral")

    if not sel_lbl:
        st.info("Select a period.")
    else:
        # Título del panel derecho: solo el valor del periodo
        st.markdown(f"<div class='period-label'>{sel_lbl}</div>", unsafe_allow_html=True)

        # Activo puntual para KPIs a la derecha
        if tmode == "Semestral":
            active_side = _filter_for_timeframe(df, "Semestral", sel_sem=sel_lbl)
        elif tmode == "Anual":
            active_side = _filter_for_timeframe(df, "Anual", sel_year=int(sel_lbl))
        else:
            active_side = _filter_for_timeframe(df, "Intersemestral", sel_inter_label=sel_lbl)

        total_act = int(active_side[IDCOL].nunique()) if not active_side.empty else 0

        mint = "#00A896"; mint_dark = "#004d47"; gauge_bg = "#E8FAF7"

        # ---- Gauge PhD% ----
        dcol = col_degree(active_side)
        phds = 0
        if total_act and dcol:
            active_side["Degree_norm"] = normalize_degree(active_side[dcol])
            phds = int(active_side[active_side["Degree_norm"]=="PhD"][IDCOL].nunique())
        pct_phd = round(100*phds/total_act,1) if total_act else 0.0
        fig_gauge_phd = go.Figure(go.Indicator(
            mode="gauge", value=pct_phd,
            gauge={'axis':{'range':[0,100]}, 'bar':{'color':mint}, 'bgcolor':gauge_bg}
        ))
        fig_gauge_phd.add_annotation(x=0.5,y=0.40,xref="paper",yref="paper",text="PhD%:",showarrow=False,
                                     font=dict(color=mint_dark,size=13))
        fig_gauge_phd.add_annotation(x=0.5,y=0.0,xref="paper",yref="paper",text=f"{pct_phd:.1f}%",showarrow=False,
                                     font=dict(color=mint_dark,size=18))
        fig_gauge_phd.update_layout(height=110, margin=dict(l=10,r=10,t=10,b=6))
        st.plotly_chart(fig_gauge_phd, use_container_width=True)

        # ---- Gauge International% ----
        ncol = col_nationality(active_side)
        pct_int = 0.0
        if total_act and ncol:
            nat = active_side[ncol].astype(str).str.strip()
            is_col = nat.eq("Colombian")
            is_tbd = nat.str.upper().eq("TBD")
            is_empty = nat.eq("")
            is_int = (~is_col) & (~is_tbd) & (~is_empty)
            intl_cnt = int(active_side[is_int][IDCOL].nunique())
            pct_int = round(100*intl_cnt/total_act,1)
        fig_gauge_int = go.Figure(go.Indicator(
            mode="gauge", value=pct_int,
            gauge={'axis':{'range':[0,100]}, 'bar':{'color':mint}, 'bgcolor':gauge_bg}
        ))
        fig_gauge_int.add_annotation(x=0.5,y=0.40,xref="paper",yref="paper",text="International%:",showarrow=False,
                                     font=dict(color=mint_dark,size=13))
        fig_gauge_int.add_annotation(x=0.5,y=0.0,xref="paper",yref="paper",text=f"{pct_int:.1f}%",showarrow=False,
                                     font=dict(color=mint_dark,size=18))
        fig_gauge_int.update_layout(height=110, margin=dict(l=10, r=10, t=10, b=6))
        st.plotly_chart(fig_gauge_int, use_container_width=True)

        # ---- 100% stacked gender bar ----
        gcol = col_gender(active_side)
        if gcol and total_act>0:
            male = int(active_side[active_side[gcol].astype(str)=="Male"][IDCOL].nunique())
            female = int(active_side[active_side[gcol].astype(str)=="Female"][IDCOL].nunique())
        else:
            male = female = 0
        pct_m = round(100*male/total_act,1) if total_act else 0.0
        pct_f = round(100*female/total_act,1) if total_act else 0.0
        fig_gender = go.Figure()
        fig_gender.add_trace(go.Bar(x=[pct_m], y=[" "], orientation="h", name="Male",
                                    text=[f"Male {pct_m}%"], textposition="inside", insidetextanchor="middle"))
        fig_gender.add_trace(go.Bar(x=[pct_f], y=[" "], orientation="h", name="Female",
                                    text=[f"Female {pct_f}%"], textposition="inside", insidetextanchor="middle"))
        fig_gender.update_layout(barmode="stack", showlegend=False,
                                 xaxis=dict(range=[0,100], visible=False),
                                 yaxis=dict(visible=False),
                                 height=100, margin=dict(l=10,r=10,t=18,b=12))
        st.plotly_chart(fig_gender, use_container_width=True)

        # ---- Age buckets ----
        labels_age = ["Under 30","31-40","41-50","51-60","over 61"]
        if not active_side.empty and "Age" in active_side.columns:
            age = pd.to_numeric(active_side["Age"], errors="coerce")
            active_side = active_side.assign(
                Age_bucket=pd.cut(age, bins=[-np.inf,29,40,50,60,np.inf], labels=labels_age)
            )
            age_counts = (active_side.groupby("Age_bucket")[IDCOL].nunique()
                          .reindex(labels_age, fill_value=0).reset_index(name="Count"))
        else:
            age_counts = pd.DataFrame({"Age_bucket": labels_age, "Count":[0]*len(labels_age)})
        fig_age = px.bar(age_counts, x="Count", y="Age_bucket", orientation="h", text="Count")
        fig_age.update_traces(marker_color="#00A896", textposition="outside", texttemplate="%{text}")
        fig_age.update_xaxes(range=[0,35], title=None)
        fig_age.update_yaxes(title=None, autorange="reversed")
        fig_age.update_layout(height=200, margin=dict(l=10, r=10, t=0, b=12))
        st.plotly_chart(fig_age, use_container_width=True)

# ============================
# ROW 1: %PhD over time (left) + PhD by region (right) + detalle bajo la barra
# ============================
st.markdown("---")

# ---- Time series según timeframe actual ----
tmode_ts = st.session_state.get("time_mode_side", "Semestral")
IDCOL = col_id(df)
ncol_global = col_nationality(df)
dcol_global = col_degree(df)
labels_ts, phd_ts, intl_ts = build_time_series(df, tmode_ts, IDCOL, dcol_global, ncol_global)

# periodo actual = exactamente el del sidebar (si existe entre las etiquetas)
sel_lbl = st.session_state.get("sel_tf_label")
if labels_ts:
    if tmode_ts == "Anual":
        period_current = sel_lbl if sel_lbl in labels_ts else labels_ts[-1]
    elif tmode_ts == "Intersemestral":
        period_current = sel_lbl if (sel_lbl in labels_ts and _is_inter_label(sel_lbl)) else labels_ts[-1]
    else:
        period_current = sel_lbl if (sel_lbl in labels_ts and _is_semester_label(sel_lbl)) else labels_ts[-1]
else:
    period_current = None

row1_left, row1_right = st.columns([6, 4])

# Ranges by mode
if mode_now == "Part-time":
    y_min_phd, y_max_phd = 0, 30
    line_h, bar_h = 280, 220
else:
    y_min_phd, y_max_phd = 70, 100
    line_h, bar_h = 280, 220

# Override para Intersemestral: eje 0..100
if tmode_ts == "Intersemestral":
    y_min_phd, y_max_phd = 0, 100

with row1_left:
    df_pct_phd = pd.DataFrame({"Label": labels_ts, "Percent": phd_ts})
    title_phd = "% of Full-time Faculty with PhD" if mode_now == "Full-time" else "% of Part-time Professors with PhD"
    fig_phd = px.line(df_pct_phd, x="Label", y="Percent", markers=True, text="Percent", title=title_phd)
    fig_phd.update_traces(line=dict(color="#00A896", width=3), marker=dict(size=7, color="#00A896"),
                          texttemplate="%{y:.1f}%", textposition="top center")
    fig_phd.update_xaxes(type="category", categoryorder="array", categoryarray=labels_ts, tickangle=0, title=None)
    fig_phd.update_yaxes(range=[y_min_phd, y_max_phd], title=None)
    if period_current in labels_ts:
        pos = labels_ts.index(period_current)
        fig_phd.add_shape(type="rect", xref="x", yref="paper",
                          x0=pos - 0.4, x1=pos + 0.4, y0=0, y1=1,
                          fillcolor="#D0E5F5", opacity=0.35, line_width=0)
    fig_phd.update_layout(height=line_h, margin=dict(l=10, r=10, t=40, b=40), showlegend=False)
    st.plotly_chart(fig_phd, use_container_width=True)

with row1_right:
    # Usa el periodo_current para la barra por región
    if period_current is None:
        active_p = df.iloc[0:0].copy()
    else:
        if tmode_ts == "Anual":
            active_p = _filter_for_timeframe(df, "Anual", sel_year=int(period_current))
        else:
            active_p = df[df["Periodo"].astype(str).eq(str(period_current))].copy()

    dcol_here = col_degree(df)
    if dcol_here is not None and not active_p.empty:
        active_p["Degree_norm"] = normalize_degree(active_p[dcol_here])
        phd_now_all = active_p[active_p["Degree_norm"].eq("PhD")].copy()
        region_col = "Region were degree was obtained" if "Region were degree was obtained" in phd_now_all.columns else None
        if region_col:
            reg = phd_now_all[region_col].astype(str).str.strip()
            mask_valid_region = (~reg.eq("")) & (~reg.str.upper().eq("TBD"))
            phd_for_regions = phd_now_all[mask_valid_region].copy()
        else:
            phd_for_regions = pd.DataFrame(columns=phd_now_all.columns)
    else:
        phd_for_regions = pd.DataFrame(columns=df.columns)

    total_phd_valid = int(phd_for_regions[IDCOL].nunique()) if not phd_for_regions.empty else 0
    if "International Degree" in phd_for_regions.columns:
        phd_int = int(phd_for_regions[phd_for_regions["International Degree"].astype(str).str.strip().str.lower().eq("yes")][IDCOL].nunique())
    else:
        phd_int = 0

    if not phd_for_regions.empty and "Region were degree was obtained" in phd_for_regions.columns:
        reg_counts = (phd_for_regions.groupby("Region were degree was obtained")[IDCOL]
                      .nunique().sort_values(ascending=False)
                      .reset_index().rename(columns={"Region were degree was obtained": "Region", IDCOL: "Count"}))
    else:
        reg_counts = pd.DataFrame({"Region": [], "Count": []})

    title_phd_bar = f"{total_phd_valid} professors with a PhD, {phd_int} obtained it abroad" if phd_int else f"{total_phd_valid} professors with a PhD"
    fig_phd_reg = px.bar(reg_counts, x="Count", y="Region", orientation="h", title=title_phd_bar, text="Count")
    fig_phd_reg.update_traces(marker_color="#00A896", textposition="outside", texttemplate="%{text}")
    fig_phd_reg.update_xaxes(title=None, dtick=1)
    fig_phd_reg.update_yaxes(title=None, autorange="reversed")
    fig_phd_reg.update_layout(height=bar_h, margin=dict(l=10, r=10, t=50, b=6))
    st.plotly_chart(fig_phd_reg, use_container_width=True)

    # ---- Detalle bajo la barra (botón) ----
    cols_posibles = {
        "Full Name": ["Full Name","Full-Name","Full_Name","Profesor","First Name"],
        "Highest Earned Degree": ["Highest Earned Degree","Highest Degree","TÍTULO"],
        "University": ["University","University Name"],
        "Region were degree was obtained": ["Region were degree was obtained","Region"],
        "Year": ["Year","Year Earned ","Year Degree","Year Earned"]
    }
    def pick_cols(df_, mapping):
        out = {}
        for new, opts in mapping.items():
            for c in opts:
                if c in df_.columns:
                    out[new] = df_[c]; break
            if new not in out:
                out[new] = pd.Series([""]*len(df_), index=df_.index)
        return pd.DataFrame(out)

    detalle_phd = pick_cols(phd_for_regions, cols_posibles) if not phd_for_regions.empty else pd.DataFrame(columns=list(cols_posibles.keys()))
    if not detalle_phd.empty:
        try:
            pop = st.popover("🔎 Ver detalle de profesores con PhD")
        except AttributeError:
            pop = st.expander("🔎 Ver detalle de profesores con PhD")
        with pop:
            st.dataframe(detalle_phd.reset_index(drop=True), use_container_width=True)

# ============================
# ROW 2: % International over time (left) + nacionalidades (right) + detalle
# ============================
st.markdown("---")
row2_left, row2_right = st.columns([6, 4])

if mode_now == "Part-time":
    y_min_int, y_max_int = 0, 10
    line_h2, bar_h2 = 260, 220
else:
    y_min_int, y_max_int = 0, 40
    line_h2, bar_h2 = 350, 300

# Override para Intersemestral: eje 0..100
if tmode_ts == "Intersemestral":
    y_min_int, y_max_int = 0, 100

with row2_left:
    df_pct_int = pd.DataFrame({"Label": labels_ts, "Percent": intl_ts})
    title_int = "% of International Full-time Faculty" if mode_now == "Full-time" else "% of International Part-time Faculty"
    fig_int = px.line(df_pct_int, x="Label", y="Percent", markers=True, text="Percent", title=title_int)
    fig_int.update_traces(line=dict(color="#2EC4B6", width=3), marker=dict(size=7, color="#2EC4B6"),
                          texttemplate="%{y:.1f}%", textposition="top center")
    fig_int.update_xaxes(type="category", categoryorder="array", categoryarray=labels_ts, tickangle=0, title=None)
    fig_int.update_yaxes(range=[y_min_int, y_max_int], title=None)
    if period_current in labels_ts:
        pos = labels_ts.index(period_current)
        fig_int.add_shape(type="rect", xref="x", yref="paper",
                          x0=pos - 0.4, x1=pos + 0.4, y0=0, y1=1,
                          fillcolor="#D0E5F5", opacity=0.35, line_width=0)
    fig_int.update_layout(height=line_h2, margin=dict(l=10, r=10, t=40, b=40), showlegend=False)
    st.plotly_chart(fig_int, use_container_width=True)

with row2_right:
    nat_col = col_nationality(df)
    if nat_col and period_current:
        if tmode_ts == "Anual":
            active_p2 = _filter_for_timeframe(df, "Anual", sel_year=int(period_current))
        else:
            active_p2 = df[df["Periodo"].astype(str).eq(str(period_current))].copy()

        nat = active_p2[nat_col].astype(str).str.strip()
        is_col   = nat.eq("Colombian")
        is_tbd   = nat.str.upper().eq("TBD")
        is_empty = nat.eq("")
        intl_now = active_p2[(~is_col) & (~is_tbd) & (~is_empty)].copy()

        IDCOL = col_id(df)
        nat_counts = (intl_now.groupby(nat_col)[IDCOL]
                      .nunique().sort_values(ascending=False)
                      .reset_index().rename(columns={nat_col:"Nationality", IDCOL:"Count"}))
        total_intl = int(intl_now[IDCOL].nunique()) if not intl_now.empty else 0
        n_nats = int(nat_counts["Nationality"].nunique()) if not nat_counts.empty else 0
    else:
        nat_counts = pd.DataFrame({"Nationality":[], "Count":[]})
        total_intl = 0; n_nats = 0

    title_nat = f"{total_intl} international Faculty. {n_nats} different nationalities"
    fig_nat = px.bar(nat_counts, x="Count", y="Nationality", orientation="h", title=title_nat, text="Count")
    fig_nat.update_traces(marker_color="#2EC4B6", textposition="outside", texttemplate="%{text}")
    fig_nat.update_xaxes(title=None, dtick=1)
    fig_nat.update_yaxes(title=None, autorange="reversed")
    fig_nat.update_layout(height=bar_h2, margin=dict(l=10, r=10, t=50, b=6))
    st.plotly_chart(fig_nat, use_container_width=True)

    # ---- Detalle bajo la barra (botón) ----
    if not intl_now.empty:
        cols_nat = {
            "Full Name": ["Full Name","Full-Name","Full_Name","Profesor","First Name"],
            "Nationality": ["Nationality","Country of Birth"]
        }
        def pick_cols2(df_, mapping):
            out = {}
            for new, opts in mapping.items():
                for c in opts:
                    if c in df_.columns:
                        out[new] = df_[c]; break
                if new not in out:
                    out[new] = pd.Series([""]*len(df_), index=df_.index)
            return pd.DataFrame(out)

        detalle_nat = pick_cols2(intl_now, cols_nat)
        try:
            pop2 = st.popover("🔎 Ver detalle de nacionalidad (profesores)")
        except AttributeError:
            pop2 = st.expander("🔎 Ver detalle de nacionalidad (profesores)")
        with pop2:
            st.dataframe(detalle_nat.reset_index(drop=True), use_container_width=True)
