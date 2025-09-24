# ======================= Faculty Qualifications (full app) =======================
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import re
from io import BytesIO

# ------------------------ PAGE CONFIG & STYLES ------------------------
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
        .header-title { color:#21877D; font-weight:700; text-align:center; font-size:32px; }
        .header-btn { background-color:#21877D; padding:8px 16px; border:none; border-radius:8px; cursor:pointer; font-size:14px; display:inline-block; }
        a.header-btn, a.header-btn:link, a.header-btn:visited, a.header-btn:hover, a.header-btn:active { color:#ffffff !important; text-decoration:none !important; }
        .scroll-wrap-600 { max-height:600px; overflow-y:auto; }
        .scroll-wrap-400 { max-height:400px; overflow-y:auto; }
        .scroll-wrap-program { max-height:520px; overflow-y:auto; }
        </style>
        """,
        unsafe_allow_html=True
    )

# DOWNLOAD BUTTONS — minimal style
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

# ------------------------ HEADER ------------------------
st.markdown('<div class="header-title">Full-time Faculty Qalifications</div>', unsafe_allow_html=True)

# ------------------------ DATA LOAD ------------------------
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

# ------------------------ CONSTANTS & HELPERS ------------------------
MINT = "#1FA89B"
SUPPORTING = "#7FD3FF"
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
    if "inter" in s:
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

# —— utilidades de descarga ——
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

# ================== SENSITIVITY HELPERS ==================
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

def apply_ops_to_aggs(agg_ps: pd.DataFrame, agg_tipo: pd.DataFrame, ops: list, member_all_label="All") -> tuple[pd.DataFrame, pd.DataFrame]:
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

# ================== HISTORY (timeframe-aware) ==================
def _period_sort_key(p: str) -> tuple[int,int]:
    y = extract_year_from_period(p) or -1
    suf = period_suffix(p)
    try:
        suf_i = int(suf) if suf is not None else 0
    except Exception:
        suf_i = 0
    return (y, suf_i)

def build_time_axis_for_history(df_hist: pd.DataFrame):
    # Se apoya en st.session_state["time_mode"]
    time_mode = st.session_state.get("time_mode", "Semestral")

    if "_SEM" not in df_hist.columns:
        sc = _get_any(df_hist, "Semestre","Periodo","Periodo Académico","Periodo academico")
        sem = df_hist[sc].astype(str).str.strip() if sc else pd.Series([], dtype=str)
    else:
        sem = df_hist["_SEM"].astype(str).str.strip()

    if time_mode == "Semestral":
        regs = sorted(
            {s for s in sem.dropna().unique() if period_suffix(s) in {"10","20"}},
            key=_period_sort_key
        )
        x_labels = regs
    elif time_mode == "Anual":
        years = sorted({extract_year_from_period(s) for s in sem if extract_year_from_period(s)}, key=int)
        x_labels = years
    else:  # Intersemestral
        inter = sorted(
            {f"{extract_year_from_period(s)} Intersemestral" for s in sem if "inter" in str(s).lower() and extract_year_from_period(s)},
            key=lambda x: int(str(x).split()[0])
        )
        x_labels = inter

    x_map = {lab: i for i, lab in enumerate(x_labels)}
    return "_SEM", x_labels, x_map

def transform_for_time_mode_ps(df_ps: pd.DataFrame):
    time_mode = st.session_state.get("time_mode", "Semestral")
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
    # Intersemestral
    base = base[~base["_INTER_LABEL"].isna()].copy()
    g = base.groupby(["_INTER_LABEL"] + [c for c in base.columns if c.startswith("_") and c not in {"_SEM","_YEAR","_INTER_LABEL"}], dropna=False).sum(numeric_only=True).reset_index()
    if "P" in g and "S" in g:
        g["P_share"] = (g["P"] / (g["P"] + g["S"]).replace(0, pd.NA)) * 100
    return g.rename(columns={"_INTER_LABEL":"_SEM"})

def transform_for_time_mode_tipo(df_tipo: pd.DataFrame, share_col_name: str):
    time_mode = st.session_state.get("time_mode", "Semestral")
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
    # Intersemestral
    base = base[~base["_INTER_LABEL"].isna()].copy()
    keys = ["_INTER_LABEL"] + [c for c in base.columns if c.startswith("_") and c not in {"_SEM","_YEAR","_INTER_LABEL"}]
    g = base.groupby(keys, dropna=False)[cats].sum().reset_index()
    den = (g[cats].sum(axis=1)).replace(0, pd.NA)
    if share_col_name == "SA_share":
        g["SA_share"] = (g["SA"] / den) * 100
    else:
        g["OTHER_share"] = (g["OTHER"] / den) * 100
    return g.rename(columns={"_INTER_LABEL":"_SEM"})

def draw_history(fig_title, level_name, level_values, metric_kind, total_series_builders, agg_ps_all, agg_tipo_all, x_labels, x_map, sel_x):
    palette = px.colors.qualitative.Safe + px.colors.qualitative.Bold + px.colors.qualitative.Pastel
    color_map = {a: palette[i % len(palette)] for i, a in enumerate(level_values)}

    st.markdown(f"<h4 style='margin:0 0 6px 0; font-weight:500;'>{fig_title}</h4>", unsafe_allow_html=True)
    sel_col, radio_col = st.columns([6,4])
    options = ["(All)", "(TOTAL)"] + level_values

    with sel_col:
        opt = st.selectbox("", options, index=0, key=f"{level_name}_filter", label_visibility="collapsed")
    with radio_col:
        metric_choice = st.radio("", ["%P", "%SA", "%OTHER"],
                                 index={"%P":0, "%SA":1, "%OTHER":2}[metric_kind],
                                 horizontal=True, key=f"metric_{level_name}", label_visibility="collapsed")

    fig = go.Figure()

    # --- series según métrica ---
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
                    marker=dict(size=6, color=color_map[a]), line=dict(width=2, color=color_map[a]),
                    hovertemplate=a + "<br>%{y:.1f}%<extra></extra>"
                ))
        elif opt == "(TOTAL)":
            sub = total_series_builders["P"].copy()
            sub["x"] = sub["_SEM"].map(x_map)
            sub = sub.sort_values("x")
            fig.add_trace(go.Scatter(
                x=sub["x"], y=sub["P_share"], mode="lines+markers", name="TOTAL",
                marker=dict(size=6, color=TOTAL_SERIES_COLOR), line=dict(width=2, color=TOTAL_SERIES_COLOR),
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
        y_min, bad_high = 40, False

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
                    marker=dict(size=6, color=color_map[a]), line=dict(width=2, color=color_map[a]),
                    hovertemplate=a + "<br>%{y:.1f}%<extra></extra>"
                ))
        elif opt == "(TOTAL)":
            sub = total_series_builders["SA"].copy()
            sub["x"] = sub["_SEM"].map(x_map)
            sub = sub.sort_values("x")
            fig.add_trace(go.Scatter(
                x=sub["x"], y=sub[share_col], mode="lines+markers", name="TOTAL",
                marker=dict(size=6, color=TOTAL_SERIES_COLOR), line=dict(width=2, color=TOTAL_SERIES_COLOR),
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
        y_min, bad_high = 20, False

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
                    marker=dict(size=6, color=color_map[a]), line=dict(width=2, color=color_map[a]),
                    hovertemplate=a + "<br>%{y:.1f}%<extra></extra>"
                ))
        elif opt == "(TOTAL)":
            sub = total_series_builders["OTHER"].copy()
            sub["x"] = sub["_SEM"].map(x_map)
            sub = sub.sort_values("x")
            fig.add_trace(go.Scatter(
                x=sub["x"], y=sub[share_col], mode="lines+markers", name="TOTAL",
                marker=dict(size=6, color=TOTAL_SERIES_COLOR), line=dict(width=2, color=TOTAL_SERIES_COLOR),
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
        y_min, bad_high = 0, True
        y_max = 40

    # Zonas “malas” y líneas de referencia
    if bad_high:
        fig.update_layout(shapes=[dict(type="rect", xref="paper", yref="y", x0=0, x1=1, y0=thr, y1=100, fillcolor="#FDE2E2", opacity=0.35, layer="below", line_width=0)])
        fig.add_hline(y=thr, line_color="#F5A3A3", line_dash="dash")
    else:
        fig.update_layout(shapes=[dict(type="rect", xref="paper", yref="y", x0=0, x1=1, y0=0, y1=thr, fillcolor="#FDE2E2", opacity=0.35, layer="below", line_width=0)])
        fig.add_hline(y=thr, line_color="red", line_dash="dash")

    # Barra vertical del período seleccionado
    if sel_x is not None:
        fig.add_vrect(x0=sel_x-0.5, x1=sel_x+0.5, fillcolor="#E8FAF7", opacity=0.5, layer="below", line_width=0)

    # Ejes y render
    tickvals = list(range(len(x_labels)))
    ticktext = [str(x) for x in x_labels]
    if metric_choice == "%OTHER":
        fig.update_layout(xaxis=dict(tickmode="array", tickvals=tickvals, ticktext=ticktext),
                          yaxis=dict(range=[y_min, y_max]))
    else:
        fig.update_layout(xaxis=dict(tickmode="array", tickvals=tickvals, ticktext=ticktext),
                          yaxis=dict(range=[y_min, 100]))
    fig.update_xaxes(title=None)
    fig.update_yaxes(title=None)
    st.plotly_chart(fig, use_container_width=True)

    # ===== Datos para descargar (lo visible) =====
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
            for a in level_values: base_cols[a] = _series_for(a, ycol)
        elif opt == "(TOTAL)":
            sub = total_series_builders["P"].set_index("_SEM")["P_share"].to_dict()
            base_cols["TOTAL"] = [sub.get(x, None) for x in x_labels]
        else:
            base_cols[opt] = _series_for(opt, ycol)
    elif metric_choice == "%SA":
        ycol = "SA_share"
        base_cols = {}
        if opt == "(All)":
            for a in level_values: base_cols[a] = _series_for(a, ycol)
        elif opt == "(TOTAL)":
            sub = total_series_builders["SA"].set_index("_SEM")[ycol].to_dict()
            base_cols["TOTAL"] = [sub.get(x, None) for x in x_labels]
        else:
            base_cols[opt] = _series_for(opt, ycol)
    else:
        ycol = "OTHER_share"
        base_cols = {}
        if opt == "(All)":
            for a in level_values: base_cols[a] = _series_for(a, ycol)
        elif opt == "(TOTAL)":
            sub = total_series_builders["OTHER"].set_index("_SEM")[ycol].to_dict()
            base_cols["TOTAL"] = [sub.get(x, None) for x in x_labels]
        else:
            base_cols[opt] = _series_for(opt, ycol)

    export_df = pd.DataFrame({"Period": x_labels, **base_cols})
    fname = f"chart_{_slugify(fig_title)}_{_slugify(metric_choice)}_{_slugify(opt)}_{_slugify(st.session_state.get('sel_label','sel'))}.xlsx"
    _download_xlsx_button(export_df, fname, key=f"dl_hist_{_slugify(fig_title)}_{metric_choice}_{_slugify(opt)}_{_slugify(st.session_state.get('sel_label','sel'))}", label="⬇️ Datos de la gráfica (Excel)")

# ============== NORMALIZACIÓN BÁSICA EN CARTELERA ==============
col_sem  = _get_any(df_car, "Semestre","Periodo","Periodo Académico","Periodo academico")
if "_SEM" not in df_car.columns and col_sem:
    df_car["_SEM"] = df_car[col_sem].astype(str).str.strip()
else:
    df_car["_SEM"] = df_car.get("_SEM", pd.Series(dtype=str))
df_car["_YEAR"] = df_car["_SEM"].map(extract_year_from_period)
df_car["_IS_INTER"] = df_car["_SEM"].str.lower().str.contains("inter", na=False)

# ================== TIMEFRAME FILTERS ==================
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

# ================== SIDEBAR (con toggle & sin claves duplicadas) ==================
SEMESTRAL_PERIODS = list_periods_semestral()
YEARS_ALL        = list_years_from_sem()
INTER_YEARS      = years_with_inter()

with st.sidebar:
    # --- Sensitivity FIRST (toggle on/off visual) ---
    st.markdown("#### Sensitivity analysis")
    sens_mode = st.toggle("Enable sensitivity mode", value=st.session_state.get("sens_mode", False), key="sens_mode")

    # Placeholder para "Apply to" — se llena UNA SOLA VEZ más abajo
    sens_member_placeholder = st.empty()

    if "sens_ops" not in st.session_state:
        st.session_state.sens_ops = []

    if sens_mode:
        st.session_state.setdefault("sens_cat_ps", "None")
        st.session_state.setdefault("sens_cat_qual", "None")

        st.selectbox("P/S category", ["None", "P", "S"], key="sens_cat_ps")
        st.selectbox("Qualification", ["None", "SA", "PA", "SP", "IP", "OTHER"], key="sens_cat_qual")

        st.number_input("Professors", min_value=1, step=1, value=1, key="sens_count")
        st.number_input("Credits per professor", min_value=0.0, step=0.5, value=3.0, key="sens_credits")

        c_add, c_reset = st.columns(2)
        with c_add:
            if st.button("Add", use_container_width=True, key="sens_add"):
                ops_to_add = []
                member_val = st.session_state.get("sens_member", "All")
                cnt   = int(st.session_state.get("sens_count", 1))
                cred  = float(st.session_state.get("sens_credits", 3.0))
                if st.session_state.get("sens_cat_ps") and st.session_state["sens_cat_ps"] != "None":
                    ops_to_add.append({
                        "scope": "PS", "cat": st.session_state["sens_cat_ps"],
                        "member": member_val, "credits": cred, "count": cnt
                    })
                if st.session_state.get("sens_cat_qual") and st.session_state["sens_cat_qual"] != "None":
                    ops_to_add.append({
                        "scope": "QUAL", "cat": st.session_state["sens_cat_qual"],
                        "member": member_val, "credits": cred, "count": cnt
                    })
                if ops_to_add:
                    st.session_state.sens_ops.extend(ops_to_add)
                    st.success("Added.")
        with c_reset:
            if st.button("Reset to original", use_container_width=True, key="sens_reset"):
                st.session_state.sens_ops = []
                st.success("Reset.")

    if not sens_mode:
        st.markdown("### Go to KPI")
        options = {
            "1 Full-time Composition": "https://facultycompositiondashboardpy-dtacyzfa3otmpbewqc5axu.streamlit.app/",
            "2 Full-time Staffing Levels": "https://facultystaffinglevelsdashboardpy-phv4t8jzbyyz5rrepqttuf.streamlit.app/",
            "3 Distribution by Academic Area": "https://facultydistributionareadashboardpy-yzwpiqdlukfdp6qcygxjhj.streamlit.app/",
            "4 Faculty Demographics": "https://facultydemographicsdashboardpy-kmsnpswxs35psbqtdtvb6y.streamlit.app/",
            "5 Full-time Faculty Questionnaire": "https://full-timefacultyactivitiespy-bbe7fmmyrxvssadnygm4fx.streamlit.app/",
            "6 Faculty Qualifications": "https://facultyqualificationspy-drvj3wpyrxvm2lrnafdwx5.streamlit.app/",
        }
        choices = list(options.keys())
        default_label = "6 Faculty Qualifications"
        default_idx = choices.index(default_label) if default_label in choices else 0
        choice = st.selectbox("Select…", choices, index=default_idx, key="kpi_nav_top")
        st.link_button("Open", options[choice], use_container_width=True)

    st.markdown('---')
    st.markdown("#### Timeframe")
    st.session_state.setdefault("time_mode", "Semestral")
    time_mode = st.radio("Timeframe", ["Semestral", "Anual", "Intersemestral"], key="time_mode", label_visibility="collapsed", horizontal=False)

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

    st.session_state["sel_label"] = sel_label

    st.session_state.setdefault("view_mode", "By Academic Area")
    view_mode = st.selectbox("View", ["By Program", "By Academic Area", "By Field"], key="view_mode")

    # Placeholder para DOWNLOAD DB (se llena más abajo)
    dl_bd_placeholder = st.empty()

# ================== DESPUÉS DEL SIDEBAR: filtros & sensibilidad ==================
df_car_base = df_car.copy()
df_fd_base  = df_fd.copy()

df_car_filt_all = filter_df_car(df_car_base, time_mode, sel_year, sel_sem)
df_fd_f         = filter_df_fd(df_fd_base, time_mode, sel_year, sel_sem)

# Sidebar: botón de descarga de la BD filtrada
if 'dl_bd_placeholder' in locals():
    safe = _sanitize_for_export(df_car_filt_all)
    dl_bd_placeholder.download_button(
        "Download DB (Excel)",
        data=_xlsx_bytes(safe),
        file_name=f"BD_Cartelera_{_slugify(sel_label)}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"dl_bd_{_slugify(sel_label)}"
    )

# --------- Sensitivity: llenar “Apply to” UNA SOLA VEZ ---------
if st.session_state.get("sens_mode", False):
    col_areaCourse = _get_any(df_car_filt_all, "Area del curso","Área del curso","Area del Curso","AREA DEL CURSO")
    col_field      = _get_any(df_car_filt_all, "Field","FIELD","Campo","Área de conocimiento")
    program_col    = _get_any(df_car_filt_all, "Program","PROGRAM","program","Materia")
    members = build_member_list_for_view(df_car_filt_all, view_mode, col_areaCourse, col_field, program_col)
    with st.sidebar:
        sens_member_placeholder.selectbox("Apply to", members, key="sens_member")

SENS = {"on": bool(st.session_state.get("sens_mode", False)), "ops": st.session_state.get("sens_ops", [])}

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

# ---------- Stylers ----------
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

def style_diverging_simple(df_: pd.DataFrame, colname: str):
    sty = pd.DataFrame('', index=df_.index, columns=df_.columns)
    if colname not in df_.columns:
        return sty
    vals = pd.to_numeric(df_[colname], errors='coerce').fillna(0.0)
    sty.loc[vals < 0, colname] = 'background-color:#FDE2E2;'
    sty.loc[vals > 0, colname] = 'background-color:#E6F4EA;'
    if df_.columns[0] in df_.columns:
        first_col = df_.columns[0]
        mask_total = df_[first_col].astype(str).str.upper().eq("TOTAL")
        for c in df_.columns:
            sty.loc[mask_total, c] = (sty.loc[mask_total, c].astype(str) + 'font-weight:700;').str.replace(';;',';', regex=False)
    return sty

# ================== PRINCIPAL ==================
st.markdown("---")
st.subheader(f"Faculty Sufficiency and Qualifications — {sel_label}")

# ====== NORMALIZACIÓN BASE PARA CARTELERA ======
if not all([col_cred, col_tipoC, col_areaCourse]):
    st.error("Missing columns in 'BD Cartelera 2020-2025': 'Credits', 'TIPO', and/or 'Academic Area (course)'.")
else:
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

    # Excluir materias específicas (si aplica)
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
        # ============================ VISTAS ============================
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

        # ------------------- BY ACADEMIC AREA -------------------
        if view_mode == "By Academic Area":
            colT, colG = st.columns([6,6], gap="large")

            # Agregaciones frame seleccionado
            agg_tipo = (fil.groupby(["_AREA","_TIPO"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["SA","PA","SP","IP","OTHER"]:
                if k not in agg_tipo.columns: agg_tipo[k] = 0.0
            agg_tipo = agg_tipo[["SA","PA","SP","IP","OTHER"]]

            agg_ps = (fil.groupby(["_AREA","_PS"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["P","S"]:
                if k not in agg_ps.columns: agg_ps[k] = 0.0
            agg_ps = agg_ps[["P","S"]]

            with colT:
                base_agg_ps = agg_ps.copy()
                base_agg_tipo = agg_tipo.copy()

                if SENS["on"] and SENS["ops"]:
                    mod_agg_ps, mod_agg_tipo = apply_ops_to_aggs(base_agg_ps, base_agg_tipo, SENS["ops"])
                else:
                    mod_agg_ps, mod_agg_tipo = base_agg_ps, base_agg_tipo

                metrics_tbl = build_percent_table("Academic Area", mod_agg_tipo, mod_agg_ps)

                if SENS["on"] and SENS["ops"]:
                    den0 = (base_agg_ps["P"] + base_agg_ps["S"]).replace(0, pd.NA)
                    den1 = (mod_agg_ps["P"] + mod_agg_ps["S"]).replace(0, pd.NA)
                    pct0 = (base_agg_ps["P"] / den0 * 100).fillna(0.0)
                    pct1 = (mod_agg_ps["P"]  / den1 * 100).fillna(0.0)
                    impP = (pct1 - pct0).reindex(mod_agg_ps.index).round(2)
                    cats = ["SA","PA","SP","IP","OTHER"]
                    d0 = base_agg_tipo[cats].sum(axis=1).replace(0, pd.NA)
                    d1 = mod_agg_tipo[cats].sum(axis=1).replace(0, pd.NA)
                    sa0 = (base_agg_tipo["SA"]/d0*100).fillna(0.0)
                    sa1 = (mod_agg_tipo["SA"] /d1*100).fillna(0.0)
                    ot0 = (base_agg_tipo["OTHER"]/d0*100).fillna(0.0)
                    ot1 = (mod_agg_tipo["OTHER"]/d1*100).fillna(0.0)
                    impSA = (sa1 - sa0).reindex(mod_agg_tipo.index).round(2)
                    impOT = (ot1 - ot0).reindex(mod_agg_tipo.index).round(2)
                    mt = metrics_tbl.set_index("Academic Area")
                    mt["Impact (Δ%P)"] = impP
                    mt["Impact (Δ%SA)"] = impSA
                    mt["Impact (Δ%OTHER)"] = impOT
                    bt = build_percent_table("Academic Area", base_agg_tipo, base_agg_ps).set_index("Academic Area")
                    mt.loc["TOTAL","Impact (Δ%P)"] = (mt.loc["TOTAL","%P"] - bt.loc["TOTAL","%P"]).round(2)
                    mt.loc["TOTAL","Impact (Δ%SA)"] = (mt.loc["TOTAL","%SA"] - bt.loc["TOTAL","%SA"]).round(2)
                    mt.loc["TOTAL","Impact (Δ%OTHER)"] = (mt.loc["TOTAL","%OTHER"] - bt.loc["TOTAL","%OTHER"]).round(2)
                    metrics_tbl = mt.reset_index()

                _download_xlsx_button(
                    metrics_tbl, f"table_ByArea_{_slugify(sel_label)}.xlsx",
                    key=f"dl_tbl_area_{_slugify(sel_label)}", label="⬇️ Download table (Excel)"
                )

                if SENS["on"] and SENS["ops"]:
                    styled_tbl = (
                        metrics_tbl.style
                        .format({"%P":"{:.1f}%","%S":"{:.1f}%","%SA":"{:.1f}%","%OTHER":"{:.1f}%","Impact (Δ%P)":"{:+.2f}","Impact (Δ%SA)":"{:+.2f}","Impact (Δ%OTHER)":"{:+.2f}"})
                        .apply(lambda df_: style_diverging_simple(df_, "Impact (Δ%P)"), axis=None)
                        .hide(axis="index")
                    )
                else:
                    styled_tbl = (
                        metrics_tbl.style
                        .format({"%P": "{:.1f}%", "%S": "{:.1f}%", "%SA": "{:.1f}%", "%OTHER": "{:.1f}%"})
                        .apply(style_percent_tables, id_col="Academic Area", axis=None)
                        .hide(axis="index")
                    )
                st.markdown(f"<div class='scroll-wrap-400'>{styled_tbl.to_html(escape=False)}</div>", unsafe_allow_html=True)

            # --- HISTÓRICOS por Área (timeframe-aware) ---
            df_hist = df_car_global.copy()
            agg_ps_all = (df_hist.groupby(["_SEM","_AREA","_PS"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["P","S"]:
                if k not in agg_ps_all.columns: agg_ps_all[k] = 0.0
            agg_ps_all["P_share"] = (agg_ps_all["P"] / (agg_ps_all["P"] + agg_ps_all["S"]).replace(0, pd.NA)) * 100
            agg_ps_all = agg_ps_all.reset_index()
            agg_tipo_all = (df_hist.groupby(["_SEM","_AREA","_TIPO"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["SA","PA","SP","IP","OTHER"]:
                if k not in agg_tipo_all.columns: agg_tipo_all[k] = 0.0
            den_all = (agg_tipo_all[["SA","PA","SP","IP","OTHER"]].sum(axis=1)).replace(0, pd.NA)
            agg_tipo_all["SA_share"]    = (agg_tipo_all["SA"]    / den_all) * 100
            agg_tipo_all["OTHER_share"] = (agg_tipo_all["OTHER"] / den_all) * 100
            agg_tipo_all = agg_tipo_all.reset_index()

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

            key_col, x_labels, x_map = build_time_axis_for_history(df_hist)
            sel_x = None
            if time_mode == "Semestral" and sel_sem and str(sel_sem) in x_map: sel_x = x_map[str(sel_sem)]
            if time_mode == "Anual" and sel_year in x_map: sel_x = x_map[sel_year]
            if time_mode == "Intersemestral":
                lab = f"{sel_year} Intersemestral"
                if lab in x_map: sel_x = x_map[lab]

            areas_all = sorted(set(agg_ps_all_tm["_AREA"].astype(str).unique()) | set(agg_tipo_all_tm["_AREA"].astype(str).unique()))
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

                    if SENS["on"] and SENS["ops"]:
                        mod_agg_ps, mod_agg_tipo = apply_ops_to_aggs(base_agg_ps, base_agg_tipo, SENS["ops"])
                    else:
                        mod_agg_ps, mod_agg_tipo = base_agg_ps, base_agg_tipo

                    metrics_tbl_f = build_percent_table("Field", mod_agg_tipo, mod_agg_ps)

                    if SENS["on"] and SENS["ops"]:
                        den0 = (base_agg_ps["P"] + base_agg_ps["S"]).replace(0, pd.NA)
                        den1 = (mod_agg_ps["P"] + mod_agg_ps["S"]).replace(0, pd.NA)
                        pct0 = (base_agg_ps["P"] / den0 * 100).fillna(0.0)
                        pct1 = (mod_agg_ps["P"]  / den1 * 100).fillna(0.0)
                        impP = (pct1 - pct0).reindex(mod_agg_ps.index).round(2)
                        cats = ["SA","PA","SP","IP","OTHER"]
                        d0 = base_agg_tipo[cats].sum(axis=1).replace(0, pd.NA)
                        d1 = mod_agg_tipo[cats].sum(axis=1).replace(0, pd.NA)
                        sa0 = (base_agg_tipo["SA"]/d0*100).fillna(0.0)
                        sa1 = (mod_agg_tipo["SA"] /d1*100).fillna(0.0)
                        ot0 = (base_agg_tipo["OTHER"]/d0*100).fillna(0.0)
                        ot1 = (mod_agg_tipo["OTHER"]/d1*100).fillna(0.0)
                        impSA = (sa1 - sa0).reindex(mod_agg_tipo.index).round(2)
                        impOT = (ot1 - ot0).reindex(mod_agg_tipo.index).round(2)
                        mt = metrics_tbl_f.set_index("Field")
                        mt["Impact (Δ%P)"] = impP
                        mt["Impact (Δ%SA)"] = impSA
                        mt["Impact (Δ%OTHER)"] = impOT
                        bt = build_percent_table("Field", base_agg_tipo, base_agg_ps).set_index("Field")
                        mt.loc["TOTAL","Impact (Δ%P)"] = (mt.loc["TOTAL","%P"] - bt.loc["TOTAL","%P"]).round(2)
                        mt.loc["TOTAL","Impact (Δ%SA)"] = (mt.loc["TOTAL","%SA"] - bt.loc["TOTAL","%SA"]).round(2)
                        mt.loc["TOTAL","Impact (Δ%OTHER)"] = (mt.loc["TOTAL","%OTHER"] - bt.loc["TOTAL","%OTHER"]).round(2)
                        metrics_tbl_f = mt.reset_index()

                    _download_xlsx_button(
                        metrics_tbl_f,
                        f"table_ByField_{_slugify(sel_label)}.xlsx",
                        key=f"dl_tbl_field_{_slugify(sel_label)}",
                        label="⬇️ Download table (Excel)"
                    )

                    if SENS["on"] and SENS["ops"]:
                        styled_tbl_f = (
                            metrics_tbl_f.style
                            .format({"%P":"{:.1f}%","%S":"{:.1f}%","%SA":"{:.1f}%","%OTHER":"{:.1f}%","Impact (Δ%P)":"{:+.2f}","Impact (Δ%SA)":"{:+.2f}","Impact (Δ%OTHER)":"{:+.2f}"})
                            .apply(lambda df_: style_diverging_simple(df_, "Impact (Δ%P)"), axis=None)
                            .hide(axis="index")
                        )
                    else:
                        styled_tbl_f = (
                            metrics_tbl_f.style
                            .format({"%P":"{:.1f}%","%S":"{:.1f}%","%SA":"{:.1f}%","%OTHER":"{:.1f}%"})
                            .apply(style_percent_tables, id_col="Field", axis=None)
                            .hide(axis="index")
                        )
                    st.markdown(f"<div class='scroll-wrap-400'>{styled_tbl_f.to_html(escape=False)}</div>", unsafe_allow_html=True)

                # HISTÓRICOS Field (timeframe-aware)
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

                tot_by_sem_f = (df_hist_f.groupby(["_SEM","_PS"])["_CRED"].sum().unstack(fill_value=0.0))
                for k in ["P","S"]:
                    if k not in tot_by_sem_f.columns: tot_by_sem_f[k] = 0.0
                tot_by_sem_f["P_share"] = (tot_by_sem_f["P"] / (tot_by_sem_f["P"] + tot_by_sem_f["S"]).replace(0, pd.NA)) * 100
                tot_by_sem_f = tot_by_sem_f.reset_index()

                agg_ps_all_tm   = transform_for_time_mode_ps(agg_ps_all_f.rename(columns={"_FIELD":"__LEVEL__"})).rename(columns={"__LEVEL__":"_FIELD"})
                agg_tipo_all_tm = transform_for_time_mode_tipo(agg_tipo_all_f.rename(columns={"_FIELD":"__LEVEL__"}), "SA_share").rename(columns={"__LEVEL__":"_FIELD"})
                agg_tipo_all_tm_other = transform_for_time_mode_tipo(agg_tipo_all_f.rename(columns={"_FIELD":"__LEVEL__"}), "OTHER_share").rename(columns={"__LEVEL__":"_FIELD"})
                agg_tipo_all_tm = agg_tipo_all_tm.drop(columns=[c for c in ["OTHER_share"] if c in agg_tipo_all_tm], errors="ignore")\
                                                 .merge(agg_tipo_all_tm_other[["_SEM","_FIELD","OTHER","SA","PA","SP","IP","OTHER_share"]], on=["_SEM","_FIELD","SA","PA","SP","IP","OTHER"], how="outer")

                tot_by_sem_P_tm = transform_for_time_mode_ps(tot_by_sem_f.copy())
                tot_by_sem_tipo_sa_tm = transform_for_time_mode_tipo(tot_by_sem_f.copy(), "SA_share")  # (no se usa, mantenido por simetría)
                tot_by_sem_tipo_tm = tot_by_sem_tipo_sa_tm  # alias

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
                                 total_series_builders={"P": tot_by_sem_P_tm, "SA": agg_tipo_all_tm, "OTHER": agg_tipo_all_tm},
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
                    if SENS["on"] and SENS["ops"]:
                        mod_agg_ps, mod_agg_tipo = apply_ops_to_aggs(base_agg_ps, base_agg_tipo, SENS["ops"])
                    else:
                        mod_agg_ps, mod_agg_tipo = base_agg_ps, base_agg_tipo

                    metrics_tbl_m = build_percent_table("Program", mod_agg_tipo, mod_agg_ps)

                    if SENS["on"] and SENS["ops"]:

                        den0 = (base_agg_ps["P"] + base_agg_ps["S"]).replace(0, pd.NA)
                        den1 = (mod_agg_ps["P"] + mod_agg_ps["S"]).replace(0, pd.NA)
                        pct0 = (base_agg_ps["P"] / den0 * 100).fillna(0.0)
                        pct1 = (mod_agg_ps["P"]  / den1 * 100).fillna(0.0)
                        impP = (pct1 - pct0).reindex(mod_agg_ps.index).round(2)

                        cats = ["SA","PA","SP","IP","OTHER"]
                        d0 = base_agg_tipo[cats].sum(axis=1).replace(0, pd.NA)
                        d1 = mod_agg_tipo[cats].sum(axis=1).replace(0, pd.NA)
                        sa0 = (base_agg_tipo["SA"]/d0*100).fillna(0.0)
                        sa1 = (mod_agg_tipo["SA"]/d1*100).fillna(0.0)
                        ot0 = (base_agg_tipo["OTHER"]/d0*100).fillna(0.0)
                        ot1 = (mod_agg_tipo["OTHER"]/d1*100).fillna(0.0)
                        mt = metrics_tbl_m.set_index("Program")
                        mt["Impact (Δ%P)"] = impP
                        mt["Impact (Δ%SA)"] = (sa1 - sa0).reindex(mod_agg_tipo.index).round(2)
                        mt["Impact (Δ%OTHER)"] = (ot1 - ot0).reindex(mod_agg_tipo.index).round(2)
                        bt = build_percent_table("Program", base_agg_tipo, base_agg_ps).set_index("Program")
                        mt.loc["TOTAL","Impact (Δ%P)"] = (mt.loc["TOTAL","%P"] - bt.loc["TOTAL","%P"]).round(2)
                        mt.loc["TOTAL","Impact (Δ%SA)"] = (mt.loc["TOTAL","%SA"] - bt.loc["TOTAL","%SA"]).round(2)
                        mt.loc["TOTAL","Impact (Δ%OTHER)"] = (mt.loc["TOTAL","%OTHER"] - bt.loc["TOTAL","%OTHER"]).round(2)
                        metrics_tbl_m = mt.reset_index()

                    _download_xlsx_button(metrics_tbl_m, f"table_ByProgram_{_slugify(sel_label)}.xlsx",
                                          key=f"dl_tbl_prog_{_slugify(sel_label)}", label="⬇️ Download table (Excel)")

                    if SENS["on"] and SENS["ops"]:
                        styled_tbl_m = (
                            metrics_tbl_m.style
                            .format({"%P":"{:.1f}%","%S":"{:.1f}%","%SA":"{:.1f}%","%OTHER":"{:.1f}%","Impact (Δ%P)":"{:+.2f}","Impact (Δ%SA)":"{:+.2f}","Impact (Δ%OTHER)":"{:+.2f}"})
                            .apply(lambda df_: style_diverging_simple(df_, "Impact (Δ%P)"), axis=None)
                            .hide(axis="index")
                        )
                    else:
                        styled_tbl_m = (
                            metrics_tbl_m.style
                            .format({"%P":"{:.1f}%","%S":"{:.1f}%","%SA":"{:.1f}%","%OTHER":"{:.1f}%"})
                            .apply(style_percent_tables, id_col="Program", axis=None)
                            .hide(axis="index")
                        )
                    st.markdown(f"<div class='scroll-wrap-program'>{styled_tbl_m.to_html(escape=False)}</div>", unsafe_allow_html=True)

                # HISTÓRICO — Program
                df_hist_m = df_car_global.copy()
                df_hist_m["_MAT"] = df_hist_m[program_col].astype(str).str.strip()
                agg_ps_all_m = (df_hist_m.groupby(["_SEM","_MAT","_PS"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
                for k in ["P","S"]:
                    if k not in agg_ps_all_m.columns: agg_ps_all_m[k] = 0.0
                agg_ps_all_m["P_share"] = (agg_ps_all_m["P"] / (agg_ps_all_m["P"] + agg_ps_all_m["S"]).replace(0, pd.NA)) * 100
                agg_ps_all_m = agg_ps_all_m.reset_index()

                tot_by_sem_m = (df_hist_m.groupby(["_SEM","_PS"])["_CRED"].sum().unstack(fill_value=0.0))
                for k in ["P","S"]:
                    if k not in tot_by_sem_m.columns: tot_by_sem_m[k] = 0.0
                tot_by_sem_m["P_share"] = (tot_by_sem_m["P"] / (tot_by_sem_m["P"] + tot_by_sem_m["S"]).replace(0, pd.NA)) * 100
                tot_by_sem_m = tot_by_sem_m.reset_index()

                agg_ps_all_tm   = transform_for_time_mode_ps(agg_ps_all_m.rename(columns={"_MAT":"__LEVEL__"})).rename(columns={"__LEVEL__":"_MAT"})
                tot_by_sem_m_tm = transform_for_time_mode_ps(tot_by_sem_m.copy())

                key_col, x_labels, x_map = build_time_axis_for_history(df_hist_m)
                sel_x = None
                if time_mode == "Semestral" and sel_sem and str(sel_sem) in x_map: sel_x = x_map[str(sel_sem)]
                if time_mode == "Anual" and sel_year in x_map: sel_x = x_map[sel_year]
                if time_mode == "Intersemestral":
                    lab = f"{sel_year} Intersemestral"
                    if lab in x_map: sel_x = x_map[lab]
                programs_all = sorted(set(agg_ps_all_tm["_MAT"].astype(str).unique()))

                with colM_R:
                    draw_history("Evolution by Academic Program",
                                 level_name="_MAT",
                                 level_values=programs_all,
                                 metric_kind="%P",
                                 total_series_builders={"P": tot_by_sem_m_tm},
                                 agg_ps_all=agg_ps_all_tm, agg_tipo_all=None,
                                 x_labels=x_labels, x_map=x_map, sel_x=sel_x)

# --------------------------
# CREDIT SUMS (EXPANDER)
# --------------------------
try:
    period_df = df_car_filt_all.copy()
    if "_CRED"  not in period_df.columns and col_cred:  period_df["_CRED"]  = pd.to_numeric(period_df[col_cred], errors="coerce").fillna(0.0)
    if "_PS"    not in period_df.columns and col_ps_C:  period_df["_PS"]    = _norm_str(period_df[col_ps_C]).map(normalize_ps)
    if "_TIPO"  not in period_df.columns and col_tipoC: period_df["_TIPO"]  = _norm_str(period_df[col_tipoC]).map(normalize_tipo)
    if "_AREA"  not in period_df.columns and col_areaCourse: period_df["_AREA"] = period_df[col_areaCourse].astype(str).str.strip()
    if "_FIELD" not in period_df.columns and col_field:      period_df["_FIELD"] = period_df[col_field].astype(str).str.strip()
    if "_MAT"   not in period_df.columns and col_prog:       period_df["_MAT"] = period_df[col_prog].astype(str).str.strip()

    view = st.session_state.view_mode if "view_mode" in st.session_state else "By Academic Area"
    if view == "By Academic Area":
        dim_col, dim_label = "_AREA", "Academic Area"
    elif view == "By Field":
        dim_col, dim_label = "_FIELD", "Field"
    else:
        dim_col, dim_label = "_MAT", "Program"

    if dim_col in period_df.columns:
        base_index = period_df.groupby(dim_col)["_CRED"].sum().sort_values(ascending=False)
        idx = base_index.index

        sum_total = base_index.rename("Credit Sum")
        sum_P  = (period_df[period_df["_PS"]   == "P"     ].groupby(dim_col)["_CRED"].sum().reindex(idx, fill_value=0.0)).rename("P Sum")
        sum_S  = (period_df[period_df["_PS"]   == "S"     ].groupby(dim_col)["_CRED"].sum().reindex(idx, fill_value=0.0)).rename("S Sum")
        sum_SA = (period_df[period_df["_TIPO"] == "SA"    ].groupby(dim_col)["_CRED"].sum().reindex(idx, fill_value=0.0)).rename("SA Sum")
        sum_PA = (period_df[period_df["_TIPO"] == "PA"    ].groupby(dim_col)["_CRED"].sum().reindex(idx, fill_value=0.0)).rename("PA Sum")
        sum_SP = (period_df[period_df["_TIPO"] == "SP"    ].groupby(dim_col)["_CRED"].sum().reindex(idx, fill_value=0.0)).rename("SP Sum")
        sum_IP = (period_df[period_df["_TIPO"] == "IP"    ].groupby(dim_col)["_CRED"].sum().reindex(idx, fill_value=0.0)).rename("IP Sum")
        sum_OT = (period_df[period_df["_TIPO"] == "OTHER" ].groupby(dim_col)["_CRED"].sum().reindex(idx, fill_value=0.0)).rename("OTHER Sum")

        tbl = pd.concat([sum_total, sum_P, sum_S, sum_SA, sum_PA, sum_SP, sum_IP, sum_OT], axis=1).fillna(0.0)

        if SENS.get("on"):
            agg_tipo = (period_df.groupby([dim_col,"_TIPO"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["SA","PA","SP","IP","OTHER"]:
                if k not in agg_tipo.columns: agg_tipo[k] = 0.0
            agg_ps = (period_df.groupby([dim_col,"_PS"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["P","S"]:
                if k not in agg_ps.columns: agg_ps[k] = 0.0
            agg_ps = agg_ps[["P","S"]]; agg_tipo = agg_tipo[["SA","PA","SP","IP","OTHER"]]

            mod_ps, mod_tipo = apply_ops_to_aggs(agg_ps, agg_tipo, SENS.get("ops", []), member_all_label="All")
            tbl["P Sum"]     = mod_ps["P"].reindex(tbl.index, fill_value=0.0)
            tbl["S Sum"]     = mod_ps["S"].reindex(tbl.index, fill_value=0.0)
            tbl["SA Sum"]    = mod_tipo["SA"].reindex(tbl.index, fill_value=0.0)
            tbl["PA Sum"]    = mod_tipo["PA"].reindex(tbl.index, fill_value=0.0)
            tbl["SP Sum"]    = mod_tipo["SP"].reindex(tbl.index, fill_value=0.0)
            tbl["IP Sum"]    = mod_tipo["IP"].reindex(tbl.index, fill_value=0.0)
            tbl["OTHER Sum"] = mod_tipo["OTHER"].reindex(tbl.index, fill_value=0.0)
            tbl["Credit Sum"]= tbl[["P Sum","S Sum"]].sum(axis=1)

        total_row = pd.DataFrame(tbl.sum(axis=0)).T
        total_row.index = ["TOTAL"]
        tbl_out = pd.concat([tbl, total_row], axis=0)

        display_label = sel_label if 'sel_label' in locals() else "Selected Period"
        with st.expander(f"Credit sums by {dim_label} — {display_label}", expanded=False):
            export_tbl = tbl_out.reset_index().rename(columns={"index": dim_label})
            _download_xlsx_button(export_tbl,
                                  f"credit_sums_{_slugify(dim_label)}_{_slugify(display_label)}.xlsx",
                                  key=f"dl_credit_sums_{_slugify(dim_label)}_{_slugify(display_label)}",
                                  label="⬇️ Descargar tabla (Excel)")
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

# --------------------------
# DETAIL TABLE + DONUT + SEARCH
# --------------------------
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
        metric_choice = st.session_state.get(cfg[view]["metric_key"], "%P")
        opt_val = st.session_state.get(key, "(All)")

        base = df_car_filt_all.copy()
        if "_AREA"  not in base.columns and col_areaCourse: base["_AREA"]  = base[col_areaCourse].astype(str).str.strip()
        if "_FIELD" not in base.columns and col_field:      base["_FIELD"] = base[col_field].astype(str).str.strip()
        if "_MAT"   not in base.columns and col_prog:       base["_MAT"]   = base[col_prog].astype(str).str.strip()
        if "_TIPO"  not in base.columns and col_tipoC:      base["_TIPO"]  = _norm_str(base[col_tipoC]).map(normalize_tipo)
        if "_PS"    not in base.columns and col_ps_C:       base["_PS"]    = _norm_str(base[col_ps_C]).map(normalize_ps)
        if "_CRED"  not in base.columns and col_cred:       base["_CRED"]  = pd.to_numeric(base[col_cred], errors="coerce").fillna(0.0)

        cL, cR = st.columns([7,5], gap="large")

        with cL:
            if metric_choice == "%P":
                table_filter = st.radio("", ["All", "Only P", "Only S"], index=0, horizontal=True, key=f"table_filt_ps_{view}_{opt_val}")
                base_tbl = base.copy()
                if opt_val not in {"(All)", "(TOTAL)"} and col_tag in base_tbl.columns:
                    base_tbl = base_tbl[base_tbl[col_tag] == opt_val].copy()
                if table_filter == "Only P": base_tbl = base_tbl[base_tbl["_PS"] == "P"]
                elif table_filter == "Only S": base_tbl = base_tbl[base_tbl["_PS"] == "S"]
            else:
                table_filter = st.radio("", ["All", "Only SA", "Only OTHER"], index=0, horizontal=True, key=f"table_filt_tipo_{view}_{opt_val}")
                base_tbl = base.copy()
                if opt_val not in {"(All)", "(TOTAL)"} and col_tag in base_tbl.columns:
                    base_tbl = base_tbl[base_tbl[col_tag] == opt_val].copy()
                if table_filter == "Only SA": base_tbl = base_tbl[base_tbl["_TIPO"] == "SA"]
                elif table_filter == "Only OTHER": base_tbl = base_tbl[base_tbl["_TIPO"] == "OTHER"]

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
                             if opt_val in {"(TOTAL)", "(All)"} else f"{n_courses} courses of {opt_val} were taught in {display_label}")
            else:
                if table_filter == "Only SA":
                    title = f"{n_courses} courses were taught in {display_label} by Scholarly Academics"
                elif table_filter == "Only OTHER":
                    title = f"{n_courses} courses were taught in {display_label} by Others"
                else:
                    title = (f"{n_courses} courses were taught in {display_label}"
                             if opt_val in {"(TOTAL)", "(All)"} else f"{n_courses} courses of {opt_val} were taught in {display_label}")

            st.markdown(f"### {title}")

            _download_xlsx_button(out, f"table_detail_{_slugify(opt_val)}_{_slugify(display_label)}.xlsx",
                                  key=f"dl_tbl_detail_{_slugify(opt_val)}_{_slugify(display_label)}",
                                  label="⬇️ Descargar tabla (Excel)")
            st.dataframe(out, use_container_width=True, hide_index=True)

        with cR:
            agg_tipo = (base.groupby([col_tag,"_TIPO"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0)) if col_tag in base.columns else pd.DataFrame()
            for k in ["SA","PA","SP","IP","OTHER"]:
                if k not in agg_tipo.columns: agg_tipo[k] = 0.0
            agg_ps = (base.groupby([col_tag,"_PS"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0)) if col_tag in base.columns else pd.DataFrame()
            for k in ["P","S"]:
                if k not in agg_ps.columns: agg_ps[k] = 0.0
            agg_ps = agg_ps[["P","S"]]; agg_tipo = agg_tipo[["SA","PA","SP","IP","OTHER"]]

            if SENS.get("on"):
                mod_ps, mod_tipo = apply_ops_to_aggs(agg_ps, agg_tipo, SENS.get("ops", []), member_all_label="All")
            else:
                mod_ps, mod_tipo = agg_ps, agg_tipo

            if opt_val in {"(TOTAL)", "(All)"} or col_tag not in base.columns:
                p_val, s_val = float(mod_ps["P"].sum() if not mod_ps.empty else 0.0), float(mod_ps["S"].sum() if not mod_ps.empty else 0.0)
                sa = float(mod_tipo["SA"].sum() if not mod_tipo.empty else 0.0)
                pa = float(mod_tipo["PA"].sum() if not mod_tipo.empty else 0.0)
                sp = float(mod_tipo["SP"].sum() if not mod_tipo.empty else 0.0)
                ip = float(mod_tipo["IP"].sum() if not mod_tipo.empty else 0.0)
                other = float(mod_tipo["OTHER"].sum() if not mod_tipo.empty else 0.0)
                title_suffix = "TOTAL"
            else:
                row_ps = mod_ps.loc[[opt_val]] if opt_val in mod_ps.index else pd.DataFrame(columns=["P","S"])
                row_q  = mod_tipo.loc[[opt_val]] if opt_val in mod_tipo.index else pd.DataFrame(columns=["SA","PA","SP","IP","OTHER"])
                p_val, s_val = float(row_ps["P"].sum() if not row_ps.empty else 0.0), float(row_ps["S"].sum() if not row_ps.empty else 0.0)
                sa = float(row_q["SA"].sum() if not row_q.empty else 0.0)
                pa = float(row_q["PA"].sum() if not row_q.empty else 0.0)
                sp = float(row_q["SP"].sum() if not row_q.empty else 0.0)
                ip = float(row_q["IP"].sum() if not row_q.empty else 0.0)
                other = float(row_q["OTHER"].sum() if not row_q.empty else 0.0)
                title_suffix = opt_val

            donut_h   = 360
            thrP = 75.0 if title_suffix == "TOTAL" else 60.0

            if metric_choice == "%P":
                den = p_val + s_val
                p_share = (p_val/den*100) if den else 0.0
                alert = (p_share < thrP)
                color_map = {"P": ("#F5A3A3" if alert else MINT), "S": "#B0B0B0"}
                fig = px.pie(names=["P","S"], values=[p_val, s_val], color=["P","S"], color_discrete_map=color_map, hole=0.55)
                fig.update_traces(textinfo="percent+label", hovertemplate="%{label}: %{percent:.1%}<extra></extra>")
                fig.update_layout(
                    title=f"% Participating Distribution — {title_suffix}",
                    height=donut_h, margin=dict(l=10, r=10, t=40, b=10),
                    legend=dict(orientation="v", yanchor="bottom", y=0.4, xanchor="center", x=0.9),
                )
                st.plotly_chart(fig, use_container_width=True)

                donut_df = pd.DataFrame({"Group": ["P","S"], "Credits": [p_val, s_val]})
                donut_df["Percent"] = (donut_df["Credits"] / max(1e-9, donut_df["Credits"].sum()))*100
                _download_xlsx_button(donut_df, f"chart_donut_PS_{_slugify(title_suffix)}_{_slugify(display_label)}.xlsx",
                                      key=f"dl_donut_ps_{_slugify(title_suffix)}_{_slugify(display_label)}", label="⬇️ Datos de la gráfica (Excel)")
            else:
                labels_all  = ["SA", "PA", "SP", "IP", "OTHER"]
                values_all  = [sa, pa, sp, ip, other]
                filtered    = [(l, v) for l, v in zip(labels_all, values_all) if v > 0]
                if filtered:
                    labels = [l for l, _ in filtered]
                    values = [v for _, v in filtered]
                    den = sum(values_all) or 1.0
                    sa_share    = sa/den*100
                    other_share = other/den*100
                    color_map = {}
                    for l in labels:
                        if l == "SA": color_map[l] = ("#F5A3A3" if sa_share < 40.0 else MINT)
                        elif l == "OTHER": color_map[l] = ("#F5A3A3" if other_share > 10.0 else "#6B7280")
                        else: color_map[l] = "#B0B0B0"

                    fig = px.pie(names=labels, values=values, color=labels, color_discrete_map=color_map, hole=0.55)
                    fig.update_traces(textinfo="percent+label", sort=False, hovertemplate="%{label}: %{percent:.1%}<extra></extra>")
                    title_txt = "%SA Distribution" if metric_choice == "%SA" else "%OTHER Distribution"
                    fig.update_layout(
                        title=f"{title_txt} — {title_suffix}",
                        height=donut_h, margin=dict(l=10, r=10, t=40, b=10),
                        legend=dict(orientation="v", yanchor="bottom", y=0.4, xanchor="center", x=0.9),
                    )
                    st.plotly_chart(fig, use_container_width=True)

                    donut_df = pd.DataFrame({"Type": labels_all, "Credits": values_all})
                    donut_df["Percent"] = (donut_df["Credits"] / max(1e-9, donut_df["Credits"].sum()))*100
                    _download_xlsx_button(donut_df, f"chart_donut_TIPO_{_slugify(title_suffix)}_{_slugify(display_label)}.xlsx",
                                          key=f"dl_donut_tipo_{_slugify(title_suffix)}_{_slugify(display_label)}", label="⬇️ Datos de la gráfica (Excel)")
                else:
                    st.caption("No hay registros de TIPO para esta métrica en este período.")
except Exception:
    pass

# --------------------------
# COUNTS — PIVOT
# --------------------------
st.markdown("---")
show_counts = st.checkbox("Show P/S counts", value=False)
if show_counts:
    st.subheader(f"Participating vs Supporting — {sel_label} (Counts & %)")

    if col_ps_fd:   df_fd_f["_PS"]   = _norm_str(df_fd_f[col_ps_fd]).map(normalize_ps)
    if col_area_fd: df_fd_f["_AREA"] = df_fd_f[col_area_fd].astype(str).str.strip()
    if col_tipo_fd: df_fd_f["_TIPO"] = _norm_str(df_fd_f[col_tipo_fd]).map(normalize_tipo)

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

    # Ajuste simple por sensitivity (solo TOTAL)
    if SENS["on"] and SENS["ops"]:
        add_P = sum(op.get("count",0) for op in SENS["ops"] if op.get("scope")=="PS" and op.get("cat")=="P")
        add_S = sum(op.get("count",0) for op in SENS["ops"] if op.get("scope")=="PS" and op.get("cat")=="S")
        table_totals_increase = {"Participating": int(add_P), "Supporting": int(add_S)}
    else:
        table_totals_increase = {"Participating": 0, "Supporting": 0}

    table["__Total__"] = table["Participating"] + table["Supporting"]

    df_counts = (
        table[["Participating", "Supporting"]].astype(int).reset_index()
    )
    total_row = pd.DataFrame([{
        row_name: "TOTAL",
        "Participating": int(df_counts["Participating"].sum()) + table_totals_increase["Participating"],
        "Supporting":    int(df_counts["Supporting"].sum())    + table_totals_increase["Supporting"],
    }])
    df_counts_out = pd.concat([df_counts, total_row], ignore_index=True)

    def _bold_total(df_):
        sty = pd.DataFrame('', index=df_.index, columns=df_.columns)
        mask = df_[row_name].astype(str).str.upper().eq("TOTAL")
        for c in df_.columns:
            sty.loc[mask, c] = 'font-weight:700;'
        return sty

    left, right = st.columns([6,6], gap="large")

    denom = table["__Total__"].replace(0, pd.NA)
    perc_df = pd.DataFrame({
        row_name: table.index,
        "%Participating": (table["Participating"] / denom * 100).round(1).fillna(0.0),
        "%Supporting":    (table["Supporting"]    / denom * 100).round(1).fillna(0.0),
    })
    if desired_order:
        for code in desired_order:
            if code not in perc_df[row_name].tolist():
                perc_df.loc[len(perc_df)] = [code, 0.0, 0.0]
        cat_order = desired_order
    else:
        cat_order = perc_df[row_name].tolist()

    chart_export = perc_df.melt(id_vars=row_name, value_vars=["%Participating", "%Supporting"],
                                var_name="Group", value_name="Percent")

    with left:
        _download_xlsx_button(df_counts_out, f"ps_counts_{_slugify(row_name)}_{_slugify(sel_label)}.xlsx",
                              key=f"dl_ps_counts_{_slugify(row_name)}_{_slugify(sel_label)}",
                              label="Descargar tabla (Excel)")
        styled_counts = (df_counts_out.style
                         .format({"Participating": "{:,.0f}", "Supporting": "{:,.0f}"})
                         .apply(_bold_total, axis=None))
        st.dataframe(styled_counts, use_container_width=True, hide_index=True)

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

        _download_xlsx_button(chart_export, f"chart_ps_perc_{_slugify(row_name)}_{_slugify(sel_label)}.xlsx",
                              key=f"dl_chart_ps_perc_{_slugify(row_name)}_{_slugify(sel_label)}",
                              label="Descargar datos (Excel)")

