# =============================================================================
#  4 · Faculty Demographics  — UASM Faculty Analytics Suite
#  Refactored: design, UX, structure.  Business logic UNCHANGED.
# =============================================================================
import re
import datetime
import numpy as np
import streamlit as st
import pandas as pd
import plotly.express as px

from suite_styles import (
    apply_page_config, inject_global_css, render_header,
    render_sidebar_nav, render_update_banner, kpi_card_row, section_divider,
    apply_chart_style, add_period_highlight,
    PALETTE, _download_link,
)

# ── Page bootstrap ─────────────────────────────────────────────────────────────
apply_page_config("Faculty Demographics · UASM")
inject_global_css()

# ── Data Load (UNCHANGED logic) ────────────────────────────────────────────────
@st.cache_data(ttl=0)
def load_fulltime():
    df = pd.read_excel("data/Faculty/BD_Faculty.xlsx", sheet_name="BD PLANTA 2020-2025")
    if "Semestre" in df.columns:
        sem      = df["Semestre"].astype(str).str.strip()
        is_inter = sem.str.contains("inter", case=False, na=False)
        df["Periodo"] = np.where(is_inter, sem.str[:4]+" Intersemestral", sem.str[:4]+sem.str[-2:])
    else:
        raw      = df.iloc[:,0].astype(str).str.strip()
        is_inter = raw.str.contains("inter", case=False, na=False)
        df["Periodo"] = np.where(is_inter, raw.str[:4]+" Intersemestral", raw.str.slice(0,4)+raw.str.slice(4,6))
    if "Academic Area" in df.columns and "AREA_PROFESOR" not in df.columns:
        df["AREA_PROFESOR"] = df["Academic Area"]
    if "ID Nr." in df.columns and "ID" not in df.columns:
        df = df.rename(columns={"ID Nr.": "ID"})
    if "Full Name" not in df.columns:
        fn = df.get("First Name",""); ln = df.get("Last Name","")
        df["Full Name"] = (pd.Series(fn).astype(str).fillna("")+" "+pd.Series(ln).astype(str).fillna("")).str.strip()
    return df, datetime.datetime.now()

@st.cache_data(ttl=0)
def load_parttime():
    df = pd.read_excel("data/Faculty/BD_Faculty.xlsx", sheet_name="Faculty Distribution")
    if "PLANTA_CATEDRA" in df.columns:
        col = df["PLANTA_CATEDRA"].astype(str).str.strip()
        col = col.str.normalize("NFKD").str.encode("ascii",errors="ignore").str.decode("ascii")
        df  = df[col.str.upper().eq("CATEDRA")].copy()
    sem      = df["Semestre"].astype(str).str.strip()
    is_inter = sem.str.contains("inter", case=False, na=False)
    df.loc[~is_inter,"Periodo"] = sem.str[:4]+sem.str[-2:]
    df.loc[is_inter, "Periodo"] = sem.str[:4]+" Intersemestral"
    if "ID Nr." not in df.columns and "ID" in df.columns:
        df = df.rename(columns={"ID":"ID Nr."})
    if "AREA_PROFESOR" not in df.columns and "Academic Area" in df.columns:
        df["AREA_PROFESOR"] = df["Academic Area"]
    return df

_ft_result = load_fulltime()
df_full, _load_ts = _ft_result if isinstance(_ft_result, tuple) else (_ft_result, None)
df_part = load_parttime()

# ── Column helpers (UNCHANGED) ─────────────────────────────────────────────────
def col_id(df_):
    return "ID Nr." if "ID Nr." in df_.columns else ("ID" if "ID" in df_.columns else None)
def col_degree(df_):
    for c in ["Highest Degree","TÍTULO"]:
        if c in df_.columns: return c
    return None
def col_gender(df_):
    for c in ["Gender","GÉNERO"]:
        if c in df_.columns: return c
    return None
def col_nationality(df_):
    for c in ["Country of Birth","Nationality"]:
        if c in df_.columns: return c
    return None
def normalize_degree(series: pd.Series) -> pd.Series:
    s       = series.astype(str).str.strip()
    is_tbd  = s.str.upper().eq("TBD")|s.eq("")|s.str.lower().eq("na")|s.str.lower().eq("none")
    s_norm  = s.str.lower().str.replace(".",  "", regex=False)
    s_norm  = s_norm.str.normalize("NFKD").str.encode("ascii",errors="ignore").str.decode("ascii")
    is_phd  = s_norm.str.contains(r"\bphd\b")|s_norm.str.contains("doctor")
    is_mast = s_norm.str.contains("master")|s_norm.str.contains(r"\bmsc\b")|s_norm.str.contains(r"\bms\b")
    is_bach = (s_norm.str.contains("bachelor")|s_norm.str.contains(r"\bbsc\b")|
               s_norm.str.contains(r"\bbs\b")|s_norm.str.contains(r"\bba\b")|s_norm.str.contains("licen"))
    out = pd.Series("Other", index=s.index, dtype=object)
    out[is_tbd]              = "TBD"
    out[~is_tbd & is_phd]   = "PhD"
    out[~is_tbd & ~is_phd & is_mast] = "Master"
    out[~is_tbd & ~is_phd & ~is_mast & is_bach] = "Bachelor"
    return out

# ── Period helpers (UNCHANGED) ─────────────────────────────────────────────────
def _years_all():
    y = []
    for d in [df_full, df_part]:
        for col in ["Periodo","Semestre"]:
            if col in d.columns:
                y.append(d[col].dropna().astype(str).str[:4])
    if not y: return []
    ys = pd.concat(y, ignore_index=True)
    ys = ys[ys.str.match(r"^\d{4}$", na=False)]
    return sorted(ys.astype(int).unique().tolist())

def _col_id(df_): return "ID Nr." if "ID Nr." in df_.columns else ("ID" if "ID" in df_.columns else None)

def _filter_for_timeframe(df_in, time_mode, sel_sem=None, sel_year=None, sel_inter_label=None):
    dfb  = df_in.copy()
    pcol = "Periodo" if "Periodo" in dfb.columns else None
    scol = "Semestre" if "Semestre" in dfb.columns else None
    idc  = _col_id(dfb)
    if time_mode == "Semestral" and sel_sem:
        target = str(sel_sem)
        mask   = pd.Series(False, index=dfb.index)
        if pcol: mask |= dfb[pcol].astype(str).eq(target)
        if scol: mask |= dfb[scol].astype(str).str.replace("-","",regex=False).str.fullmatch(target, na=False)
        dfb    = dfb[mask].copy()
    elif time_mode == "Anual" and sel_year is not None:
        y = str(sel_year)
        mask = pd.Series(False, index=dfb.index)
        if pcol: mask |= dfb[pcol].astype(str).str.startswith(y)
        if scol: mask |= dfb[scol].astype(str).str.startswith(y)
        dfb = dfb[mask].copy()
        if idc:
            sort_key = pcol if pcol else (scol if scol else None)
            if sort_key: dfb = dfb.sort_values(by=[sort_key])
            dfb = dfb.drop_duplicates(subset=[idc], keep="last")
    elif time_mode == "Intersemestral" and sel_inter_label:
        y = sel_inter_label.split()[0]
        mask = pd.Series(False, index=dfb.index)
        if scol:
            sn = dfb[scol].astype(str)
            mask |= sn.str.contains("inter",case=False,na=False)&sn.str.contains(y,na=False)
        if pcol: mask |= dfb[pcol].astype(str).eq(sel_inter_label)
        dfb = dfb[mask].copy()
    return dfb

def _is_semester_label(p): return bool(re.fullmatch(r"\d{4}(10|20)", str(p)))
def _is_inter_label(p):    return bool(re.fullmatch(r"\d{4}\s+Intersemestral", str(p)))

def _options_for_timeframe(df_src, time_mode):
    per = df_src["Periodo"].dropna().astype(str)
    if time_mode == "Semestral":       return sorted([p for p in per.unique() if _is_semester_label(p)])
    if time_mode == "Intersemestral":  return sorted([p for p in per.unique() if _is_inter_label(p)])
    return sorted(per.str[:4].unique().tolist())

def build_time_series(df_src, time_mode, idcol, degree_col, nat_col):
    labels, phd_pct, intl_pct = [], [], []

    def _phd(sub):
        if sub.empty or degree_col is None: return 0.0
        sub = sub.copy()
        s      = sub[degree_col].astype(str).str.strip()
        is_tbd = s.str.upper().eq("TBD")|s.eq("")|s.str.lower().eq("na")|s.str.lower().eq("none")
        s_norm = s.str.lower().str.replace(".",  "",regex=False)
        s_norm = s_norm.str.normalize("NFKD").str.encode("ascii",errors="ignore").str.decode("ascii")
        is_phd = s_norm.str.contains(r"\bphd\b")|s_norm.str.contains("doctor")
        sub["__deg"] = np.where(is_tbd,"TBD",np.where(is_phd,"PhD","Other"))
        tot = sub[idcol].nunique()
        return 0.0 if tot==0 else round(100*sub.loc[sub["__deg"]=="PhD",idcol].nunique()/tot,1)

    def _intl(sub):
        if sub.empty or nat_col is None: return 0.0
        nat = sub[nat_col].astype(str).str.strip()
        is_col,is_tbd,is_empty = nat.eq("Colombian"),nat.str.upper().eq("TBD"),nat.eq("")
        tot = sub[idcol].nunique()
        return 0.0 if tot==0 else round(100*sub.loc[(~is_col)&(~is_tbd)&(~is_empty),idcol].nunique()/tot,1)

    if time_mode == "Semestral":
        periods = sorted([p for p in df_src["Periodo"].dropna().astype(str).unique() if _is_semester_label(p)])
        for p in periods:
            sub = df_src[df_src["Periodo"].astype(str).eq(p)]
            labels.append(p); phd_pct.append(_phd(sub)); intl_pct.append(_intl(sub))
    elif time_mode == "Intersemestral":
        periods = sorted([p for p in df_src["Periodo"].dropna().astype(str).unique() if _is_inter_label(p)])
        for p in periods:
            sub = df_src[df_src["Periodo"].astype(str).eq(p)]
            labels.append(p); phd_pct.append(_phd(sub)); intl_pct.append(_intl(sub))
    else:
        years = sorted(df_src["Periodo"].dropna().astype(str).str[:4].unique())
        for y in years:
            sub = _filter_for_timeframe(df_src,"Anual",sel_year=int(y))
            labels.append(y); phd_pct.append(_phd(sub)); intl_pct.append(_intl(sub))
    return labels, phd_pct, intl_pct

# ── Session state ──────────────────────────────────────────────────────────────
if "modo_faculty"   not in st.session_state: st.session_state.modo_faculty   = "Full-time"
if "time_mode_side" not in st.session_state: st.session_state.time_mode_side = "Semestral"
if "sel_tf_label"   not in st.session_state: st.session_state.sel_tf_label   = None

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    render_sidebar_nav(current_key="4 · Faculty Demographics")

    st.markdown("#### 👥 Faculty Type")
    st.markdown('<div id="mode-pill">', unsafe_allow_html=True)
    mode_sidebar = st.radio(
        "Mode",["Full-time","Part-time"],
        index=0 if st.session_state.modo_faculty=="Full-time" else 1,
        horizontal=True, label_visibility="collapsed", key="demo_mode_pill",
    )
    st.markdown("</div>", unsafe_allow_html=True)

if mode_sidebar != st.session_state.modo_faculty:
    st.session_state.modo_faculty = mode_sidebar
    st.session_state.sel_tf_label = None
    st.rerun()

mode_now = st.session_state.modo_faculty
df       = df_full if mode_now=="Full-time" else df_part

with st.sidebar:
    st.markdown("#### ⏱ Timeframe")
    time_mode_side = st.radio(
        "Timeframe",["Semestral","Anual","Intersemestral"],
        key="time_mode_side", label_visibility="collapsed",
    )
    df_base    = df_full if st.session_state.get("modo_faculty","Full-time")=="Full-time" else df_part
    options_tf = _options_for_timeframe(df_base, time_mode_side)
    default_opt = options_tf[-1] if options_tf else None

    sel_label = st.selectbox(
        "Period", options_tf,
        index=(options_tf.index(st.session_state.sel_tf_label)
               if st.session_state.sel_tf_label in options_tf
               else (len(options_tf)-1 if options_tf else 0)),
        help="Select the period to filter charts and tables.",
    ) if options_tf else None

    if sel_label != st.session_state.get("sel_tf_label"):
        st.session_state.sel_tf_label = sel_label

    if sel_label:
        if time_mode_side == "Semestral":
            export_df  = _filter_for_timeframe(df_base,"Semestral",sel_sem=sel_label)
            label_time = sel_label
        elif time_mode_side == "Anual":
            export_df  = _filter_for_timeframe(df_base,"Anual",sel_year=int(sel_label))
            label_time = sel_label
        else:
            export_df  = _filter_for_timeframe(df_base,"Intersemestral",sel_inter_label=sel_label)
            label_time = sel_label
        st.markdown("<hr style='margin:10px 0;opacity:.4'>", unsafe_allow_html=True)
        _download_link(f"⬇ Download data {label_time} (Excel)", export_df, f"Export_{label_time}.xlsx")

# ── Time series data ──────────────────────────────────────────────────────────
IDCOL       = col_id(df)
tmode_ts    = st.session_state.get("time_mode_side","Semestral")
labels_ts, phd_ts, intl_ts = build_time_series(
    df, tmode_ts, IDCOL, col_degree(df), col_nationality(df)
)
period_current = st.session_state.get("sel_tf_label")

# ── Header ─────────────────────────────────────────────────────────────────────
render_header(
    "Faculty Demographics",
    subtitle="PhD attainment, international diversity, and faculty composition over time",
)
render_update_banner(last_updated=_load_ts, is_fresh=True)

# ── KPI Row ────────────────────────────────────────────────────────────────────
sel_period_text = st.session_state.get("sel_tf_label") or ""
if sel_period_text and mode_now == "Full-time":
    if tmode_ts == "Semestral":
        _active_kpi = _filter_for_timeframe(df,"Semestral",sel_sem=sel_period_text)
    elif tmode_ts == "Anual":
        _active_kpi = _filter_for_timeframe(df,"Anual",sel_year=int(sel_period_text))
    else:
        _active_kpi = _filter_for_timeframe(df,"Intersemestral",sel_inter_label=sel_period_text)
    _total_kpi = int(_active_kpi[IDCOL].nunique()) if (not _active_kpi.empty and IDCOL) else 0
    _phd_now   = (round(100*_active_kpi.loc[normalize_degree(_active_kpi[col_degree(_active_kpi)]).eq("PhD"),IDCOL].nunique()/_total_kpi,1)
                  if (_total_kpi and col_degree(_active_kpi)) else 0)
else:
    _total_kpi = 0; _phd_now = 0

kpi_card_row([
    {"value": sel_period_text or "—",  "label": "Selected Period"},
    {"value": mode_now,                "label": "Faculty Type"},
    {"value": _total_kpi,              "label": "Faculty Count"},
    {"value": f"{_phd_now}%",          "label": "PhD Holders",
     "up": _phd_now >= 40 if mode_now=="Full-time" else None},
])

section_divider()

# ── Demographics Table ─────────────────────────────────────────────────────────
st.subheader(
    "Full-time Demographics by Faculty Ranking"
    if mode_now=="Full-time" else "Part-time Demographic Table"
)
if sel_period_text and mode_now=="Full-time":
    st.markdown(f"<div class='period-label'>{sel_period_text}</div>", unsafe_allow_html=True)

col_table, col_side = st.columns([3, 1.2])

with col_table:
    if not IDCOL:
        st.error("ID column not found."); st.stop()

    tmode    = st.session_state.get("time_mode_side","Semestral")
    sel_lbl  = st.session_state.get("sel_tf_label")

    if not sel_lbl and mode_now=="Full-time":
        st.info("No data for the selected mode/period.")
    else:
        if mode_now=="Full-time":
            if tmode=="Semestral":        active = _filter_for_timeframe(df,"Semestral",sel_sem=sel_lbl)
            elif tmode=="Anual":          active = _filter_for_timeframe(df,"Anual",sel_year=int(sel_lbl))
            else:                         active = _filter_for_timeframe(df,"Intersemestral",sel_inter_label=sel_lbl)

            if "Faculty Ranking" in active.columns:
                base_order = ["Full Professor","Associate Professor","Assistant Professor","Instructor"]
                uniq = active["Faculty Ranking"].dropna().unique().tolist()
                ranking_order = [x for x in base_order if x in uniq]+[x for x in uniq if x not in base_order]
            else:
                ranking_order = []

            cols_out = ["Category"]+ranking_order+["Total"]
            groups   = {"Highest Degree","Nationality","Gender","Age"}

            def counts_by_ranking(df_sub):
                if ranking_order:
                    s = (df_sub.groupby("Faculty Ranking")[IDCOL].nunique()
                         .reindex(ranking_order,fill_value=0).astype(int))
                else:
                    s = pd.Series(dtype=int)
                s.loc["Total"] = int(s.sum()) if not s.empty else int(df_sub[IDCOL].nunique())
                return s

            rows = []
            dcol = col_degree(active)
            if dcol and not active.empty:
                active["Degree_norm"] = normalize_degree(active[dcol])
                rows.append(pd.Series({"Category":"Highest Degree",**counts_by_ranking(active[active["Degree_norm"].isin(["PhD","Master","Bachelor"])]).to_dict()}))
                for d in ["PhD","Master","Bachelor"]:
                    rows.append(pd.Series({"Category":d,**counts_by_ranking(active[active["Degree_norm"]==d]).to_dict()}))

            ncol = col_nationality(active)
            if ncol and not active.empty:
                rows.append(pd.Series({"Category":"Nationality",**counts_by_ranking(active).to_dict()}))
                rows.append(pd.Series({"Category":"Colombian",**counts_by_ranking(active[active[ncol].astype(str).eq("Colombian")]).to_dict()}))
                rows.append(pd.Series({"Category":"International",**counts_by_ranking(active[~active[ncol].astype(str).eq("Colombian")]).to_dict()}))

            gcol = col_gender(active)
            if gcol and not active.empty:
                rows.append(pd.Series({"Category":"Gender",**counts_by_ranking(active[active[gcol].astype(str).isin(["Male","Female"])]).to_dict()}))
                for g in ["Male","Female"]:
                    rows.append(pd.Series({"Category":g,**counts_by_ranking(active[active[gcol].astype(str)==g]).to_dict()}))

            if not active.empty and "Age" in active.columns:
                age = pd.to_numeric(active["Age"],errors="coerce")
                active["Age_bucket"] = pd.cut(age,bins=[-np.inf,29,40,50,60,np.inf],
                                              labels=["Under 30","31-40","41-50","51-60","over 61"])
                rows.append(pd.Series({"Category":"Age",**counts_by_ranking(active[active["Age_bucket"].notna()]).to_dict()}))
                for b in ["Under 30","31-40","41-50","51-60","over 61"]:
                    rows.append(pd.Series({"Category":b,**counts_by_ranking(active[active["Age_bucket"]==b]).to_dict()}))

            table_df = pd.DataFrame(rows).reindex(columns=cols_out).fillna(0)
            if not table_df.empty:
                numeric_cols = [c for c in (ranking_order+["Total"]) if c in table_df.columns]
                is_group  = table_df["Category"].isin(groups)
                all_zero  = (table_df[numeric_cols].sum(axis=1)==0) if numeric_cols else pd.Series(False,index=table_df.index)
                table_df  = table_df.loc[~(all_zero & ~is_group)].copy()
                for c in numeric_cols:
                    table_df[c] = pd.to_numeric(table_df[c],errors="coerce").fillna(0).astype(int)

            mint_dark  = PALETTE["primary_dark"]; mint_light = PALETTE["primary_light"]

            def _style_group_rows(df_):
                styles = pd.DataFrame('',index=df_.index,columns=df_.columns)
                mask   = df_["Category"].isin(groups)
                styles.loc[mask,["Category"]+ranking_order+["Total"]] = 'background-color:#f2f2f2;'
                styles.loc[mask,"Category"] += 'font-weight:700;'
                return styles
            def _style_total_col_bold(df_):
                styles = pd.DataFrame('',index=df_.index,columns=df_.columns)
                styles.loc[:,"Total"] = 'font-weight:700;'
                return styles
            def _style_group_totals_mint(df_):
                styles = pd.DataFrame('',index=df_.index,columns=df_.columns)
                mask   = df_["Category"].isin(groups)
                styles.loc[mask,"Total"] = f'background-color:{mint_light};color:{mint_dark};font-weight:800;'
                return styles

            if table_df.empty:
                st.info("No rows to display for this selection.")
            else:
                styled_table = (
                    table_df.style
                    .apply(_style_group_rows,      axis=None)
                    .apply(_style_total_col_bold,  axis=None)
                    .apply(_style_group_totals_mint, axis=None)
                    .format(precision=0, na_rep="")
                    .hide(axis="index")
                )
                st.dataframe(styled_table, use_container_width=True)

        else:  # Part-time — show all periods grouped
            pt_id = col_id(df_part)
            if pt_id:
                pt_sum = (
                    df_part.groupby("Periodo")[pt_id].nunique()
                    .reset_index().rename(columns={pt_id:"Faculty Count"})
                    .sort_values("Periodo")
                )
                st.dataframe(pt_sum, use_container_width=True)

# ── Small charts (right column) ───────────────────────────────────────────────
with col_side:
    if mode_now == "Full-time" and period_current:
        _check_active = _filter_for_timeframe(df,"Anual",sel_year=int(period_current)) if tmode_ts=="Anual" else df[df["Periodo"].astype(str).eq(str(period_current))].copy()
        _dcol = col_degree(_check_active) if not _check_active.empty else None
        if _dcol and not _check_active.empty:
            _check_active["Degree_norm"] = normalize_degree(_check_active[_dcol])
            _deg_counts = _check_active.groupby("Degree_norm")[IDCOL].nunique().reset_index()
            _deg_counts.columns = ["Degree","Count"]
            fig_deg = px.pie(_deg_counts,names="Degree",values="Count",hole=0.45,
                             title="Degree Distribution",
                             color_discrete_sequence=[PALETTE["primary"],PALETTE["accent1"],PALETTE["accent2"],PALETTE["neutral"]])
            apply_chart_style(fig_deg, height=260)
            st.plotly_chart(fig_deg, use_container_width=True)

section_divider()

# ── Row 1: PhD % over time (left) + PhD regions (right) ───────────────────────
st.subheader("📊 PhD Attainment Over Time")
if mode_now == "Part-time":
    y_min_phd,y_max_phd = 0,10; line_h,bar_h = 260,220
else:
    y_min_phd,y_max_phd = 0,100; line_h,bar_h = 350,300
if tmode_ts == "Intersemestral":
    y_min_phd,y_max_phd = 0,100

row1_left, row1_right = st.columns([6, 4])

with row1_left:
    df_pct_phd = pd.DataFrame({"Label":labels_ts,"Percent":phd_ts})
    title_phd  = "% Full-time Faculty with PhD" if mode_now=="Full-time" else "% Part-time Faculty with PhD"
    fig_phd    = px.line(df_pct_phd,x="Label",y="Percent",markers=True,text="Percent",title=title_phd)
    fig_phd.update_traces(line=dict(color=PALETTE["primary"],width=3),
                          marker=dict(size=7,color=PALETTE["primary"]),
                          texttemplate="%{y:.1f}%",textposition="top center")
    fig_phd.update_xaxes(type="category",categoryorder="array",categoryarray=labels_ts,tickangle=0,title=None)
    fig_phd.update_yaxes(range=[y_min_phd,y_max_phd],title=None)
    add_period_highlight(fig_phd,period_current,labels_ts)
    apply_chart_style(fig_phd,height=line_h,show_legend=False)
    st.plotly_chart(fig_phd, use_container_width=True)

with row1_right:
    if period_current is None:
        active_p = df.iloc[0:0].copy()
    else:
        active_p = (_filter_for_timeframe(df,"Anual",sel_year=int(period_current))
                    if tmode_ts=="Anual"
                    else df[df["Periodo"].astype(str).eq(str(period_current))].copy())

    dcol_here = col_degree(df)
    if dcol_here is not None and not active_p.empty:
        active_p["Degree_norm"] = normalize_degree(active_p[dcol_here])
        phd_now_all = active_p[active_p["Degree_norm"].eq("PhD")].copy()
        region_col  = "Region were degree was obtained" if "Region were degree was obtained" in phd_now_all.columns else None
        if region_col:
            reg   = phd_now_all[region_col].astype(str).str.strip()
            mask_v = (~reg.eq(""))&(~reg.str.upper().eq("TBD"))
            phd_for_regions = phd_now_all[mask_v].copy()
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
                      .nunique().sort_values(ascending=False).reset_index()
                      .rename(columns={"Region were degree was obtained":"Region",IDCOL:"Count"}))
    else:
        reg_counts = pd.DataFrame({"Region":[],"Count":[]})

    title_phd_bar = (f"{total_phd_valid} professors with PhD · {phd_int} obtained it abroad"
                     if phd_int else f"{total_phd_valid} professors with PhD")
    fig_phd_reg = px.bar(reg_counts,x="Count",y="Region",orientation="h",title=title_phd_bar,text="Count")
    fig_phd_reg.update_traces(marker_color=PALETTE["accent2"],textposition="outside",texttemplate="%{text}")
    fig_phd_reg.update_xaxes(title=None,dtick=1)
    fig_phd_reg.update_yaxes(title=None,autorange="reversed")
    apply_chart_style(fig_phd_reg,height=bar_h)
    st.plotly_chart(fig_phd_reg, use_container_width=True)

    # Detail popover
    cols_posibles = {
        "Full Name":                     ["Full Name","Full-Name","Full_Name","Profesor","First Name"],
        "Highest Earned Degree":         ["Highest Earned Degree","Highest Degree","TÍTULO"],
        "University":                    ["University","University Name"],
        "Region were degree was obtained":["Region were degree was obtained","Region"],
        "Year":                          ["Year","Year Earned ","Year Degree","Year Earned"],
    }
    def pick_cols(df_, mapping):
        out = {}
        for new,opts in mapping.items():
            for c in opts:
                if c in df_.columns: out[new]=df_[c]; break
            if new not in out: out[new]=pd.Series([""]*len(df_),index=df_.index)
        return pd.DataFrame(out)

    detalle_phd = pick_cols(phd_for_regions,cols_posibles) if not phd_for_regions.empty else pd.DataFrame(columns=list(cols_posibles.keys()))
    if not detalle_phd.empty:
        try:   pop = st.popover("🔎 PhD faculty detail")
        except AttributeError: pop = st.expander("🔎 PhD faculty detail")
        with pop:
            st.dataframe(detalle_phd.reset_index(drop=True), use_container_width=True)

section_divider()

# ── Row 2: International % over time (left) + Nationalities (right) ────────────
st.subheader("🌐 International Faculty Over Time")
if mode_now == "Part-time":
    y_min_int,y_max_int = 0,10; line_h2,bar_h2 = 260,220
else:
    y_min_int,y_max_int = 0,40; line_h2,bar_h2 = 350,300
if tmode_ts == "Intersemestral":
    y_min_int,y_max_int = 0,100

row2_left, row2_right = st.columns([6, 4])

with row2_left:
    df_pct_int = pd.DataFrame({"Label":labels_ts,"Percent":intl_ts})
    title_int  = "% International Full-time Faculty" if mode_now=="Full-time" else "% International Part-time Faculty"
    fig_int    = px.line(df_pct_int,x="Label",y="Percent",markers=True,text="Percent",title=title_int)
    fig_int.update_traces(line=dict(color=PALETTE["accent1"],width=3),
                          marker=dict(size=7,color=PALETTE["accent1"]),
                          texttemplate="%{y:.1f}%",textposition="top center")
    fig_int.update_xaxes(type="category",categoryorder="array",categoryarray=labels_ts,tickangle=0,title=None)
    fig_int.update_yaxes(range=[y_min_int,y_max_int],title=None)
    add_period_highlight(fig_int,period_current,labels_ts)
    apply_chart_style(fig_int,height=line_h2,show_legend=False)
    st.plotly_chart(fig_int, use_container_width=True)

with row2_right:
    nat_col = col_nationality(df)
    if nat_col and period_current:
        active_p2 = (_filter_for_timeframe(df,"Anual",sel_year=int(period_current))
                     if tmode_ts=="Anual"
                     else df[df["Periodo"].astype(str).eq(str(period_current))].copy())
        nat      = active_p2[nat_col].astype(str).str.strip()
        is_col   = nat.eq("Colombian"); is_tbd = nat.str.upper().eq("TBD"); is_empty = nat.eq("")
        intl_now = active_p2[(~is_col)&(~is_tbd)&(~is_empty)].copy()
        nat_counts = (intl_now.groupby(nat_col)[IDCOL].nunique().sort_values(ascending=False)
                      .reset_index().rename(columns={nat_col:"Nationality",IDCOL:"Count"}))
        total_intl = int(intl_now[IDCOL].nunique()) if not intl_now.empty else 0
        n_nats     = int(nat_counts["Nationality"].nunique()) if not nat_counts.empty else 0
    else:
        nat_counts = pd.DataFrame({"Nationality":[],"Count":[]}); total_intl=0; n_nats=0

    title_nat = f"{total_intl} international faculty · {n_nats} nationalities"
    fig_nat = px.bar(nat_counts,x="Count",y="Nationality",orientation="h",title=title_nat,text="Count")
    fig_nat.update_traces(marker_color=PALETTE["accent1"],textposition="outside",texttemplate="%{text}")
    fig_nat.update_xaxes(title=None,dtick=1)
    fig_nat.update_yaxes(title=None,autorange="reversed")
    apply_chart_style(fig_nat,height=bar_h2)
    st.plotly_chart(fig_nat, use_container_width=True)

    if not intl_now.empty:
        cols_nat = {
            "Full Name":  ["Full Name","Full-Name","Full_Name","Profesor","First Name"],
            "Nationality":["Nationality","Country of Birth"],
        }
        def pick_cols2(df_, mapping):
            out = {}
            for new,opts in mapping.items():
                for c in opts:
                    if c in df_.columns: out[new]=df_[c]; break
                if new not in out: out[new]=pd.Series([""]*len(df_),index=df_.index)
            return pd.DataFrame(out)
        detalle_nat = pick_cols2(intl_now, cols_nat)
        try:   pop2 = st.popover("🔎 International faculty detail")
        except AttributeError: pop2 = st.expander("🔎 International faculty detail")
        with pop2:
            st.dataframe(detalle_nat.reset_index(drop=True), use_container_width=True)
