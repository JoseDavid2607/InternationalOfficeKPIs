# =============================================================================
#  3 · Distribution by Academic Area  — UASM Faculty Analytics Suite
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
    PALETTE, RANKING_PALETTE, _download_link,
)

# ── Page bootstrap ─────────────────────────────────────────────────────────────
apply_page_config("Distribution by Academic Area · UASM")
inject_global_css()

# ── Helpers (UNCHANGED logic) ─────────────────────────────────────────────────
def _is_sem_label(p: str) -> bool:
    return bool(re.fullmatch(r"\d{4}-(10|20)", str(p)))

def _is_inter_label(p: str) -> bool:
    return bool(re.fullmatch(r"\d{4}\s+Intersemestral", str(p)))

def _display_label_sem(p_internal: str) -> str:
    return str(p_internal).replace("-", "")

def _filter_for_timeframe(df_in: pd.DataFrame, time_mode: str, value):
    if value is None:
        return df_in.iloc[0:0].copy()
    dfb  = df_in.copy()
    pcol = "Periodo"
    if time_mode == "Semestral":
        sem_internal = f"{str(value)[:4]}-{str(value)[-2:]}"
        return dfb[dfb[pcol].astype(str).eq(sem_internal)].copy()
    if time_mode == "Anual":
        y = str(value)
        dfy = dfb[dfb[pcol].astype(str).str.startswith(y)].copy()
        if "ID" in dfy.columns:
            dfy["__Year"] = dfy[pcol].astype(str).str[:4]
            dfy = dfy.sort_values(by=[pcol]).drop_duplicates(subset=["ID","__Year"], keep="last")
            dfy = dfy.drop(columns=["__Year"])
        return dfy
    return dfb[dfb[pcol].astype(str).eq(str(value))].copy()

# ── Data Load (UNCHANGED logic) ────────────────────────────────────────────────
@st.cache_data(ttl=0)
def load_fulltime():
    df = pd.read_excel("data/Faculty/BD_Faculty.xlsx", sheet_name="BD PLANTA 2020-2025")
    if "Semestre" in df.columns:
        sem      = df["Semestre"].astype(str).str.strip()
        is_inter = sem.str.contains("inter", case=False, na=False)
        df["Periodo"] = np.where(is_inter, sem.str[:4]+" Intersemestral", sem.str[:4]+"-"+sem.str[-2:])
    else:
        raw      = df.iloc[:,0].astype(str).str.strip()
        is_inter = raw.str.contains("inter", case=False, na=False)
        df["Periodo"] = np.where(is_inter, raw.str[:4]+" Intersemestral", raw.str.slice(0,4)+"-"+raw.str.slice(4,6))
    if "Academic Area" in df.columns and "AREA_PROFESOR" not in df.columns:
        df["AREA_PROFESOR"] = df["Academic Area"]
    if "ID Nr." in df.columns and "ID" not in df.columns:
        df = df.rename(columns={"ID Nr.": "ID"})
    if "Full Name" not in df.columns:
        fn = df["First Name"] if "First Name" in df.columns else ""
        ln = df["Last Name"]  if "Last Name"  in df.columns else ""
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
    df.loc[~is_inter, "Periodo"] = sem.str[:4]+"-"+sem.str[-2:]
    df.loc[is_inter,  "Periodo"] = sem.str[:4]+" Intersemestral"
    if "ID Nr." in df.columns and "ID" not in df.columns:
        df = df.rename(columns={"ID Nr.": "ID"})
    if "AREA_PROFESOR" not in df.columns and "Academic Area" in df.columns:
        df["AREA_PROFESOR"] = df["Academic Area"]
    return df

_ft_result = load_fulltime()
df_full, _load_ts = _ft_result if isinstance(_ft_result, tuple) else (_ft_result, None)
df_part = load_parttime()

# ── Session state ─────────────────────────────────────────────────────────────
if "modo_faculty" not in st.session_state:
    st.session_state.modo_faculty = "Full-time"

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    render_sidebar_nav(current_key="3 · Distribution by Academic Area")

    st.markdown("#### 👥 Faculty Type")
    st.markdown('<div id="mode-pill">', unsafe_allow_html=True)
    mode_sidebar = st.radio(
        "Mode", ["Full-time", "Part-time"],
        index=0 if st.session_state.modo_faculty=="Full-time" else 1,
        horizontal=True, label_visibility="collapsed", key="area_mode_pill",
    )
    st.markdown("</div>", unsafe_allow_html=True)

if mode_sidebar != st.session_state.modo_faculty:
    st.session_state.modo_faculty = mode_sidebar
    st.rerun()

df = df_full if st.session_state.modo_faculty == "Full-time" else df_part

# ── Period selectors (UNCHANGED logic) ────────────────────────────────────────
all_periods   = df["Periodo"].dropna().astype(str).unique().tolist()
sem_labels    = sorted([p for p in all_periods if _is_sem_label(p)])
inter_labels  = sorted([p for p in all_periods if _is_inter_label(p)])
year_labels   = sorted(pd.Series(all_periods).str[:4].unique().tolist())

with st.sidebar:
    st.markdown("#### ⏱ Timeframe")
    tmode_now = st.radio("Timeframe", ["Semestral","Anual","Intersemestral"],
                         key="area_tmode", label_visibility="collapsed")

    if tmode_now == "Semestral":
        vis_sem  = [_display_label_sem(p) for p in sem_labels]
        idx      = len(vis_sem)-1 if vis_sem else 0
        sel_vis  = st.selectbox("Period", vis_sem, index=idx if vis_sem else None)
        sel_label  = sel_vis
        sel_value  = sel_vis
        internal   = f"{sel_vis[:4]}-{sel_vis[-2:]}" if sel_vis else None
    elif tmode_now == "Intersemestral":
        idx        = len(inter_labels)-1 if inter_labels else 0
        sel_label  = st.selectbox("Period", inter_labels, index=idx if inter_labels else None)
        sel_value  = sel_label; internal = sel_label
    else:
        idx        = len(year_labels)-1 if year_labels else 0
        sel_label  = st.selectbox("Year", year_labels, index=idx if year_labels else None)
        sel_value  = sel_label; internal = sel_label

    IDCOL = "ID" if "ID" in df.columns else ("ID Nr." if "ID Nr." in df.columns else df.columns[0])
    mode_key = f"{'FT' if st.session_state.modo_faculty=='Full-time' else 'PT'}_{tmode_now}"
    _download_link(
        f"⬇ Download dataset (Excel)",
        df, f"BD_{'FT' if st.session_state.modo_faculty=='Full-time' else 'PT'}.xlsx",
    )

# ── Area palette (UNCHANGED logic) ─────────────────────────────────────────────
areas_all          = sorted(df["AREA_PROFESOR"].dropna().unique().tolist()) if "AREA_PROFESOR" in df.columns else []
areas_palette_order = areas_all
color_map_area     = {a: RANKING_PALETTE[i % len(RANKING_PALETTE)] for i, a in enumerate(areas_palette_order)}

# ── Pivot table (UNCHANGED logic) ─────────────────────────────────────────────
if tmode_now == "Semestral":
    pivot_src = df[df["Periodo"].astype(str).apply(_is_sem_label)]
    pivot_area = pivot_src.pivot_table(index="AREA_PROFESOR", columns="Periodo",
                                       values=IDCOL, aggfunc="nunique", fill_value=0)
elif tmode_now == "Intersemestral":
    pivot_src  = df[df["Periodo"].astype(str).apply(_is_inter_label)]
    pivot_area = pivot_src.pivot_table(index="AREA_PROFESOR", columns="Periodo",
                                       values=IDCOL, aggfunc="nunique", fill_value=0)
else:
    tmp = df.copy()
    tmp["__Year"] = tmp["Periodo"].astype(str).str[:4]
    tmp = tmp.sort_values("Periodo").drop_duplicates(subset=[IDCOL,"__Year"], keep="last")
    pivot_area = tmp.pivot_table(index="AREA_PROFESOR", columns="__Year",
                                 values=IDCOL, aggfunc="nunique", fill_value=0)
pivot_area.loc["Total"] = pivot_area.sum()

# ── Header ─────────────────────────────────────────────────────────────────────
faculty_type = st.session_state.modo_faculty
render_header(
    f"Distribution by Academic Area — {faculty_type}",
    subtitle="Distribution and evolution of faculty across academic areas",
)
render_update_banner(last_updated=_load_ts, is_fresh=True)

# ── KPI Row ────────────────────────────────────────────────────────────────────
df_kpi    = _filter_for_timeframe(df, tmode_now, sel_value)
total_kpi = int(df_kpi[IDCOL].nunique()) if not df_kpi.empty else 0
n_areas   = int(df_kpi["AREA_PROFESOR"].nunique()) if "AREA_PROFESOR" in df_kpi.columns else 0
top_area  = df_kpi["AREA_PROFESOR"].value_counts().idxmax() if (not df_kpi.empty and "AREA_PROFESOR" in df_kpi.columns) else "—"

kpi_card_row([
    {"value": sel_label or "—", "label": "Selected Period"},
    {"value": total_kpi,        "label": f"Total {faculty_type} Faculty"},
    {"value": n_areas,          "label": "Academic Areas"},
    {"value": top_area,         "label": "Largest Area"},
])

section_divider()

# ── Pivot area table ───────────────────────────────────────────────────────────
st.subheader("📋 Faculty Count by Academic Area")

def _bold_total_row(df_):
    s = pd.DataFrame("", index=df_.index, columns=df_.columns)
    if "Total" in df_.index:
        s.loc["Total",:] = "font-weight:700;"
    return s

def _highlight_last_total(df_):
    s = pd.DataFrame("", index=df_.index, columns=df_.columns)
    if len(df_.columns)>0 and "Total" in df_.index:
        s.loc["Total", df_.columns[-1]] = (
            f"background-color:{PALETTE['primary_light']};"
            f"color:{PALETTE['primary_dark']};font-weight:700;"
        )
    return s

st.dataframe(
    pivot_area.style.apply(_bold_total_row, axis=None).apply(_highlight_last_total, axis=None).format(precision=0),
    use_container_width=True,
)
_download_link(
    "⬇ Download table (Excel)",
    pivot_area.reset_index().rename(columns={"AREA_PROFESOR":"Academic Area"}),
    f"Area_Distribution_{tmode_now}.xlsx",
)

section_divider()

# ── Charts: Line (left) + Donut (right) ────────────────────────────────────────
st.subheader("📈 Area Evolution & Period Distribution")

# Build base for line (UNCHANGED logic)
if tmode_now == "Semestral":
    base_line = df[df["Periodo"].astype(str).apply(_is_sem_label)].copy()
    base_line["X"] = base_line["Periodo"].astype(str).map(_display_label_sem)
    x_labels  = sorted(base_line["X"].unique().tolist())
    x_to_filter = sel_label
elif tmode_now == "Intersemestral":
    base_line = df[df["Periodo"].astype(str).apply(_is_inter_label)].copy()
    base_line["X"] = base_line["Periodo"].astype(str)
    x_labels  = sorted(base_line["X"].unique().tolist())
    x_to_filter = sel_label
else:
    tmp = df.copy()
    tmp["__Year"] = tmp["Periodo"].astype(str).str[:4]
    tmp = tmp.sort_values("Periodo").drop_duplicates(subset=[IDCOL,"__Year"], keep="last")
    base_line = tmp.rename(columns={"__Year":"X"}).copy()
    x_labels  = sorted(base_line["X"].unique().tolist())
    x_to_filter = str(sel_label)

totals_period = (
    base_line.groupby("X")[IDCOL].nunique()
    .reindex(x_labels).fillna(0).astype(int).reset_index(name="Total")
)

show_all_key = f'ver_todas_{mode_key}'
area_sel_key = f'area_sel_{mode_key}'
if show_all_key not in st.session_state: st.session_state[show_all_key] = True
if area_sel_key not in st.session_state: st.session_state[area_sel_key] = "Select..."

colL, colR = st.columns([3, 2])

with colL:
    areas_current = sorted([a for a in pivot_area.index if a != "Total"])
    if st.session_state[area_sel_key] not in ["Select...", *areas_current]:
        st.session_state[area_sel_key] = "Select..."

    st.checkbox("Show all lines", key=show_all_key)
    st.selectbox("Select academic area:", ["Select...", *areas_current], key=area_sel_key)

    show_all     = st.session_state[show_all_key]
    area_sel_val = st.session_state[area_sel_key]

    if show_all:
        df_counts = (
            base_line.groupby(["X","AREA_PROFESOR"])[IDCOL]
            .nunique().reset_index(name="Count")
        )
        df_counts["X"] = pd.Categorical(df_counts["X"], categories=x_labels, ordered=True)
        df_counts = df_counts.merge(totals_period, on="X", how="left")
        df_counts["Pct"] = (df_counts["Count"] / df_counts["Total"].replace(0, pd.NA)).fillna(0)

        fig_line = px.line(
            df_counts, x="X", y="Pct", color="AREA_PROFESOR", markers=True,
            category_orders={"X": x_labels, "AREA_PROFESOR": areas_palette_order},
            color_discrete_map=color_map_area,
        )
        fig_line.update_traces(mode="lines+markers", line=dict(width=2),
                               hovertemplate="<b>%{x}</b><br>%{fullData.name}: %{y:.1%}<extra></extra>")
        fig_line.update_xaxes(type="category", categoryorder="array", categoryarray=x_labels, title=None)
        fig_line.update_yaxes(rangemode="tozero", tickformat=".0%", title=None)
        apply_chart_style(fig_line, height=380, show_legend=False)
        add_period_highlight(fig_line, x_to_filter, x_labels)
        st.plotly_chart(fig_line, use_container_width=True)
    else:
        if area_sel_val == "Select...":
            st.info("Select an academic area or enable 'Show all lines'.")
        else:
            df_area = (
                base_line[base_line["AREA_PROFESOR"]==area_sel_val]
                .groupby("X")[IDCOL].nunique()
                .reindex(x_labels).fillna(0).astype(int).reset_index(name="Count")
            )
            df_line = df_area.merge(totals_period, on="X", how="left")
            df_line["Pct"] = (df_line["Count"] / df_line["Total"].replace(0, pd.NA)).fillna(0)
            df_line["X"]   = pd.Categorical(df_line["X"], categories=x_labels, ordered=True)

            fig_line = px.line(
                df_line, x="X", y="Pct", markers=True,
                title=f"Evolution (% share) — {area_sel_val}",
                color_discrete_sequence=[color_map_area.get(area_sel_val, PALETTE["accent1"])],
            )
            fig_line.update_traces(mode="lines+markers", line=dict(width=2),
                                   hovertemplate="<b>%{x}</b><br>%{y:.1%}<extra></extra>")
            fig_line.update_xaxes(type="category", categoryorder="array", categoryarray=x_labels, title=None)
            fig_line.update_yaxes(rangemode="tozero", tickformat=".0%", title=None)
            apply_chart_style(fig_line, height=380, show_legend=False)
            add_period_highlight(fig_line, x_to_filter, x_labels)
            st.plotly_chart(fig_line, use_container_width=True)

with colR:
    st.markdown(
        f"<div style='text-align:center;font-weight:700;color:{PALETTE[\"primary\"]}'>"
        f"Distribution — {sel_label or ''}</div>",
        unsafe_allow_html=True,
    )
    df_donut = _filter_for_timeframe(df, tmode_now, sel_value)
    if df_donut.empty:
        st.info("No data for the selected period.")
    else:
        dist     = df_donut.groupby("AREA_PROFESOR")[IDCOL].nunique().sort_values(ascending=False)
        donut_df = pd.DataFrame({"Area": dist.index, "Value": dist.values})

        fig_donut = px.pie(
            donut_df, names="Area", values="Value", hole=0.45,
            color="Area", color_discrete_map=color_map_area,
            category_orders={"Area": areas_palette_order},
        )
        apply_chart_style(fig_donut, height=380, show_legend=True)
        st.plotly_chart(fig_donut, use_container_width=True)

        fname_donut = (
            f"Donut_{'FT' if st.session_state.modo_faculty=='Full-time' else 'PT'}"
            f"_{tmode_now}_{str(sel_label).replace(' ','_')}.xlsx"
        )
        _download_link("⬇ Download distribution (Excel)", donut_df, fname_donut)

section_divider()

# ── Detail Table ───────────────────────────────────────────────────────────────
detail = _filter_for_timeframe(df, tmode_now, sel_value)
cols_prefer_ft = ["Full Name","AREA_PROFESOR","Faculty Ranking","Faculty Qualific.","P/S"]
cols_prefer_pt = ["Profesor","AREA_PROFESOR","PLANTA_CATEDRA","TIPO","P/S"]
prefer_cols    = cols_prefer_ft if st.session_state.modo_faculty == "Full-time" else cols_prefer_pt
cols_to_show   = [c for c in prefer_cols if c in detail.columns]
detail_out     = detail[cols_to_show].drop_duplicates().reset_index(drop=True)
n_detail       = int(detail[IDCOL].nunique()) if IDCOL in detail.columns else len(detail_out)

st.subheader(f"👥 Faculty Detail — {sel_label}")
st.markdown(
    f"There are **{n_detail}** "
    f"{'full-time' if st.session_state.modo_faculty=='Full-time' else 'part-time'} "
    f"faculty members in **{sel_label}**"
)
st.dataframe(detail_out, use_container_width=True)

fname_det = f"Detail_{'FT' if st.session_state.modo_faculty=='Full-time' else 'PT'}_{tmode_now}_{str(sel_label).replace(' ','_')}.xlsx"
_download_link("⬇ Download detail (Excel)", detail_out, fname_det)
