# =============================================================================
#  1 · Full-time Faculty Composition  — UASM Faculty Analytics Suite
#  Refactored: design, UX, structure.  Business logic UNCHANGED.
# =============================================================================
import re
import datetime
import requests
import streamlit as st
import pandas as pd
import plotly.express as px

from suite_styles import (
    apply_page_config, inject_global_css, render_header,
    render_sidebar_nav, render_update_banner, kpi_card_row, section_divider,
    apply_chart_style, add_period_highlight,
    PALETTE, RANKING_PALETTE,
    _xlsx_bytes, _download_link,
)

# ── Page bootstrap ─────────────────────────────────────────────────────────────
apply_page_config("Full-time Faculty Composition · UASM")
inject_global_css()

# ── Data Load ──────────────────────────────────────────────────────────────────
DRIVE_FILE_ID = "1rPDVrdIxBFMrf0VkBmLtdUmbhvT4dku-"

@st.cache_data(ttl=300)
def load_data():
    # ── Business logic UNCHANGED ──
    url = f"https://drive.google.com/uc?export=download&id={DRIVE_FILE_ID}"
    output_path = "/tmp/BD_Faculty.xlsx"
    response = requests.get(url, stream=True)
    with open(output_path, "wb") as f:
        for chunk in response.iter_content(chunk_size=32768):
            if chunk:
                f.write(chunk)

    df_ = pd.read_excel(output_path, sheet_name="BD PLANTA 2020-2025")

    def _norm_per(val):
        s = str(val).strip()
        m_inter = re.search(r'((?:19|20)\d{2}).{0,6}inter', s, flags=re.IGNORECASE)
        if m_inter:
            return f"{m_inter.group(1)} Intersemestral"
        m = re.search(r'((?:19|20)\d{2})\D?(\d{2})', s)
        if m:
            return f"{m.group(1)}-{m.group(2)}"
        return None

    first_col = df_.columns[0]
    df_["Periodo"] = df_[first_col].map(_norm_per)
    valid = df_["Periodo"].astype(str).str.match(
        r'^(?:19|20)\d{2}-(10|20)$|^(?:19|20)\d{2}\sIntersemestral$'
    )
    df_ = df_.loc[valid].copy()
    if "ID Nr." in df_.columns and "ID" not in df_.columns:
        df_ = df_.rename(columns={"ID Nr.": "ID"})

    def _key(p):
        s = str(p); y = int(s[:4])
        suf = 30 if "Intersemestral" in s else int(s[-2:])
        return (y, suf)
    df_ = df_.sort_values(by="Periodo", key=lambda s: s.map(_key))
    return df_, datetime.datetime.now()

df, _load_ts = load_data()
# unwrap tuple if cached result is tuple
if isinstance(df, tuple):
    df, _load_ts = df

# ── Ranking & color setup (UNCHANGED logic) ────────────────────────────────────
base_order = [
    "Full Professor", "Associate Professor", "Assistant Professor", "Instructor",
    "Adjunct Faculty", "Distinguished Practitioner", "Emeritus Professor",
]
uniq_ranks    = df["Faculty Ranking"].dropna().astype(str).unique().tolist()
ranking_order = [x for x in base_order if x in uniq_ranks] + [x for x in uniq_ranks if x not in base_order]
if "Faculty Ranking" in df.columns:
    df["Faculty Ranking"] = pd.Categorical(df["Faculty Ranking"], categories=ranking_order, ordered=True)
color_map_rk = {rk: RANKING_PALETTE[i % len(RANKING_PALETTE)] for i, rk in enumerate(ranking_order)}

# ── Period helpers (UNCHANGED logic) ──────────────────────────────────────────
all_periods   = df["Periodo"].astype(str).unique().tolist()
sem_periods   = [p for p in all_periods if re.fullmatch(r'(?:19|20)\d{2}-(10|20)', p)]
inter_periods = [p for p in all_periods if re.fullmatch(r'(?:19|20)\d{2}\sIntersemestral', p)]
years         = sorted(pd.Series(all_periods).str[:4].unique().tolist())

def periods_for_tables():
    if tmode == "Semestral":      return sem_periods
    if tmode == "Intersemestral": return inter_periods
    return years

def df_active_for_selection():
    if sel_period_internal is None:
        return df.iloc[0:0].copy()
    if tmode in ("Semestral", "Intersemestral"):
        return df[df["Periodo"].astype(str).eq(sel_period_internal)].copy()
    y   = str(sel_period_internal)
    dfa = df[df["Periodo"].astype(str).str.startswith(y)].copy()
    return dfa.sort_values("Periodo").drop_duplicates(subset=["ID"], keep="last")

def pivot_counts_by_ranking():
    cols = periods_for_tables()
    if tmode in ("Semestral", "Intersemestral"):
        return pd.pivot_table(
            df[df["Periodo"].isin(cols)],
            index="Faculty Ranking", columns="Periodo",
            values="ID", aggfunc="count", fill_value=0,
        ).reindex(ranking_order)
    out = {rk: {y: 0 for y in cols} for rk in ranking_order}
    for y in cols:
        dfa = df[df["Periodo"].astype(str).str.startswith(str(y))].copy()
        dfa = dfa.sort_values("Periodo").drop_duplicates(subset=["ID"], keep="last")
        for rk in ranking_order:
            out[rk][y] = int(dfa[dfa["Faculty Ranking"] == rk]["ID"].count())
    return pd.DataFrame(out).T

def line_source_all():
    cols = periods_for_tables()
    if tmode in ("Semestral", "Intersemestral"):
        dat = (
            df[df["Periodo"].isin(cols)]
            .groupby(["Periodo","Faculty Ranking"])["ID"]
            .count().reset_index(name="Count")
        )
        dat["Periodo"] = pd.Categorical(dat["Periodo"], categories=cols, ordered=True)
        return dat, cols
    rows = []
    for y in cols:
        dfa = df[df["Periodo"].astype(str).str.startswith(str(y))].copy()
        dfa = dfa.sort_values("Periodo").drop_duplicates(subset=["ID"], keep="last")
        cts = (
            dfa.groupby("Faculty Ranking")["ID"].count()
               .reindex(ranking_order, fill_value=0).reset_index()
               .rename(columns={"ID":"Count"})
        )
        cts["Periodo"] = str(y); rows.append(cts)
    out = pd.concat(rows, ignore_index=True) if rows else pd.DataFrame(columns=["Faculty Ranking","Count","Periodo"])
    out["Periodo"] = pd.Categorical(out["Periodo"], categories=cols, ordered=True)
    return out, cols

def line_source_single(rank):
    cols = periods_for_tables()
    if tmode in ("Semestral", "Intersemestral"):
        dat = (
            df[df["Periodo"].isin(cols) & (df["Faculty Ranking"] == rank)]
            .groupby("Periodo")["ID"].count()
            .reindex(cols, fill_value=0).reset_index(name="Count")
        )
        return dat, cols
    vals = []
    for y in cols:
        dfa = df[df["Periodo"].astype(str).str.startswith(str(y))].copy()
        dfa = dfa.sort_values("Periodo").drop_duplicates(subset=["ID"], keep="last")
        vals.append({"Periodo": str(y), "Count": int(dfa[dfa["Faculty Ranking"] == rank]["ID"].count())})
    return pd.DataFrame(vals), cols

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    render_sidebar_nav(current_key="1 · Full-time Composition")

    st.sidebar.markdown("#### ⏱ Timeframe")
    tmode = st.radio("", ["Semestral", "Anual", "Intersemestral"], key="ft_comp_timeframe")

    if tmode == "Semestral":
        vis = [p.replace("-", "") for p in sem_periods]
        idx = len(vis) - 1 if vis else 0
        sel_vis = st.selectbox("Period", vis, index=idx if vis else None)
        sel_period_internal = sem_periods[vis.index(sel_vis)] if vis else None
        sel_period_label    = sel_vis
    elif tmode == "Anual":
        idx = len(years) - 1 if years else 0
        sel_period_internal = st.selectbox("Year", years, index=idx if years else None)
        sel_period_label    = sel_period_internal
    else:
        idx = len(inter_periods) - 1 if inter_periods else 0
        sel_period_internal = st.selectbox("Period", inter_periods, index=idx if inter_periods else None)
        sel_period_label    = sel_period_internal

    st.markdown("<hr style='margin:10px 0;opacity:.4'>", unsafe_allow_html=True)
    _download_link("⬇ Download full dataset (Excel)", df, "FT_Base_Completa.xlsx")

# ── Header ─────────────────────────────────────────────────────────────────────
render_header(
    "Full-time Faculty Composition",
    subtitle="Evolution and distribution of full-time faculty by ranking",
)
render_update_banner(last_updated=_load_ts, is_fresh=True)

# ── KPI Row ─────────────────────────────────────────────────────────────────────
active_kpi = df_active_for_selection()
total_now  = int(active_kpi["ID"].count()) if not active_kpi.empty else 0
n_ranks    = int(active_kpi["Faculty Ranking"].nunique()) if not active_kpi.empty else 0
top_rank   = (
    active_kpi["Faculty Ranking"].value_counts().idxmax()
    if not active_kpi.empty else "—"
)
kpi_card_row([
    {"value": sel_period_label or "—", "label": "Selected Period"},
    {"value": total_now,               "label": "Total Faculty"},
    {"value": n_ranks,                 "label": "Active Rankings"},
    {"value": top_rank,                "label": "Largest Ranking"},
])

section_divider()

# ── Pivot Table ────────────────────────────────────────────────────────────────
pivot = pivot_counts_by_ranking().reindex(ranking_order)
pivot.loc["Total"] = pivot.sum(numeric_only=True)

st.subheader("📋 Faculty Count by Ranking")

def _bold_total(df_):
    s = pd.DataFrame("", index=df_.index, columns=df_.columns)
    if "Total" in df_.index:
        s.loc["Total", :] = "font-weight:700;"
    return s

def _highlight_last(df_):
    s = pd.DataFrame("", index=df_.index, columns=df_.columns)
    if len(df_.columns) > 0 and "Total" in df_.index:
        s.loc["Total", df_.columns[-1]] = (
            f"background-color:{PALETTE['primary_light']};"
            f"color:{PALETTE['primary_dark']};font-weight:700;"
        )
    return s

st.dataframe(
    pivot.style.apply(_bold_total, axis=None).apply(_highlight_last, axis=None).format(precision=0),
    use_container_width=True,
)
_download_link(
    "⬇ Download table (Excel)",
    pivot.reset_index().rename(columns={"index": "Faculty Ranking"}),
    f"FT_Composition_{tmode}.xlsx",
)

section_divider()

# ── Charts: Evolution (left) + Composition bar (right) ─────────────────────────
periods_sorted = periods_for_tables()
st.session_state.setdefault("show_all", True)
st.session_state.setdefault("single_ranking", "Select...")

def on_select_ranking():
    if st.session_state.single_ranking != "Select...":
        st.session_state.show_all = False

def on_toggle_show_all():
    if st.session_state.show_all:
        st.session_state.single_ranking = "Select..."

st.subheader("📈 Evolution & Composition")
col_left, col_right = st.columns([3, 2])

# ── RIGHT: Bar chart for selected period ───────────────────────────────────────
with col_right:
    st.markdown(
        f"<div style='text-align:center;font-weight:700;font-size:1rem;"
        f"color:{PALETTE[\"primary\"]};margin-bottom:4px'>"
        f"Composition — {sel_period_label}</div>",
        unsafe_allow_html=True,
    )
    if tmode in ("Semestral", "Intersemestral"):
        dfbar = df[df["Periodo"].astype(str).eq(sel_period_internal)]
    else:
        y_   = str(sel_period_internal)
        dfa  = df[df["Periodo"].astype(str).str.startswith(y_)].copy()
        dfbar = dfa.sort_values("Periodo").drop_duplicates(subset=["ID"], keep="last")

    bar_counts = (
        dfbar.groupby("Faculty Ranking")["ID"].count()
             .reindex(ranking_order).fillna(0).reset_index()
    )
    bar_counts.columns = ["Faculty Ranking", "Count"]
    total_bar = int(bar_counts["Count"].sum())
    st.metric("Total Faculty", total_bar)

    fig_bar = px.bar(
        bar_counts, x="Count", y="Faculty Ranking", orientation="h",
        text="Count", color="Faculty Ranking",
        color_discrete_map=color_map_rk,
        category_orders={"Faculty Ranking": ranking_order[::-1]},
    )
    fig_bar.update_xaxes(range=[0, max(1, int(bar_counts["Count"].max() or 0)) + 5], title=None)
    fig_bar.update_yaxes(title=None)
    fig_bar.update_traces(textposition="outside")
    apply_chart_style(fig_bar, height=360)
    st.plotly_chart(fig_bar, use_container_width=True)

# ── LEFT: Line chart (evolution) ───────────────────────────────────────────────
with col_left:
    st.checkbox("Show all lines", key="show_all", on_change=on_toggle_show_all)
    st.selectbox(
        "Filter to ranking:",
        ["Select..."] + ranking_order,
        key="single_ranking",
        on_change=on_select_ranking,
    )

    fig_line = None
    xcats    = periods_sorted

    if st.session_state.show_all:
        data_long, xcats = line_source_all()
        y_max = max(1, int(data_long["Count"].max()) if not data_long.empty else 0)
        fig_line = px.line(
            data_long, x="Periodo", y="Count", color="Faculty Ranking",
            markers=True, title="Evolution — all rankings",
            color_discrete_map=color_map_rk,
            category_orders={"Periodo": xcats, "Faculty Ranking": ranking_order},
        )
        fig_line.update_yaxes(range=[0, y_max + 1], title=None)
        fig_line.update_xaxes(type="category", categoryorder="array", categoryarray=xcats, title=None)
        apply_chart_style(fig_line, height=400, show_legend=False)
    else:
        rk = st.session_state.single_ranking
        if rk != "Select...":
            data_single, xcats = line_source_single(rk)
            y_max = max(1, int(data_single["Count"].max()) if not data_single.empty else 0)
            fig_line = px.line(
                data_single, x="Periodo", y="Count", markers=True,
                title=f"Evolution — {rk}",
                color_discrete_sequence=[color_map_rk.get(rk, PALETTE["accent1"])],
                category_orders={"Periodo": xcats},
            )
            fig_line.update_yaxes(range=[0, y_max + 1], title=None)
            fig_line.update_xaxes(type="category", categoryorder="array", categoryarray=xcats, title=None)
            apply_chart_style(fig_line, height=400)
        else:
            st.info("Select a ranking to visualize its evolution.")

    if fig_line is not None:
        add_period_highlight(fig_line, sel_period_internal, list(xcats))
        st.plotly_chart(fig_line, use_container_width=True)

section_divider()

# ── Detail Table ───────────────────────────────────────────────────────────────
st.subheader("🔍 Faculty Detail")
active = df_active_for_selection()
selected_ranking = (
    None
    if st.session_state.show_all or st.session_state.single_ranking == "Select..."
    else st.session_state.single_ranking
)

if selected_ranking:
    detail_df  = active[active["Faculty Ranking"] == selected_ranking].copy()
    title_txt  = (
        f"**{len(detail_df)}** · {selected_ranking} "
        f"— period **{sel_period_label}**"
    )
else:
    detail_df  = active.copy()
    title_txt  = f"**{len(detail_df)}** Full-time Faculty — period **{sel_period_label}**"

st.markdown(title_txt)

detail_cols = [
    "Periodo","ID","ID Nr.","Full Name","Academic Area",
    "Faculty Ranking","Subcategorization","Faculty Qualific.","P/S",
    "Highest Earned Degree","Year","University","Normal professional Resp.",
]
show_cols = [c for c in detail_cols if c in detail_df.columns]
st.dataframe(detail_df[show_cols], use_container_width=True)
_download_link(
    "⬇ Download detail (Excel)",
    detail_df[show_cols],
    f"FT_Composition_Detail_{sel_period_label}.xlsx",
)
