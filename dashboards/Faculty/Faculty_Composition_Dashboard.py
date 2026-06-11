# ===========================================================================
#  Full-time Faculty Composition · UASM
# ===========================================================================
import streamlit as st
import pandas as pd
import plotly.express as px
import re
import io, base64
import requests

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Full-time Faculty Composition · UASM",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Design tokens ─────────────────────────────────────────────────────────────
_P = {
    "primary":       "#21877D",
    "primary_dark":  "#004d47",
    "primary_light": "#dff7f2",
    "accent1":       "#2EC4B6",
    "accent2":       "#00A896",
    "accent3":       "#56D6C9",
    "danger":        "#E63946",
    "success":       "#06D6A0",
    "neutral":       "#C7C7C7",
    "highlight":     "#D0E5F5",
    "text_muted":    "#6B7280",
    "border":        "#D1E8E4",
}

# ── Global CSS ─────────────────────────────────────────────────────────────────
st.markdown(
    """
    <style>

    .suite-header{
        display:flex;
        flex-direction:column;
        align-items:center;
        padding:16px 24px 12px;
        margin-top:-35px;
        background:linear-gradient(135deg,#004d47 0%,#21877D 60%,#2EC4B6 100%);
        border-radius:12px;
        box-shadow:0 2px 8px rgba(0,77,71,.18);
        margin-bottom:14px;
    }

    .sh-super{
        font-size:11px;
        font-weight:700;
        letter-spacing:2px;
        color:#56D6C9;
        text-transform:uppercase;
        margin-bottom:2px;
    }

    .sh-title{
        font-size:26px;
        font-weight:800;
        color:#fff;
        text-align:center;
        line-height:1.2;
    }

    .sh-sub{
        font-size:13px;
        color:rgba(255,255,255,.75);
        margin-top:4px;
        text-align:center;
    }

    .kpi-row{
        display:flex;
        gap:12px;
        margin-bottom:18px;
        flex-wrap:wrap;
    }

    .kpi-card{
        flex:1;
        min-width:120px;
        background:#F8FFFE;
        border:1px solid #D1E8E4;
        border-radius:10px;
        padding:12px 14px;
        text-align:center;
        box-shadow:0 1px 4px rgba(0,77,71,.07);
    }

    .kv{
        font-size:28px;
        font-weight:800;
        color:#21877D;
        line-height:1.1;
    }

    .kl{
        font-size:11px;
        font-weight:600;
        color:#6B7280;
        text-transform:uppercase;
        letter-spacing:.5px;
        margin-top:3px;
    }

    .sec-sep{
        border:none;
        border-top:1px solid #D1E8E4;
        margin:16px 0;
        opacity:.6;
    }

    .period-label{
        text-align:center;
        font-weight:700;
        font-size:1.05rem;
        color:#21877D;
    }

    a.dl-min,
    a.dl-min:link,
    a.dl-min:visited{
        color:#00A896 !important;
        text-decoration:underline !important;
        font-size:13px;
        display:inline-block;
        margin-top:6px;
    }

    a.dl-min:hover{
        opacity:.85;
    }

    div.stDownloadButton > button{
        background:transparent !important;
        border:none !important;
        box-shadow:none !important;
        color:#21877D !important;
        font-size:13px !important;
        padding:0 !important;
        text-decoration:underline !important;
    }

    div.stDownloadButton{
        margin:2px 0 8px 0;
    }

    thead th{
        background:#dff7f2 !important;
        color:#004d47 !important;
        font-weight:700 !important;
    }

    section[data-testid="stSidebar"]{
        background:#F0F7F6 !important;
    }

    #mode-pill [role="radiogroup"]{
        display:flex;
        gap:8px;
        margin-top:0;
    }

    #mode-pill [role="radio"]{
        flex:1;
        justify-content:center;
        border:1px solid #d0d4d9;
        border-radius:999px;
        padding:8px 12px;
        background:#f0f2f6;
        color:#666;
        font-weight:600;
        cursor:pointer;
        text-align:center;
    }

    #mode-pill [role="radio"][aria-checked="true"]{
        background:#dff7f2;
        color:#004d47;
        border-color:#8fd7cc;
    }

    #mode-pill [data-baseweb="radio"] input{
        display:none !important;
    }

    /* ===== BOTONES SIDEBAR ===== */

    .modern-btn{
        background:#FFFFFF;
        border:1px solid #D1E8E4;
        border-radius:10px;
        padding:12px 14px;
        color:#374151 !important;
        font-size:14px;
        font-weight:600;
        text-decoration:none !important;
        display:block;
        text-align:center;
        margin-bottom:10px;
        transition:all .2s ease;
        box-shadow:0 1px 3px rgba(0,0,0,.04);
    }

    .modern-btn:hover{
        background:#F8FFFE;
        border-color:#B7DCD6;
    }

    div[data-testid="stButton"] button{
        background:#FFFFFF !important;
        border:1px solid #D1E8E4 !important;
        border-radius:10px !important;
        color:#374151 !important;
        font-size:14px !important;
        font-weight:600 !important;
        height:48px !important;
        box-shadow:0 1px 3px rgba(0,0,0,.04) !important;
    }

    div[data-testid="stButton"] button:hover{
        background:#F8FFFE !important;
        border-color:#B7DCD6 !important;
    }

    </style>
    """,
    unsafe_allow_html=True,
)

# ── Inline helpers ─────────────────────────────────────────────────────────────
import io as _io, base64 as _b64

def _xlsx_bytes(df, sheet_name="Data"):
    buf = _io.BytesIO()
    with pd.ExcelWriter(buf) as w:
        df.to_excel(w, index=False, sheet_name=sheet_name[:31])
    buf.seek(0)
    return buf.getvalue()

def _download_link(label, df, filename):
    b64 = _b64.b64encode(_xlsx_bytes(df)).decode()
    href = ("data:application/vnd.openxmlformats-officedocument"
            ".spreadsheetml.sheet;base64," + b64)
    st.markdown(
        '<a class="dl-min" download="' + filename + '" href="' + href + '">' + label + '</a>',
        unsafe_allow_html=True,
    )

def _render_header(title, subtitle=""):
    sub = '<div class="sh-sub">' + subtitle + '</div>' if subtitle else ""
    st.markdown(
        '<div class="suite-header">'
        '<div class="sh-super">UASM \u00b7 Faculty Analytics</div>'
        '<div class="sh-title">' + title + '</div>' + sub +
        '</div>',
        unsafe_allow_html=True,
    )

def _kpi_row(cards):
    html = '<div class="kpi-row">'
    for c in cards:
        html += ('<div class="kpi-card"><div class="kv">' + str(c.get("v", "\u2014")) +
                 '</div><div class="kl">' + str(c.get("l", "")) + '</div></div>')
    html += "</div>"
    st.markdown(html, unsafe_allow_html=True)

def _sec_div():
    st.markdown('<hr class="sec-sep">', unsafe_allow_html=True)
    
def _highlight_band(fig, label, all_labels):
    if label in all_labels:
        pos = all_labels.index(label)
        fig.add_shape(
            type="rect", xref="x", yref="paper",
            x0=pos - 0.4, x1=pos + 0.4, y0=0, y1=1,
            fillcolor=_P["highlight"], opacity=0.35, line_width=0,
        )
    return fig

# ── Header ─────────────────────────────────────────────────────────────────────
_render_header("Full-time Faculty Composition", "Evolution and distribution of full-time faculty by ranking")

DRIVE_FILE_ID = "1rPDVrdIxBFMrf0VkBmLtdUmbhvT4dku-"

@st.cache_data(ttl=300)
def load_data():
    url = f"https://drive.google.com/uc?export=download&id={DRIVE_FILE_ID}"
    output_path = "/tmp/BD_Faculty.xlsx"

    import requests as _req
    response = _req.get(url, stream=True)
    with open(output_path, "wb") as f:
        for chunk in response.iter_content(chunk_size=32768):
            if chunk:
                f.write(chunk)

    # ── Robust sheet detection: never breaks when the sheet is renamed ──
    xls = pd.ExcelFile(output_path)
    possible = ["BD_PLANTA", "BD PLANTA", "BD_PLANTA 2020-2025",
                "BD PLANTA 2020-2025", "BD PLANTA 2020-2026",
                "BD_PLANTA 2020-2025", "BD_PLANTA 2020-2026"]
    sheet_found = next((s for s in possible if s in xls.sheet_names), None)
    if sheet_found is None:
        # Fall back to first sheet as last resort
        sheet_found = xls.sheet_names[0]
    df_ = pd.read_excel(output_path, sheet_name=sheet_found)

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
        s = str(p)
        y = int(s[:4])
        suf = 30 if "Intersemestral" in s else int(s[-2:])
        return (y, suf)
    df_ = df_.sort_values(by="Periodo", key=lambda s: s.map(_key))
    return df_

df = load_data()

# =============================
# UTILS
# =============================

# =============================
# SIDEBAR — Nav + Refresh + Timeframe
# =============================

all_periods   = df["Periodo"].astype(str).unique().tolist()
sem_periods   = [p for p in all_periods if re.fullmatch(r'(?:19|20)\d{2}-(10|20)', p)]
inter_periods = [p for p in all_periods if re.fullmatch(r'(?:19|20)\d{2}\sIntersemestral', p)]
years         = sorted(pd.Series(all_periods).str[:4].unique().tolist())

with st.sidebar:

    # ── Logo UASM ──────────────────────────────────────────────────────────
    st.markdown(
        """
        <div style="text-align:center; padding: 10px 0 4px 0;">
            <img src="https://uniandes.edu.co/sites/default/files/logo-uniandes.png"
                 onerror="this.style.display='none'"
                 style="max-width:120px; height:auto; margin-bottom:6px;" />
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("""
    <div style="
        text-align:center;
        padding-top:0px;
        padding-bottom:20px;
    ">
        <h1 style="
            color:#004d47;
            font-size:22px;
            font-weight:800;
            margin-bottom:0px;
        ">
            UASM Faculty KPIs
        </h1>
        <div style="
            color:#6B7280;
            font-size:12px;
            letter-spacing:1px;
            text-transform:uppercase;
        ">
            Analytics Dashboard
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("#### Timeframe")
    tmode = st.radio("", ["Semestral", "Anual", "Intersemestral"], key="ft_comp_timeframe")

    if tmode == "Semestral":
        vis = [p.replace("-", "") for p in sem_periods]
        idx = len(vis) - 1 if vis else 0
        sel_vis = st.selectbox("Periodo", vis, index=idx if vis else None)
        sel_period_internal = sem_periods[vis.index(sel_vis)] if vis else None
        sel_period_label    = sel_vis
    elif tmode == "Anual":
        idx = len(years) - 1 if years else 0
        sel_period_internal = st.selectbox("Periodo", years, index=idx if years else None)
        sel_period_label    = sel_period_internal
    else:
        idx = len(inter_periods) - 1 if inter_periods else 0
        sel_period_internal = st.selectbox("Periodo", inter_periods, index=idx if inter_periods else None)
        sel_period_label    = sel_period_internal

    xlsx_data = _xlsx_bytes(df)
    b64 = _b64.b64encode(xlsx_data).decode()

    st.markdown(
        f"""
        <a class="modern-btn"
           download="FT_Base_Completa.xlsx"
           href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}">
           ⭳ Descargar Base Completa
        </a>
        """,
        unsafe_allow_html=True
    )

# =============================
# RANKING ORDER & COLOR MAP
# =============================
base_order = [
    "Full Professor", "Associate Professor", "Assistant Professor", "Instructor",
    "Adjunct Faculty", "Distinguished Practitioner", "Emeritus Professor"
]
uniq_ranks    = df["Faculty Ranking"].dropna().astype(str).unique().tolist()
ranking_order = [x for x in base_order if x in uniq_ranks] + [x for x in uniq_ranks if x not in base_order]

if "Faculty Ranking" in df.columns:
    df["Faculty Ranking"] = pd.Categorical(df["Faculty Ranking"], categories=ranking_order, ordered=True)

palette = [
    "#037C70","#27BDAE","#4FFF98","#FFD166",
    "#F4A261","#E76F51","#9D4EDD","#6D597A",
    "#118AB2","#073B4C","#8AC926","#FF70A6"
]
color_map_rk = {rk: palette[i % len(palette)] for i, rk in enumerate(ranking_order)}

# =============================
# HELPERS
# =============================
def periods_for_tables():
    if tmode == "Semestral":       return sem_periods
    if tmode == "Intersemestral":  return inter_periods
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
        return (
            pd.pivot_table(
                df[df["Periodo"].isin(cols)],
                index="Faculty Ranking", columns="Periodo",
                values="ID", aggfunc="count", fill_value=0
            ).reindex(ranking_order)
        )
    out = {rk: {y: 0 for y in cols} for rk in ranking_order}
    for y in cols:
        dfa = df[df["Periodo"].astype(str).str.startswith(str(y))].copy()
        if dfa.empty: continue
        dfa = dfa.sort_values("Periodo").drop_duplicates(subset=["ID"], keep="last")
        for rk, v in dfa.groupby("Faculty Ranking")["ID"].count().items():
            out.setdefault(rk, {})[y] = int(v)
    return pd.DataFrame(out).T.reindex(ranking_order).reindex(columns=cols, fill_value=0)

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
        cts["Periodo"] = str(y)
        rows.append(cts)
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

def highlight_current_period(fig, current_period, xcats):
    if not current_period or current_period not in xcats: return
    pos = xcats.index(current_period)
    fig.add_shape(
        type="rect", xref="x", yref="paper",
        x0=pos-0.4, x1=pos+0.4, y0=0, y1=1,
        fillcolor="#D0E5F5", opacity=0.35, line_width=0
    )

# =============================
# PIVOT TABLE
# =============================
pivot = pivot_counts_by_ranking().reindex(ranking_order)
pivot.loc["Total"] = pivot.sum(numeric_only=True)

st.subheader("Number of Full-time Faculty by Ranking")

def _bold_total(df_):
    s = pd.DataFrame("", index=df_.index, columns=df_.columns)
    if "Total" in df_.index: s.loc["Total", :] = "font-weight:700;"
    return s

def _highlight_last(df_):
    s = pd.DataFrame("", index=df_.index, columns=df_.columns)
    if len(df_.columns) > 0 and "Total" in df_.index:
        s.loc["Total", df_.columns[-1]] = "background-color:#dff7f2;color:#004d47;font-weight:700;"
    return s

st.dataframe(
    pivot.style.apply(_bold_total, axis=None).apply(_highlight_last, axis=None).format(precision=0),
    use_container_width=True
)
_download_link(
    "Descargar tabla (Excel)",
    pivot.reset_index().rename(columns={"index":"Faculty Ranking"}),
    f"FT_Composition_{tmode}.xlsx"
)

# =============================
# CHARTS
# =============================
periods_sorted = periods_for_tables()
st.session_state.setdefault("show_all", True)
st.session_state.setdefault("single_ranking", "Select...")

def on_select_ranking():
    if st.session_state.single_ranking != "Select...":
        st.session_state.show_all = False

def on_toggle_show_all():
    if st.session_state.show_all:
        st.session_state.single_ranking = "Select..."

st.header("Evolution & composition")
col_left, col_right = st.columns(2)

with col_right:
    st.subheader("Composition by period")
    st.markdown(
        f"<div style='text-align:center;font-weight:800;font-size:2rem;padding-top:4px;'>{sel_period_label}</div>",
        unsafe_allow_html=True
    )
    if tmode in ("Semestral", "Intersemestral"):
        dfbar = df[df["Periodo"].astype(str).eq(sel_period_internal)]
    else:
        y    = str(sel_period_internal)
        dfa  = df[df["Periodo"].astype(str).str.startswith(y)].copy()
        dfbar = dfa.sort_values("Periodo").drop_duplicates(subset=["ID"], keep="last")

    bar_counts = (
        dfbar.groupby("Faculty Ranking")["ID"].count()
             .reindex(ranking_order).fillna(0).reset_index()
    )
    bar_counts.columns = ["Faculty Ranking", "Count"]
    st.metric("Total Faculty:", int(bar_counts["Count"].sum()))

    fig_bar = px.bar(
        bar_counts, x="Count", y="Faculty Ranking", orientation="h",
        text="Count", color="Faculty Ranking", color_discrete_map=color_map_rk,
        category_orders={"Faculty Ranking": ranking_order[::-1]}
    )
    fig_bar.update_xaxes(range=[0, max(1, int(bar_counts["Count"].max() or 0)) + 5], title=None)
    fig_bar.update_yaxes(title=None)
    fig_bar.update_traces(textposition="outside")
    st.plotly_chart(fig_bar, use_container_width=True)

with col_left:
    st.subheader("Evolution of rankings")
    st.checkbox("Show all lines", key="show_all", on_change=on_toggle_show_all)
    st.selectbox("Select a ranking:", ["Select..."] + ranking_order,
                 key="single_ranking", on_change=on_select_ranking)

    fig_line = None
    xcats    = periods_sorted

    if st.session_state.show_all:
        data_long, xcats = line_source_all()
        y_max = max(1, int(data_long["Count"].max()) if not data_long.empty else 0)
        fig_line = px.line(
            data_long, x="Periodo", y="Count", color="Faculty Ranking",
            markers=True, title="Evolution — all rankings",
            color_discrete_map=color_map_rk,
            category_orders={"Periodo": xcats, "Faculty Ranking": ranking_order}
        )
        fig_line.update_yaxes(range=[0, y_max + 1], title=None)
        fig_line.update_xaxes(type="category", categoryorder="array", categoryarray=xcats, title=None)
        fig_line.update_layout(height=550, showlegend=False)
    else:
        rk = st.session_state.single_ranking
        if rk != "Select...":
            data_single, xcats = line_source_single(rk)
            y_max = max(1, int(data_single["Count"].max()) if not data_single.empty else 0)
            fig_line = px.line(
                data_single, x="Periodo", y="Count", markers=True,
                title=f"Evolution — {rk}",
                color_discrete_sequence=[color_map_rk.get(rk, "#00A896")],
                category_orders={"Periodo": xcats}
            )
            fig_line.update_yaxes(range=[0, y_max + 1], title=None)
            fig_line.update_xaxes(type="category", categoryorder="array", categoryarray=xcats, title=None)
            fig_line.update_layout(height=480)
        else:
            st.info("Select a ranking to visualize its evolution.")

    if fig_line is not None:
        _highlight_band(fig_line, sel_period_internal, list(xcats))
        st.plotly_chart(fig_line, use_container_width=True)

# =============================
# DETAIL TABLE
# =============================
st.subheader("Faculty Detail")
active = df_active_for_selection()
selected_ranking = (
    None if st.session_state.show_all or st.session_state.single_ranking == "Select..."
    else st.session_state.single_ranking
)

if selected_ranking:
    detail_df = active[active["Faculty Ranking"] == selected_ranking].copy()
    title_txt = f"### **{len(detail_df)}** **{selected_ranking}** in period **{sel_period_label}**"
else:
    detail_df = active.copy()
    title_txt = f"### **{len(detail_df)}** Full-time Faculty in period **{sel_period_label}**"

st.markdown(title_txt)
detail_cols = [
    "Periodo","ID","ID Nr.","Full Name","Academic Area",
    "Faculty Ranking","Subcategorization","Faculty Qualific.","P/S",
    "Highest Earned Degree","Year","University","Normal professional Resp."
]
show_cols = [c for c in detail_cols if c in detail_df.columns]
st.dataframe(detail_df[show_cols], use_container_width=True)
_download_link(
    "Descargar detalle (Excel)",
    detail_df[show_cols],
    f"FT_Composition_Detail_{sel_period_label}.xlsx"
)
