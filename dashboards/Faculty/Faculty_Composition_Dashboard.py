import streamlit as st
import pandas as pd
import plotly.express as px
import webbrowser
import os
import re
import io, base64
import numpy as np

# =============================
# PAGE CONFIG
# =============================
st.set_page_config(
    page_title="Dashboard",
    page_icon=":bar_chart:",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =============================
# INSTITUTIONAL HEADER
# =============================
header = st.container()
with header:
    st.markdown(
        """
        <style>
        .header-title {
            color: #21877D;
            font-weight: bold;
            text-align:center;
            font-size:32px;
        }
        .header-btn {
            background-color:#21877D;
            color:white !important;
            padding:8px 16px;
            border:none;
            border-radius:8px;
            cursor:pointer;
            font-size:14px;
            text-decoration:none !important;
            display:inline-block;
        }
        .header-btn:hover {
            background-color:#1a6b62;
        }
        a.dl-min, a.dl-min:link, a.dl-min:visited {
          color:#1FA89B !important; text-decoration:underline !important;
          font-size:13px; display:inline-block; margin-top:6px;
        }
        a.dl-min:hover { opacity:.85; }
        </style>
        """,
        unsafe_allow_html=True
    )

with st.container():
    cols = st.columns([1, 3, 1], gap="small")
    with cols[0]:
        st.markdown(
            '<a href="https://facultyqualificationspy-drvj3wpyrxvm2lrnafdwx5.streamlit.app/" class="header-btn" target="_self">⬅ Previous KPI</a>',
            unsafe_allow_html=True
        )
    with cols[1]:
        st.markdown('<div class="header-title">Full-time Faculty Composition</div>', unsafe_allow_html=True)
    with cols[2]:
        st.markdown(
            '<a href="https://facultystaffinglevelsdashboardpy-phv4t8jzbyyz5rrepqttuf.streamlit.app/" class="header-btn" target="_self">➡ Next KPI</a>',
            unsafe_allow_html=True
        )
# =============================
# LOAD DATA (con soporte Intersemestral)
# =============================
@st.cache_data
def load_data():
    df_ = pd.read_excel(
        r"data/Faculty/BD_Faculty.xlsx",
        sheet_name="BD PLANTA 2020-2025"
    )

    # Construye Periodo robusto: YYYY-SS o "YYYY Intersemestral"
    def _norm_per(val):
        s = str(val).strip()
        m_inter = re.search(r'((?:19|20)\d{2}).{0,6}inter', s, flags=re.IGNORECASE)
        if m_inter:
            return f"{m_inter.group(1)} Intersemestral"
        m = re.search(r'((?:19|20)\d{2})\D?(\d{2})', s)  # 202010 / 2020-10 / 2020_10
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

    # Orden estable de Periodo para series
    def _key(p):
        s = str(p)
        y = int(s[:4])
        suf = 30 if "Intersemestral" in s else int(s[-2:])
        return (y, suf)
    df_ = df_.sort_values(by="Periodo", key=lambda s: s.map(_key))

    return df_

df = load_data()

# ========= Descarga utils =========
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
    st.markdown(f'<a class="dl-min" download="{filename}" href="{href}">{label}</a>', unsafe_allow_html=True)

# =============================
# SIDE NAV
# =============================

LINKS = {
    "1 Full-time Composition": "https://facultycompositiondashboardpy-dtacyzfa3otmpbewqc5axu.streamlit.app/",
    "2 Full-time Staffing Levels": "https://facultystaffinglevelsdashboardpy-phv4t8jzbyyz5rrepqttuf.streamlit.app/",
    "3 Distribution by Academic Area": "https://facultydistributionareadashboardpy-yzwpiqdlukfdp6qcygxjhj.streamlit.app/",
    "4 Faculty Demographics": "https://facultydemographicsdashboardpy-kmsnpswxs35psbqtdtvb6y.streamlit.app/",
    "5 Full-time Faculty Questionnaire": "https://full-timefacultyactivitiespy-bbe7fmmyrxvssadnygm4fx.streamlit.app/",
    "6 Faculty Qualifications": "https://facultyqualificationspy-drvj3wpyrxvm2lrnafdwx5.streamlit.app/",
    "Open main HTML menu": None
}

choices = [k for k,v in LINKS.items() if v]  # solo los que tienen URL
sel = st.selectbox("Ir a KPI:", choices)
st.link_button(f"Abrir: {sel}", LINKS[sel], use_container_width=True)

options = {
    "Select...": None,
    "1 Full-time Composition": "https://facultycompositiondashboardpy-dtacyzfa3otmpbewqc5axu.streamlit.app/",
    "2 Full-time Staffing Levels": "https://facultystaffinglevelsdashboardpy-phv4t8jzbyyz5rrepqttuf.streamlit.app/",
    "3 Distribution by Academic Area": "https://facultydistributionareadashboardpy-yzwpiqdlukfdp6qcygxjhj.streamlit.app/",
    "4 Faculty Demographics": "https://facultydemographicsdashboardpy-kmsnpswxs35psbqtdtvb6y.streamlit.app/",
    "5 Full-time Faculty Questionnaire": "https://full-timefacultyactivitiespy-bbe7fmmyrxvssadnygm4fx.streamlit.app/",
    "6 Faculty Qualifications": "https://facultyqualificationspy-drvj3wpyrxvm2lrnafdwx5.streamlit.app/",
    "Open main HTML menu": "web/KPIs/Faculty/Web KPIs - Faculty.html"
}
choice = st.sidebar.selectbox("📊 Go to KPI:", list(options.keys()))
target = options.get(choice)
if target:
    if target.endswith(".html"):
        abs_path = os.path.abspath(target)
        webbrowser.open(f"file:///{abs_path}")
        st.success("The Faculty menu was opened in a new browser tab.")
    else:
        st.markdown(f'<meta http-equiv="refresh" content="0; url={target}" />', unsafe_allow_html=True)

# =============================
# TIMEFRAME + PERIODO (SIDEBAR)
# =============================
all_periods = df["Periodo"].astype(str).unique().tolist()
sem_periods  = [p for p in all_periods if re.fullmatch(r'(?:19|20)\d{2}-(10|20)', p)]
inter_periods = [p for p in all_periods if re.fullmatch(r'(?:19|20)\d{2}\sIntersemestral', p)]
years = sorted(pd.Series(all_periods).str[:4].unique().tolist())

with st.sidebar:
    st.markdown("---")
    st.markdown("#### Timeframe")
    tmode = st.radio("", ["Semestral", "Anual", "Intersemestral"], key="ft_comp_timeframe")

    if tmode == "Semestral":
        vis = [p.replace("-", "") for p in sem_periods]  # sin guion
        idx = len(vis) - 1 if vis else 0
        sel_vis = st.selectbox("Periodo", vis, index=idx if vis else None)
        sel_period_internal = sem_periods[vis.index(sel_vis)] if vis else None
        sel_period_label = sel_vis
    elif tmode == "Anual":
        idx = len(years) - 1 if years else 0
        sel_period_internal = st.selectbox("Periodo", years, index=idx if years else None)
        sel_period_label = sel_period_internal
    else:
        idx = len(inter_periods) - 1 if inter_periods else 0
        sel_period_internal = st.selectbox("Periodo", inter_periods, index=idx if inter_periods else None)
        sel_period_label = sel_period_internal

    # ---- Descarga BD completa (sidebar)
    _download_link("Descargar base completa (Excel) — Full-time", df, "FT_Base_Completa.xlsx")

# =============================
# RANKING ORDER & COLOR MAP (basado en lo que realmente existe)
# =============================
base_order = [
    "Full Professor", "Associate Professor", "Assistant Professor", "Instructor",
    "Adjunct Faculty", "Distinguished Practitioner", "Emeritus Professor"
]
uniq_ranks = df["Faculty Ranking"].dropna().astype(str).unique().tolist()
ranking_order = [x for x in base_order if x in uniq_ranks] + [x for x in uniq_ranks if x not in base_order]

if "Faculty Ranking" in df.columns:
    df['Faculty Ranking'] = pd.Categorical(df['Faculty Ranking'], categories=ranking_order, ordered=True)

creative_palette = [
    "#037C70", "#27BDAE", "#4FFF98", "#FFD166",
    "#F4A261", "#E76F51", "#9D4EDD", "#6D597A",
    "#118AB2", "#073B4C", "#8AC926", "#FF70A6"
]
color_map_rk = {rk: creative_palette[i % len(creative_palette)] for i, rk in enumerate(ranking_order)}

# =============================
# HELPERS Timeframe
# =============================
def periods_for_tables():
    if tmode == "Semestral":
        return sem_periods
    if tmode == "Intersemestral":
        return inter_periods
    return years  # anual

def df_active_for_selection():
    """Subconjunto activo según timeframe + periodo seleccionado."""
    if sel_period_internal is None:
        return df.iloc[0:0].copy()
    if tmode in ("Semestral", "Intersemestral"):
        return df[df["Periodo"].astype(str).eq(sel_period_internal)].copy()
    # Anual: última aparición por ID en el año
    y = str(sel_period_internal)
    dfa = df[df["Periodo"].astype(str).str.startswith(y)].copy()
    dfa = dfa.sort_values("Periodo").drop_duplicates(subset=["ID"], keep="last")
    return dfa

def pivot_counts_by_ranking():
    """Tabla de conteos (filas=ranking, columnas=periodos según timeframe)."""
    cols = periods_for_tables()
    if tmode in ("Semestral", "Intersemestral"):
        piv = (
            pd.pivot_table(
                df[df["Periodo"].isin(cols)],
                index="Faculty Ranking", columns="Periodo", values="ID",
                aggfunc="count", fill_value=0
            ).reindex(ranking_order)
        )
        return piv
    else:
        # Anual: conteo de IDs únicos (última aparición en el año)
        out = {rk: {y: 0 for y in cols} for rk in ranking_order}
        for y in cols:
            dfa = df[df["Periodo"].astype(str).str.startswith(str(y))].copy()
            if dfa.empty:
                continue
            dfa = dfa.sort_values("Periodo").drop_duplicates(subset=["ID"], keep="last")
            cts = dfa.groupby("Faculty Ranking")["ID"].count()
            for rk, v in cts.items():
                out.setdefault(rk, {})[y] = int(v)
        piv = pd.DataFrame(out).T.reindex(ranking_order).reindex(columns=cols, fill_value=0)
        return piv

def line_source_all():
    """Fuente para líneas (todas las jerarquías) según timeframe."""
    cols = periods_for_tables()
    if tmode in ("Semestral", "Intersemestral"):
        dat = (
            df[df["Periodo"].isin(cols)]
            .groupby(["Periodo", "Faculty Ranking"])["ID"]
            .count()
            .reset_index(name="Count")
        )
        dat["Periodo"] = pd.Categorical(dat["Periodo"], categories=cols, ordered=True)
        return dat, cols
    else:
        rows = []
        for y in cols:
            dfa = df[df["Periodo"].astype(str).str.startswith(str(y))].copy()
            dfa = dfa.sort_values("Periodo").drop_duplicates(subset=["ID"], keep="last")
            cts = (
                dfa.groupby("Faculty Ranking")["ID"]
                   .count()
                   .reindex(ranking_order, fill_value=0)
                   .reset_index()
                   .rename(columns={"ID":"Count"})
            )
            cts["Periodo"] = str(y)
            rows.append(cts)
        out = pd.concat(rows, ignore_index=True) if rows else pd.DataFrame(columns=["Faculty Ranking","Count","Periodo"])
        out["Periodo"] = pd.Categorical(out["Periodo"], categories=cols, ordered=True)
        return out, cols

def line_source_single(rank):
    """Fuente para línea de 1 ranking según timeframe."""
    cols = periods_for_tables()
    if tmode in ("Semestral", "Intersemestral"):
        dat = (
            df[df["Periodo"].isin(cols) & (df["Faculty Ranking"]==rank)]
            .groupby("Periodo")["ID"].count()
            .reindex(cols, fill_value=0)
            .reset_index(name="Count")
        )
        return dat, cols
    else:
        vals = []
        for y in cols:
            dfa = df[df["Periodo"].astype(str).str.startswith(str(y))].copy()
            dfa = dfa.sort_values("Periodo").drop_duplicates(subset=["ID"], keep="last")
            cnt = int(dfa[dfa["Faculty Ranking"]==rank]["ID"].count())
            vals.append({"Periodo": str(y), "Count": cnt})
        dat = pd.DataFrame(vals)
        return dat, cols

def sel_key_for_band():
    """Etiqueta exacta que existe en el eje X para resaltar."""
    return sel_period_internal

# =============================
# PIVOT (COUNTS) + STYLES (respeta timeframe)
# =============================
pivot = pivot_counts_by_ranking()
pivot = pivot.reindex(ranking_order)
pivot.loc['Total'] = pivot.sum(numeric_only=True)

st.subheader("Number of Full-time Faculty by Ranking")

def _bold_total_row(df_):
    styles = pd.DataFrame('', index=df_.index, columns=df_.columns)
    if 'Total' in df_.index:
        styles.loc['Total', :] = 'font-weight:700;'
    return styles

def _highlight_total_latest(df_):
    styles = pd.DataFrame('', index=df_.index, columns=df_.columns)
    if len(df_.columns) > 0 and 'Total' in df_.index:
        last_col = df_.columns[-1]
        styles.loc['Total', last_col] = 'background-color:#dff7f2; color:#004d47; font-weight:700;'
    return styles

styled_pivot = (
    pivot
    .style
    .apply(_bold_total_row, axis=None)
    .apply(_highlight_total_latest, axis=None)
    .format(precision=0)
)
st.dataframe(styled_pivot, use_container_width=True)

# Descarga (debajo de la tabla)
_download_link("Descargar tabla (Excel)", pivot.reset_index().rename(columns={"index":"Faculty Ranking"}), f"FT_Composition_{tmode}.xlsx")

# =============================
# SHARED STATE (selección de líneas)
# =============================
periods_sorted = periods_for_tables()  # ahora depende del timeframe

st.session_state.setdefault("show_all", True)
st.session_state.setdefault("single_ranking", "Select...")

def on_select_ranking():
    if st.session_state.single_ranking != "Select...":
        st.session_state.show_all = False

def on_toggle_show_all():
    if st.session_state.show_all:
        st.session_state.single_ranking = "Select..."

# =============================
# Helper: light blue band for current period
# =============================
def highlight_current_period(fig, current_period, xcats):
    if (current_period is None) or (current_period not in xcats):
        return
    pos = xcats.index(current_period)
    fig.add_shape(
        type="rect",
        xref="x", yref="paper",
        x0=pos - 0.4, x1=pos + 0.4, y0=0, y1=1,
        fillcolor="#D0E5F5", opacity=0.35, line_width=0
    )

# =============================
# LAYOUT
# =============================
st.header("Evolution & composition")
col_left, col_right = st.columns(2)

# =============================
# RIGHT: HORIZONTAL BARS (sin flechas; usa periodo del sidebar)
# =============================
with col_right:
    st.subheader("Composition by period")
    st.markdown(
        f"<div style='text-align:center; font-weight:800; font-size:2rem; padding-top:4px;'>{sel_period_label}</div>",
        unsafe_allow_html=True
    )

    # Fuente para barras (según timeframe)
    if tmode in ("Semestral", "Intersemestral"):
        dfbar = df[df["Periodo"].astype(str).eq(sel_period_internal)]
    else:
        # Año: última aparición del año por ID
        y = str(sel_period_internal)
        dfa = df[df["Periodo"].astype(str).str.startswith(y)].copy()
        dfa = dfa.sort_values("Periodo").drop_duplicates(subset=["ID"], keep="last")
        dfbar = dfa

    bar_counts = (
        dfbar.groupby("Faculty Ranking")["ID"].count()
             .reindex(ranking_order).fillna(0).reset_index()
    )
    bar_counts.columns = ["Faculty Ranking", "Count"]

    total_prof = int(bar_counts["Count"].sum())
    st.metric(f"Total Faculty:", total_prof)

    fig_bar = px.bar(
        bar_counts,
        x="Count",
        y="Faculty Ranking",
        orientation="h",
        text="Count",
        color="Faculty Ranking",
        color_discrete_map=color_map_rk,
        category_orders={"Faculty Ranking": ranking_order[::-1]}  # <<< invertir el orden
    )
    xmax = int(bar_counts["Count"].max() or 0)
    fig_bar.update_xaxes(range=[0, max(1, xmax) + 5], title=None)
    fig_bar.update_yaxes(title=None)
    fig_bar.update_traces(textposition="outside")
    st.plotly_chart(fig_bar, use_container_width=True)

# =============================
# LEFT: LINES (usa timeframe; resalta periodo seleccionado)
# =============================
with col_left:
    st.subheader("Evolution of rankings")

    st.checkbox(
        "Show all lines",
        key="show_all",
        on_change=on_toggle_show_all
    )
    st.selectbox(
        "Select a ranking:",
        ["Select..."] + ranking_order,
        key="single_ranking",
        on_change=on_select_ranking
    )

    fig_line = None
    xcats = periods_sorted
    sel_for_band = sel_key_for_band()

    if st.session_state.show_all:
        data_long, xcats = line_source_all()
        y_max = int(data_long["Count"].max()) if not data_long.empty else 0
        y_max = max(1, y_max)

        fig_line = px.line(
            data_long,
            x="Periodo",
            y="Count",
            color="Faculty Ranking",
            markers=True,
            title="Evolution — all rankings",
            color_discrete_map=color_map_rk,               # <<< MISMO MAPA QUE BARRAS
            category_orders={"Periodo": xcats, "Faculty Ranking": ranking_order}
        )
        fig_line.update_yaxes(range=[0, y_max + 1], title=None)
        fig_line.update_xaxes(type="category", categoryorder="array", categoryarray=xcats, title=None)
        fig_line.update_layout(height=550, showlegend=False)
    else:
        rk = st.session_state.single_ranking
        if rk != "Select...":
            data_single, xcats = line_source_single(rk)
            y_max = int(data_single["Count"].max()) if not data_single.empty else 0
            y_max = max(1, y_max)

            fig_line = px.line(
                data_single,
                x="Periodo",
                y="Count",
                markers=True,
                title=f"Evolution — {rk}",
                color_discrete_sequence=[color_map_rk.get(rk, "#00A896")],  # <<< MISMO COLOR QUE EN BARRAS
                category_orders={"Periodo": xcats}
            )
            fig_line.update_yaxes(range=[0, y_max + 1], title=None)
            fig_line.update_xaxes(type="category", categoryorder="array", categoryarray=xcats, title=None)
            fig_line.update_layout(height=480)
        else:
            st.info("Select a ranking to visualize its evolution")

    if fig_line is not None:
        highlight_current_period(fig_line, sel_for_band, list(xcats))
        st.plotly_chart(fig_line, use_container_width=True)

# =============================
# DETAIL TABLE (sin selectores/flechas; usa periodo del sidebar)
# =============================
st.subheader("Faculty Detail")

active = df_active_for_selection()
selected_ranking = (None if st.session_state.show_all or st.session_state.single_ranking == "Select..."
                    else st.session_state.single_ranking)

if selected_ranking:
    detail_df = active[active["Faculty Ranking"] == selected_ranking].copy()
    title_txt = f"### **{len(detail_df)}** **{selected_ranking}** in period **{sel_period_label}**"
else:
    detail_df = active.copy()
    title_txt = f"### **{len(detail_df)}** Full-time Faculty in period **{sel_period_label}**"

st.markdown(title_txt)

detail_cols = [
    "Periodo", "ID", "ID Nr.", "Full Name", "Academic Area",
    "Faculty Ranking", "Subcategorization", "Faculty Qualific.", "P/S",
    "Highest Earned Degree", "Year", "University", "Normal professional Resp."
]
show_cols = [c for c in detail_cols if c in detail_df.columns]
st.dataframe(detail_df[show_cols], use_container_width=True)

# Descarga (debajo de la tabla de detalle)
_download_link("Descargar detalle (Excel)", detail_df[show_cols], f"FT_Composition_Detail_{sel_period_label}.xlsx")



