# ===========================================================================
#  UASM · Faculty Analytics Suite  ·  App multipágina (un solo archivo)
#  Página 1: Full-time Faculty Composition
#  Página 2: Full-time Faculty Staffing Levels
# ===========================================================================
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import re
import io
import base64
import time
import requests

# ===========================================================================
# 1) CONFIGURACIÓN GLOBAL (una sola vez para toda la app)
# ===========================================================================
st.set_page_config(
    page_title="UASM Faculty Analytics",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS compartido por ambas páginas ───────────────────────────────────────
st.markdown(
    "<style>"
    ".suite-header{display:flex;flex-direction:column;margin-top:-35px;align-items:center;"
    "padding:16px 24px 12px;"
    "background:linear-gradient(135deg,#004d47 0%,#21877D 60%,#2EC4B6 100%);"
    "border-radius:12px;box-shadow:0 2px 8px rgba(0,77,71,.18);margin-bottom:14px;}"
    ".sh-super{font-size:11px;font-weight:700;letter-spacing:2px;"
    "color:#56D6C9;text-transform:uppercase;margin-bottom:2px;}"
    ".sh-title{font-size:26px;font-weight:800;color:#fff;text-align:center;line-height:1.2;}"
    ".sh-sub{font-size:13px;color:rgba(255,255,255,.75);margin-top:4px;text-align:center;}"
    ".kpi-row{display:flex;gap:12px;margin-bottom:18px;flex-wrap:wrap;}"
    ".kpi-card{flex:1;min-width:120px;background:#F8FFFE;border:1px solid #D1E8E4;"
    "border-radius:10px;padding:12px 14px;text-align:center;"
    "box-shadow:0 1px 4px rgba(0,77,71,.07);}"
    ".kv{font-size:28px;font-weight:800;color:#21877D;line-height:1.1;}"
    ".kl{font-size:11px;font-weight:600;color:#6B7280;"
    "text-transform:uppercase;letter-spacing:.5px;margin-top:3px;}"
    ".sec-sep{border:none;border-top:1px solid #D1E8E4;margin:16px 0;opacity:.6;}"
    ".period-label{text-align:center;font-weight:700;font-size:1.05rem;color:#21877D;}"
    "a.dl-min,a.dl-min:link,a.dl-min:visited{color:#00A896 !important;"
    "text-decoration:underline !important;font-size:13px;"
    "display:inline-block;margin-top:6px;}"
    "a.dl-min:hover{opacity:.85;}"
    "div.stDownloadButton>button{background:transparent !important;"
    "border:none !important;box-shadow:none !important;"
    "color:#21877D !important;font-size:13px !important;"
    "padding:0 !important;text-decoration:underline !important;}"
    "div.stDownloadButton{margin:2px 0 8px 0;}"
    "thead th{background:#dff7f2 !important;color:#004d47 !important;"
    "font-weight:700 !important;}"
    "section[data-testid='stSidebar']{background:#F0F7F6 !important;}"
    "#mode-pill [role='radiogroup']{display:flex;gap:8px;margin-top:0;}"
    "#mode-pill [role='radio']{flex:1;justify-content:center;"
    "border:1px solid #d0d4d9;border-radius:999px;padding:8px 12px;"
    "background:#f0f2f6;color:#666;font-weight:600;cursor:pointer;text-align:center;}"
    "#mode-pill [role='radio'][aria-checked='true']{"
    "background:#dff7f2;color:#004d47;border-color:#8fd7cc;}"
    "#mode-pill [data-baseweb='radio'] input{display:none !important;}"
    ".modern-btn{background:#FFFFFF;border:1px solid #D1E8E4;"
    "border-radius:10px;padding:12px 14px;color:#374151 !important;"
    "font-size:14px;font-weight:600;text-decoration:none !important;"
    "display:block;text-align:center;margin-bottom:10px;"
    "transition:all .2s ease;box-shadow:0 1px 3px rgba(0,0,0,.04);}"
    ".modern-btn:hover{background:#F8FFFE;border-color:#B7DCD6;}"
    "div[data-testid='stButton'] button{background:#FFFFFF !important;"
    "border:1px solid #D1E8E4 !important;border-radius:10px !important;"
    "color:#374151 !important;font-size:14px !important;"
    "font-weight:600 !important;height:48px !important;"
    "box-shadow:0 1px 3px rgba(0,0,0,.04) !important;}"
    "div[data-testid='stButton'] button:hover{"
    "background:#F8FFFE !important;border-color:#B7DCD6 !important;}"
    "</style>",
    unsafe_allow_html=True,
)

# ===========================================================================
# 2) HELPERS COMPARTIDOS
# ===========================================================================
def _xlsx_bytes(df, sheet_name="Data"):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf) as w:
        df.to_excel(w, index=False, sheet_name=sheet_name[:31])
    buf.seek(0)
    return buf.getvalue()


def _download_link(label, df, filename):
    b64 = base64.b64encode(_xlsx_bytes(df)).decode()
    href = ("data:application/vnd.openxmlformats-officedocument"
            ".spreadsheetml.sheet;base64," + b64)
    st.markdown(
        f'<a class="dl-min" download="{filename}" href="{href}">{label}</a>',
        unsafe_allow_html=True,
    )


def _render_header(title, subtitle=""):
    sub = f'<div class="sh-sub">{subtitle}</div>' if subtitle else ""
    st.markdown(
        f'<div class="suite-header"><div class="sh-super">UASM · Faculty Analytics</div>'
        f'<div class="sh-title">{title}</div>{sub}</div>',
        unsafe_allow_html=True,
    )


def _highlight_band(fig, label, all_labels, color="#D0E5F5"):
    if label in all_labels:
        pos = all_labels.index(label)
        fig.add_shape(type="rect", xref="x", yref="paper",
                      x0=pos - 0.4, x1=pos + 0.4, y0=0, y1=1,
                      fillcolor=color, opacity=0.35, line_width=0)


# ===========================================================================
# 3) CARGA DE DATOS (compartida por ambas páginas)
# ===========================================================================
SHEET_ID = "1PZkqgtvct5LFNWVUEkA5fuglvqvAuMxseSq10MV9ji8"


@st.cache_data(ttl=300)
def load_data():
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"

    resp = None
    last_err = None
    # Reintenta hasta 3 veces: la primera descarga en frío contra Google a
    # veces redirige a un host googleusercontent.com que puede tardar.
    for attempt in range(3):
        try:
            resp = requests.get(url, timeout=60)
            break
        except Exception as e:
            last_err = e
            time.sleep(2)

    if resp is None:
        st.error(f"🌐 No se pudo conectar con Google Sheets tras varios intentos: {last_err}")
        st.stop()

    if resp.status_code != 200 or "text/html" in resp.headers.get("Content-Type", ""):
        st.error("🔒 El Google Sheet no está compartido públicamente. "
                 "Ve a Compartir → Cualquiera con el enlace → Lector.")
        st.stop()

    raw = io.BytesIO(resp.content)
    try:
        xls = pd.ExcelFile(raw)
    except Exception:
        st.error("❌ El archivo descargado no es un Excel válido.")
        st.stop()

    tab = next((s for s in xls.sheet_names if "planta" in s.lower()), xls.sheet_names[0])
    raw.seek(0)
    df_ = pd.read_excel(raw, sheet_name=tab)

    def _norm_per(val):
        if pd.isna(val):
            return None
        try:
            s = str(int(float(str(val).strip())))
        except (ValueError, OverflowError):
            s = str(val).strip()
        if re.search(r'inter', s, re.IGNORECASE):
            y = re.search(r'((?:19|20)\d{2})', s)
            return f"{y.group(1)} Intersemestral" if y else None
        m = re.fullmatch(r'((?:19|20)\d{2})(10|20)', s) or \
            re.search(r'((?:19|20)\d{2})[^0-9]+(10|20)', s)
        return f"{m.group(1)}-{m.group(2)}" if m else None

    src = "Periodo" if "Periodo" in df_.columns else df_.columns[0]
    df_["Periodo"] = df_[src].map(_norm_per)
    df_ = df_[df_["Periodo"].astype(str).str.match(
        r'^(?:19|20)\d{2}-(10|20)$|^(?:19|20)\d{2}\s+Intersemestral$'
    )].copy()

    if "ID Nr." in df_.columns and "ID" not in df_.columns:
        df_ = df_.rename(columns={"ID Nr.": "ID"})
    if "ID" not in df_.columns and "ID Nr." in df_.columns:
        df_["ID"] = df_["ID Nr."]

    def _key(p):
        s = str(p)
        return (int(s[:4]), 30 if "Intersemestral" in s else int(s[-2:]))

    return df_.sort_values("Periodo", key=lambda c: c.map(_key))


df = load_data()

# ===========================================================================
# 4) SIDEBAR — encabezado común (logo + título), visible en todas las páginas
# ===========================================================================
with st.sidebar:
    col_logo, col_title = st.columns([1, 3])
    with col_logo:
        st.image("web/imagenes/logo.png", width=65)
    with col_title:
        st.markdown(
            '<div style="padding-top:10px;color:#004d47;font-size:24px;'
            'font-weight:800;line-height:1.1;">UASM Faculty KPIs</div>',
            unsafe_allow_html=True,
        )
        st.caption("Analytics Dashboard")
    st.markdown("---")


# ===========================================================================
# 5) PÁGINA 1 — Full-time Faculty Composition
# ===========================================================================
def page_composition():
    all_periods = df["Periodo"].astype(str).unique().tolist()
    sem_periods = [p for p in all_periods if re.fullmatch(r'(?:19|20)\d{2}-(10|20)', p)]
    inter_periods = [p for p in all_periods if re.fullmatch(r'(?:19|20)\d{2}\sIntersemestral', p)]
    years = sorted(pd.Series(all_periods).str[:4].unique().tolist())

    # ── Sidebar específico de esta página ──────────────────────────────
    with st.sidebar:
        st.markdown("#### Timeframe")
        tmode = st.radio("", ["Semestral", "Anual", "Intersemestral"], key="ft_comp_timeframe")

        if tmode == "Semestral":
            vis = [p.replace("-", "") for p in sem_periods]
            sel_vis = st.selectbox("Periodo", vis, index=len(vis) - 1 if vis else 0,
                                    key="ft_comp_periodo_sem")
            sel_period_internal = sem_periods[vis.index(sel_vis)] if vis else None
            sel_period_label = sel_vis
        elif tmode == "Anual":
            sel_period_internal = st.selectbox("Periodo", years, index=len(years) - 1 if years else 0,
                                                key="ft_comp_periodo_anual")
            sel_period_label = sel_period_internal
        else:
            sel_period_internal = st.selectbox("Periodo", inter_periods,
                                                index=len(inter_periods) - 1 if inter_periods else 0,
                                                key="ft_comp_periodo_inter")
            sel_period_label = sel_period_internal

        st.markdown("---")
        xlsx_data = _xlsx_bytes(df)
        b64 = base64.b64encode(xlsx_data).decode()
        st.markdown(
            f'<a class="modern-btn" download="FT_Base_Completa.xlsx" '
            f'href="data:application/vnd.openxmlformats-officedocument.'
            f'spreadsheetml.sheet;base64,{b64}">⭳ Descargar Base Completa</a>',
            unsafe_allow_html=True)

    _render_header("Full-time Faculty Composition",
                   "Evolution and distribution of full-time faculty by ranking")

    # ── Ranking order & color map ───────────────────────────────────────
    base_order = ["Full Professor", "Associate Professor", "Assistant Professor",
                  "Instructor", "Adjunct Faculty", "Distinguished Practitioner", "Emeritus Professor"]
    uniq_ranks = df["Faculty Ranking"].dropna().astype(str).unique().tolist()
    ranking_order = [x for x in base_order if x in uniq_ranks] + \
                     [x for x in uniq_ranks if x not in base_order]
    if "Faculty Ranking" in df.columns:
        df["Faculty Ranking"] = pd.Categorical(df["Faculty Ranking"],
                                                categories=ranking_order, ordered=True)

    palette = ["#037C70", "#27BDAE", "#4FFF98", "#FFD166",
               "#F4A261", "#E76F51", "#9D4EDD", "#6D597A",
               "#118AB2", "#073B4C", "#8AC926", "#FF70A6"]
    color_map_rk = {rk: palette[i % len(palette)] for i, rk in enumerate(ranking_order)}

    # ── Helpers de filtrado por periodo ────────────────────────────────
    def periods_for_tables():
        if tmode == "Semestral":
            return sem_periods
        if tmode == "Intersemestral":
            return inter_periods
        return years

    def df_active():
        if sel_period_internal is None:
            return df.iloc[0:0].copy()
        if tmode in ("Semestral", "Intersemestral"):
            return df[df["Periodo"].astype(str).eq(sel_period_internal)].copy()
        dfa = df[df["Periodo"].astype(str).str.startswith(str(sel_period_internal))].copy()
        return dfa.sort_values("Periodo").drop_duplicates(subset=["ID"], keep="last")

    def pivot_counts():
        cols = periods_for_tables()
        if tmode in ("Semestral", "Intersemestral"):
            return pd.pivot_table(
                df[df["Periodo"].isin(cols)],
                index="Faculty Ranking", columns="Periodo",
                values="ID", aggfunc="count", fill_value=0
            ).reindex(ranking_order)
        out = {rk: {y: 0 for y in cols} for rk in ranking_order}
        for y in cols:
            dfa = df[df["Periodo"].astype(str).str.startswith(str(y))].copy()
            if dfa.empty:
                continue
            dfa = dfa.sort_values("Periodo").drop_duplicates(subset=["ID"], keep="last")
            for rk, v in dfa.groupby("Faculty Ranking")["ID"].count().items():
                out.setdefault(rk, {})[y] = int(v)
        return pd.DataFrame(out).T.reindex(ranking_order).reindex(columns=cols, fill_value=0)

    def line_source_all():
        cols = periods_for_tables()
        if tmode in ("Semestral", "Intersemestral"):
            dat = (df[df["Periodo"].isin(cols)]
                   .groupby(["Periodo", "Faculty Ranking"])["ID"]
                   .count().reset_index(name="Count"))
            dat["Periodo"] = pd.Categorical(dat["Periodo"], categories=cols, ordered=True)
            return dat, cols
        rows = []
        for y in cols:
            dfa = df[df["Periodo"].astype(str).str.startswith(str(y))].copy()
            dfa = dfa.sort_values("Periodo").drop_duplicates(subset=["ID"], keep="last")
            cts = (dfa.groupby("Faculty Ranking")["ID"].count()
                      .reindex(ranking_order, fill_value=0).reset_index()
                      .rename(columns={"ID": "Count"}))
            cts["Periodo"] = str(y)
            rows.append(cts)
        out = pd.concat(rows, ignore_index=True) if rows else \
              pd.DataFrame(columns=["Faculty Ranking", "Count", "Periodo"])
        out["Periodo"] = pd.Categorical(out["Periodo"], categories=cols, ordered=True)
        return out, cols

    def line_source_single(rank):
        cols = periods_for_tables()
        if tmode in ("Semestral", "Intersemestral"):
            dat = (df[df["Periodo"].isin(cols) & (df["Faculty Ranking"] == rank)]
                   .groupby("Periodo")["ID"].count()
                   .reindex(cols, fill_value=0).reset_index(name="Count"))
            return dat, cols
        vals = []
        for y in cols:
            dfa = df[df["Periodo"].astype(str).str.startswith(str(y))].copy()
            dfa = dfa.sort_values("Periodo").drop_duplicates(subset=["ID"], keep="last")
            vals.append({"Periodo": str(y),
                         "Count": int(dfa[dfa["Faculty Ranking"] == rank]["ID"].count())})
        return pd.DataFrame(vals), cols

    # ── Pivot table ──────────────────────────────────────────────────────
    pivot = pivot_counts().reindex(ranking_order)
    pivot.loc["Total"] = pivot.sum(numeric_only=True)

    st.subheader("Number of Full-time Faculty by Ranking")

    def _bold_total(df_):
        s = pd.DataFrame("", index=df_.index, columns=df_.columns)
        if "Total" in df_.index:
            s.loc["Total", :] = "font-weight:700;"
        return s

    def _highlight_last(df_):
        s = pd.DataFrame("", index=df_.index, columns=df_.columns)
        if len(df_.columns) > 0 and "Total" in df_.index:
            s.loc["Total", df_.columns[-1]] = "background-color:#dff7f2;color:#004d47;font-weight:700;"
        return s

    st.dataframe(
        pivot.style.apply(_bold_total, axis=None).apply(_highlight_last, axis=None).format(precision=0),
        use_container_width=True)
    _download_link("Descargar tabla (Excel)",
                   pivot.reset_index().rename(columns={"index": "Faculty Ranking"}),
                   f"FT_Composition_{tmode}.xlsx")

    # ── Charts ───────────────────────────────────────────────────────────
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
        st.markdown(f"<div style='text-align:center;font-weight:800;font-size:2rem;"
                    f"padding-top:4px;'>{sel_period_label}</div>", unsafe_allow_html=True)

        if tmode in ("Semestral", "Intersemestral"):
            dfbar = df[df["Periodo"].astype(str).eq(sel_period_internal)]
        else:
            dfa = df[df["Periodo"].astype(str).str.startswith(str(sel_period_internal))].copy()
            dfbar = dfa.sort_values("Periodo").drop_duplicates(subset=["ID"], keep="last")

        bar_counts = (dfbar.groupby("Faculty Ranking")["ID"].count()
                           .reindex(ranking_order).fillna(0).reset_index())
        bar_counts.columns = ["Faculty Ranking", "Count"]
        st.metric("Total Faculty:", int(bar_counts["Count"].sum()))

        fig_bar = px.bar(bar_counts, x="Count", y="Faculty Ranking", orientation="h",
                         text="Count", color="Faculty Ranking", color_discrete_map=color_map_rk,
                         category_orders={"Faculty Ranking": ranking_order[::-1]})
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
        xcats = periods_sorted

        if st.session_state.show_all:
            data_long, xcats = line_source_all()
            y_max = max(1, int(data_long["Count"].max()) if not data_long.empty else 0)
            fig_line = px.line(data_long, x="Periodo", y="Count", color="Faculty Ranking",
                               markers=True, title="Evolution — all rankings",
                               color_discrete_map=color_map_rk,
                               category_orders={"Periodo": xcats, "Faculty Ranking": ranking_order})
            fig_line.update_yaxes(range=[0, y_max + 1], title=None)
            fig_line.update_xaxes(type="category", categoryorder="array", categoryarray=xcats, title=None)
            fig_line.update_layout(height=550, showlegend=False)
        else:
            rk = st.session_state.single_ranking
            if rk != "Select...":
                data_single, xcats = line_source_single(rk)
                y_max = max(1, int(data_single["Count"].max()) if not data_single.empty else 0)
                fig_line = px.line(data_single, x="Periodo", y="Count", markers=True,
                                   title=f"Evolution — {rk}",
                                   color_discrete_sequence=[color_map_rk.get(rk, "#00A896")],
                                   category_orders={"Periodo": xcats})
                fig_line.update_yaxes(range=[0, y_max + 1], title=None)
                fig_line.update_xaxes(type="category", categoryorder="array", categoryarray=xcats, title=None)
                fig_line.update_layout(height=480)
            else:
                st.info("Select a ranking to visualize its evolution.")

        if fig_line is not None:
            _highlight_band(fig_line, sel_period_internal, list(xcats))
            st.plotly_chart(fig_line, use_container_width=True)

    # ── Detail table ─────────────────────────────────────────────────────
    st.subheader("Faculty Detail")
    active = df_active()
    selected_ranking = (None if st.session_state.show_all or
                        st.session_state.single_ranking == "Select..."
                        else st.session_state.single_ranking)

    if selected_ranking:
        detail_df = active[active["Faculty Ranking"] == selected_ranking].copy()
        title_txt = f"### **{len(detail_df)}** **{selected_ranking}** in period **{sel_period_label}**"
    else:
        detail_df = active.copy()
        title_txt = f"### **{len(detail_df)}** Full-time Faculty in period **{sel_period_label}**"

    st.markdown(title_txt)
    detail_cols = ["Periodo", "ID", "ID Nr.", "Full Name", "Academic Area",
                   "Faculty Ranking", "Subcategorization", "Faculty Qualific.", "P/S",
                   "Highest Earned Degree", "Year", "University", "Normal professional Resp."]
    show_cols = [c for c in detail_cols if c in detail_df.columns]
    st.dataframe(detail_df[show_cols], use_container_width=True)
    _download_link("Descargar detalle (Excel)", detail_df[show_cols],
                   f"FT_Composition_Detail_{sel_period_label}.xlsx")


# ===========================================================================
# 6) PÁGINA 2 — Full-time Faculty Staffing Levels
# ===========================================================================
def page_staffing():
    all_periods = sorted(df["Periodo"].astype(str).unique().tolist())
    sem_periods = [p for p in all_periods if re.fullmatch(r'(?:19|20)\d{2}-(10|20)', p)]

    # ── Sidebar específico de esta página ──────────────────────────────
    with st.sidebar:
        st.markdown("#### Select Semester")
        vis_opts = [p.replace("-", "") for p in sem_periods]
        idx = len(vis_opts) - 1 if vis_opts else 0
        sel_vis = st.selectbox("", vis_opts, index=idx if vis_opts else None, key="ft_staff_periodo")
        sel_period_internal = sem_periods[vis_opts.index(sel_vis)] if vis_opts else None
        sel_period_label = sel_vis

        _b64_dl = base64.b64encode(_xlsx_bytes(df)).decode()
        st.markdown(
            '<a class="modern-btn" download="FT_Base_Completa.xlsx" '
            f'href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{_b64_dl}">'
            '⭳ Descargar Base Completa</a>',
            unsafe_allow_html=True,
        )

    _render_header("Full-time Faculty Staffing Levels", "New entrants, leavers, and headcount evolution")

    # ── Helpers (solo semestral) ────────────────────────────────────────
    def perlist_sem():
        return sem_periods

    def final_count_series_sem(df_):
        return df_.groupby("Periodo")["ID"].nunique()

    def in_out_counts_sem(df_, label):
        counts = {}
        for p in sem_periods:
            flat = p.replace("-", "")
            counts[p] = int(df_["Notes"].astype(str).str.contains(
                fr"\b{label}\s+IN\s+\(?{flat}\)?\b", case=False, na=False
            ).sum())
        return pd.Series(counts)

    # ── Staffing summary table ──────────────────────────────────────────
    cols_summary = perlist_sem()
    fin_ser = final_count_series_sem(df).reindex(cols_summary, fill_value=0)
    new_ser = in_out_counts_sem(df, "IN").reindex(cols_summary, fill_value=0)
    out_ser = in_out_counts_sem(df, "OUT").reindex(cols_summary, fill_value=0)

    rows = []
    for i, key in enumerate(cols_summary):
        new_hires = int(new_ser.get(key, 0))
        leavers = int(out_ser.get(key, 0))
        if i == 0:
            start_val = int(fin_ser.iloc[0]) - new_hires + leavers
        else:
            start_val = int(fin_ser.iloc[i - 1])
        rows.append({
            "Start": int(start_val),
            "New": new_hires,
            "Leavers": leavers,
            "Final": int(fin_ser.iloc[i])
        })

    summary_df = pd.DataFrame(rows, index=cols_summary).T

    st.subheader("New entrants and leavers")

    def _bold_final_row(df_):
        styles = pd.DataFrame('', index=df_.index, columns=df_.columns)
        if 'Final' in df_.index:
            styles.loc['Final', :] = 'font-weight:700;'
        return styles

    def _highlight_final_latest(df_):
        styles = pd.DataFrame('', index=df_.index, columns=df_.columns)
        if len(df_.columns) > 0 and 'Final' in df_.index:
            last_col = df_.columns[-1]
            styles.loc['Final', last_col] = 'background-color:#dff7f2; color:#004d47; font-weight:700;'
        return styles

    styled_summary = (
        summary_df
        .style
        .apply(_bold_final_row, axis=None)
        .apply(_highlight_final_latest, axis=None)
        .format(precision=0)
    )
    st.dataframe(styled_summary, use_container_width=True)

    sum_left, sum_right = st.columns([1, 5])
    with sum_left:
        simple_tbl = summary_df.reset_index().rename(columns={"index": "Metric"})
        _download_link("Descargar tabla (Excel)", simple_tbl, "FT_New_Leavers_Semestral.xlsx")

    # ── Charts layout ────────────────────────────────────────────────────
    areas = sorted(df.get("Academic Area", pd.Series(dtype=object)).dropna().unique().tolist())
    col_left, col_right = st.columns([3, 2])

    with col_right:
        st.markdown(f"<div style='text-align:center;font-weight:700'>Period: {sel_period_label}</div>",
                   unsafe_allow_html=True)

        current_period = sel_period_internal or ""
        flat_list = [current_period.replace("-", "")] if current_period else []

        pat_in = "|".join([re.escape(f) for f in flat_list]) if flat_list else r"$^"
        df_in = df[df["Notes"].astype(str).str.contains(fr"\bIN\s+IN\s+\(?({pat_in})\)?\b", case=False, na=False)]
        df_out = df[df["Notes"].astype(str).str.contains(fr"\bOUT\s+IN\s+\(?({pat_in})\)?\b", case=False, na=False)]

        new_by_area = df_in.groupby("Academic Area")["ID"].nunique().reindex(areas, fill_value=0)
        left_by_area = df_out.groupby("Academic Area")["ID"].nunique().reindex(areas, fill_value=0)
        net_by_area = (new_by_area - left_by_area).astype(int)
        order = net_by_area.sort_values(ascending=True).index

        ret_vals = left_by_area.reindex(order).astype(int)
        new_vals = new_by_area.reindex(order).astype(int)

        fig_tornado = go.Figure()
        fig_tornado.add_trace(go.Bar(
            y=order, x=-ret_vals, orientation="h",
            name="Leavers", marker_color="#C0392B",
            text=ret_vals, texttemplate="%{text}", textposition="inside",
            insidetextanchor="middle", textfont=dict(size=14, color="white"),
            hovertemplate="Area: %{y}<br>Leavers: %{customdata}<extra></extra>",
            customdata=ret_vals
        ))
        fig_tornado.add_trace(go.Bar(
            y=order, x=new_vals, orientation="h",
            name="New", marker_color="#56d6c9",
            text=new_vals, texttemplate="%{text}", textposition="inside",
            insidetextanchor="middle", textfont=dict(size=14, color="white"),
            hovertemplate="Area: %{y}<br>New: %{customdata}<extra></extra>",
            customdata=new_vals
        ))

        fig_tornado.update_xaxes(showticklabels=False, showgrid=False, zeroline=False, showline=False)
        fig_tornado.update_yaxes(autorange="reversed")
        fig_tornado.update_layout(
            title=f"New vs Leavers by Area — {sel_period_label}",
            barmode="relative",
            height=max(360, 24 * len(areas)),
            margin=dict(l=10, r=10, t=20, b=80),
            legend=dict(orientation="h", y=-0.25, yanchor="top", x=0.5, xanchor="center"),
            xaxis_title=None, yaxis_title=None
        )
        st.plotly_chart(fig_tornado, use_container_width=True)

    with col_left:
        st.markdown("### Evolution of Faculty (Start vs Final)")

        x_periods = list(summary_df.columns)
        y_start = pd.to_numeric(summary_df.loc["Start"], errors="coerce").tolist() if "Start" in summary_df.index else []
        y_final = pd.to_numeric(summary_df.loc["Final"], errors="coerce").tolist() if "Final" in summary_df.index else []

        fig_line = go.Figure()
        fig_line.add_trace(go.Scatter(
            x=x_periods, y=y_start, mode="lines+markers",
            name="Start", line=dict(shape="linear", width=3, dash="dot")
        ))
        fig_line.add_trace(go.Scatter(
            x=x_periods, y=y_final, mode="lines+markers",
            name="Final", line=dict(shape="linear", width=3, color="#003366")
        ))

        for p, s, f in zip(x_periods, y_start, y_final):
            if pd.isna(s) or pd.isna(f):
                continue
            arrowcolor = "green" if f > s else ("red" if f < s else None)
            if not arrowcolor:
                continue
            fig_line.add_annotation(
                x=p, y=f, ax=p, ay=s,
                xref="x", yref="y", axref="x", ayref="y",
                showarrow=True, arrowhead=2, arrowsize=1.2, arrowwidth=2, arrowcolor=arrowcolor
            )

        fig_line.update_xaxes(type="category", tickangle=45, showgrid=True, title=None)
        fig_line.update_yaxes(showgrid=True, zeroline=False, title=None)
        fig_line.update_layout(
            height=380,
            margin=dict(l=10, r=10, t=10, b=80),
            legend=dict(orientation="h", y=-0.25, yanchor="top", x=0.5, xanchor="center"),
        )

        sel_for_band = sel_period_internal
        if sel_for_band in x_periods:
            pos = x_periods.index(sel_for_band)
            fig_line.add_shape(
                type="rect", xref="x", yref="paper",
                x0=pos - 0.4, x1=pos + 0.4, y0=0, y1=1,
                fillcolor="#D0E5F5", opacity=0.35, line_width=0
            )

        st.plotly_chart(fig_line, use_container_width=True)

    # ── Faculty details (semestre seleccionado) ─────────────────────────
    st.markdown("### View Faculty details")

    active = df[df["Periodo"].astype(str).eq(sel_period_internal)].copy()

    st.markdown(
        f"<div style='text-align:center;font-size:34px;font-weight:bold'>Active Full-time Faculty {sel_period_label}</div>",
        unsafe_allow_html=True
    )
    total_act = active["ID"].nunique()

    c0, c1 = st.columns([1, 2])
    c0.markdown(f"<div style='text-align:right;font-size:56px;font-weight:bold'>{total_act}</div>",
               unsafe_allow_html=True)

    gen = active["Gender"].value_counts()
    df_gen = pd.DataFrame({
        "Gender": ["Male", "Female"],
        "P": [
            round(gen.get("Male", 0) / total_act * 100, 1) if total_act else 0,
            round(gen.get("Female", 0) / total_act * 100, 1) if total_act else 0
        ],
        "Bar": [" ", " "]
    })
    figG = px.bar(
        df_gen, x="P", y="Bar", color="Gender", text="Gender",
        color_discrete_map={"Male": "#003366", "Female": "#56d6c9"},
        orientation="h"
    )
    figG.update_traces(texttemplate='%{text} %{x}%', textposition="inside", textfont_size=20, width=0.7)
    figG.update_layout(showlegend=False, xaxis_visible=False, yaxis_visible=False,
                       height=100, margin=dict(l=10, r=10, t=10, b=10),
                       yaxis=dict(domain=[0.2, 1]))
    c1.plotly_chart(figG, use_container_width=True)

    # ── Full table ───────────────────────────────────────────────────────
    st.markdown("### Complete Full-time table")
    cols_full = [
        "ID Nr.", "ID", "First Name", "Last Name",
        "Date of First Appointment to the School", "Academic Area",
        "Highest Degree", "Year", "Region were degree was obtained",
        "International Degree", "% devoted to Mission", "Faculty Ranking",
        "Subcategorization",
        "Country of Birth", "Double Nationality", "Date of Birth",
        "Age", "Gender", "Faculty Qualific.", "P/S",
        "Normal professional Resp.", "Notes"
    ]
    full = active[[c for c in cols_full if c in active.columns]].copy().reset_index(drop=True)

    if "Date of First Appointment to the School" in full.columns:
        full["Date of First Appointment to the School"] = pd.to_datetime(
            full["Date of First Appointment to the School"], errors="coerce"
        ).dt.date
    if "Date of Birth" in full.columns:
        full["Date of Birth"] = pd.to_datetime(full["Date of Birth"], errors="coerce").dt.date
    if "Year" in full.columns:
        full["Year"] = full["Year"].astype(str).str.extract(r'(\d{4})')

    with st.expander("Show complete table"):
        _download_link("Descargar tabla completa (Excel)", full, f"FT_Complete_Table_{sel_period_label}.xlsx")
        full.index += 1
        st.dataframe(full, use_container_width=False)

    # ── Professor trajectory (PLANTA only) ──────────────────────────────
    def _c(df0, *names):
        if df0 is None or df0.empty:
            return None
        cmap = {str(c).strip().casefold(): c for c in df0.columns}
        for n in names:
            key = str(n).strip().casefold()
            if key in cmap:
                return cmap[key]
        return None

    period_col = _c(df, "Periodo")
    id_col = _c(df, "ID Nr.", "ID Nr", "ID")
    fn_col = _c(df, "First Name")
    ln_col = _c(df, "Last Name")
    area_col = _c(df, "Academic Area", "AREA_PROFESOR")
    deg_col = _c(df, "Highest Degree")
    rank_col = _c(df, "Faculty Ranking")
    subc_col = _c(df, "Subcategorization")
    age_col = _c(df, "Age")
    qual_col = _c(df, "Faculty Qualific.")
    ps_col = _c(df, "P/S", "P - S")
    resp_col = _c(df, "Normal professional Resp.")
    notes_col = _c(df, "Notes")

    vals = sorted(df[period_col].dropna().astype(str).unique().tolist()) if period_col else []
    last_period = vals[-1] if vals else ""
    st.markdown(
        f"<div style='font-size:18px;font-weight:700;margin-top:18px'>"
        f"Search professor trajectory (PLANTA only) 2020-10 – {last_period}"
        f"</div>",
        unsafe_allow_html=True
    )

    if not all([period_col, id_col, fn_col, ln_col]):
        st.warning("Required columns were not found in the PLANTA sheet.")
    else:
        df_ids = df[[id_col, fn_col, ln_col]].copy().dropna(subset=[id_col])
        df_ids = df_ids.drop_duplicates(subset=[id_col], keep="last")
        df_ids["label"] = (
            df_ids[fn_col].astype(str).str.strip() + " " +
            df_ids[ln_col].astype(str).str.strip() +
            " — ID: " + df_ids[id_col].astype(str).str.strip()
        )
        df_ids = df_ids.sort_values("label")

        sel_label = st.selectbox(
            "Select a professor (PLANTA):",
            options=["(Select...)"] + df_ids["label"].tolist(),
            index=0,
            key="ft_staff_trajectory_select"
        )

        if sel_label and sel_label != "(Select...)":
            m = re.search(r"ID:\s*(.+)$", sel_label)
            chosen_id = m.group(1).strip() if m else None

            traj = df[df[id_col].astype(str).str.strip() == chosen_id].copy()

            out_cols_raw = [
                (period_col, "Periodo"),
                (id_col, "ID Nr."),
                (fn_col, "First Name"),
                (ln_col, "Last Name"),
                (area_col, "Academic Area"),
                (deg_col, "Highest Degree"),
                (rank_col, "Faculty Ranking"),
                (subc_col, "Subcategorization"),
                (age_col, "Age"),
                (qual_col, "Faculty Qualific."),
                (ps_col, "P/S"),
                (resp_col, "Normal professional Resp."),
                (notes_col, "Notes"),
            ]

            out_df = pd.DataFrame({
                new: (traj[orig] if (orig in traj.columns) else pd.Series([""] * len(traj), index=traj.index))
                for orig, new in out_cols_raw
            })
            out_df.columns = pd.Index(out_df.columns).map(str)

            OUT_COLOR = "#8B0000"
            IN_COLOR = "#00796B"

            def _matches_tag_for_period(note_upper: str, tag: str, flat_period: str) -> bool:
                pat = rf'\b{tag}\s+IN\s+\(?((?:19|20)\d{{2}}[-_/ ]?\d{{2}})\)?\b'
                m2 = re.search(pat, note_upper, flags=re.IGNORECASE)
                if not m2:
                    return False
                per_txt = m2.group(1)
                per_flat = re.sub(r'\D', '', per_txt)
                return flat_period and (per_flat == flat_period)

            def _color_in_out(row: pd.Series):
                per = str(row.get("Periodo", ""))
                note_upper = str(row.get("Notes", "")).upper()
                flat = re.sub(r'\D', '', per)

                is_out = _matches_tag_for_period(note_upper, "OUT", flat) or ("OUT IN" in note_upper)
                is_in = _matches_tag_for_period(note_upper, "IN", flat) or ("IN IN" in note_upper)

                if is_out:
                    return [f'color:{OUT_COLOR};font-weight:700;' for _ in row.index]
                if is_in:
                    return [f'color:{IN_COLOR};font-weight:700;' for _ in row.index]
                return ['' for _ in row.index]

            st.dataframe(
                out_df.reset_index(drop=True).style.apply(_color_in_out, axis=1).hide(axis="index"),
                use_container_width=True
            )

            _download_link("Descargar trayectoria (Excel)", out_df, f"Trajectory_{chosen_id}.xlsx")


# ===========================================================================
# 7) NAVEGACIÓN MULTIPÁGINA (todo en un solo archivo)
# ===========================================================================
pages = [
    st.Page(page_composition, title="Composition", icon="🎓", url_path="composition", default=True),
    st.Page(page_staffing, title="Staffing Levels", icon="📊", url_path="staffing"),
]

pg = st.navigation(pages, position="top")
pg.run()
