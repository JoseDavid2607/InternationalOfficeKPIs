# ===========================================================================
#  UASM · Faculty Analytics Suite  ·  App multipágina (un solo archivo)
#  Página 1: Full-time Faculty Composition
#  Página 2: Full-time Faculty Staffing Levels
#  Página 3: Distribution by Academic Area
#  Página 4: Faculty Demographics
# ===========================================================================
from __future__ import annotations

import streamlit as st
import pandas as pd
import numpy as np
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
def _download_workbook_bytes() -> bytes:
    """Descarga el Google Sheet completo (una sola vez, cacheado 5 min) y
    devuelve los bytes del .xlsx. Todas las páginas leen de aquí para evitar
    descargas repetidas y minimizar el riesgo de timeouts."""
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"

    resp = None
    last_err = None
    # Reintenta hasta 3 veces: la primera descarga en frío contra Google a
    # veces redirige a un host googleusercontent.com que puede tardar.
    for attempt in range(3):
        try:
            resp = requests.get(url, timeout=60)
            if resp.status_code == 200 and "text/html" not in resp.headers.get("Content-Type", ""):
                return resp.content
            last_err = RuntimeError(f"HTTP {resp.status_code}")
        except Exception as e:
            last_err = e
        time.sleep(2)

    st.error(f"🌐 No se pudo conectar con Google Sheets tras varios intentos: {last_err}")
    st.stop()


@st.cache_data(ttl=300)
def load_data():
    raw = io.BytesIO(_download_workbook_bytes())
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


# ── Loaders específicos: página "Distribution by Area" ─────────────────────
# (mismo workbook, pero conservan exactamente la lógica original de esa página)
@st.cache_data(ttl=0)
def area_load_fulltime() -> pd.DataFrame:
    raw = io.BytesIO(_download_workbook_bytes())
    df_ = pd.read_excel(raw, sheet_name="BD_PLANTA")

    sem = df_["Semestre"].astype(str).str.strip() if "Semestre" in df_.columns else df_.iloc[:, 0].astype(str).str.strip()
    is_inter = sem.str.contains("inter", case=False, na=False)
    df_["Periodo"] = np.where(is_inter, sem.str[:4] + " Intersemestral", sem.str[:4] + "-" + sem.str[-2:])

    if "Academic Area" in df_.columns and "AREA_PROFESOR" not in df_.columns:
        df_["AREA_PROFESOR"] = df_["Academic Area"]
    if "ID Nr." in df_.columns and "ID" not in df_.columns:
        df_ = df_.rename(columns={"ID Nr.": "ID"})
    if "Full Name" not in df_.columns:
        fn = df_.get("First Name", "").astype(str).fillna("")
        ln = df_.get("Last Name", "").astype(str).fillna("")
        df_["Full Name"] = (fn + " " + ln).str.strip()
    return df_


@st.cache_data(ttl=0)
def area_load_parttime() -> pd.DataFrame:
    raw = io.BytesIO(_download_workbook_bytes())
    df_ = pd.read_excel(raw, sheet_name="Faculty Distribution")

    if "PLANTA_CATEDRA" in df_.columns:
        col = df_["PLANTA_CATEDRA"].astype(str).str.strip()
        col = col.str.normalize("NFKD").str.encode("ascii", errors="ignore").str.decode("ascii")
        df_ = df_[col.str.upper().eq("CATEDRA")].copy()

    sem = df_["Semestre"].astype(str).str.strip()
    is_inter = sem.str.contains("inter", case=False, na=False)
    df_.loc[~is_inter, "Periodo"] = sem.str[:4] + "-" + sem.str[-2:]
    df_.loc[is_inter, "Periodo"] = sem.str[:4] + " Intersemestral"

    if "ID Nr." in df_.columns and "ID" not in df_.columns:
        df_ = df_.rename(columns={"ID Nr.": "ID"})
    if "AREA_PROFESOR" not in df_.columns and "Academic Area" in df_.columns:
        df_["AREA_PROFESOR"] = df_["Academic Area"]
    return df_


# ── Loaders específicos: página "Demographics" ──────────────────────────────
# Nota: esta página usa un formato de Periodo sin guion ("YYYY10"/"YYYY Intersemestral"),
# distinto al de las demás páginas — se conserva igual que en el script original.
@st.cache_data(ttl=0)
def demo_load_fulltime() -> pd.DataFrame:
    raw = io.BytesIO(_download_workbook_bytes())
    df_ = pd.read_excel(raw, sheet_name="BD_PLANTA")

    if "Semestre" in df_.columns:
        sem = df_["Semestre"].astype(str).str.strip()
    else:
        sem = df_.iloc[:, 0].astype(str).str.strip()
    is_inter = sem.str.contains("inter", case=False, na=False)
    df_["Periodo"] = np.where(is_inter, sem.str[:4] + " Intersemestral", sem.str[:4] + sem.str[-2:])

    if "Academic Area" in df_.columns and "AREA_PROFESOR" not in df_.columns:
        df_["AREA_PROFESOR"] = df_["Academic Area"]
    if "ID Nr." in df_.columns and "ID" not in df_.columns:
        df_ = df_.rename(columns={"ID Nr.": "ID"})
    if "Full Name" not in df_.columns:
        fn = df_.get("First Name", "").astype(str).fillna("")
        ln = df_.get("Last Name", "").astype(str).fillna("")
        df_["Full Name"] = (fn + " " + ln).str.strip()
    return df_


@st.cache_data(ttl=0)
def demo_load_parttime() -> pd.DataFrame:
    raw = io.BytesIO(_download_workbook_bytes())
    df_ = pd.read_excel(raw, sheet_name="Faculty Distribution")

    if "PLANTA_CATEDRA" in df_.columns:
        col = df_["PLANTA_CATEDRA"].astype(str).str.strip()
        col = col.str.normalize("NFKD").str.encode("ascii", errors="ignore").str.decode("ascii")
        df_ = df_[col.str.upper().eq("CATEDRA")].copy()

    sem = df_["Semestre"].astype(str).str.strip()
    is_inter = sem.str.contains("inter", case=False, na=False)
    df_.loc[~is_inter, "Periodo"] = sem.str[:4] + sem.str[-2:]
    df_.loc[is_inter, "Periodo"] = sem.str[:4] + " Intersemestral"

    if "ID Nr." not in df_.columns and "ID" in df_.columns:
        df_ = df_.rename(columns={"ID": "ID Nr."})
    if "AREA_PROFESOR" not in df_.columns and "Academic Area" in df_.columns:
        df_["AREA_PROFESOR"] = df_["Academic Area"]
    return df_


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
# 7) PÁGINA 3 — Distribution by Academic Area
# ===========================================================================
def page_area():
    MINT = "#00A896"
    HIGHLIGHT = "#D0E5F5"
    PALETTE = [
        "#056D62", "#1CDFCB", "#FF7F50", "#9B59B6", "#F4A261",
        "#1B6CA8", "#0EAD69", "#E76F51", "#3D5A80", "#8D99AE",
        "#78A7A2", "#F6BD60", "#6D597A", "#43AA8B", "#277DA1",
    ]

    def is_sem_label(p: str) -> bool:
        return bool(re.fullmatch(r"\d{4}-(10|20)", str(p)))


    def is_inter_label(p: str) -> bool:
        return bool(re.fullmatch(r"\d{4}\s+Intersemestral", str(p)))


    def display_label_sem(p_internal: str) -> str:
        return str(p_internal).replace("-", "")


    def filter_for_timeframe(df_in: pd.DataFrame, time_mode: str, value) -> pd.DataFrame:
        """
        value:
          - Semestral: visible 'YYYY10'/'YYYY20', internamente 'YYYY-10/20'
          - Anual: 'YYYY' -> incluye ambos semestres + intersemestral, deduplicado por profesor/año
          - Intersemestral: 'YYYY Intersemestral'
        """
        if value is None:
            return df_in.iloc[0:0].copy()

        dfb = df_in.copy()

        if time_mode == "Semestral":
            sem_internal = f"{str(value)[:4]}-{str(value)[-2:]}"
            return dfb[dfb["Periodo"].astype(str).eq(sem_internal)].copy()

        if time_mode == "Anual":
            y = str(value)
            dfy = dfb[dfb["Periodo"].astype(str).str.startswith(y)].copy()
            if "ID" in dfy.columns:
                dfy["__Year"] = dfy["Periodo"].astype(str).str[:4]
                dfy = dfy.sort_values(by=["Periodo"]).drop_duplicates(subset=["ID", "__Year"], keep="last")
                dfy = dfy.drop(columns=["__Year"])
            return dfy

        return dfb[dfb["Periodo"].astype(str).eq(str(value))].copy()



    _render_header("Distribution by Academic Area", "Faculty distribution and evolution across academic areas")

    df_full = area_load_fulltime()
    df_part = area_load_parttime()

    st.session_state.setdefault("modo_faculty", "Full-time")


    # ── Sidebar ──────────────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("#### Faculty Type")
        st.markdown('<div id="mode-pill">', unsafe_allow_html=True)
        mode_sidebar = st.radio(
            "Mode", ["Full-time", "Part-time"],
            index=0 if st.session_state.modo_faculty == "Full-time" else 1,
            horizontal=True, label_visibility="collapsed", key="mode_pill_radio",
        )
        st.markdown('</div>', unsafe_allow_html=True)

        if mode_sidebar != st.session_state.modo_faculty:
            st.session_state.modo_faculty = mode_sidebar
            st.rerun()

        df_base = df_full if st.session_state.modo_faculty == "Full-time" else df_part

        st.markdown("#### Timeframe")
        tmode = st.radio("Timeframe", ["Semestral", "Anual", "Intersemestral"],
                          key="time_mode_side", label_visibility="collapsed")

        all_periods = sorted(df_base["Periodo"].astype(str).dropna().unique().tolist())

        if tmode == "Semestral":
            period_opts = [p for p in all_periods if is_sem_label(p)]
            visible_opts = [display_label_sem(p) for p in period_opts]
            default_idx = len(period_opts) - 1 if period_opts else 0
            sel_visible = st.selectbox("Periodo", visible_opts, index=default_idx if period_opts else None)
            sel_value = period_opts[visible_opts.index(sel_visible)] if period_opts else None
            sel_label = sel_visible
        elif tmode == "Anual":
            years = sorted(pd.Series(all_periods).astype(str).str[:4].unique().tolist())
            default_idx = len(years) - 1 if years else 0
            sel_value = st.selectbox("Periodo", years, index=default_idx if years else None)
            sel_label = sel_value
        else:
            inters = [p for p in all_periods if is_inter_label(p)]
            default_idx = len(inters) - 1 if inters else 0
            sel_value = st.selectbox("Periodo", inters, index=default_idx if inters else None)
            sel_label = sel_value

        st.session_state["sel_tf_mode"] = tmode
        st.session_state["sel_tf_value"] = sel_value
        st.session_state["sel_tf_label"] = sel_label

        export_df = filter_for_timeframe(df_base, tmode, sel_value)
        fname = f"{'FT' if st.session_state.modo_faculty == 'Full-time' else 'PT'}_{tmode}_{str(sel_label).replace(' ', '_')}.xlsx"
        b64_dl = base64.b64encode(_xlsx_bytes(export_df)).decode()
        st.markdown(
            f'<a class="modern-btn" download="{fname}" '
            f'href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_dl}">'
            f'⭳ Descargar Base — {st.session_state.modo_faculty} — {sel_label}</a>',
            unsafe_allow_html=True,
        )


    # ── Active dataset ───────────────────────────────────────────────────────────
    df = df_full.copy() if st.session_state.modo_faculty == "Full-time" else df_part.copy()

    tmode_now = st.session_state.get("sel_tf_mode", "Semestral")
    sel_value = st.session_state.get("sel_tf_value")
    sel_label = st.session_state.get("sel_tf_label")
    IDCOL = "ID"


    # ── Pivot table by academic area ────────────────────────────────────────────
    df_view = df.copy()

    if tmode_now == "Semestral":
        df_view = df_view[df_view["Periodo"].astype(str).apply(is_sem_label)].copy()
        df_view["Periodo_display"] = df_view["Periodo"].astype(str).map(display_label_sem)
        pivot_area = pd.pivot_table(df_view, index="AREA_PROFESOR", columns="Periodo_display",
                                     values=IDCOL, aggfunc="nunique", fill_value=0).sort_index()
        col_order = sorted(df_view["Periodo_display"].unique().tolist())

    elif tmode_now == "Intersemestral":
        df_view = df_view[df_view["Periodo"].astype(str).apply(is_inter_label)].copy()
        df_view["Periodo_display"] = df_view["Periodo"].astype(str)
        pivot_area = pd.pivot_table(df_view, index="AREA_PROFESOR", columns="Periodo_display",
                                     values=IDCOL, aggfunc="nunique", fill_value=0).sort_index()
        col_order = sorted(df_view["Periodo_display"].unique().tolist())

    else:  # Anual
        df_view["__Year"] = df_view["Periodo"].astype(str).str[:4]
        df_view = df_view.sort_values(by=["Periodo"]).drop_duplicates(subset=[IDCOL, "__Year"], keep="last")
        pivot_area = pd.pivot_table(df_view, index="AREA_PROFESOR", columns="__Year",
                                     values=IDCOL, aggfunc="nunique", fill_value=0).sort_index()
        col_order = sorted(pivot_area.columns.astype(str).tolist())

    areas_palette_order = [a for a in pivot_area.index if a != "Total"]
    color_map_area = {a: PALETTE[i % len(PALETTE)] for i, a in enumerate(areas_palette_order)}

    pivot_area.loc["Total"] = pivot_area.sum(numeric_only=True)


    def style_bold_total(df_):
        styles = pd.DataFrame('', index=df_.index, columns=df_.columns)
        if "Total" in df_.index:
            styles.loc["Total", :] = "font-weight:700;"
        return styles


    def style_total_lastcell(df_):
        styles = pd.DataFrame("", index=df_.index, columns=df_.columns)
        if len(df_.columns) > 0 and "Total" in df_.index:
            styles.loc["Total", df_.columns[-1]] = "background-color:#dff7f2; color:#00A896; font-weight:700;"
        return styles


    styled_area = (
        pivot_area[col_order].style
        .apply(style_bold_total, axis=None)
        .apply(style_total_lastcell, axis=None)
        .format(precision=0)
    )

    st.subheader(f"{st.session_state.modo_faculty} Faculty count")
    st.dataframe(styled_area, use_container_width=True)

    pivot_download = pivot_area[col_order].reset_index()
    fname_pvt = f"Pivot_{'FT' if st.session_state.modo_faculty == 'Full-time' else 'PT'}_{tmode_now}.xlsx"
    _download_link("Descargar tabla (Excel)", pivot_download, fname_pvt)


    # ── Charts: evolution line + donut ──────────────────────────────────────────
    st.markdown(f"### Evolution by Academic Area — Number of {st.session_state.modo_faculty} Faculty")

    mode_key = "ft" if st.session_state.modo_faculty == "Full-time" else "pt"
    show_all_key = f"ver_todas_{mode_key}"
    area_sel_key = f"area_sel_{mode_key}"

    st.session_state.setdefault(show_all_key, True)
    st.session_state.setdefault(area_sel_key, "Select...")

    if tmode_now == "Semestral":
        base_line = df[df["Periodo"].astype(str).apply(is_sem_label)].copy()
        base_line["X"] = base_line["Periodo"].astype(str).map(display_label_sem)
        x_to_filter = sel_label
    elif tmode_now == "Intersemestral":
        base_line = df[df["Periodo"].astype(str).apply(is_inter_label)].copy()
        base_line["X"] = base_line["Periodo"].astype(str)
        x_to_filter = sel_label
    else:  # Anual
        tmp = df.copy()
        tmp["__Year"] = tmp["Periodo"].astype(str).str[:4]
        tmp = tmp.sort_values(by=["Periodo"]).drop_duplicates(subset=[IDCOL, "__Year"], keep="last")
        base_line = tmp.rename(columns={"__Year": "X"}).copy()
        x_to_filter = str(sel_label)

    x_labels = sorted(base_line["X"].unique().tolist())

    totals_period = (
        base_line.groupby("X")[IDCOL].nunique()
        .reindex(x_labels).fillna(0).astype(int)
        .reset_index(name="Total")
    )


    def highlight_current(fig, labels, current):
        if current in labels:
            pos = labels.index(current)
            fig.add_shape(type="rect", xref="x", yref="paper", x0=pos - 0.4, x1=pos + 0.4, y0=0, y1=1,
                          fillcolor=HIGHLIGHT, opacity=0.35, line_width=0)
        return fig


    colL, colR = st.columns([3, 2])

    with colL:
        areas_current = sorted([a for a in pivot_area.index if a != "Total"])
        if st.session_state[area_sel_key] not in ["Select...", *areas_current]:
            st.session_state[area_sel_key] = "Select..."

        st.checkbox("Show all lines", key=show_all_key)
        st.selectbox("Select academic area:", ["Select...", *areas_current], key=area_sel_key)

        show_all = st.session_state[show_all_key]
        area_sel_val = st.session_state[area_sel_key]

        if show_all:
            df_counts = base_line.groupby(["X", "AREA_PROFESOR"])[IDCOL].nunique().reset_index(name="Count")
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
            fig_line.update_layout(showlegend=False)
            fig_line = highlight_current(fig_line, x_labels, x_to_filter)
            st.plotly_chart(fig_line, use_container_width=True)

        elif area_sel_val == "Select...":
            st.info("Select an academic area or enable 'Show all lines'.")

        else:
            df_area = (
                base_line[base_line["AREA_PROFESOR"] == area_sel_val]
                .groupby("X")[IDCOL].nunique()
                .reindex(x_labels).fillna(0).astype(int)
                .reset_index(name="Count")
            )
            df_line = df_area.merge(totals_period, on="X", how="left")
            df_line["Pct"] = (df_line["Count"] / df_line["Total"].replace(0, pd.NA)).fillna(0)
            df_line["X"] = pd.Categorical(df_line["X"], categories=x_labels, ordered=True)

            fig_line = px.line(
                df_line, x="X", y="Pct", markers=True, title=f"Evolution (% share) — {area_sel_val}",
                color_discrete_sequence=[color_map_area.get(area_sel_val, MINT)],
            )
            fig_line.update_traces(mode="lines+markers", line=dict(width=2),
                                    hovertemplate="<b>%{x}</b><br>%{y:.1%}<extra></extra>")
            fig_line.update_xaxes(type="category", categoryorder="array", categoryarray=x_labels, title=None)
            fig_line.update_yaxes(rangemode="tozero", tickformat=".0%", title=None)
            fig_line.update_layout(showlegend=False)
            fig_line = highlight_current(fig_line, x_labels, x_to_filter)
            st.plotly_chart(fig_line, use_container_width=True)

    with colR:
        st.markdown(f"##### Distribution by academic area — {sel_label or ''}")

        df_donut = filter_for_timeframe(df, tmode_now, sel_value)

        if df_donut.empty:
            st.info("No data for the selected period.")
        else:
            dist = df_donut.groupby("AREA_PROFESOR")[IDCOL].nunique().sort_values(ascending=False)
            donut_df = pd.DataFrame({"Area": dist.index, "Value": dist.values})

            fig_donut = px.pie(
                donut_df, names="Area", values="Value", hole=0.45, color="Area",
                color_discrete_map=color_map_area, category_orders={"Area": areas_palette_order},
            )
            st.plotly_chart(fig_donut, use_container_width=True)

            fname_donut = f"Donut_{'FT' if st.session_state.modo_faculty == 'Full-time' else 'PT'}_{tmode_now}_{str(sel_label).replace(' ', '_')}.xlsx"
            _download_link("Descargar tabla (Excel)", donut_df, fname_donut)


    # ── Detail table ─────────────────────────────────────────────────────────────
    detail = filter_for_timeframe(df, tmode_now, sel_value)

    cols_prefer_ft = ["Full Name", "AREA_PROFESOR", "Faculty Ranking", "Faculty Qualific.", "P/S"]
    cols_prefer_pt = ["Profesor", "AREA_PROFESOR", "PLANTA_CATEDRA", "TIPO", "P/S"]
    prefer_cols = cols_prefer_ft if st.session_state.modo_faculty == "Full-time" else cols_prefer_pt

    cols_to_show = [c for c in prefer_cols if c in detail.columns]
    detail_out = detail[cols_to_show].drop_duplicates().reset_index(drop=True)

    count_label = int(detail[IDCOL].nunique()) if IDCOL in detail.columns else len(detail_out)
    faculty_word = "full-time" if st.session_state.modo_faculty == "Full-time" else "part-time"
    st.markdown(f"### There are {count_label} {faculty_word} Faculty in **{sel_label}**")
    st.dataframe(detail_out, use_container_width=True)

    fname_det = f"Detail_{'FT' if st.session_state.modo_faculty == 'Full-time' else 'PT'}_{tmode_now}_{str(sel_label).replace(' ', '_')}.xlsx"
    _download_link("Descargar tabla (Excel)", detail_out, fname_det)


# ===========================================================================
# 8) PÁGINA 4 — Faculty Demographics
# ===========================================================================
def page_demographics():
    COLORS = {
        "primary": "#21877D", "primary_dark": "#004d47", "primary_light": "#dff7f2",
        "accent1": "#2EC4B6", "accent2": "#00A896", "accent3": "#56D6C9",
        "highlight": "#D0E5F5",
    }

    _render_header("Faculty Demographics", "PhD attainment, international diversity, and composition over time")

    df_full = demo_load_fulltime()
    df_part = demo_load_parttime()

    # ── Timeframe helpers ──────────────────────────────────────────────────────
    def col_id(df_: pd.DataFrame) -> str | None:
        return "ID Nr." if "ID Nr." in df_.columns else ("ID" if "ID" in df_.columns else None)


    def col_degree(df_: pd.DataFrame) -> str | None:
        return next((c for c in ["Highest Degree", "TÍTULO"] if c in df_.columns), None)


    def col_gender(df_: pd.DataFrame) -> str | None:
        return next((c for c in ["Gender", "GÉNERO"] if c in df_.columns), None)


    def col_nationality(df_: pd.DataFrame) -> str | None:
        return next((c for c in ["Country of Birth", "Nationality"] if c in df_.columns), None)


    def is_semester_label(p: str) -> bool:
        return bool(re.fullmatch(r"\d{4}(10|20)", str(p)))


    def is_inter_label(p: str) -> bool:
        return bool(re.fullmatch(r"\d{4}\s+Intersemestral", str(p)))


    def normalize_degree(series: pd.Series) -> pd.Series:
        s = series.astype(str).str.strip()
        is_tbd = s.str.upper().eq("TBD") | s.eq("") | s.str.lower().isin(["na", "none"])
        s_norm = s.str.lower().str.replace(".", "", regex=False)
        s_norm = s_norm.str.normalize("NFKD").str.encode("ascii", errors="ignore").str.decode("ascii")
        is_phd = s_norm.str.contains(r"\bphd\b") | s_norm.str.contains("doctor")
        is_master = s_norm.str.contains("master") | s_norm.str.contains(r"\bmsc\b") | s_norm.str.contains(r"\bms\b")
        is_bachelor = (s_norm.str.contains("bachelor") | s_norm.str.contains(r"\bbsc\b")
                       | s_norm.str.contains(r"\bbs\b") | s_norm.str.contains(r"\bba\b")
                       | s_norm.str.contains("licen"))
        out = pd.Series("Other", index=s.index, dtype=object)
        out[is_tbd] = "TBD"
        out[~is_tbd & is_phd] = "PhD"
        out[~is_tbd & ~is_phd & is_master] = "Master"
        out[~is_tbd & ~is_phd & ~is_master & is_bachelor] = "Bachelor"
        return out


    def filter_for_timeframe(df_in: pd.DataFrame, time_mode: str, sel_sem: str | None = None,
                              sel_year: int | None = None, sel_inter_label: str | None = None) -> pd.DataFrame:
        dfb = df_in.copy()
        pcol = "Periodo" if "Periodo" in dfb.columns else None
        scol = "Semestre" if "Semestre" in dfb.columns else None
        idc = col_id(dfb)

        if time_mode == "Semestral" and sel_sem:
            target = str(sel_sem)
            mask = pd.Series(False, index=dfb.index)
            if pcol:
                mask |= dfb[pcol].astype(str).eq(target)
            if scol:
                mask |= dfb[scol].astype(str).str.replace("-", "", regex=False).str.fullmatch(target, na=False)
            dfb = dfb[mask].copy()

        elif time_mode == "Anual" and sel_year is not None:
            y = str(sel_year)
            mask = pd.Series(False, index=dfb.index)
            if pcol:
                mask |= dfb[pcol].astype(str).str.startswith(y)
            if scol:
                mask |= dfb[scol].astype(str).str.startswith(y)
            dfb = dfb[mask].copy()
            if idc:
                sort_key = pcol or scol
                if sort_key:
                    dfb = dfb.sort_values(by=[sort_key])
                dfb = dfb.drop_duplicates(subset=[idc], keep="last")

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


    def options_for_timeframe(df_src: pd.DataFrame, time_mode: str) -> list:
        per = df_src["Periodo"].dropna().astype(str)
        if time_mode == "Semestral":
            return sorted([p for p in per.unique() if is_semester_label(p)])
        if time_mode == "Intersemestral":
            return sorted([p for p in per.unique() if is_inter_label(p)])
        return sorted(per.str[:4].unique().tolist())


    def build_time_series(df_src: pd.DataFrame, time_mode: str, idcol: str,
                           degree_col: str | None, nat_col: str | None):
        labels, phd_pct, intl_pct = [], [], []

        def _phd_pct(sub):
            if sub.empty or degree_col is None:
                return 0.0
            sub = sub.copy()
            sub["__deg"] = normalize_degree(sub[degree_col])
            tot = sub[idcol].nunique()
            return 0.0 if tot == 0 else round(100 * sub.loc[sub["__deg"] == "PhD", idcol].nunique() / tot, 1)

        def _intl_pct(sub):
            if sub.empty or nat_col is None:
                return 0.0
            nat = sub[nat_col].astype(str).str.strip()
            is_valid = ~nat.eq("Colombian") & ~nat.str.upper().eq("TBD") & ~nat.eq("")
            tot = sub[idcol].nunique()
            return 0.0 if tot == 0 else round(100 * sub.loc[is_valid, idcol].nunique() / tot, 1)

        if time_mode in ("Semestral", "Intersemestral"):
            label_filter = is_semester_label if time_mode == "Semestral" else is_inter_label
            periods = sorted([p for p in df_src["Periodo"].dropna().astype(str).unique() if label_filter(p)])
            for p in periods:
                sub = df_src[df_src["Periodo"].astype(str).eq(p)]
                labels.append(p)
                phd_pct.append(_phd_pct(sub))
                intl_pct.append(_intl_pct(sub))
        else:
            years = sorted(df_src["Periodo"].dropna().astype(str).str[:4].unique())
            for y in years:
                sub = filter_for_timeframe(df_src, "Anual", sel_year=int(y))
                labels.append(y)
                phd_pct.append(_phd_pct(sub))
                intl_pct.append(_intl_pct(sub))

        return labels, phd_pct, intl_pct


    AGE_LABELS = ["Under 30", "31-40", "41-50", "51-60", "over 61"]
    AGE_BINS = [-np.inf, 29, 40, 50, 60, np.inf]


    def age_buckets(series: pd.Series) -> pd.Series:
        return pd.cut(pd.to_numeric(series, errors="coerce"), bins=AGE_BINS, labels=AGE_LABELS)


    # ── Session defaults ────────────────────────────────────────────────────────
    st.session_state.setdefault("modo_faculty", "Full-time")
    st.session_state.setdefault("time_mode_side", "Semestral")
    st.session_state.setdefault("sel_tf_label", None)


    # ── Sidebar ─────────────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("#### Faculty Type")
        st.markdown('<div id="mode-pill">', unsafe_allow_html=True)
        mode_sidebar = st.radio(
            "Mode", ["Full-time", "Part-time"],
            index=0 if st.session_state.modo_faculty == "Full-time" else 1,
            horizontal=True, label_visibility="collapsed", key="mode_pill_radio",
        )
        if mode_sidebar != st.session_state.modo_faculty:
            st.session_state.modo_faculty = mode_sidebar
            st.session_state.sel_tf_label = None
            st.rerun()

        st.markdown("#### Timeframe")
        time_mode_side = st.radio(
            "Timeframe", ["Semestral", "Anual", "Intersemestral"],
            key="time_mode_side", label_visibility="collapsed",
        )

        df_base = df_full if st.session_state.modo_faculty == "Full-time" else df_part
        options_tf = options_for_timeframe(df_base, time_mode_side)

        sel_label = st.selectbox(
            "Periodo", options_tf,
            index=(options_tf.index(st.session_state.sel_tf_label)
                   if st.session_state.sel_tf_label in options_tf else len(options_tf) - 1),
        ) if options_tf else None

        if sel_label != st.session_state.get("sel_tf_label"):
            st.session_state.sel_tf_label = sel_label

        if sel_label:
            if time_mode_side == "Semestral":
                export_df = filter_for_timeframe(df_base, "Semestral", sel_sem=sel_label)
            elif time_mode_side == "Anual":
                export_df = filter_for_timeframe(df_base, "Anual", sel_year=int(sel_label))
            else:
                export_df = filter_for_timeframe(df_base, "Intersemestral", sel_inter_label=sel_label)
            label_time = sel_label
        else:
            export_df = df_base.iloc[0:0].copy()
            label_time = ""

        modo_dl = st.session_state.get("modo_faculty", "Full-time")
        fname = f"{'FT' if modo_dl == 'Full-time' else 'PT'}_{time_mode_side}_{str(label_time).replace(' ', '_')}.xlsx"
        b64_dl = base64.b64encode(_xlsx_bytes(export_df)).decode()
        st.markdown(
            f'<a class="modern-btn" download="{fname}" '
            f'href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_dl}">'
            f'⭳ Descargar Base — {modo_dl} — {label_time}</a>',
            unsafe_allow_html=True,
        )


    # ── Active dataset ──────────────────────────────────────────────────────────
    mode_now = st.session_state.get("modo_faculty", "Full-time")
    df = (df_full if mode_now == "Full-time" else df_part).copy()
    if "ID Nr." not in df.columns and "ID" in df.columns:
        df["ID Nr."] = df["ID"]

    sel_period_text = st.session_state.get("sel_tf_label") or ""
    st.subheader("Full-time demographics by Faculty ranking" if mode_now == "Full-time" else "Part-time demographic table")
    if sel_period_text and mode_now == "Full-time":
        st.markdown(f"<div class='period-label'>{sel_period_text}</div>", unsafe_allow_html=True)

    col_table, col_side = st.columns([3, 1.2])


    # ── Main table ──────────────────────────────────────────────────────────────
    with col_table:
        IDCOL = col_id(df)
        if not IDCOL:
            st.error("ID column not found.")
            st.stop()

        tmode = st.session_state.get("time_mode_side", "Semestral")
        sel_lbl = st.session_state.get("sel_tf_label")
        GROUPS = {"Highest Degree", "Nationality", "Gender", "Age"}

        if not sel_lbl and mode_now == "Full-time":
            st.info("No data for the selected mode/period.")

        elif mode_now == "Full-time":
            if tmode == "Semestral":
                active = filter_for_timeframe(df, "Semestral", sel_sem=sel_lbl)
            elif tmode == "Anual":
                active = filter_for_timeframe(df, "Anual", sel_year=int(sel_lbl))
            else:
                active = filter_for_timeframe(df, "Intersemestral", sel_inter_label=sel_lbl)

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
            dcol = col_degree(active)
            if dcol and not active.empty:
                active["Degree_norm"] = normalize_degree(active[dcol])
                rows.append(pd.Series({"Category": "Highest Degree",
                                        **counts_by_ranking(active[active["Degree_norm"].isin(["PhD", "Master", "Bachelor"])]).to_dict()}))
                for d in ["PhD", "Master", "Bachelor"]:
                    rows.append(pd.Series({"Category": d, **counts_by_ranking(active[active["Degree_norm"] == d]).to_dict()}))

            ncol = col_nationality(active)
            if ncol and not active.empty:
                rows.append(pd.Series({"Category": "Nationality", **counts_by_ranking(active).to_dict()}))
                rows.append(pd.Series({"Category": "Colombian", **counts_by_ranking(active[active[ncol].astype(str).eq("Colombian")]).to_dict()}))
                rows.append(pd.Series({"Category": "International", **counts_by_ranking(active[~active[ncol].astype(str).eq("Colombian")]).to_dict()}))

            gcol = col_gender(active)
            if gcol and not active.empty:
                rows.append(pd.Series({"Category": "Gender", **counts_by_ranking(active[active[gcol].astype(str).isin(["Male", "Female"])]).to_dict()}))
                for g in ["Male", "Female"]:
                    rows.append(pd.Series({"Category": g, **counts_by_ranking(active[active[gcol].astype(str) == g]).to_dict()}))

            if not active.empty and "Age" in active.columns:
                active["Age_bucket"] = age_buckets(active["Age"])
                rows.append(pd.Series({"Category": "Age", **counts_by_ranking(active[active["Age_bucket"].notna()]).to_dict()}))
                for b in AGE_LABELS:
                    rows.append(pd.Series({"Category": b, **counts_by_ranking(active[active["Age_bucket"] == b]).to_dict()}))

            table_df = pd.DataFrame(rows).reindex(columns=cols_out).fillna(0)

            if not table_df.empty:
                numeric_cols = [c for c in (ranking_order + ["Total"]) if c in table_df.columns]
                is_group = table_df["Category"].isin(GROUPS)
                all_zero = (table_df[numeric_cols].sum(axis=1) == 0) if numeric_cols else pd.Series(False, index=table_df.index)
                table_df = table_df.loc[~(all_zero & ~is_group)].copy()
                for c in numeric_cols:
                    table_df[c] = pd.to_numeric(table_df[c], errors="coerce").fillna(0).astype(int)

            if table_df.empty:
                st.info("No rows to display for this selection.")
            else:
                mint_light = "#dff7f2"
                mint_dark = "#004d47"

                def style_group_rows(df_):
                    styles = pd.DataFrame('', index=df_.index, columns=df_.columns)
                    mask = df_["Category"].isin(GROUPS)
                    styles.loc[mask, ["Category"] + ranking_order + ["Total"]] = 'background-color:#f2f2f2;'
                    styles.loc[mask, "Category"] += 'font-weight:700;'
                    return styles

                def style_total_col(df_):
                    styles = pd.DataFrame('', index=df_.index, columns=df_.columns)
                    styles.loc[:, "Total"] = 'font-weight:700;'
                    return styles

                def style_group_totals(df_):
                    styles = pd.DataFrame('', index=df_.index, columns=df_.columns)
                    mask = df_["Category"].isin(GROUPS)
                    styles.loc[mask, "Total"] = f'background-color:{mint_light}; color:{mint_dark}; font-weight:800;'
                    return styles

                styled_table = (
                    table_df.style
                    .apply(style_group_rows, axis=None)
                    .apply(style_total_col, axis=None)
                    .apply(style_group_totals, axis=None)
                    .format(precision=0, na_rep="")
                    .hide(axis="index")
                )
                st.dataframe(styled_table, use_container_width=True, height=48 + 33 * (len(table_df) + 1))

        else:  # Part-time
            active = df.copy()
            if tmode == "Semestral":
                active = active[active["Periodo"].astype(str).apply(is_semester_label)].copy()
                keys = sorted(active["Periodo"].dropna().astype(str).unique().tolist())
            elif tmode == "Intersemestral":
                active = active[active["Periodo"].astype(str).apply(is_inter_label)].copy()
                keys = sorted(active["Periodo"].dropna().astype(str).unique().tolist())
            else:
                active["__Year"] = active["Periodo"].astype(str).str[:4]
                active = active.sort_values(by=["Periodo"]).drop_duplicates(subset=[IDCOL, "__Year"], keep="last")
                keys = sorted(active["__Year"].dropna().astype(str).unique().tolist())

            def counts_by_key(df_sub: pd.DataFrame) -> pd.Series:
                group_col = "__Year" if tmode == "Anual" else "Periodo"
                return df_sub.groupby(group_col)[IDCOL].nunique().reindex(keys, fill_value=0).astype(int)

            deg_col = "TÍTULO" if "TÍTULO" in active.columns else ("Highest Degree" if "Highest Degree" in active.columns else None)
            nat_col = "Nationality" if "Nationality" in active.columns else ("Country of Birth" if "Country of Birth" in active.columns else None)
            gen_col = "GÉNERO" if "GÉNERO" in active.columns else ("Gender" if "Gender" in active.columns else None)

            rows = []
            GROUPS_PT = {"Highest Degree", "Nationality", "Gender", "Age"}

            if deg_col and not active.empty:
                active["Degree_norm"] = normalize_degree(active[deg_col])
                rows.append(pd.Series({"Category": "Highest Degree", **counts_by_key(active).to_dict()}))
                for d in ["PhD", "Master", "Bachelor", "TBD"]:
                    rows.append(pd.Series({"Category": d, **counts_by_key(active[active["Degree_norm"] == d]).to_dict()}))

            if nat_col and not active.empty:
                nat = active[nat_col].astype(str).str.strip()
                is_tbd_nat = nat.str.upper().eq("TBD")
                is_col = nat.eq("Colombian")
                is_int = ~is_col & ~is_tbd_nat & nat.ne("")
                rows.append(pd.Series({"Category": "Nationality", **counts_by_key(active).to_dict()}))
                rows.append(pd.Series({"Category": "Colombian", **counts_by_key(active[is_col]).to_dict()}))
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
                active = active.assign(Age_bucket=age_buckets(active["Age"]))
                is_tbd_age = active["Age_bucket"].isna()
                rows.append(pd.Series({"Category": "Age", **counts_by_key(active).to_dict()}))
                for b in AGE_LABELS:
                    rows.append(pd.Series({"Category": b, **counts_by_key(active[active["Age_bucket"] == b]).to_dict()}))
                rows.append(pd.Series({"Category": "TBD (Age)", **counts_by_key(active[is_tbd_age]).to_dict()}))

            ycol = "Years Industry experience"
            if ycol in active.columns:
                yseries = pd.to_numeric(active[ycol], errors="coerce")
                group_col = "__Year" if tmode == "Anual" else "Periodo"
                avg_series = active.assign(**{ycol: yseries}).groupby(group_col)[ycol].mean().reindex(keys).round(1)
                rows.append(pd.Series({"Category": "Avg years of Work Exp.", **avg_series.to_dict()}))
                rows.append(pd.Series({"Category": "TBD (Work Exp.)", **counts_by_key(active[active[ycol].isna()]).to_dict()}))

            cols_out = ["Category"] + keys
            table_df = pd.DataFrame(rows).reindex(columns=cols_out).fillna(0)

            if not table_df.empty:
                is_group_or_avg = table_df["Category"].isin(GROUPS_PT) | table_df["Category"].eq("Avg years of Work Exp.")
                all_zero = table_df[keys].sum(axis=1) == 0
                table_df = table_df.loc[~(all_zero & ~is_group_or_avg)].copy()

            if table_df.empty:
                st.info("No rows to display.")
                st.stop()

            mask_avg = table_df["Category"].eq("Avg years of Work Exp.")
            display_df = table_df.copy().astype(object)
            for c in keys:
                display_df.loc[mask_avg, c] = pd.to_numeric(table_df.loc[mask_avg, c], errors="coerce").map(
                    lambda x: "" if pd.isna(x) else f"{x:.1f}")
                display_df.loc[~mask_avg, c] = pd.to_numeric(table_df.loc[~mask_avg, c], errors="coerce").fillna(0).astype(int).map(str)

            blue_light, blue_dark = "#e6f0fb", "#184a90"
            red_light, red_dark = "#f8d7da", "#721c24"

            def style_gray(df_):
                styles = pd.DataFrame('', index=df_.index, columns=df_.columns)
                gray_rows = df_["Category"].isin(GROUPS_PT) | df_["Category"].eq("Avg years of Work Exp.")
                styles.loc[gray_rows, ["Category"] + keys] = 'background-color:#f2f2f2; font-weight:700;'
                return styles

            def style_last_col(df_):
                styles = pd.DataFrame('', index=df_.index, columns=df_.columns)
                if keys:
                    mask = df_["Category"].isin(GROUPS_PT)
                    styles.loc[mask, keys[-1]] = f'background-color:{blue_light}; color:{blue_dark}; font-weight:800;'
                return styles

            def style_tbd(df_):
                styles = pd.DataFrame('', index=df_.index, columns=df_.columns)
                for i, row in df_.iterrows():
                    if "TBD" in str(row["Category"]):
                        for c in keys:
                            v = pd.to_numeric(table_df.loc[i, c], errors="coerce")
                            if pd.notna(v) and v > 0:
                                styles.at[i, c] = f'background-color:{red_light}; color:{red_dark}; font-weight:800;'
                return styles

            styled_table = (
                display_df.style
                .apply(style_gray, axis=None)
                .apply(style_last_col, axis=None)
                .apply(style_tbd, axis=None)
                .hide(axis="index")
            )
            st.dataframe(styled_table, use_container_width=True, height=48 + 33 * (len(display_df) + 1))


    # ── Side KPI charts ─────────────────────────────────────────────────────────
    with col_side:
        IDCOL = col_id(df)
        if not IDCOL:
            st.stop()

        sel_lbl = st.session_state.get("sel_tf_label")
        tmode = st.session_state.get("time_mode_side", "Semestral")

        if not sel_lbl:
            st.info("Select a period.")
        else:
            st.markdown(f"<div class='period-label'>{sel_lbl}</div>", unsafe_allow_html=True)

            if tmode == "Semestral":
                active_side = filter_for_timeframe(df, "Semestral", sel_sem=sel_lbl)
            elif tmode == "Anual":
                active_side = filter_for_timeframe(df, "Anual", sel_year=int(sel_lbl))
            else:
                active_side = filter_for_timeframe(df, "Intersemestral", sel_inter_label=sel_lbl)

            total_act = int(active_side[IDCOL].nunique()) if not active_side.empty else 0
            mint, mint_dark, gauge_bg = "#00A896", "#004d47", "#E8FAF7"

            def make_gauge(pct: float, label: str):
                fig = go.Figure(go.Indicator(
                    mode="gauge", value=pct,
                    gauge={'axis': {'range': [0, 100]}, 'bar': {'color': mint}, 'bgcolor': gauge_bg},
                ))
                fig.add_annotation(x=0.5, y=0.40, xref="paper", yref="paper", text=label,
                                    showarrow=False, font=dict(color=mint_dark, size=13))
                fig.add_annotation(x=0.5, y=0.0, xref="paper", yref="paper", text=f"{pct:.1f}%",
                                    showarrow=False, font=dict(color=mint_dark, size=18))
                fig.update_layout(height=110, margin=dict(l=10, r=10, t=10, b=6))
                return fig

            dcol = col_degree(active_side)
            phds = 0
            if total_act and dcol:
                active_side["Degree_norm"] = normalize_degree(active_side[dcol])
                phds = int(active_side[active_side["Degree_norm"] == "PhD"][IDCOL].nunique())
            pct_phd = round(100 * phds / total_act, 1) if total_act else 0.0
            st.plotly_chart(make_gauge(pct_phd, "PhD%:"), use_container_width=True)

            ncol = col_nationality(active_side)
            pct_int = 0.0
            if total_act and ncol:
                nat = active_side[ncol].astype(str).str.strip()
                is_int = ~nat.eq("Colombian") & ~nat.str.upper().eq("TBD") & ~nat.eq("")
                pct_int = round(100 * int(active_side[is_int][IDCOL].nunique()) / total_act, 1)
            st.plotly_chart(make_gauge(pct_int, "International%:"), use_container_width=True)

            gcol = col_gender(active_side)
            male = int(active_side[active_side[gcol].astype(str) == "Male"][IDCOL].nunique()) if gcol and total_act else 0
            female = int(active_side[active_side[gcol].astype(str) == "Female"][IDCOL].nunique()) if gcol and total_act else 0
            pct_m = round(100 * male / total_act, 1) if total_act else 0.0
            pct_f = round(100 * female / total_act, 1) if total_act else 0.0

            fig_gender = go.Figure()
            fig_gender.add_trace(go.Bar(x=[pct_m], y=[" "], orientation="h", name="Male",
                                         text=[f"Male {pct_m}%"], textposition="inside", insidetextanchor="middle"))
            fig_gender.add_trace(go.Bar(x=[pct_f], y=[" "], orientation="h", name="Female",
                                         text=[f"Female {pct_f}%"], textposition="inside", insidetextanchor="middle"))
            fig_gender.update_layout(barmode="stack", showlegend=False,
                                      xaxis=dict(range=[0, 100], visible=False), yaxis=dict(visible=False),
                                      height=100, margin=dict(l=10, r=10, t=18, b=12))
            st.plotly_chart(fig_gender, use_container_width=True)

            if not active_side.empty and "Age" in active_side.columns:
                active_side = active_side.assign(Age_bucket=age_buckets(active_side["Age"]))
                age_counts = (active_side.groupby("Age_bucket")[IDCOL].nunique()
                              .reindex(AGE_LABELS, fill_value=0).reset_index(name="Count"))
            else:
                age_counts = pd.DataFrame({"Age_bucket": AGE_LABELS, "Count": [0] * len(AGE_LABELS)})

            fig_age = px.bar(age_counts, x="Count", y="Age_bucket", orientation="h", text="Count")
            fig_age.update_traces(marker_color="#00A896", textposition="outside", texttemplate="%{text}")
            fig_age.update_xaxes(range=[0, 35], title=None)
            fig_age.update_yaxes(title=None, autorange="reversed")
            fig_age.update_layout(height=200, margin=dict(l=10, r=10, t=0, b=12))
            st.plotly_chart(fig_age, use_container_width=True)


    # ── Row 1: % PhD over time + PhD by region ─────────────────────────────────
    st.markdown("---")

    tmode_ts = st.session_state.get("time_mode_side", "Semestral")
    IDCOL = col_id(df)
    labels_ts, phd_ts, intl_ts = build_time_series(df, tmode_ts, IDCOL, col_degree(df), col_nationality(df))

    sel_lbl = st.session_state.get("sel_tf_label")
    if labels_ts:
        if tmode_ts == "Anual":
            period_current = sel_lbl if sel_lbl in labels_ts else labels_ts[-1]
        elif tmode_ts == "Intersemestral":
            period_current = sel_lbl if (sel_lbl in labels_ts and is_inter_label(sel_lbl)) else labels_ts[-1]
        else:
            period_current = sel_lbl if (sel_lbl in labels_ts and is_semester_label(sel_lbl)) else labels_ts[-1]
    else:
        period_current = None

    row1_left, row1_right = st.columns([6, 4])

    if mode_now == "Part-time":
        y_min_phd, y_max_phd, line_h, bar_h = 0, 30, 280, 220
    else:
        y_min_phd, y_max_phd, line_h, bar_h = 70, 100, 280, 220
    if tmode_ts == "Intersemestral":
        y_min_phd, y_max_phd = 0, 100


    def highlight_current(fig, labels, current):
        if current in labels:
            pos = labels.index(current)
            fig.add_shape(type="rect", xref="x", yref="paper", x0=pos - 0.4, x1=pos + 0.4, y0=0, y1=1,
                          fillcolor=COLORS["highlight"], opacity=0.35, line_width=0)
        return fig


    with row1_left:
        df_pct_phd = pd.DataFrame({"Label": labels_ts, "Percent": phd_ts})
        title_phd = "% of Full-time Faculty with PhD" if mode_now == "Full-time" else "% of Part-time Professors with PhD"
        fig_phd = px.line(df_pct_phd, x="Label", y="Percent", markers=True, text="Percent", title=title_phd)
        fig_phd.update_traces(line=dict(color="#00A896", width=3), marker=dict(size=7, color="#00A896"),
                              texttemplate="%{y:.1f}%", textposition="top center")
        fig_phd.update_xaxes(type="category", categoryorder="array", categoryarray=labels_ts, tickangle=0, title=None)
        fig_phd.update_yaxes(range=[y_min_phd, y_max_phd], title=None)
        fig_phd = highlight_current(fig_phd, labels_ts, period_current)
        fig_phd.update_layout(height=line_h, margin=dict(l=10, r=10, t=40, b=40), showlegend=False)
        st.plotly_chart(fig_phd, use_container_width=True)

    with row1_right:
        if period_current is None:
            active_p = df.iloc[0:0].copy()
        elif tmode_ts == "Anual":
            active_p = filter_for_timeframe(df, "Anual", sel_year=int(period_current))
        else:
            active_p = df[df["Periodo"].astype(str).eq(str(period_current))].copy()

        dcol_here = col_degree(df)
        phd_for_regions = pd.DataFrame(columns=df.columns)
        if dcol_here is not None and not active_p.empty:
            active_p["Degree_norm"] = normalize_degree(active_p[dcol_here])
            phd_now_all = active_p[active_p["Degree_norm"].eq("PhD")].copy()
            region_col = "Region were degree was obtained" if "Region were degree was obtained" in phd_now_all.columns else None
            if region_col:
                reg = phd_now_all[region_col].astype(str).str.strip()
                mask_valid_region = ~reg.eq("") & ~reg.str.upper().eq("TBD")
                phd_for_regions = phd_now_all[mask_valid_region].copy()

        total_phd_valid = int(phd_for_regions[IDCOL].nunique()) if not phd_for_regions.empty else 0
        phd_int = 0
        if "International Degree" in phd_for_regions.columns:
            phd_int = int(phd_for_regions[phd_for_regions["International Degree"].astype(str).str.strip().str.lower().eq("yes")][IDCOL].nunique())

        if not phd_for_regions.empty and "Region were degree was obtained" in phd_for_regions.columns:
            reg_counts = (phd_for_regions.groupby("Region were degree was obtained")[IDCOL].nunique()
                          .sort_values(ascending=False).reset_index()
                          .rename(columns={"Region were degree was obtained": "Region", IDCOL: "Count"}))
        else:
            reg_counts = pd.DataFrame({"Region": [], "Count": []})

        title_phd_bar = f"{total_phd_valid} professors with a PhD, {phd_int} obtained it abroad" if phd_int else f"{total_phd_valid} professors with a PhD"
        fig_phd_reg = px.bar(reg_counts, x="Count", y="Region", orientation="h", title=title_phd_bar, text="Count")
        fig_phd_reg.update_traces(marker_color="#00A896", textposition="outside", texttemplate="%{text}")
        fig_phd_reg.update_xaxes(title=None, dtick=1)
        fig_phd_reg.update_yaxes(title=None, autorange="reversed")
        fig_phd_reg.update_layout(height=bar_h, margin=dict(l=10, r=10, t=50, b=6))
        st.plotly_chart(fig_phd_reg, use_container_width=True)

        def pick_cols(df_, mapping):
            out = {}
            for new, opts in mapping.items():
                match = next((c for c in opts if c in df_.columns), None)
                out[new] = df_[match] if match else pd.Series([""] * len(df_), index=df_.index)
            return pd.DataFrame(out)

        if not phd_for_regions.empty:
            detalle_phd = pick_cols(phd_for_regions, {
                "Full Name": ["Full Name", "Full-Name", "Full_Name", "Profesor", "First Name"],
                "Highest Earned Degree": ["Highest Earned Degree", "Highest Degree", "TÍTULO"],
                "University": ["University", "University Name"],
                "Region were degree was obtained": ["Region were degree was obtained", "Region"],
                "Year": ["Year", "Year Earned ", "Year Degree", "Year Earned"],
            })
            popover = st.popover if hasattr(st, "popover") else st.expander
            with popover("🔎 Ver detalle de profesores con PhD"):
                st.dataframe(detalle_phd.reset_index(drop=True), use_container_width=True)


    # ── Row 2: % International over time + nationalities ───────────────────────
    st.markdown("---")
    row2_left, row2_right = st.columns([6, 4])

    if mode_now == "Part-time":
        y_min_int, y_max_int, line_h2, bar_h2 = 0, 10, 260, 220
    else:
        y_min_int, y_max_int, line_h2, bar_h2 = 0, 40, 350, 300
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
        fig_int = highlight_current(fig_int, labels_ts, period_current)
        fig_int.update_layout(height=line_h2, margin=dict(l=10, r=10, t=40, b=40), showlegend=False)
        st.plotly_chart(fig_int, use_container_width=True)

    with row2_right:
        nat_col = col_nationality(df)
        intl_now = pd.DataFrame(columns=df.columns)
        if nat_col and period_current:
            if tmode_ts == "Anual":
                active_p2 = filter_for_timeframe(df, "Anual", sel_year=int(period_current))
            else:
                active_p2 = df[df["Periodo"].astype(str).eq(str(period_current))].copy()

            nat = active_p2[nat_col].astype(str).str.strip()
            is_valid = ~nat.eq("Colombian") & ~nat.str.upper().eq("TBD") & ~nat.eq("")
            intl_now = active_p2[is_valid].copy()

            nat_counts = (intl_now.groupby(nat_col)[IDCOL].nunique().sort_values(ascending=False)
                          .reset_index().rename(columns={nat_col: "Nationality", IDCOL: "Count"}))
            total_intl = int(intl_now[IDCOL].nunique()) if not intl_now.empty else 0
            n_nats = int(nat_counts["Nationality"].nunique()) if not nat_counts.empty else 0
        else:
            nat_counts = pd.DataFrame({"Nationality": [], "Count": []})
            total_intl = n_nats = 0

        title_nat = f"{total_intl} international Faculty. {n_nats} different nationalities"
        fig_nat = px.bar(nat_counts, x="Count", y="Nationality", orientation="h", title=title_nat, text="Count")
        fig_nat.update_traces(marker_color="#2EC4B6", textposition="outside", texttemplate="%{text}")
        fig_nat.update_xaxes(title=None, dtick=1)
        fig_nat.update_yaxes(title=None, autorange="reversed")
        fig_nat.update_layout(height=bar_h2, margin=dict(l=10, r=10, t=50, b=6))
        st.plotly_chart(fig_nat, use_container_width=True)

        if not intl_now.empty:
            detalle_nat = pick_cols(intl_now, {
                "Full Name": ["Full Name", "Full-Name", "Full_Name", "Profesor", "First Name"],
                "Nationality": ["Nationality", "Country of Birth"],
            })
            popover2 = st.popover if hasattr(st, "popover") else st.expander
            with popover2("🔎 Ver detalle de nacionalidad (profesores)"):
                st.dataframe(detalle_nat.reset_index(drop=True), use_container_width=True)


# ===========================================================================
# 9) NAVEGACIÓN MULTIPÁGINA (todo en un solo archivo)
# ===========================================================================
pages = [
    st.Page(page_composition, title="Composition", icon="🎓", url_path="composition", default=True),
    st.Page(page_staffing, title="Staffing Levels", icon="📊", url_path="staffing"),
    st.Page(page_area, title="By Area", icon="🏛️", url_path="area"),
    st.Page(page_demographics, title="Demographics", icon="🧑‍🤝‍🧑", url_path="demographics"),
]

pg = st.navigation(pages, position="top")
pg.run()
