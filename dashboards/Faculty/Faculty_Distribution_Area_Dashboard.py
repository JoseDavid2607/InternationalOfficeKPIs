# ===========================================================================
#  Distribution by Academic Area · UASM
#  Self-contained · no external module dependencies
# ===========================================================================
import streamlit as st
import pandas as pd
import plotly.express as px
import io, base64
import numpy as np
import re

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Distribution by Academic Area · UASM",
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
    "<style>"
    ".suite-header{display:flex;flex-direction:column;margin-top:-30px;align-items:center;"
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
    "border:1px solid #D1E8E4;border-radius:8px;padding:6px 14px;"
    "margin-bottom:14px;font-size:13px;}"
    "background:#06D6A0;flex-shrink:0;}"
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


import requests as _requests

DRIVE_FILE_ID = "1rPDVrdIxBFMrf0VkBmLtdUmbhvT4dku-"

@st.cache_data(ttl=300)
def _download_excel() -> str:
    """Download BD_Faculty.xlsx from Google Drive to /tmp and return local path."""
    url = f"https://drive.google.com/uc?export=download&id={DRIVE_FILE_ID}"
    path = "/tmp/BD_Faculty.xlsx"
    response = _requests.get(url, stream=True)
    with open(path, "wb") as f:
        for chunk in response.iter_content(chunk_size=32768):
            if chunk:
                f.write(chunk)
    return path

# ── Helper functions (from original) ────────────────────────────────────────
def _is_sem_label(p: str) -> bool:
    return bool(re.fullmatch(r"\d{4}-(10|20)", str(p)))

def _is_inter_label(p: str) -> bool:
    return bool(re.fullmatch(r"\d{4}\s+Intersemestral", str(p)))

def _display_label_sem(p_internal: str) -> str:
    # para mostrar semestres sin guion
    return str(p_internal).replace("-", "")

def _filter_for_timeframe(df_in: pd.DataFrame, time_mode: str, value: str | int | None):
    """
    value:
      - Semestral: visible 'YYYY10'/'YYYY20' pero internamente mapeamos a 'YYYY-10/20'
      - Anual: 'YYYY' -> incluir ambos semestres y el intersemestral, deduplicando por profesor/año
      - Intersemestral: 'YYYY Intersemestral'
    """
    if value is None:
        return df_in.iloc[0:0].copy()

    dfb = df_in.copy()
    pcol = "Periodo"

    if time_mode == "Semestral":
        sem_internal = f"{str(value)[:4]}-{str(value)[-2:]}"
        return dfb[dfb[pcol].astype(str).eq(sem_internal)].copy()

    if time_mode == "Anual":
        y = str(value)
        dfy = dfb[dfb[pcol].astype(str).str.startswith(y)].copy()
        if "ID" in dfy.columns:
            dfy["__Year"] = dfy[pcol].astype(str).str[:4]
            dfy = dfy.sort_values(by=[pcol]).drop_duplicates(subset=["ID", "__Year"], keep="last")
            dfy = dfy.drop(columns=["__Year"])
        return dfy

    # Intersemestral
    inter_label = str(value)
    return dfb[dfb[pcol].astype(str).eq(inter_label)].copy()

# ========= DATA LOAD =========
@st.cache_data(ttl=0)

@st.cache_data(ttl=0)
def load_fulltime():
    df = pd.read_excel(_download_excel(), sheet_name="BD_PLANTA")
    # Periodo soportando intersemestral
    if "Semestre" in df.columns:
        sem = df["Semestre"].astype(str).str.strip()
        is_inter = sem.str.contains("inter", case=False, na=False)
        df["Periodo"] = np.where(is_inter, sem.str[:4] + " Intersemestral", sem.str[:4] + "-" + sem.str[-2:])
    else:
        raw = df.iloc[:, 0].astype(str).str.strip()
        is_inter = raw.str.contains("inter", case=False, na=False)
        df["Periodo"] = np.where(is_inter, raw.str[:4] + " Intersemestral", raw.str.slice(0, 4) + "-" + raw.str.slice(4, 6))
    if "Academic Area" in df.columns and "AREA_PROFESOR" not in df.columns:
        df["AREA_PROFESOR"] = df["Academic Area"]
    if "ID Nr." in df.columns and "ID" not in df.columns:
        df = df.rename(columns={"ID Nr.": "ID"})
    if "Full Name" not in df.columns:
        fn = df["First Name"] if "First Name" in df.columns else ""
        ln = df["Last Name"] if "Last Name" in df.columns else ""
        df["Full Name"] = (pd.Series(fn).astype(str).fillna("") + " " + pd.Series(ln).astype(str).fillna("")).str.strip()
    return df

@st.cache_data(ttl=0)
def load_parttime():
    df = pd.read_excel(_download_excel(), sheet_name="Faculty Distribution")
    # NO excluir intersemestral; construir Periodo consistente
    if "PLANTA_CATEDRA" in df.columns:
        col = df["PLANTA_CATEDRA"].astype(str).str.strip()
        col = col.str.normalize("NFKD").str.encode("ascii", errors="ignore").str.decode("ascii")
        df = df[col.str.upper().eq("CATEDRA")].copy()
    sem = df["Semestre"].astype(str).str.strip()
    is_inter = sem.str.contains("inter", case=False, na=False)
    df.loc[~is_inter, "Periodo"] = sem.str[:4] + "-" + sem.str[-2:]
    df.loc[is_inter,  "Periodo"] = sem.str[:4] + " Intersemestral"
    if "ID Nr." in df.columns and "ID" not in df.columns:
        df = df.rename(columns={"ID Nr.": "ID"})
    if "AREA_PROFESOR" not in df.columns and "Academic Area" in df.columns:
        df["AREA_PROFESOR"] = df["Academic Area"]
    return df

# ── Header ─────────────────────────────────────────────────────────────────────
_render_header("Distribution by Academic Area", "Faculty distribution and evolution across academic areas")

# ── Sidebar navigation ─────────────────────────────────────────────────────────

df_full = load_fulltime()
df_part = load_parttime()

# ── Session state initialisation (must run before sidebar reads it) ──
if "modo_faculty" not in st.session_state:
    st.session_state.modo_faculty = "Full-time"


# ========= SIDEBAR (KPI -> Faculty Type -> Timeframe -> Periodo -> Descarga BD) =========

with st.sidebar:

    # ── Logo UASM ──────────────────────────────────────────────────────────
    col_logo, col_title = st.columns([1, 3])

    with col_logo:
        st.image("web/imagenes/logo.png", width=65)

    with col_title:
        st.markdown(
            """
            <div style="
                padding-top:10px;
                color:#004d47;
                font-size:24px;
                font-weight:800;
                line-height:1.1;
            ">
                UASM Faculty KPIs
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.caption("Analytics Dashboard")

    st.markdown("---")

    st.markdown("#### Faculty Type")
    st.markdown('<div id="mode-pill">', unsafe_allow_html=True)
    mode_sidebar = st.radio(
        "Mode", ["Full-time", "Part-time"],
        index=0 if st.session_state.modo_faculty == "Full-time" else 1,
        horizontal=True, label_visibility="collapsed", key="mode_pill_radio"
    )
    st.markdown('</div>', unsafe_allow_html=True)

    if mode_sidebar != st.session_state.modo_faculty:
        st.session_state.modo_faculty = mode_sidebar
        st.rerun()

    # dataset base según modo
    df_base = df_full if st.session_state.modo_faculty == "Full-time" else df_part

    # 3) Timeframe
    st.markdown("#### Timeframe")
    tmode = st.radio("", ["Semestral", "Anual", "Intersemestral"], key="time_mode_side")

    # 4) Selector de Periodo dependiente
    all_periods = sorted(df_base["Periodo"].astype(str).dropna().unique().tolist())
    if tmode == "Semestral":
        period_opts = [p for p in all_periods if _is_sem_label(p)]
        visible_opts = [_display_label_sem(p) for p in period_opts]  # sin '-'
        default_idx = len(period_opts) - 1 if period_opts else 0
        sel_visible = st.selectbox("Periodo", visible_opts, index=default_idx if period_opts else None)
        sel_value   = period_opts[visible_opts.index(sel_visible)] if period_opts else None  # interno
        sel_label   = sel_visible
    elif tmode == "Anual":
        years = sorted(pd.Series(all_periods).astype(str).str[:4].unique().tolist())
        default_idx = len(years) - 1 if years else 0
        sel_value = st.selectbox("Periodo", years, index=default_idx if years else None)  # 'YYYY'
        sel_label = sel_value
    else:
        inters = [p for p in all_periods if _is_inter_label(p)]
        default_idx = len(inters) - 1 if inters else 0
        sel_value = st.selectbox("Periodo", inters, index=idx if (idx:=default_idx) or inters else None)  # 'YYYY Intersemestral'
        sel_label = sel_value

    st.session_state["sel_tf_mode"]  = tmode
    st.session_state["sel_tf_value"] = sel_value    # interno (ver _filter_for_timeframe)
    st.session_state["sel_tf_label"] = sel_label    # visible

    # 5) Descarga BD completa (según timeframe/periodo seleccionado)
    export_df = _filter_for_timeframe(df_base, tmode, sel_value)
    fname = f"{'FT' if st.session_state.modo_faculty=='Full-time' else 'PT'}_{tmode}_{str(sel_label).replace(' ','_')}.xlsx"
    _b64_dl = _b64.b64encode(_xlsx_bytes(export_df)).decode()
    st.markdown(
        f'<a class="modern-btn" download="{fname}" '
        f'href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{_b64_dl}">'
        f'⭳ Descargar Base — {st.session_state.modo_faculty} — {sel_label}</a>',
        unsafe_allow_html=True,
    )

# ========= PICK DATA BY MODE =========

# ========= PICK DATA BY MODE =========
MINT = "#00A896"
PALETTE = [
    "#056D62", "#1CDFCB", "#FF7F50", "#9B59B6", "#F4A261",
    "#1B6CA8", "#0EAD69", "#E76F51", "#3D5A80", "#8D99AE",
    "#78A7A2", "#F6BD60", "#6D597A", "#43AA8B", "#277DA1"
]
df = df_full.copy() if st.session_state.modo_faculty == "Full-time" else df_part.copy()



# ========= TOP PIVOT TABLE (depende del timeframe) =========
tmode_now   = st.session_state.get("sel_tf_mode", "Semestral")
sel_value   = st.session_state.get("sel_tf_value", None)   # interno
sel_label   = st.session_state.get("sel_tf_label", None)   # visible
IDCOL = "ID"

# Construir tabla pivote por timeframe
df_view = df.copy()
if tmode_now == "Semestral":
    df_view = df_view[df_view["Periodo"].astype(str).apply(_is_sem_label)].copy()
    # Etiquetas visibles sin '-'
    df_view["Periodo_display"] = df_view["Periodo"].astype(str).map(_display_label_sem)
    pivot_area = pd.pivot_table(
        df_view, index="AREA_PROFESOR", columns="Periodo_display",
        values=IDCOL, aggfunc="nunique", fill_value=0
    ).sort_index()
    col_order = sorted(df_view["Periodo_display"].unique().tolist())

elif tmode_now == "Intersemestral":
    df_view = df_view[df_view["Periodo"].astype(str).apply(_is_inter_label)].copy()
    df_view["Periodo_display"] = df_view["Periodo"].astype(str)
    pivot_area = pd.pivot_table(
        df_view, index="AREA_PROFESOR", columns="Periodo_display",
        values=IDCOL, aggfunc="nunique", fill_value=0
    ).sort_index()
    col_order = sorted(df_view["Periodo_display"].unique().tolist())

else:  # Anual: deduplicar por profesor/año
    df_view["__Year"] = df_view["Periodo"].astype(str).str[:4]
    df_view = df_view.sort_values(by=["Periodo"]).drop_duplicates(subset=[IDCOL, "__Year"], keep="last")
    pivot_area = pd.pivot_table(
        df_view, index="AREA_PROFESOR", columns="__Year",
        values=IDCOL, aggfunc="nunique", fill_value=0
    ).sort_index()
    col_order = sorted(pivot_area.columns.astype(str).tolist())

# ----- Color map ÚNICO y consistente por área (para línea y dona) -----
areas_palette_order = [a for a in pivot_area.index if a != "Total"]
color_map_area = {a: PALETTE[i % len(PALETTE)] for i, a in enumerate(areas_palette_order)}

# Total row
pivot_area.loc["Total"] = pivot_area.sum(numeric_only=True)

def _bold_total_row(df_):
    styles = pd.DataFrame('', index=df_.index, columns=df_.columns)
    if "Total" in df_.index:
        styles.loc["Total", :] = "font-weight:700;"
    return styles

def _highlight_total_lastcell(df_):
    styles = pd.DataFrame("", index=df_.index, columns=df_.columns)
    if len(df_.columns) > 0 and "Total" in df_.index:
        last_col = df_.columns[-1]
        styles.loc["Total", last_col] = "background-color:#dff7f2; color:#00A896; font-weight:700;"
    return styles

styled_area = (
    pivot_area[col_order]  # asegurar orden de columnas
    .style
    .apply(_bold_total_row, axis=None)
    .apply(_highlight_total_lastcell, axis=None)
    .format(precision=0)
)

st.subheader(f"{st.session_state.modo_faculty} Faculty count")
st.dataframe(styled_area, use_container_width=True)

# Descarga de la tabla pivote
pivot_download = pivot_area[col_order].reset_index()
fname_pvt = f"Pivot_{'FT' if st.session_state.modo_faculty=='Full-time' else 'PT'}_{tmode_now}.xlsx"
_download_link("Descargar tabla (Excel)", pivot_download, fname_pvt)

# ========= CHARTS =========
st.markdown(f"### Evolution by Academic Area — Number of {st.session_state.modo_faculty} Faculty")

# Claves de estado por modo
mode_key = 'ft' if st.session_state.modo_faculty == 'Full-time' else 'pt'
show_all_key   = f'ver_todas_{mode_key}'
area_sel_key   = f'area_sel_{mode_key}'

# Defaults
if show_all_key not in st.session_state:
    st.session_state[show_all_key] = True
if area_sel_key not in st.session_state:
    st.session_state[area_sel_key] = "Select..."

# Universos y ejes X para la línea según timeframe
if tmode_now == "Semestral":
    base_line = df[df["Periodo"].astype(str).apply(_is_sem_label)].copy()
    base_line["X"] = base_line["Periodo"].astype(str).map(_display_label_sem)  # visible sin '-'
    x_labels = sorted(base_line["X"].unique().tolist())
    group_key = "Periodo"          # para conteos
    x_to_filter = sel_label        # visible label (YYYY10/20)

elif tmode_now == "Intersemestral":
    base_line = df[df["Periodo"].astype(str).apply(_is_inter_label)].copy()
    base_line["X"] = base_line["Periodo"].astype(str)
    x_labels = sorted(base_line["X"].unique().tolist())
    group_key = "Periodo"
    x_to_filter = sel_label

else:  # Anual
    tmp = df.copy()
    tmp["__Year"] = tmp["Periodo"].astype(str).str[:4]
    tmp = tmp.sort_values(by=["Periodo"]).drop_duplicates(subset=[IDCOL, "__Year"], keep="last")
    base_line = tmp.rename(columns={"__Year": "X"}).copy()
    x_labels = sorted(base_line["X"].unique().tolist())
    group_key = "X"
    x_to_filter = str(sel_label)

# Totales por X
totals_period = (
    base_line.groupby("X")[IDCOL]
    .nunique()
    .reindex(x_labels)
    .fillna(0)
    .astype(int)
    .reset_index(name="Total")
)

# Layout
colL, colR = st.columns([3, 2])

# ----------------- LEFT: LINE (% EVOLUTION) -----------------
with colL:
    # Áreas válidas
    areas_current = sorted([a for a in pivot_area.index if a != "Total"])
    if st.session_state[area_sel_key] not in ["Select...", *areas_current]:
        st.session_state[area_sel_key] = "Select..."

    st.checkbox("Show all lines", key=show_all_key)
    area_options = ["Select...", *areas_current]
    st.selectbox("Select academic area:", area_options, key=area_sel_key)

    show_all = st.session_state[show_all_key]
    area_sel_val = st.session_state[area_sel_key]

    if show_all:
        df_counts = (
            base_line.groupby(["X", "AREA_PROFESOR"])[IDCOL]
            .nunique()
            .reset_index(name="Count")
        )
        df_counts["X"] = pd.Categorical(df_counts["X"], categories=x_labels, ordered=True)
        df_counts = df_counts.merge(totals_period, on="X", how="left")
        df_counts["Pct"] = (df_counts["Count"] / df_counts["Total"].replace(0, pd.NA)).fillna(0)

        fig_line = px.line(
            df_counts, x="X", y="Pct", color="AREA_PROFESOR",
            markers=True,
            category_orders={"X": x_labels, "AREA_PROFESOR": areas_palette_order},
            color_discrete_map=color_map_area   # <<< mismo mapa que la dona
        )
        fig_line.update_traces(mode="lines+markers", line=dict(width=2),
                               hovertemplate="<b>%{x}</b><br>%{fullData.name}: %{y:.1%}<extra></extra>")
        fig_line.update_xaxes(type="category", categoryorder="array", categoryarray=x_labels, title=None)
        fig_line.update_yaxes(rangemode="tozero", tickformat=".0%", title=None)
        fig_line.update_layout(showlegend=False)

        # Franja azul resaltando el periodo seleccionado del sidebar
        if x_to_filter in x_labels:
            pos_sel = x_labels.index(x_to_filter)
            fig_line.add_shape(
                type="rect", xref="x", yref="paper",
                x0=pos_sel - 0.4, x1=pos_sel + 0.4, y0=0, y1=1,
                fillcolor="#D0E5F5", opacity=0.35, line_width=0
            )

        st.plotly_chart(fig_line, use_container_width=True)
    else:
        if area_sel_val == "Select...":
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
                df_line, x="X", y="Pct",
                markers=True, title=f"Evolution (% share) — {area_sel_val}",
                color_discrete_sequence=[color_map_area.get(area_sel_val, MINT)]  # <<< color consistente
            )
            fig_line.update_traces(mode="lines+markers", line=dict(width=2),
                                   hovertemplate="<b>%{x}</b><br>%{y:.1%}<extra></extra>")
            fig_line.update_xaxes(type="category", categoryorder="array", categoryarray=x_labels, title=None)
            fig_line.update_yaxes(rangemode="tozero", tickformat=".0%", title=None)
            fig_line.update_layout(showlegend=False)

            if x_to_filter in x_labels:
                pos_sel = x_labels.index(x_to_filter)
                fig_line.add_shape(
                    type="rect", xref="x", yref="paper",
                    x0=pos_sel - 0.4, x1=pos_sel + 0.4, y0=0, y1=1,
                    fillcolor="#D0E5F5", opacity=0.35, line_width=0
                )

            st.plotly_chart(fig_line, use_container_width=True)

# ----------------- RIGHT: DONUT (periodo seleccionado) -----------------
with colR:
    st.markdown(f"##### Distribution by academic area — {sel_label if sel_label else ''}")

    if tmode_now == "Anual":
        df_donut = _filter_for_timeframe(df, "Anual", sel_value)
    else:
        df_donut = _filter_for_timeframe(df, tmode_now, sel_value)

    if df_donut.empty:
        st.info("No data for the selected period.")
    else:
        dist = (
            df_donut.groupby("AREA_PROFESOR")[IDCOL]
            .nunique()
            .sort_values(ascending=False)
        )
        donut_df = pd.DataFrame({"Area": dist.index, "Value": dist.values})

        fig_donut = px.pie(
            donut_df,
            names="Area", values="Value", hole=0.45,
            color="Area",
            color_discrete_map=color_map_area,  # <<< mismo mapa que la línea
            category_orders={"Area": areas_palette_order}
        )
        st.plotly_chart(fig_donut, use_container_width=True)

        # 👉 Botón de descarga de la tabla usada en la dona
        fname_donut = f"Donut_{'FT' if st.session_state.modo_faculty=='Full-time' else 'PT'}_{tmode_now}_{str(sel_label).replace(' ','_')}.xlsx"
        _download_link("Descargar tabla (Excel)", donut_df, fname_donut)

# ========= DETAIL TABLE (del periodo seleccionado) =========
if tmode_now == "Anual":
    detail = _filter_for_timeframe(df, "Anual", sel_value)
else:
    detail = _filter_for_timeframe(df, tmode_now, sel_value)

cols_prefer_ft = ["Full Name", "AREA_PROFESOR", "Faculty Ranking", "Faculty Qualific.", "P/S"]
cols_prefer_pt = ["Profesor", "AREA_PROFESOR", "PLANTA_CATEDRA", "TIPO", "P/S"]

prefer_cols = cols_prefer_ft if st.session_state.modo_faculty == "Full-time" else cols_prefer_pt
cols_to_show = [c for c in prefer_cols if c in detail.columns]
detail_out = detail[cols_to_show].drop_duplicates().reset_index(drop=True)

st.markdown(f"### There are {int(detail[IDCOL].nunique()) if IDCOL in detail.columns else len(detail_out)} {'full-time' if st.session_state.modo_faculty=='Full-time' else 'part-time'} Faculty in **{sel_label}**")
st.dataframe(detail_out, use_container_width=True)

# Descarga de la tabla de detalle
fname_det = f"Detail_{'FT' if st.session_state.modo_faculty=='Full-time' else 'PT'}_{tmode_now}_{str(sel_label).replace(' ','_')}.xlsx"
_download_link("Descargar tabla (Excel)", detail_out, fname_det)
