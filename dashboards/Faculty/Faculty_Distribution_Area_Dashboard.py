# ===========================================================================
#  Distribution by Academic Area · UASM
# ===========================================================================
import base64
import re
import numpy as np
import pandas as pd
import plotly.express as px
import requests
import streamlit as st

st.set_page_config(page_title="Distribution by Academic Area · UASM", page_icon="🎓",
                    layout="wide", initial_sidebar_state="expanded")

MINT = "#00A896"
HIGHLIGHT = "#D0E5F5"
PALETTE = [
    "#056D62", "#1CDFCB", "#FF7F50", "#9B59B6", "#F4A261",
    "#1B6CA8", "#0EAD69", "#E76F51", "#3D5A80", "#8D99AE",
    "#78A7A2", "#F6BD60", "#6D597A", "#43AA8B", "#277DA1",
]

# ── Global CSS ──────────────────────────────────────────────────────────────
st.markdown("""
<style>
.suite-header{display:flex;flex-direction:column;margin-top:-35px;align-items:center;
  padding:16px 24px 12px;background:linear-gradient(135deg,#004d47 0%,#21877D 60%,#2EC4B6 100%);
  border-radius:12px;box-shadow:0 2px 8px rgba(0,77,71,.18);margin-bottom:14px;}
.sh-super{font-size:11px;font-weight:700;letter-spacing:2px;color:#56D6C9;
  text-transform:uppercase;margin-bottom:2px;}
.sh-title{font-size:26px;font-weight:800;color:#fff;text-align:center;line-height:1.2;}
.sh-sub{font-size:13px;color:rgba(255,255,255,.75);margin-top:4px;text-align:center;}
a.dl-min,a.dl-min:link,a.dl-min:visited{color:#00A896 !important;text-decoration:underline !important;
  font-size:13px;display:inline-block;margin-top:6px;}
a.dl-min:hover{opacity:.85;}
thead th{background:#dff7f2 !important;color:#004d47 !important;font-weight:700 !important;}
section[data-testid='stSidebar']{background:#F0F7F6 !important;}
#mode-pill [role='radiogroup']{display:flex;gap:8px;margin-top:0;}
#mode-pill [role='radio']{flex:1;justify-content:center;border:1px solid #d0d4d9;
  border-radius:999px;padding:8px 12px;background:#f0f2f6;color:#666;font-weight:600;
  cursor:pointer;text-align:center;}
#mode-pill [role='radio'][aria-checked='true']{background:#dff7f2;color:#004d47;border-color:#8fd7cc;}
#mode-pill [data-baseweb='radio'] input{display:none !important;}
.modern-btn{background:#FFFFFF;border:1px solid #D1E8E4;border-radius:10px;padding:12px 14px;
  color:#374151 !important;font-size:14px;font-weight:600;text-decoration:none !important;
  display:block;text-align:center;margin-bottom:10px;transition:all .2s ease;
  box-shadow:0 1px 3px rgba(0,0,0,.04);}
.modern-btn:hover{background:#F8FFFE;border-color:#B7DCD6;}
</style>
""", unsafe_allow_html=True)


# ── Helpers ─────────────────────────────────────────────────────────────────
def xlsx_bytes(df: pd.DataFrame, sheet_name: str = "Data") -> bytes:
    import io
    buf = io.BytesIO()
    with pd.ExcelWriter(buf) as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    return buf.getvalue()


def download_link(label: str, df: pd.DataFrame, filename: str):
    b64 = base64.b64encode(xlsx_bytes(df)).decode()
    href = f"data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}"
    st.markdown(f'<a class="dl-min" download="{filename}" href="{href}">{label}</a>', unsafe_allow_html=True)


def render_header(title: str, subtitle: str = ""):
    sub = f'<div class="sh-sub">{subtitle}</div>' if subtitle else ""
    st.markdown(
        f'<div class="suite-header"><div class="sh-super">UASM · Faculty Analytics</div>'
        f'<div class="sh-title">{title}</div>{sub}</div>',
        unsafe_allow_html=True,
    )


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


# ── Data loading (Google Sheets export) ─────────────────────────────────────
SHEET_ID = "1PZkqgtvct5LFNWVUEkA5fuglvqvAuMxseSq10MV9ji8"
SHEET_EXPORT_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"


@st.cache_data(ttl=300)
def download_excel() -> str:
    path = "/tmp/BD_Faculty_demo.xlsx"
    content = requests.get(SHEET_EXPORT_URL).content
    with open(path, "wb") as f:
        f.write(content)
    return path


@st.cache_data(ttl=0)
def load_fulltime() -> pd.DataFrame:
    df = pd.read_excel(download_excel(), sheet_name="BD_PLANTA")

    sem = df["Semestre"].astype(str).str.strip() if "Semestre" in df.columns else df.iloc[:, 0].astype(str).str.strip()
    is_inter = sem.str.contains("inter", case=False, na=False)
    df["Periodo"] = np.where(is_inter, sem.str[:4] + " Intersemestral", sem.str[:4] + "-" + sem.str[-2:])

    if "Academic Area" in df.columns and "AREA_PROFESOR" not in df.columns:
        df["AREA_PROFESOR"] = df["Academic Area"]
    if "ID Nr." in df.columns and "ID" not in df.columns:
        df = df.rename(columns={"ID Nr.": "ID"})
    if "Full Name" not in df.columns:
        fn = df.get("First Name", "").astype(str).fillna("")
        ln = df.get("Last Name", "").astype(str).fillna("")
        df["Full Name"] = (fn + " " + ln).str.strip()
    return df


@st.cache_data(ttl=0)
def load_parttime() -> pd.DataFrame:
    df = pd.read_excel(download_excel(), sheet_name="Faculty Distribution")

    if "PLANTA_CATEDRA" in df.columns:
        col = df["PLANTA_CATEDRA"].astype(str).str.strip()
        col = col.str.normalize("NFKD").str.encode("ascii", errors="ignore").str.decode("ascii")
        df = df[col.str.upper().eq("CATEDRA")].copy()

    sem = df["Semestre"].astype(str).str.strip()
    is_inter = sem.str.contains("inter", case=False, na=False)
    df.loc[~is_inter, "Periodo"] = sem.str[:4] + "-" + sem.str[-2:]
    df.loc[is_inter, "Periodo"] = sem.str[:4] + " Intersemestral"

    if "ID Nr." in df.columns and "ID" not in df.columns:
        df = df.rename(columns={"ID Nr.": "ID"})
    if "AREA_PROFESOR" not in df.columns and "Academic Area" in df.columns:
        df["AREA_PROFESOR"] = df["Academic Area"]
    return df


render_header("Distribution by Academic Area", "Faculty distribution and evolution across academic areas")

df_full = load_fulltime()
df_part = load_parttime()

st.session_state.setdefault("modo_faculty", "Full-time")


# ── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    col_logo, col_title = st.columns([1, 3])
    with col_logo:
        st.image("web/imagenes/logo.png", width=65)
    with col_title:
        st.markdown(
            '<div style="padding-top:10px;color:#004d47;font-size:24px;font-weight:800;'
            'line-height:1.1;">UASM Faculty KPIs</div>',
            unsafe_allow_html=True,
        )
        st.caption("Analytics Dashboard")

    st.markdown("---")
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
    b64_dl = base64.b64encode(xlsx_bytes(export_df)).decode()
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
download_link("Descargar tabla (Excel)", pivot_download, fname_pvt)


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
        download_link("Descargar tabla (Excel)", donut_df, fname_donut)


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
download_link("Descargar tabla (Excel)", detail_out, fname_det)
