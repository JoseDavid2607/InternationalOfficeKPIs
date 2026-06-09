# ===========================================================================
#  Full-time Faculty Staffing Levels · UASM
#  Self-contained · no external module dependencies
# ===========================================================================
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import re
import io, base64
import numpy as np

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Full-time Faculty Staffing Levels · UASM",
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

_NAV = {
    "1 Full-time Composition":           "https://facultycompositiondashboardpy-dtacyzfa3otmpbewqc5axu.streamlit.app/",
    "2 Full-time Staffing Levels":       "https://facultystaffinglevelsdashboardpy-phv4t8jzbyyz5rrepqttuf.streamlit.app/",
    "3 Distribution by Academic Area":   "https://facultydistributionareadashboardpy-yzwpiqdlukfdp6qcygxjhj.streamlit.app/",
    "4 Faculty Demographics":            "https://facultydemographicsdashboardpy-kmsnpswxs35psbqtdtvb6y.streamlit.app/",
    "5 Full-time Faculty Questionnaire": "https://full-timefacultyactivitiespy-bbe7fmmyrxvssadnygm4fx.streamlit.app/",
    "6 Faculty Qualifications":          "https://facultyqualificationspy-drvj3wpyrxvm2lrnafdwx5.streamlit.app/",
}

# ── Global CSS ─────────────────────────────────────────────────────────────────
st.markdown(
    "<style>"
    ".suite-header{display:flex;flex-direction:column;align-items:center;"
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
    ".upd-banner{display:flex;align-items:center;gap:10px;background:#dff7f2;"
    "border:1px solid #D1E8E4;border-radius:8px;padding:6px 14px;"
    "margin-bottom:14px;font-size:13px;}"
    ".upd-dot{width:8px;height:8px;border-radius:50%;"
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
    "</style>",
    unsafe_allow_html=True,
)

# ── Inline helpers ─────────────────────────────────────────────────────────────
import io as _io, base64 as _b64, datetime as _dt_mod

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

def _render_update_banner():
    t = _dt_mod.datetime.now().strftime("%d %b %Y \u00b7 %H:%M")
    st.markdown(
        '<div class="upd-banner"><span class="upd-dot"></span>'
        '<b>Last updated:</b>&nbsp;' + t + '&nbsp;\u00b7&nbsp;Data is current</div>',
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

def _nav_sidebar(current):
    st.sidebar.markdown(
        "<div style='font-size:11px;font-weight:700;letter-spacing:1.5px;"
        "color:#6B7280;text-transform:uppercase;margin-bottom:6px'>Navigation</div>",
        unsafe_allow_html=True,
    )
    choices = list(_NAV.keys())
    idx = choices.index(current) if current in choices else 0
    sel = st.sidebar.selectbox(
        "Go to dashboard", choices, index=idx, label_visibility="collapsed"
    )
    st.sidebar.link_button("🔗 Open Dashboard", _NAV[sel], use_container_width=True)
    st.sidebar.markdown("<hr style='margin:10px 0;opacity:.4'>", unsafe_allow_html=True)
    if st.sidebar.button("🔄 Update Data", use_container_width=True, key="upd_data_btn"):
        st.cache_data.clear()
        st.rerun()
    st.sidebar.markdown("<hr style='margin:10px 0;opacity:.4'>", unsafe_allow_html=True)

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
_render_header("Full-time Faculty Staffing Levels", "New entrants, leavers, and headcount evolution")
_render_update_banner()

# ── Sidebar navigation ─────────────────────────────────────────────────────────
_nav_sidebar("2 Full-time Staffing Levels")
@st.cache_data
def load_data():
    df_ = pd.read_excel("data/Faculty/BD_Faculty.xlsx", sheet_name="BD PLANTA 2020-2025")

    # Build 'Periodo' soportando intersemestral, pero nos quedaremos con semestral
    def _norm_per(val):
        s = str(val).strip()
        m_inter = re.search(r'((?:19|20)\d{2}).{0,6}inter', s, flags=re.IGNORECASE)
        if m_inter:
            return f"{m_inter.group(1)} Intersemestral"
        m = re.search(r'((?:19|20)\d{2})\D?(\d{2})', s)  # 202010, 2020-10, 2020_10, etc.
        if m:
            return f"{m.group(1)}-{m.group(2)}"
        return None

    base_col = df_.columns[0]
    df_["Periodo"] = df_[base_col].map(_norm_per)

    # Mantener solo semestres YYYY-10 / YYYY-20
    valid_sem = df_["Periodo"].astype(str).str.match(r'^(?:19|20)\d{2}-(10|20)$')
    df_ = df_.loc[valid_sem].copy()

    # Asegurar columna ID coherente
    if "ID Nr." in df_.columns and "ID" not in df_.columns:
        df_ = df_.rename(columns={"ID Nr.": "ID"})
    if "ID" not in df_.columns and "ID Nr." in df_.columns:
        df_["ID"] = df_["ID Nr."]

    return df_

df = load_data()

# Periodos semestrales
all_periods = sorted(df["Periodo"].astype(str).unique().tolist())          # YYYY-SS
sem_periods = [p for p in all_periods if re.fullmatch(r'(?:19|20)\d{2}-(10|20)', p)]

# Sidebar
with st.sidebar:
    st.markdown("### 📊 Go to KPI:")
    # Solo entradas con URL http(s)
    choices = [k for k, u in _NAV.items() if isinstance(u, str) and (u.startswith("http://") or u.startswith("https://"))]

    # Navigation handled by _nav_sidebar() above

    # ---- Selector SOLO SEMESTRE (visible sin guion)
    st.markdown("---")
    st.markdown("#### Select Semester")
    vis_opts = [p.replace("-", "") for p in sem_periods]  # mostrar sin guion
    idx = len(vis_opts) - 1 if vis_opts else 0
    sel_vis = st.selectbox("", vis_opts, index=idx if vis_opts else None)
    sel_period_internal = sem_periods[vis_opts.index(sel_vis)] if vis_opts else None  # 'YYYY-10/20'
    sel_period_label = sel_vis  # 'YYYY10' / 'YYYY20'

    # ---- Descarga BD completa
    _download_link("Descargar base completa (Excel) — Full-time", df, "FT_Base_Completa.xlsx")

# =============================
# HELPERS (solo semestral)
# =============================
def perlist_sem():
    return sem_periods

def final_count_series_sem(df_):
    # Final por periodo puntual (semestral)
    return df_.groupby("Periodo")["ID"].nunique()

def in_out_counts_sem(df_, label):
    """Cuenta IN/OUT por semestre exacto usando 'Notes' con IN IN YYYYSS u OUT IN YYYYSS."""
    counts = {}
    for p in sem_periods:
        flat = p.replace("-", "")
        counts[p] = int(df_["Notes"].astype(str).str.contains(
            fr"\b{label}\s+IN\s+\(?{flat}\)?\b", case=False, na=False
        ).sum())
    return pd.Series(counts)

# =============================
# STAFFING SUMMARY TABLE (solo semestral)
# =============================
cols_summary = perlist_sem()
fin_ser = final_count_series_sem(df).reindex(cols_summary, fill_value=0)
new_ser = in_out_counts_sem(df, "IN").reindex(cols_summary, fill_value=0)
out_ser = in_out_counts_sem(df, "OUT").reindex(cols_summary, fill_value=0)

rows = []
for i, key in enumerate(cols_summary):
    new_hires = int(new_ser.get(key, 0))
    leavers   = int(out_ser.get(key, 0))
    if i == 0:
        start_val = int(fin_ser.iloc[0]) - new_hires + leavers
    else:
        start_val = int(fin_ser.iloc[i-1])
    rows.append({
        "Start": int(start_val),
        "New": new_hires,
        "Leavers": leavers,
        "Final": int(fin_ser.iloc[i])
    })

summary_df = pd.DataFrame(rows, index=cols_summary).T  # filas=metricas, columnas=semestres

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

# ---- Descarga (pegada a la izquierda)
sum_left, sum_right = st.columns([1,5])
with sum_left:
    simple_tbl = summary_df.reset_index().rename(columns={"index": "Metric"})
    _download_link("Descargar tabla (Excel)", simple_tbl, "FT_New_Leavers_Semestral.xlsx")

# =============================
# CHARTS LAYOUT (tornado right; line left) — ambos obedecen el semestre seleccionado
# =============================
areas = sorted(df.get("Academic Area", pd.Series(dtype=object)).dropna().unique().tolist())
col_left, col_right = st.columns([3, 2])

# ---------- RIGHT: TORNADO (solo semestre seleccionado) ----------
with col_right:
    st.markdown(f"<div style='text-align:center;font-weight:700'>Period: {sel_period_label}</div>", unsafe_allow_html=True)

    current_period = sel_period_internal or ""
    flat_list = [current_period.replace("-", "")] if current_period else []

    pat_in  = "|".join([re.escape(f) for f in flat_list]) if flat_list else r"$^"
    df_in   = df[df["Notes"].astype(str).str.contains(fr"\bIN\s+IN\s+\(?({pat_in})\)?\b",  case=False, na=False)]
    df_out  = df[df["Notes"].astype(str).str.contains(fr"\bOUT\s+IN\s+\(?({pat_in})\)?\b", case=False, na=False)]

    new_by_area  = df_in.groupby("Academic Area")["ID"].nunique().reindex(areas, fill_value=0)
    left_by_area = df_out.groupby("Academic Area")["ID"].nunique().reindex(areas, fill_value=0)
    net_by_area  = (new_by_area - left_by_area).astype(int)
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

# ---------- LEFT: LINE (usa summary_df semestral) ----------
with col_left:
    st.markdown("### Evolution of Faculty (Start vs Final)")

    x_periods = list(summary_df.columns)  # semestres
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

    # Resaltar banda del semestre seleccionado (coincide con x_periods: 'YYYY-SS')
    sel_for_band = sel_period_internal
    if sel_for_band in x_periods:
        pos = x_periods.index(sel_for_band)
        fig_line.add_shape(
            type="rect", xref="x", yref="paper",
            x0=pos - 0.4, x1=pos + 0.4, y0=0, y1=1,
            fillcolor="#D0E5F5", opacity=0.35, line_width=0
        )

    st.plotly_chart(fig_line, use_container_width=True)

# =============================
# FACULTY DETAILS (solo semestre del sidebar)
# =============================
st.markdown("### View Faculty details")

active = df[df["Periodo"].astype(str).eq(sel_period_internal)].copy()

# ---------- DEMOGRAPHICS ----------
st.markdown(
    f"<div style='text-align:center;font-size:34px;font-weight:bold'>Active Full-time Faculty {sel_period_label}</div>",
    unsafe_allow_html=True
)
total_act = active["ID"].nunique()

c0, c1 = st.columns([1, 2])
c0.markdown(f"<div style='text-align:right;font-size:56px;font-weight:bold'>{total_act}</div>", unsafe_allow_html=True)

gen = active["Gender"].value_counts()
df_gen = pd.DataFrame({
    "Gender": ["Male", "Female"],
    "P": [
        round(gen.get("Male", 0)   / total_act * 100, 1) if total_act else 0,
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

# ---------- FULL TABLE ----------
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

# =============================
# PROFESSOR TRAJECTORY (PLANTA ONLY) — from BD PLANTA 2020-2025
# =============================
def _c(df0, *names):
    """Case-insensitive column resolver."""
    if df0 is None or df0.empty:
        return None
    cmap = {str(c).strip().casefold(): c for c in df0.columns}
    for n in names:
        key = str(n).strip().casefold()
        if key in cmap:
            return cmap[key]
    return None

period_col = _c(df, "Periodo")
id_col     = _c(df, "ID Nr.", "ID Nr", "ID")
fn_col     = _c(df, "First Name")
ln_col     = _c(df, "Last Name")
area_col   = _c(df, "Academic Area", "AREA_PROFESOR")
deg_col    = _c(df, "Highest Degree")
rank_col   = _c(df, "Faculty Ranking")
subc_col   = _c(df, "Subcategorization")
age_col    = _c(df, "Age")
qual_col   = _c(df, "Faculty Qualific.")
ps_col     = _c(df, "P/S", "P - S")
resp_col   = _c(df, "Normal professional Resp.")
notes_col  = _c(df, "Notes")

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
        index=0
    )

    if sel_label and sel_label != "(Select...)":
        m = re.search(r"ID:\s*(.+)$", sel_label)
        chosen_id = m.group(1).strip() if m else None

        traj = df[df[id_col].astype(str).str.strip() == chosen_id].copy()

        out_cols_raw = [
            (period_col, "Periodo"),
            (id_col,     "ID Nr."),
            (fn_col,     "First Name"),
            (ln_col,     "Last Name"),
            (area_col,   "Academic Area"),
            (deg_col,    "Highest Degree"),
            (rank_col,   "Faculty Ranking"),
            (subc_col,   "Subcategorization"),
            (age_col,    "Age"),
            (qual_col,   "Faculty Qualific."),
            (ps_col,     "P/S"),
            (resp_col,   "Normal professional Resp."),
            (notes_col,  "Notes"),
        ]

        out_df = pd.DataFrame({
            new: (traj[orig] if (orig in traj.columns) else pd.Series([""] * len(traj), index=traj.index))
            for orig, new in out_cols_raw
        })
        out_df.columns = pd.Index(out_df.columns).map(str)

        OUT_COLOR = "#8B0000"
        IN_COLOR  = "#00796B"

        def _matches_tag_for_period(note_upper: str, tag: str, flat_period: str) -> bool:
            pat = rf'\b{tag}\s+IN\s+\(?((?:19|20)\d{{2}}[-_/ ]?\d{{2}})\)?\b'
            m2 = re.search(pat, note_upper, flags=re.IGNORECASE)
            if not m2:
                return False
            per_txt  = m2.group(1)
            per_flat = re.sub(r'\D', '', per_txt)
            return flat_period and (per_flat == flat_period)

        def _color_in_out(row: pd.Series):
            per        = str(row.get("Periodo", ""))
            note_upper = str(row.get("Notes", "")).upper()
            flat       = re.sub(r'\D', '', per)

            is_out = _matches_tag_for_period(note_upper, "OUT", flat) or ("OUT IN" in note_upper)
            is_in  = _matches_tag_for_period(note_upper, "IN",  flat) or ("IN IN"  in note_upper)

            if is_out:
                return [f'color:{OUT_COLOR};font-weight:700;' for _ in row.index]
            if is_in:
                return [f'color:{IN_COLOR};font-weight:700;' for _ in row.index]
            return ['' for _ in row.index]

        st.dataframe(
            out_df.reset_index(drop=True).style.apply(_color_in_out, axis=1).hide(axis="index"),
            use_container_width=True
        )

        # ---- Botón de descarga minimalista (igual a los otros)
        _download_link("Descargar trayectoria (Excel)", out_df, f"Trajectory_{chosen_id}.xlsx")






