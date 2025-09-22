import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
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

# ===== Utils descarga (xlsx minimalista)
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
    st.markdown(
        f'<a class="dl-min" download="{filename}" href="{href}">{label}</a>',
        unsafe_allow_html=True
    )

# ===== CSS básico (header + link descarga)
st.markdown("""
<style>
.header-title { color:#21877D; font-weight:bold; text-align:center; font-size:32px; }
.header-btn {
  background-color:#21877D; color:white !important; padding:8px 16px;
  border:none; border-radius:8px; cursor:pointer; font-size:14px; text-decoration:none !important;
}
.header-btn:hover { background-color:#1a6b62; }
a.dl-min, a.dl-min:link, a.dl-min:visited {
  color:#1FA89B !important; text-decoration:underline !important;
  font-size:13px; display:inline-block; margin-top:6px;
}
a.dl-min:hover { opacity:.85; }
</style>
""", unsafe_allow_html=True)

# =============================
# INSTITUTIONAL HEADER
# =============================
with st.container():
    cols = st.columns([1, 3, 1], gap="small")
    with cols[0]:
        st.markdown('<a href="http://157.253.69.67:8501" class="header-btn" target="_self">⬅ Previous KPI</a>', unsafe_allow_html=True)
    with cols[1]:
        st.markdown('<div class="header-title">Full-time Faculty Staffing Levels</div>', unsafe_allow_html=True)
    with cols[2]:
        st.markdown('<a href="http://157.253.69.67:8503" class="header-btn" target="_self">➡ Next KPI</a>', unsafe_allow_html=True)

# =============================
# NAV EN SIDEBAR + SOLO SEMESTRE + DESCARGA BD
# =============================
options = {
    "Select...": None,
    "1 Full-time Composition": "http://157.253.69.67:8501",
    "2 Full-time Staffing Levels": "http://157.253.69.67:8502",
    "3 Distribution by Academic Area": "http://157.253.69.67:8503",
    "4 Faculty Demographics": "http://157.253.69.67:8504",
    "5 Full-time Faculty Questionnaire": "http://157.253.69.67:8505",
    "6 Faculty Qualifications": "http://157.253.69.67:8506",
    "Open main HTML menu": "web/KPIs/Faculty/Web KPIs - Faculty.html"
}

# =============================
# DATA LOAD
# =============================
@st.cache_data
def load_data():
    df_ = pd.read_excel(r"data/Faculty/BD_Faculty.xlsx", sheet_name="BD PLANTA 2020-2025")

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
    choice = st.selectbox("📊 Go to KPI:", list(options.keys()))
    target = options.get(choice)
    if target:
        if target.endswith(".html"):
            abs_path = os.path.abspath(target)
            webbrowser.open(f"file:///{abs_path}")
            st.success("The Faculty menu was opened in a new browser tab.")
        else:
            st.markdown(f'<meta http-equiv="refresh" content="0; url={target}" />', unsafe_allow_html=True)

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
