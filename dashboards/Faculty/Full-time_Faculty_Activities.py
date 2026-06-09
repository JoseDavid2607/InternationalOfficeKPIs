# ===========================================================================
#  Full-time Faculty Activities · UASM
#  Self-contained · no external module dependencies
# ===========================================================================
import streamlit as st
import pandas as pd
import numpy as np
from typing import Optional, Tuple, List, Dict
import plotly.express as px
from io import BytesIO
import base64

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Full-time Faculty Activities · UASM",
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

EXCEL_PATH = _download_excel()

# ── Helper functions (from original) ────────────────────────────────────────
def resolve_column(df: pd.DataFrame, target: str) -> Optional[str]:
    t = target.strip().casefold()
    for c in df.columns:
        if c.strip().casefold() == t:
            return c
    return None

@st.cache_data(ttl=0)

@st.cache_data(ttl=0)
def load_fulltime():
    df = pd.read_excel(EXCEL_PATH, sheet_name="BD PLANTA 2020-2025")
    raw = df.iloc[:, 0].astype(str)
    df["Periodo"] = raw.str.slice(0, 4) + "-" + raw.str.slice(4, 6)
    if "ID Nr." in df.columns and "ID" not in df.columns:
        df = df.rename(columns={"ID Nr.": "ID"})
    df.columns = df.columns.str.strip()
    return df

@st.cache_data(ttl=0)
def load_questionnaire():
    df = pd.read_excel(EXCEL_PATH, sheet_name="Faculty_questionnaire")
    df.columns = df.columns.str.strip()
    ycol = resolve_column(df, "Year")
    if ycol:
        df[ycol] = pd.to_numeric(df[ycol], errors="coerce").astype("Int64")
    if "ID Nr." in df.columns and "ID" not in df.columns:
        df = df.rename(columns={"ID Nr.": "ID"})
    return df

@st.cache_data(ttl=0)
def load_courses_sheets():
    """Load sheets for: Credit granted courses / Non-credit granted courses (name tolerant)."""
    xls = pd.ExcelFile(EXCEL_PATH)
    sheets = xls.sheet_names

    def pick_sheet(candidates: List[str]) -> Optional[str]:
        lowmap = {s.lower(): s for s in sheets}
        for cand in candidates:
            if cand.lower() in lowmap: return lowmap[cand.lower()]
        for s in sheets:
            s_low = s.lower()
            if any(cand.lower() in s_low for cand in candidates):
                return s
        return None

    credit_candidates = [
        "Creditd granted courses", "Credited granted courses", "Credit granted courses",
        "Credit granted course", "Credit granted", "Credit courses", "Creditd courses"
    ]
    noncredit_candidates = [
        "Non-credit granted courses", "Non credit granted courses", "Non-credit courses",
        "Non credit courses", "Noncredit granted courses"
    ]

    sh_credit = pick_sheet(credit_candidates)
    sh_noncr  = pick_sheet(noncredit_candidates)

    df_credit = pd.read_excel(xls, sheet_name=sh_credit) if sh_credit else pd.DataFrame()
    df_noncr  = pd.read_excel(xls, sheet_name=sh_noncr)  if sh_noncr  else pd.DataFrame()
    if not df_credit.empty: df_credit.columns = df_credit.columns.str.strip()
    if not df_noncr.empty:  df_noncr.columns  = df_noncr.columns.str.strip()
    return df_credit, df_noncr, sh_credit, sh_noncr

# ── Header ─────────────────────────────────────────────────────────────────────
_render_header("Full-time Faculty Activities", "Questionnaire-based engagement summary 2020–2025")
_render_update_banner()

# ── Sidebar navigation ─────────────────────────────────────────────────────────
_nav_sidebar("5 Full-time Faculty Questionnaire")

df_full = load_fulltime()
df_q    = load_questionnaire()
df_credit_sheet, df_noncredit_sheet, credit_sheet_name, noncredit_sheet_name = load_courses_sheets()

# ================= SIDEBAR: NAVIGATION (selector + Open) =================
with st.sidebar:
    st.markdown("### 📊 Go to KPI:")
    
    choices = [k for k, u in _NAV.items() if isinstance(u, str) and (u.startswith("http://") or u.startswith("https://"))]
    default_label = "5 Full-time Faculty Questionnaire"  # etiqueta de este KPI en tu lista
    default_idx = choices.index(default_label) if default_label in choices else 0

    sel = st.selectbox("Select…", choices, index=default_idx)
    st.link_button("Open", _NAV[sel], use_container_width=True)

#================= CONSTANTS ==================================================
TOT_PROFESSORS = 64           # denominator for % (donuts)
MINT      = "#56D6C9"          # mint for "YES"
MINT_DARK = "#1FA89B"          # darker mint (center text)
GREY      = "#C7C7C7"          # grey for "NO"
DONUT_H   = 160                # height of each donut
#================= YEARS (fixed 2020–2025) ====================================
YEARS = [2020, 2021, 2022, 2023, 2024, 2025]
def _norm(s: pd.Series) -> pd.Series:
    return s.astype(str).str.strip().str.lower()

def _year_filter(df: pd.DataFrame, year: int) -> Tuple[pd.DataFrame, Optional[str]]:
    ycol = resolve_column(df, "Year")
    if not ycol:
        return pd.DataFrame(), None
    d = df[df[ycol].astype("Int64") == year]
    return d, ycol

def ft_second_sem_count(full_df: pd.DataFrame, year: int) -> Optional[int]:
    """Full-time total for the 2nd term; fallback to the last term of that year."""
    y = str(year)
    period = f"{y}-20"
    dfy = full_df[full_df["Periodo"] == period]
    if dfy.empty:
        by_year = full_df[full_df["Periodo"].str.startswith(y + "-")]
        if by_year.empty:
            return None
        last_p = sorted(by_year["Periodo"].unique())[-1]
        dfy = full_df[full_df["Periodo"] == last_p]
    idcol = "ID" if "ID" in dfy.columns else ("ID Nr." if "ID Nr." in dfy.columns else None)
    return int(dfy[idcol].nunique()) if idcol else int(len(dfy))

def count_yes(df: pd.DataFrame, col: str, year: int) -> Optional[int]:
    d, ycol = _year_filter(df, year)
    if ycol is None or d.empty:
        return None
    c = resolve_column(df, col)
    if not c:
        return None
    return int(_norm(d[c]).eq("yes").sum())

def count_contains(df: pd.DataFrame, col: str, patt: str, year: int, unique_by_id: bool = True) -> Optional[int]:
    d, ycol = _year_filter(df, year)
    if ycol is None or d.empty:
        return None
    c = resolve_column(df, col)
    if not c:
        return None
    mask = _norm(d[c]).str.contains(patt, regex=False, na=False)
    sub = d[mask]
    if sub.empty:
        return 0
    if unique_by_id and "ID" in sub.columns:
        return int(sub["ID"].nunique())
    return int(len(sub))

#================= MANUAL OVERRIDES (table + donuts <=2024) ===================
# Keys map to row indicators for consistency.
KEY1 = "total_ft"
KEY2 = "postdoc"
KEY3 = "editorial_boards"
KEY4 = "reviewers"
KEY5 = "boards_directors"
KEY6 = "teaching_credit_abroad"
KEY7 = "teaching_nondegree_abroad"
KEY8 = "admin_positions"
KEY9 = "execed"

MANUAL_OVERRIDE = {
    2022: { KEY3: 12, KEY4: 37, KEY5: 15, KEY6: 13, KEY7: 8 },
    2024: { KEY2: 5,  KEY3: 24, KEY4: 15, KEY5: 13, KEY6: 12, KEY7: 5 }
}

# From 2025 on, only DB-derived values (no override) for these keys:
AUTO_KEYS_2025_ON = {KEY1, KEY2, KEY3, KEY4, KEY5, KEY6, KEY7}

def apply_override(year: int, key: str, computed: Optional[int]) -> Optional[int]:
    if year >= 2025 and key in AUTO_KEYS_2025_ON:
        return computed
    v = MANUAL_OVERRIDE.get(year, {}).get(key, None)
    return v if v is not None else computed

# Persistent manual entries for rows 8 & 9 (admin / execed)
if "manual_admin" not in st.session_state:
    st.session_state.manual_admin = {y: None for y in YEARS}
    st.session_state.manual_admin[2024] = 8
    st.session_state.manual_admin[2025] = 17
if "manual_execed" not in st.session_state:
    st.session_state.manual_execed = {y: None for y in YEARS}
    st.session_state.manual_execed[2025] = 46

with st.sidebar:
    st.markdown("---")
    y_edit = st.selectbox("Year to edit (rows 8 and 9)", YEARS, index=len(YEARS)-1)
    edit_admin = st.checkbox("Edit Administrative Positions", value=(st.session_state.manual_admin[y_edit] is not None))
    if edit_admin:
        st.session_state.manual_admin[y_edit] = st.number_input(
            "Administrative Positions value", min_value=0, step=1,
            value=st.session_state.manual_admin[y_edit] if st.session_state.manual_admin[y_edit] is not None else 0
        )
    edit_exec = st.checkbox("Edit ExecEd", value=(st.session_state.manual_execed[y_edit] is not None))
    if edit_exec:
        st.session_state.manual_execed[y_edit] = st.number_input(
            "ExecEd value", min_value=0, step=1,
            value=st.session_state.manual_execed[y_edit] if st.session_state.manual_execed[y_edit] is not None else 0
        )

#================= TABLE BUILD (2020–2025) ====================================
ROWS = [
    "Total Full-time Faculty",
    "Number of Full-time Faculty with Postdoc",
    "Number of Faculty in Editorial Boards",
    "Number of Reviewers in Academic Journals",
    "Number of Faculty in Boards of Directors",
    "Number Teaching Credit Granting Courses Abroad",
    "Number Teaching Non-degree Credit Granting Courses Abroad",
    "Number of Faculty in Administrative Positions",
    "Number of Full-time Faculty Teaching in Executive Education"
]

def compute_metrics_for_year(y: int) -> Dict[str, Optional[int]]:
    """Single source of truth for both the table and the donuts."""
    c1 = apply_override(y, KEY1, ft_second_sem_count(df_full, y))
    c2 = apply_override(y, KEY2, count_yes(df_q, "Q6", y))
    c3a = count_contains(df_q, "Q36", "editorial", y, True)
    c3b = count_contains(df_q, "Q36", "editioral", y, True)
    c3  = None if (c3a is None and c3b is None) else (int(c3a or 0) + int(c3b or 0))
    c3  = apply_override(y, KEY3, c3)
    c4  = apply_override(y, KEY4, count_contains(df_q, "Q36", "journal", y, True))
    c5  = apply_override(y, KEY5, count_contains(df_q, "Q5",  "director", y, True))
    c6  = apply_override(y, KEY6, count_yes(df_q, "Q16", y))
    c7  = apply_override(y, KEY7, count_yes(df_q, "Q18", y))
    c8  = st.session_state.manual_admin.get(y)
    c9  = st.session_state.manual_execed.get(y)
    return {
        KEY1: c1, KEY2: c2, KEY3: c3, KEY4: c4, KEY5: c5, KEY6: c6, KEY7: c7, KEY8: c8, KEY9: c9
    }

data = { "Indicator": ROWS }
for y in YEARS:
    m = compute_metrics_for_year(y)
    def show(v):
        return "" if (v is None or (isinstance(v, (int, float, np.integer, np.floating)) and float(v) == 0.0)) else int(v)
    data[str(y)] = [
        show(m[KEY1]), show(m[KEY2]), show(m[KEY3]), show(m[KEY4]), show(m[KEY5]),
        show(m[KEY6]), show(m[KEY7]), show(m[KEY8]), show(m[KEY9])
    ]
table = pd.DataFrame(data)

#================= LAYOUT: LEFT (TABLE) / RIGHT (DONUTS) ======================
colL, colR = st.columns([7,5], gap="large")

#================= LEFT: TABLE ===============================================
with colL:
    st.subheader("Summary 2020–2025")

    def _style(df_):
        styles = pd.DataFrame('', index=df_.index, columns=df_.columns)
        year_cols = [c for c in df_.columns if c.isdigit()]
        last_year = str(max(int(c) for c in year_cols)) if year_cols else None
        if len(df_.index) > 0 and year_cols:
            styles.loc[df_.index[0], year_cols] += 'font-weight:800;'
        if last_year is not None and len(df_.index) > 0:
            styles.loc[df_.index[0], last_year] += 'background-color:#E8FAF7; color:#21877D; font-weight:800;'
        return styles

    styled = table.style.apply(_style, axis=None).hide(axis="index")
    st.dataframe(styled, use_container_width=True, height=48 + 33*(len(table)+1), hide_index=True)
    _download_link(
        "Descargar Excel Summary 2020–2025",
        table,  # la tabla base (no el Styler)
        f"summary_{YEARS[0]}_{YEARS[-1]}.xlsx"
    )

#================= RIGHT: RESPONSE KPI + DONUTS ===============================
with colR:
    # Year nav
    if "year_idx" not in st.session_state:
        st.session_state.year_idx = len(YEARS) - 1
    cL, cC, cR = st.columns([1, 3, 1])
    with cL:
        if st.button("◀", key="yr_prev"):
            if st.session_state.year_idx > 0:
                st.session_state.year_idx -= 1
    with cR:
        if st.button("▶", key="yr_next"):
            if st.session_state.year_idx < len(YEARS) - 1:
                st.session_state.year_idx += 1
    y_sel = YEARS[st.session_state.year_idx]
    with cC:
        st.markdown(f"<div style='text-align:center;font-weight:800'>Year: {y_sel}</div>", unsafe_allow_html=True)

    # Response rate (simple % of respondents over total professors)
    d_y, ycol = _year_filter(df_q, y_sel)
    n_resp = int(d_y["ID"].nunique()) if (ycol and "ID" in d_y.columns) else int(len(d_y))
    rate = (n_resp / TOT_PROFESSORS * 100.0) if TOT_PROFESSORS else 0.0
    st.markdown(f"### Response rate: {rate:.1f}%")

    # Shared legend (YES/NO)
    st.markdown(
        f"<div class='legend-center'>"
        f"<div class='legend-item'><span class='legend-swatch' style='background:{MINT}'></span> YES</div>"
        f"<div class='legend-item'><span class='legend-swatch' style='background:{GREY}'></span> NO</div>"
        f"</div>",
        unsafe_allow_html=True
    )

    # Donut helper
    def donut_fig(title: str, yes_count: Optional[int], total: int = TOT_PROFESSORS, height: int = DONUT_H):
        yv = int(yes_count or 0)
        nv = max(total - yv, 0)
        dfp = pd.DataFrame({"Status": ["YES", "NO"], "Value": [yv, nv]})
        fig = px.pie(
            dfp, names="Status", values="Value", hole=0.65,
            title=title, color="Status",
            color_discrete_map={"YES": MINT, "NO": GREY}
        )
        fig.update_traces(textinfo="none", hovertemplate="%{label}: %{value} of " + str(total))
        pct = (yv / total * 100.0) if total else 0.0
        fig.add_annotation(x=0.5, y=0.5, text=f"{pct:.0f}%", showarrow=False,
                           font=dict(size=18, color=MINT_DARK))
        fig.update_layout(margin=dict(l=6,r=6,t=26,b=6), showlegend=False,
                          height=height, title_font_size=12)
        return fig

    # >>>>>> Donut values come from the SAME source as the table <<<<<<
    metrics_sel = compute_metrics_for_year(y_sel)
    postdoc   = metrics_sel[KEY2] or 0
    editorial = metrics_sel[KEY3] or 0
    reviewers = metrics_sel[KEY4] or 0
    boards    = metrics_sel[KEY5] or 0
    admin     = metrics_sel[KEY8] or 0
    execed    = metrics_sel[KEY9] or 0

    c1, c2, c3 = st.columns(3)
    with c1:
        st.plotly_chart(donut_fig("Faculty with Postdoc", postdoc), use_container_width=True)
        st.plotly_chart(donut_fig("Reviewers in Journals", reviewers), use_container_width=True)
    with c2:
        st.plotly_chart(donut_fig("Editorial Boards", editorial), use_container_width=True)
        st.plotly_chart(donut_fig("Boards of Directors", boards), use_container_width=True)
    with c3:
        st.plotly_chart(donut_fig("Faculty in Administrative Positions", admin), use_container_width=True)
        st.plotly_chart(donut_fig("Faculty Teaching in ExecEd",       execed), use_container_width=True)

#================= COURSE TABLES (from dedicated sheets) ======================
st.markdown("---")
st.subheader("Courses taught by Full-time Faculty Abroad")

# Reuse selected year
y_sel = YEARS[st.session_state.year_idx]

def _extract_year_from_text(v) -> Optional[int]:
    if pd.isna(v): return None
    if isinstance(v, (int, np.integer)): return int(v)
    if isinstance(v, (float, np.floating)) and np.isfinite(v): return int(v)
    m = re.search(r'(19|20)\d{2}', str(v))
    return int(m.group(0)) if m else None

def _find_fullname(df: pd.DataFrame) -> Optional[str]:
    for cand in ["Full name", "Full Name", "Fullname", "Name", "Faculty", "Professor", "Profesor", "Nombre"]:
        col = resolve_column(df, cand)
        if col: return col
    f = resolve_column(df, "First Name")
    l = resolve_column(df, "Last Name")
    if f and l:
        df["__FULLNAME__"] = df[f].fillna("").astype(str).str.strip() + " " + df[l].fillna("").astype(str).str.strip()
        return "__FULLNAME__"
    return None

def flatten_granted_courses_precise(df_src: pd.DataFrame, y_sel: int) -> pd.DataFrame:
    """Return rows (Professor, Course, University, Year delivered) filtered by y_sel, using 'Please specify - ...' columns."""
    if df_src.empty:
        return pd.DataFrame(columns=["Professor", "Course", "University", "Year delivered"])
    df = df_src.copy()
    df.columns = df.columns.str.strip()

    name_col = _find_fullname(df)
    year_col = resolve_column(df, "Year")  # not globally filtering by this; prefer Year:N when present

    low = {c: c.lower().strip() for c in df.columns}
    rx_course = re.compile(r'^please\s*specify\s*-\s*course\s*(?:name|title)\s*(?::\s*|\s+)?(\d+)?\s*:?\s*$', re.I)
    rx_uni    = re.compile(r'^please\s*specify\s*-\s*university\s*(?::\s*|\s+)?(\d+)?\s*:?\s*$', re.I)
    rx_year_i = re.compile(r'^please\s*specify\s*-\s*year\s*(?::\s*|\s+)?(\d+)?\s*:?\s*$', re.I)

    course_cols, uni_cols, year_item_cols = {}, {}, {}
    next_course_idx = 1
    next_uni_idx    = 1

    for col, lw in low.items():
        mc = rx_course.match(lw)
        if mc:
            n = int(mc.group(1)) if mc.group(1) else next_course_idx
            course_cols[n] = col
            if not mc.group(1): next_course_idx += 1
            continue
        mu = rx_uni.match(lw)
        if mu:
            n = int(mu.group(1)) if mu.group(1) else next_uni_idx
            uni_cols[n] = col
            if not mu.group(1): next_uni_idx += 1
            continue
        my = rx_year_i.match(lw)
        if my:
            n = int(my.group(1)) if my.group(1) else 1
            year_item_cols[n] = col
            continue

    idxs = sorted(set(course_cols.keys()) | set(uni_cols.keys()) | set(year_item_cols.keys()))
    if not idxs:
        return pd.DataFrame(columns=["Professor", "Course", "University", "Year delivered"])

    rows = []
    for _, r in df.iterrows():
        prof = str(r.get(name_col, "")).strip() if (name_col and name_col in r.index and pd.notna(r.get(name_col))) else ""
        year_numeric = None
        if year_col and year_col in r.index and pd.notna(r.get(year_col)):
            yv = pd.to_numeric(r.get(year_col), errors="coerce")
            if pd.notna(yv): year_numeric = int(yv)

        for n in idxs:
            ccol = course_cols.get(n); ucol = uni_cols.get(n); ycol = year_item_cols.get(n) or year_item_cols.get(1)
            course = str(r.get(ccol, "")).strip() if (ccol and ccol in r.index and pd.notna(r.get(ccol))) else ""
            uni    = str(r.get(ucol, "")).strip() if (ucol and ucol in r.index and pd.notna(r.get(ucol))) else ""

            year_text_display = ""
            if ycol and ycol in r.index and pd.notna(r.get(ycol)):
                year_text_display = str(r.get(ycol)).strip()

            year_text_numeric = _extract_year_from_text(r.get(ycol)) if (ycol and ycol in r.index) else None
            year_val = year_numeric if year_numeric is not None else year_text_numeric
            if year_val is None or int(year_val) != int(y_sel):
                continue
            if not (course or uni):
                continue

            rows.append({
                "Professor": prof,
                "Course": course,
                "University": uni,
                "Year delivered": year_text_display
            })
    return pd.DataFrame(rows).drop_duplicates().reset_index(drop=True)

# Two tables side by side
cTa, cTb = st.columns(2)

def _counts(df: pd.DataFrame) -> tuple[int, int]:
    if df.empty:
        return 0, 0
    profs = df["Professor"].astype(str).str.strip()
    profs = profs[profs != ""]
    return int(profs.nunique()), int(len(df))

with cTa:
    title_credit = credit_sheet_name if credit_sheet_name else "Credit granted courses"
    if df_credit_sheet.empty:
        st.markdown(f"#### {title_credit} — 0 Faculty members taught courses abroad")
        st.info(f"Sheet '{title_credit}' was not found.")
    else:
        df_credit_flat = flatten_granted_courses_precise(df_credit_sheet, y_sel)
        n_prof, n_courses = _counts(df_credit_flat)
        st.markdown(f"#### {n_prof} Faculty members taught {n_courses} {title_credit} abroad")
        if df_credit_flat.empty:
            st.info(f"No records for {y_sel}.")
        else:
            st.dataframe(df_credit_flat, use_container_width=True, hide_index=True)

            # 1) hoja completa (sin filtrar) que alimenta la tabla:
            safe_name_credit = (credit_sheet_name or "Credit granted courses").lower().replace(" ", "_")
            _download_link("Descargar Excel Credit granted",
                        df_credit_sheet,
                        f"{safe_name_credit}_full.xlsx")

with cTb:
    title_noncr = noncredit_sheet_name if noncredit_sheet_name else "Non-credit granted courses"
    if df_noncredit_sheet.empty:
        st.markdown(f"#### {title_noncr} — 0 Faculty members taught courses abroad")
        st.info(f"Sheet '{title_noncr}' was not found.")
    else:
        df_noncr_flat = flatten_granted_courses_precise(df_noncredit_sheet, y_sel)
        n_prof_nc, n_courses_nc = _counts(df_noncr_flat)
        st.markdown(f"#### {n_prof_nc} Faculty members taught {n_courses_nc} {title_noncr} abroad")
        if df_noncr_flat.empty:
            st.info(f"No records for {y_sel}.")
        else:
            st.dataframe(df_noncr_flat, use_container_width=True, hide_index=True)

            # hoja completa (sin filtrar) que alimenta la tabla:
            safe_name_noncr = (noncredit_sheet_name or "Non-credit granted courses").lower().replace(" ", "_")
            _download_link("Descargar Excel Non-credit granted",
                        df_noncredit_sheet,

                        f"{safe_name_noncr}_full.xlsx")
