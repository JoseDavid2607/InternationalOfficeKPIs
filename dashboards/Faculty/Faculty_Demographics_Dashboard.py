# ===========================================================================
#  Faculty Demographics · UASM
# ===========================================================================
import re
import base64
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import streamlit as st

st.set_page_config(page_title="Faculty Demographics · UASM", page_icon="🎓",
                    layout="wide", initial_sidebar_state="expanded")

# ── Design tokens & global CSS ────────────────────────────────────────────
COLORS = {
    "primary": "#21877D", "primary_dark": "#004d47", "primary_light": "#dff7f2",
    "accent1": "#2EC4B6", "accent2": "#00A896", "accent3": "#56D6C9",
    "highlight": "#D0E5F5",
}

st.markdown("""
<style>
.suite-header{display:flex;flex-direction:column;margin-top:-35px;align-items:center;
  padding:16px 24px 12px;background:linear-gradient(135deg,#004d47 0%,#21877D 60%,#2EC4B6 100%);
  border-radius:12px;box-shadow:0 2px 8px rgba(0,77,71,.18);margin-bottom:14px;}
.sh-super{font-size:11px;font-weight:700;letter-spacing:2px;color:#56D6C9;
  text-transform:uppercase;margin-bottom:2px;}
.sh-title{font-size:26px;font-weight:800;color:#fff;text-align:center;line-height:1.2;}
.sh-sub{font-size:13px;color:rgba(255,255,255,.75);margin-top:4px;text-align:center;}
.sec-sep{border:none;border-top:1px solid #D1E8E4;margin:16px 0;opacity:.6;}
.period-label{text-align:center;font-weight:700;font-size:1.05rem;color:#21877D;}
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
div[data-testid='stButton'] button{background:#FFFFFF !important;border:1px solid #D1E8E4 !important;
  border-radius:10px !important;color:#374151 !important;font-size:14px !important;
  font-weight:600 !important;height:48px !important;box-shadow:0 1px 3px rgba(0,0,0,.04) !important;}
div[data-testid='stButton'] button:hover{background:#F8FFFE !important;border-color:#B7DCD6 !important;}
</style>
""", unsafe_allow_html=True)


# ── Generic helpers ────────────────────────────────────────────────────────
def xlsx_bytes(df: pd.DataFrame, sheet_name: str = "Data") -> bytes:
    import io
    buf = io.BytesIO()
    with pd.ExcelWriter(buf) as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    return buf.getvalue()


def render_header(title: str, subtitle: str = ""):
    sub = f'<div class="sh-sub">{subtitle}</div>' if subtitle else ""
    st.markdown(
        f'<div class="suite-header"><div class="sh-super">UASM · Faculty Analytics</div>'
        f'<div class="sh-title">{title}</div>{sub}</div>',
        unsafe_allow_html=True,
    )


render_header("Faculty Demographics", "PhD attainment, international diversity, and composition over time")


# ── Data loading (Google Drive) ────────────────────────────────────────────
DRIVE_FILE_ID = "1rPDVrdIxBFMrf0VkBmLtdUmbhvT4dku-"
DRIVE_URLS = [
    "https://drive.usercontent.google.com/download",  # current Google endpoint, avoids the HTML gate
    "https://drive.google.com/uc?export=download",     # legacy endpoint, kept as fallback
]


def _extract_confirm_token(resp: requests.Response) -> str | None:
    for key, value in resp.cookies.items():
        if key.startswith("download_warning"):
            return value
    text = resp.text or ""
    m = re.search(r'confirm=([0-9A-Za-z_-]+)', text) or re.search(r'name="confirm"\s+value="([0-9A-Za-z_-]+)"', text)
    return m.group(1) if m else None


def _try_download(url: str) -> bytes:
    session = requests.Session()
    params = {"id": DRIVE_FILE_ID, "export": "download", "confirm": "t"}
    response = session.get(url, params=params, stream=True)

    if "text/html" in response.headers.get("Content-Type", ""):
        token = _extract_confirm_token(response)
        if token:
            params["confirm"] = token
            response = session.get(url, params=params, stream=True)

    return response.content


@st.cache_data(ttl=300)
def download_excel() -> str | None:
    """Download the faculty workbook from Google Drive to /tmp. Returns None on failure."""
    path = "/tmp/BD_Faculty_demo.xlsx"

    for url in DRIVE_URLS:
        try:
            content = _try_download(url)
        except Exception:
            continue
        if content.startswith(b"PK"):
            with open(path, "wb") as f:
                f.write(content)
            return path

    return None


def load_excel_or_stop() -> str:
    path = download_excel()
    if path is None:
        st.error(
            "❌ No se pudo descargar el archivo de Google Drive.\n\n"
            "Esto casi siempre significa que el archivo **no está compartido "
            "públicamente**. Verifica en Google Drive:\n\n"
            "1. Clic derecho sobre `BD_Faculty.xlsx` → **Compartir**.\n"
            "2. En 'Acceso general', selecciona **'Cualquiera con el enlace'**.\n"
            "3. Rol: **Lector**.\n\n"
            f"ID de archivo usado: `{DRIVE_FILE_ID}`"
        )
        st.stop()
    return path


@st.cache_data(ttl=0)
def load_fulltime() -> pd.DataFrame:
    df = pd.read_excel(load_excel_or_stop(), sheet_name="BD_PLANTA")

    if "Semestre" in df.columns:
        sem = df["Semestre"].astype(str).str.strip()
    else:
        sem = df.iloc[:, 0].astype(str).str.strip()
    is_inter = sem.str.contains("inter", case=False, na=False)
    df["Periodo"] = np.where(is_inter, sem.str[:4] + " Intersemestral", sem.str[:4] + sem.str[-2:])

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
    df = pd.read_excel(load_excel_or_stop(), sheet_name="Faculty Distribution")

    if "PLANTA_CATEDRA" in df.columns:
        col = df["PLANTA_CATEDRA"].astype(str).str.strip()
        col = col.str.normalize("NFKD").str.encode("ascii", errors="ignore").str.decode("ascii")
        df = df[col.str.upper().eq("CATEDRA")].copy()

    sem = df["Semestre"].astype(str).str.strip()
    is_inter = sem.str.contains("inter", case=False, na=False)
    df.loc[~is_inter, "Periodo"] = sem.str[:4] + sem.str[-2:]
    df.loc[is_inter, "Periodo"] = sem.str[:4] + " Intersemestral"

    if "ID Nr." not in df.columns and "ID" in df.columns:
        df = df.rename(columns={"ID": "ID Nr."})
    if "AREA_PROFESOR" not in df.columns and "Academic Area" in df.columns:
        df["AREA_PROFESOR"] = df["Academic Area"]
    return df


df_full = load_fulltime()
df_part = load_parttime()


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
    b64_dl = base64.b64encode(xlsx_bytes(export_df)).decode()
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
