# ======================= Faculty Qualifications (full app) =======================
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import re
import math
import numpy as np
from io import BytesIO

# ------------------------ PAGE CONFIG & STYLES ------------------------
st.set_page_config(
    page_title="Faculty Qualifications",
    page_icon=":bar_chart:",
    layout="wide",
    initial_sidebar_state="expanded",
)
with st.container():
    st.markdown(
        """
        <style>
        .header-title { color:#21877D; font-weight:700; text-align:center; font-size:32px; }
        .header-btn { background-color:#21877D; padding:8px 16px; border:none; border-radius:8px; cursor:pointer; font-size:14px; display:inline-block; }
        a.header-btn, a.header-btn:link, a.header-btn:visited, a.header-btn:hover, a.header-btn:active { color:#ffffff !important; text-decoration:none !important; }
        .scroll-wrap-600 { max-height:600px; overflow-y:auto; }
        .scroll-wrap-400 { max-height:400px; overflow-y:auto; }
        .scroll-wrap-program { max-height:520px; overflow-y:auto; }
        </style>
        """,
        unsafe_allow_html=True
    )

# DOWNLOAD BUTTONS — minimal style
st.markdown("""
<style>
div.stDownloadButton > button {
  background: transparent !important;
  border: none !important;
  box-shadow: none !important;
  color: #21877D !important;
  font-size: 13px !important;
  padding: 0 !important;
  text-decoration: underline !important;
}
div.stDownloadButton { margin: 2px 0 8px 0; }
div.stDownloadButton > button:hover { opacity: 0.9; }
</style>
""", unsafe_allow_html=True)

# ------------------------ HEADER ------------------------
st.markdown('<div class="header-title">Full-time Faculty Qualifications</div>', unsafe_allow_html=True)

# ------------------------ DATA LOAD ------------------------
@st.cache_data(ttl=0)
def load_faculty_distribution():
    xls = pd.ExcelFile("data/Faculty/BD_Faculty.xlsx")
    df = pd.read_excel(xls, sheet_name="Faculty Distribution")
    df.columns = df.columns.str.strip()
    return df

@st.cache_data(ttl=0)
def load_cartelera():
    xls = pd.ExcelFile("data/Faculty/BD_Faculty.xlsx")
    df = pd.read_excel(xls, sheet_name="BD Cartelera 2020-2025")
    df.columns = df.columns.str.strip()
    return df

@st.cache_data(ttl=0)
def _load_planta_sheet():
    try:
        xls = pd.ExcelFile("data/Faculty/BD_Faculty.xlsx")
        dfp = pd.read_excel(xls, sheet_name="BD PLANTA 2020-2025")
        dfp.columns = dfp.columns.str.strip()
        return dfp
    except Exception:
        return pd.DataFrame()

df_planta = _load_planta_sheet()
df_fd  = load_faculty_distribution()
df_car = load_cartelera()

# ------------------------ CONSTANTS & HELPERS ------------------------
MINT = "#1FA89B"
SUPPORTING = "#7FD3FF"
TOTAL_SERIES_COLOR = "#D09E33"

def _resolve(df: pd.DataFrame, target: str):
    t = target.strip().casefold()
    for c in df.columns:
        if c.strip().casefold() == t:
            return c
    return None

def _norm_str(s: pd.Series) -> pd.Series:
    return s.astype(str).str.strip().str.lower()

def normalize_ps(val: str) -> str:
    v = str(val).strip().lower()
    if v in {"p","participating","participante","participating faculty"}:
        return "P"
    if v in {"s","supporting","soporte","supporting faculty"}:
        return "S"
    return ""

def normalize_tipo(val: str) -> str:
    v = str(val).strip().lower()
    if v in {"sa","scholarly academics","scholarly academic"}:
        return "SA"
    if v in {"pa","practice academics","practice academic"}:
        return "PA"
    if v in {"sp","scholarly practitioners","scholarly practitioner"}:
        return "SP"
    if v in {"ip","instructional practitioners","instructional practitioner"}:
        return "IP"
    if v in {"o","other","others","otro","otros"}:
        return "OTHER"
    m = re.search(r"\b(sa|pa|sp|ip|o|other)\b", v)
    if m:
        code = m.group(1).upper()
        return "OTHER" if code in {"O","OTHER"} else code
    return "OTHER"

def _get_any(df: pd.DataFrame, *cands) -> str | None:
    for c in cands:
        got = _resolve(df, c)
        if got:
            return c
    return None

def extract_year_from_period(p: str) -> int | None:
    if p is None:
        return None
    m = re.search(r"(19|20)\d{2}", str(p))
    return int(m.group(0)) if m else None

def period_suffix(p: str) -> str | None:
    m = re.search(r"(?:19|20)\d{2}[-_/ ]?(\d+)", str(p))
    return m.group(1) if m else None

def is_regular_period(p) -> bool:
    s = str(p).strip().lower()
    if "inter" in s:
        return False
    suf = period_suffix(s)
    return (suf in {"10", "20"}) or (suf is None)

def list_periods_semestral():
    sem_col = _get_any(df_car, "Semestre", "Periodo", "Periodo Académico", "Periodo academico")
    vals = []
    if sem_col:
        vals = df_car[sem_col].dropna().astype(str).str.strip().tolist()
    regs = [v for v in vals if is_regular_period(v) and period_suffix(v) in {"10","20"}]
    def sort_key(p):
        y = extract_year_from_period(p) or -1
        suf = int(period_suffix(p) or 0)
        return (y, suf)
    return sorted(sorted(set(regs)), key=sort_key)

def list_years_from_sem():
    sem_col = _get_any(df_car, "Semestre", "Periodo", "Periodo Académico", "Periodo academico")
    years = set()
    if sem_col:
        for s in df_car[sem_col].dropna().astype(str):
            y = extract_year_from_period(s)
            if y:
                years.add(y)
    ycol_fd = _get_any(df_fd, "Year", "Año")
    if ycol_fd:
        for y in pd.to_numeric(df_fd[ycol_fd], errors="coerce").dropna().astype(int):
            years.add(int(y))
    return sorted(years)

def years_with_inter():
    sem_col = _get_any(df_car, "Semestre", "Periodo", "Periodo Académico", "Periodo academico")
    inter = set()
    if sem_col:
        for s in df_car[sem_col].dropna().astype(str):
            if "inter" in s.lower():
                y = extract_year_from_period(s)
                if y:
                    inter.add(y)
    return sorted(inter)

def _slugify(s: str) -> str:
    return re.sub(r'[^A-Za-z0-9]+', '_', str(s)).strip('_')

# —— utilidades de descarga ——
def _sanitize_for_export(df: pd.DataFrame) -> pd.DataFrame:
    return df[[c for c in df.columns if not str(c).startswith("_")]].copy()

def _xlsx_bytes(df: pd.DataFrame, sheet_name: str = "Data") -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf) as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    buf.seek(0)
    return buf.getvalue()

def _download_xlsx_button(df: pd.DataFrame, fname: str, key: str, label: str = "Download Excel"):
    safe = _sanitize_for_export(df)
    clean = re.sub(r"[^\w\sÁÉÍÓÚÜÑáéíóúüñ().%/-]+", "", label).strip()
    st.download_button(
        clean,
        data=_xlsx_bytes(safe),
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=key,
        use_container_width=False
    )

# ================== SENSITIVITY HELPERS ==================
def build_member_list_for_view(df_period: pd.DataFrame, view_mode: str, col_areaCourse, col_field, program_col) -> list[str]:
    if view_mode == "By Academic Area" and col_areaCourse:
        items = sorted(df_period[col_areaCourse].astype(str).str.strip().dropna().unique().tolist())
    elif view_mode == "By Field" and col_field:
        items = sorted(df_period[col_field].astype(str).str.strip().dropna().unique().tolist())
    elif view_mode == "By Program" and program_col:
        items = sorted(df_period[program_col].astype(str).str.strip().dropna().unique().tolist())
    else:
        items = []
    return ["All"] + items

def apply_ops_to_aggs(agg_ps: pd.DataFrame, agg_tipo: pd.DataFrame, ops: list, member_all_label="All") -> tuple[pd.DataFrame, pd.DataFrame]:
    mod_ps = agg_ps.copy()
    mod_tipo = agg_tipo.copy()
    for op in ops or []:
        scope = op.get("scope")
        cat = op.get("cat")
        member = op.get("member", member_all_label)
        delta = float(op.get("credits", 0.0)) * int(op.get("count", 0))
        if delta == 0:
            continue
        if scope == "PS":
            if cat not in ["P", "S"]:
                continue
            if cat not in mod_ps.columns:
                mod_ps[cat] = 0.0
            if member == member_all_label:
                mod_ps[cat] = (mod_ps[cat] + delta).clip(lower=0.0)
            else:
                if member in mod_ps.index:
                    mod_ps.at[member, cat] = max(0.0, float(mod_ps.at[member, cat]) + delta)
        elif scope == "QUAL":
            cats = ["SA","SP","IP","PA","OTHER"]
            if cat not in cats:
                continue
            if cat not in mod_tipo.columns:
                mod_tipo[cat] = 0.0
            if member == member_all_label:
                mod_tipo[cat] = (mod_tipo[cat] + delta).clip(lower=0.0)
            else:
                if member in mod_tipo.index:
                    mod_tipo.at[member, cat] = max(0.0, float(mod_tipo.at[member, cat]) + delta)
    return mod_ps, mod_tipo

# ===== Cálculo de “profesores necesarios” (3 créditos) por fila =====
def _needed_for_pctP(p: float, s: float, target_pct: float, credits_each: float = 3.0) -> int:
    # (p + c*n)/(p + s + c*n) >= t  ->  n >= (t*s - (1-t)*p) / (c*(1-t))
    t = target_pct / 100.0
    denom = credits_each * (1 - t)
    if denom <= 0:
        return 0
    rhs = (t * s - (1 - t) * p) / denom
    return max(0, math.ceil(rhs))

def _needed_for_pctSA(sa: float, rest: float, target_pct: float, credits_each: float = 3.0) -> int:
    # (sa + c*n)/(sa + rest + c*n) >= t -> n >= (t*rest - (1-t)*sa) / (c*(1-t))
    t = target_pct / 100.0
    denom = credits_each * (1 - t)
    if denom <= 0:
        return 0
    rhs = (t * rest - (1 - t) * sa) / denom
    return max(0, math.ceil(rhs))

def _needed_for_other_leq10(other: float, rest: float, credits_each: float = 3.0) -> int:
    # other/(other + rest + c*n) <= 0.10  ->  n >= (0.9*other - 0.1*rest) / (0.3*c) = (9*other - rest)/(3*c)
    num = 9*other - rest
    denom = 3 * credits_each
    if denom <= 0:
        return 0
    rhs = num / denom
    return max(0, math.ceil(rhs))

def _objective_targets(obj: str) -> tuple[str, float, float]:
    # devuelve etiqueta y targets por scope (by_area, overall)
    if obj == "%P":   return ("%P", 60.0, 75.0)
    if obj == "%SA":  return ("%SA", 40.0, 40.0)
    return ("%OTHER", 10.0, 10.0)

# ====== NUEVOS helpers para Overall/Impacto y secundarios ======
def _needed_for_overall_if_only_this_area_changes(obj: str, totals: dict[str, float], area_vals: dict[str, float], target_overall: float, credits_each: float = 3.0) -> int | None:
    eps = 1e-9
    t = target_overall / 100.0
    Ptot = totals.get("P",0.0);  Stot = totals.get("S",0.0)
    SA   = totals.get("SA",0.0); PA  = totals.get("PA",0.0)
    SP   = totals.get("SP",0.0); IP  = totals.get("IP",0.0)
    OT   = totals.get("OTHER",0.0)
    TQ   = SA + PA + SP + IP + OT
    if obj == "%P":
        den = Ptot + Stot
        if den <= eps: return 0
        rhs = (t*den - Ptot) / (credits_each*(1 - t))
        return max(0, math.ceil(rhs))
    if obj == "%SA":
        if TQ <= eps: return 0
        rhs = (t*TQ - SA) / (credits_each*(1 - t))
        return max(0, math.ceil(rhs))
    if TQ <= eps: return 0
    need_credits = (OT - 0.10*TQ) / 0.90
    need_n = 0 if need_credits <= 0 else math.ceil(need_credits / credits_each)
    OT_a = area_vals.get("OTHER", 0.0)
    max_remove_n = math.floor(OT_a / credits_each)
    return max(0, need_n) if need_n <= max_remove_n else None

def _impact_pp_area(obj: str, area_vals: dict[str,float], credits_each: float = 3.0) -> tuple[float,float]:
    eps = 1e-9
    P = area_vals.get("P",0.0); S = area_vals.get("S",0.0)
    SA = area_vals.get("SA",0.0); PA = area_vals.get("PA",0.0)
    SP = area_vals.get("SP",0.0); IP = area_vals.get("IP",0.0)
    OT = area_vals.get("OTHER",0.0)
    denPS = P + S
    denQ  = SA + PA + SP + IP + OT
    if obj == "%P":
        if denPS <= eps: return (0.0, 0.0)
        up   = ((P + credits_each) / (denPS + credits_each) - (P / denPS)) * 100.0
        down = ((max(0.0, P - credits_each)) / max(eps, denPS - credits_each) - (P / denPS)) * 100.0 if denPS > credits_each else 0.0
        return (round(up,2), round(down,2))
    if obj == "%SA":
        if denQ <= eps: return (0.0, 0.0)
        up   = ((SA + credits_each) / (denQ + credits_each) - (SA / denQ)) * 100.0
        down = ((max(0.0, SA - credits_each)) / max(eps, denQ - credits_each) - (SA / denQ)) * 100.0 if denQ > credits_each else 0.0
        return (round(up,2), round(down,2))
    if denQ <= eps: return (0.0, 0.0)
    up   = ((OT + credits_each) / (denQ + credits_each) - (OT / denQ)) * 100.0
    down = ((max(0.0, OT - credits_each)) / max(eps, denQ - credits_each) - (OT / denQ)) * 100.0 if denQ > credits_each else 0.0
    return (round(up,2), round(down,2))

def _impact_pp_overall_if_area_changes(obj: str, totals: dict[str,float], credits_each: float = 3.0) -> tuple[float,float]:
    eps = 1e-9
    P = totals.get("P",0.0); S = totals.get("S",0.0)
    SA = totals.get("SA",0.0); PA = totals.get("PA",0.0)
    SP = totals.get("SP",0.0); IP = totals.get("IP",0.0)
    OT = totals.get("OTHER",0.0)
    denPS = P + S
    denQ  = SA + PA + SP + IP + OT
    if obj == "%P":
        if denPS <= eps: return (0.0, 0.0)
        up   = ((P + credits_each) / (denPS + credits_each) - (P / denPS)) * 100.0
        down = ((max(0.0, P - credits_each)) / max(eps, denPS - credits_each) - (P / denPS)) * 100.0 if denPS > credits_each else 0.0
        return (round(up,2), round(down,2))
    if obj == "%SA":
        if denQ <= eps: return (0.0, 0.0)
        up   = ((SA + credits_each) / (denQ + credits_each) - (SA / denQ)) * 100.0
        down = ((max(0.0, SA - credits_each)) / max(eps, denQ - credits_each) - (SA / denQ)) * 100.0 if denQ > credits_each else 0.0
        return (round(up,2), round(down,2))
    if denQ <= eps: return (0.0, 0.0)
    up   = ((OT + credits_each) / (denQ + credits_each) - (OT / denQ)) * 100.0
    down = ((max(0.0, OT - credits_each)) / max(eps, denQ - credits_each) - (OT / denQ)) * 100.0 if denQ > credits_each else 0.0
    return (round(up,2), round(down,2))

# Secundarios para tablas "Needed"
def _needed_S_less_for_pctP_area(p: float, s: float, target_pct: float, credits_each=3.0) -> int:
    # P/(P + S - c*n) >= t  ->  n >= (t*(P+S) - P)/(t*c)
    t = target_pct/100.0
    if t <= 0: return 0
    den = t*credits_each
    rhs = (t*(p+s) - p)/den
    return max(0, math.ceil(rhs))

def _needed_S_less_for_pctP_overall(totals, area_vals, target_overall: float, credits_each=3.0) -> int | None:
    t = target_overall/100.0
    Ptot = totals.get("P",0.0); Stot = totals.get("S",0.0)
    if t <= 0: return 0
    need = (t*(Ptot+Stot) - Ptot) / (t*credits_each)
    need_n = 0 if need <= 0 else math.ceil(need)
    S_a = area_vals.get("S",0.0)
    max_remove = math.floor(S_a/credits_each)
    return need_n if need_n <= max_remove else None

def _needed_OTHERS_less_for_SA_area(sa, rest, target_pct, credits_each=3.0) -> int:
    # SA/(SA + rest - c*n) >= t -> n >= (t*(SA+rest) - SA)/(t*c)
    t = target_pct/100.0
    if t <= 0: return 0
    rhs = (t*(sa+rest) - sa)/(t*credits_each)
    return max(0, math.ceil(rhs))

def _needed_OTHERS_less_for_SA_overall(totals, area_vals, target_overall, credits_each=3.0) -> int | None:
    SA = totals.get("SA",0.0); PA=totals.get("PA",0.0); SP=totals.get("SP",0.0); IP=totals.get("IP",0.0); OT=totals.get("OTHER",0.0)
    TQ = SA+PA+SP+IP+OT; rest = TQ - SA
    t = target_overall/100.0
    if t <= 0: return 0
    need = (t*(SA+rest) - SA)/(t*credits_each)
    need_n = 0 if need <= 0 else math.ceil(need)
    rest_a = max(0.0, area_vals.get("PA",0.0)+area_vals.get("SP",0.0)+area_vals.get("IP",0.0)+area_vals.get("OTHER",0.0))
    max_remove = math.floor(rest_a/credits_each)
    return need_n if need_n <= max_remove else None

def _needed_OTHERS_more_for_OTHER_area(other, rest, target_pct, credits_each=3.0) -> int:
    # OTHER/(OTHER + rest + c*n) <= t -> c*n >= (OTHER - t*(OTHER+rest))/t
    t = target_pct/100.0
    if t <= 0: return 0
    need_credits = (other - t*(other+rest))/t
    need = 0 if need_credits <= 0 else math.ceil(need_credits/credits_each)
    return max(0, need)

def _needed_OTHERS_more_for_OTHER_overall(totals, target_overall, credits_each=3.0) -> int:
    OT = totals.get("OTHER",0.0); SA=totals.get("SA",0.0); PA=totals.get("PA",0.0); SP=totals.get("SP",0.0); IP=totals.get("IP",0.0)
    TQ = SA+PA+SP+IP+OT
    t = target_overall/100.0
    if t <= 0: return 0
    need_credits = (OT - t*TQ)/t
    return 0 if need_credits <= 0 else math.ceil(need_credits/credits_each)

# Heatmap de impacto (+Δ p.p.)
def _style_impact_heatmap(df_: pd.DataFrame, label_col: str, value_col: str, overall_mode: bool = False):
    sty = pd.DataFrame('', index=df_.index, columns=df_.columns)
    vals = pd.to_numeric(df_[value_col], errors="coerce")
    is_total = df_[label_col].astype(str).str.upper().eq("TOTAL")

    if overall_mode:
        # todo naranja en Overall
        sty[value_col] = 'background-color:#FFA500;'
        return sty

    base = vals[~is_total]
    if len(base) == 0 or base.max() == base.min():
        sty[value_col] = 'background-color:#FFA500;'
        return sty

    vmin, vmax = float(base.min()), float(base.max())

    def color_for(v):
        if pd.isna(v): return ''
        if vmin == vmax: return 'background-color:#FFA500;'
        t = (v - vmin)/(vmax - vmin)
        if t < 0.5:
            g = int(77 + (221-77)*(t/0.5))   # rojo->amarillo
            return f'background-color: rgb(255,{g},77);'
        else:
            r = int(255 - (178)*((t-0.5)/0.5))
            g = int(221 - (37)*((t-0.5)/0.5))
            return f'background-color: rgb({r},{g},77);'

    for i in df_.index:
        if is_total.loc[i]:
            sty.at[i, value_col] = 'background-color:#FFA500;'
        else:
            sty.at[i, value_col] = color_for(vals.loc[i])

    return sty

# ================== HISTORY (timeframe-aware) ==================
def _period_sort_key(p: str) -> tuple[int,int]:
    y = extract_year_from_period(p) or -1
    suf = period_suffix(p)
    try:
        suf_i = int(suf) if suf is not None else 0
    except Exception:
        suf_i = 0
    return (y, suf_i)


def build_time_axis_for_history(df_hist: pd.DataFrame):
    time_mode = st.session_state.get("time_mode", "Semestral")
    if "_SEM" not in df_hist.columns:
        sc = _get_any(df_hist, "Semestre","Periodo","Periodo Académico","Periodo academico")
        sem = df_hist[sc].astype(str).str.strip() if sc else pd.Series([], dtype=str)
    else:
        sem = df_hist["_SEM"].astype(str).str.strip()
    if time_mode == "Semestral":
        regs = sorted(
            {s for s in sem.dropna().unique() if period_suffix(s) in {"10","20"}},
            key=_period_sort_key
        )
        x_labels = regs
    elif time_mode == "Anual":
        years = sorted({extract_year_from_period(s) for s in sem if extract_year_from_period(s)}, key=int)
        x_labels = years
    else:  # Intersemestral
        inter = sorted(
            {f"{extract_year_from_period(s)} Intersemestral" for s in sem if "inter" in str(s).lower() and extract_year_from_period(s)},
            key=lambda x: int(str(x).split()[0])
        )
        x_labels = inter
    x_map = {lab: i for i, lab in enumerate(x_labels)}
    return "_SEM", x_labels, x_map


def transform_for_time_mode_ps(df_ps: pd.DataFrame):
    time_mode = st.session_state.get("time_mode", "Semestral")
    base = df_ps.copy()
    base["_YEAR"] = base["_SEM"].map(extract_year_from_period)
    base["_INTER_LABEL"] = base["_SEM"].map(lambda s: f"{extract_year_from_period(s)} Intersemestral" if "inter" in str(s).lower() else None)
    if time_mode == "Semestral":
        return base
    if time_mode == "Anual":
        need_cols = [c for c in base.columns if c not in {"P_share"}]
        g = base[need_cols].groupby(["_YEAR"] + [c for c in base.columns if c.startswith("_") and c not in {"_SEM","_YEAR","_INTER_LABEL"}], dropna=False).sum(numeric_only=True).reset_index()
        if "P" in g and "S" in g:
            g["P_share"] = (g["P"] / (g["P"] + g["S"]).replace(0, pd.NA)) * 100
        return g.rename(columns={"_YEAR":"_SEM"})
    # Intersemestral
    base = base[~base["_INTER_LABEL"].isna()].copy()
    g = base.groupby(["_INTER_LABEL"] + [c for c in base.columns if c.startswith("_") and c not in {"_SEM","_YEAR","_INTER_LABEL"}], dropna=False).sum(numeric_only=True).reset_index()
    if "P" in g and "S" in g:
        g["P_share"] = (g["P"] / (g["P"] + g["S"]).replace(0, pd.NA)) * 100
    return g.rename(columns={"_INTER_LABEL":"_SEM"})


def transform_for_time_mode_tipo(df_tipo: pd.DataFrame, share_col_name: str):
    time_mode = st.session_state.get("time_mode", "Semestral")
    base = df_tipo.copy()
    base["_YEAR"] = base["_SEM"].map(extract_year_from_period)
    base["_INTER_LABEL"] = base["_SEM"].map(lambda s: f"{extract_year_from_period(s)} Intersemestral" if "inter" in str(s).lower() else None)
    cats = ["SA","PA","SP","IP","OTHER"]
    if time_mode == "Semestral":
        return base
    if time_mode == "Anual":
        keys = ["_YEAR"] + [c for c in base.columns if c.startswith("_") and c not in {"_SEM","_YEAR","_INTER_LABEL"}]
        g = base.groupby(keys, dropna=False)[cats].sum().reset_index()
        den = (g[cats].sum(axis=1)).replace(0, pd.NA)
        if share_col_name == "SA_share":
            g["SA_share"] = (g["SA"] / den) * 100
        else:
            g["OTHER_share"] = (g["OTHER"] / den) * 100
        return g.rename(columns={"_YEAR":"_SEM"})
    # Intersemestral
    base = base[~base["_INTER_LABEL"].isna()].copy()
    keys = ["_INTER_LABEL"] + [c for c in base.columns if c.startswith("_") and c not in {"_SEM","_YEAR","_INTER_LABEL"}]
    g = base.groupby(keys, dropna=False)[cats].sum().reset_index()
    den = (g[cats].sum(axis=1)).replace(0, pd.NA)
    if share_col_name == "SA_share":
        g["SA_share"] = (g["SA"] / den) * 100
    else:
        g["OTHER_share"] = (g["OTHER"] / den) * 100
    return g.rename(columns={"_INTER_LABEL":"_SEM"})


# === aplicar sensibilidad sobre series históricas SOLO en el período seleccionado ===
def apply_sensitivity_to_history(
    agg_ps_tm: pd.DataFrame,
    agg_tipo_tm: pd.DataFrame,
    tot_ps_tm: pd.DataFrame,
    tot_tipo_tm: pd.DataFrame,
    level_name: str,
    sel_label_value,  # etiqueta exacta seleccionada en el eje X
    ops: list,
    member_all_label="All"
):
    if not ops or sel_label_value is None:
        return agg_ps_tm, agg_tipo_tm, tot_ps_tm, tot_tipo_tm

    ps = agg_ps_tm.copy()
    tq = agg_tipo_tm.copy()
    tps = tot_ps_tm.copy()
    ttq = tot_tipo_tm.copy()

    for k in ["P","S"]:
        if k not in ps.columns:  ps[k] = 0.0
        if k not in tps.columns: tps[k] = 0.0
    for k in ["SA","PA","SP","IP","OTHER"]:
        if k not in tq.columns:  tq[k] = 0.0
        if k not in ttq.columns: ttq[k] = 0.0

    mask_period_ps  = ps["_SEM"].eq(sel_label_value)
    mask_period_tq  = tq["_SEM"].eq(sel_label_value)
    mask_period_tps = tps["_SEM"].eq(sel_label_value)
    mask_period_ttq = ttq["_SEM"].eq(sel_label_value)

    for op in ops:
        scope = op.get("scope")
        cat   = op.get("cat")
        member= op.get("member", member_all_label)
        delta = float(op.get("credits", 0.0)) * int(op.get("count", 0))
        if delta == 0:
            continue

        if scope == "PS" and cat in ["P","S"]:
            if member == member_all_label:
                ps.loc[mask_period_ps, cat]  = (ps.loc[mask_period_ps, cat].astype(float)  + delta).clip(lower=0.0)
            else:
                m = mask_period_ps & ps[level_name].eq(member)
                ps.loc[m, cat] = (ps.loc[m, cat].astype(float) + delta).clip(lower=0.0)
            tps.loc[mask_period_tps, cat] = (tps.loc[mask_period_tps, cat].astype(float) + delta).clip(lower=0.0)

        if scope == "QUAL" and cat in ["SA","PA","SP","IP","OTHER"]:
            if member == member_all_label:
                tq.loc[mask_period_tq, cat]  = (tq.loc[mask_period_tq, cat].astype(float)  + delta).clip(lower=0.0)
            else:
                m = mask_period_tq & tq[level_name].eq(member)
                tq.loc[m, cat] = (tq.loc[m, cat].astype(float) + delta).clip(lower=0.0)
            ttq.loc[mask_period_ttq, cat] = (ttq.loc[mask_period_ttq, cat].astype(float) + delta).clip(lower=0.0)

    den_ps = (ps.loc[mask_period_ps, "P"].astype(float) + ps.loc[mask_period_ps, "S"].astype(float)).replace(0, pd.NA)
    ps.loc[mask_period_ps, "P_share"] = (ps.loc[mask_period_ps, "P"] / den_ps * 100).fillna(0.0)

    cats = ["SA","PA","SP","IP","OTHER"]
    den_q = tq.loc[mask_period_tq, cats].sum(axis=1).replace(0, pd.NA)
    if "SA_share" in tq.columns:
        tq.loc[mask_period_tq, "SA_share"]    = (tq.loc[mask_period_tq, "SA"]    / den_q * 100).fillna(0.0)
    if "OTHER_share" in tq.columns:
        tq.loc[mask_period_tq, "OTHER_share"] = (tq.loc[mask_period_tq, "OTHER"] / den_q * 100).fillna(0.0)

    den_tps = (tps.loc[mask_period_tps, "P"].astype(float) + tps.loc[mask_period_tps, "S"].astype(float)).replace(0, pd.NA)
    tps.loc[mask_period_tps, "P_share"] = (tps.loc[mask_period_tps, "P"] / den_tps * 100).fillna(0.0)

    den_ttq = ttq.loc[mask_period_ttq, cats].sum(axis=1).replace(0, pd.NA)
    if "SA_share" in ttq.columns:
        ttq.loc[mask_period_ttq, "SA_share"]    = (ttq.loc[mask_period_ttq, "SA"]    / den_ttq * 100).fillna(0.0)
    if "OTHER_share" in ttq.columns:
        ttq.loc[mask_period_ttq, "OTHER_share"] = (ttq.loc[mask_period_ttq, "OTHER"] / den_ttq * 100).fillna(0.0)

    return ps, tq, tps, ttq


# --------- Gráfica histórica ---------
def draw_history(fig_title, level_name, level_values, metric_kind, total_series_builders, agg_ps_all, agg_tipo_all, x_labels, x_map, sel_x):
    palette = px.colors.qualitative.Safe + px.colors.qualitative.Bold + px.colors.qualitative.Pastel
    color_map = {a: palette[i % len(palette)] for i, a in enumerate(level_values)}
    st.markdown(f"<h4 style='margin:0 0 6px 0; font-weight:500;'>{fig_title}</h4>", unsafe_allow_html=True)
    sel_col, radio_col = st.columns([6,4])
    options = ["(All)", "(TOTAL)"] + level_values
    with sel_col:
        opt = st.selectbox("", options, index=0, key=f"{level_name}_filter", label_visibility="collapsed")
    with radio_col:
        metric_choice = st.radio("", ["%P", "%SA", "%OTHER"], index={"%P":0, "%SA":1, "%OTHER":2}[metric_kind], horizontal=True, key=f"metric_{level_name}", label_visibility="collapsed")

    fig = go.Figure()

    if metric_choice == "%P":
        thr = 75 if opt == "(TOTAL)" else 60
        if opt == "(All)":
            for a in level_values:
                sub = agg_ps_all[(agg_ps_all[level_name] == a)].copy()
                sub["x"] = sub["_SEM"].map(x_map)
                sub = sub.sort_values("x")
                if sub.empty: continue
                fig.add_trace(go.Scatter(
                    x=sub["x"], y=sub["P_share"], mode="lines+markers", name=a,
                    marker=dict(size=6, color=color_map[a]), line=dict(width=2, color=color_map[a]),
                    hovertemplate=a + "<br>%{y:.1f}%<extra></extra>"
                ))
        elif opt == "(TOTAL)":
            sub = total_series_builders["P"].copy()
            sub["x"] = sub["_SEM"].map(x_map)
            sub = sub.sort_values("x")
            fig.add_trace(go.Scatter(
                x=sub["x"], y=sub["P_share"], mode="lines+markers", name="TOTAL",
                marker=dict(size=6, color=TOTAL_SERIES_COLOR), line=dict(width=2, color=TOTAL_SERIES_COLOR),
                hovertemplate="TOTAL<br>%{y:.1f}%<extra></extra>"
            ))
        else:
            sub = agg_ps_all[(agg_ps_all[level_name] == opt)].copy()
            sub["x"] = sub["_SEM"].map(x_map)
            sub = sub.sort_values("x")
            fig.add_trace(go.Scatter(
                x=sub["x"], y=sub["P_share"], mode="lines+markers", name=opt,
                marker=dict(size=6, color=MINT), line=dict(width=2, color=MINT),
                hovertemplate=opt + "<br>%{y:.1f}%<extra></extra>"
            ))
        y_min, bad_high = 40, False

    elif metric_choice == "%SA":
        thr = 40
        share_col = "SA_share"
        if opt == "(All)":
            for a in level_values:
                sub = agg_tipo_all[(agg_tipo_all[level_name] == a)].copy()
                sub["x"] = sub["_SEM"].map(x_map)
                sub = sub.sort_values("x")
                if sub.empty: continue
                fig.add_trace(go.Scatter(
                    x=sub["x"], y=sub[share_col], mode="lines+markers", name=a,
                    marker=dict(size=6, color=color_map[a]), line=dict(width=2, color=color_map[a]),
                    hovertemplate=a + "<br>%{y:.1f}%<extra></extra>"
                ))
        elif opt == "(TOTAL)":
            sub = total_series_builders["SA"].copy()
            sub["x"] = sub["_SEM"].map(x_map)
            sub = sub.sort_values("x")
            fig.add_trace(go.Scatter(
                x=sub["x"], y=sub[share_col], mode="lines+markers", name="TOTAL",
                marker=dict(size=6, color=TOTAL_SERIES_COLOR), line=dict(width=2, color=TOTAL_SERIES_COLOR),
                hovertemplate="TOTAL<br>%{y:.1f}%<extra></extra>"
            ))
        else:
            sub = agg_tipo_all[(agg_tipo_all[level_name] == opt)].copy()
            sub["x"] = sub["_SEM"].map(x_map)
            sub = sub.sort_values("x")
            fig.add_trace(go.Scatter(
                x=sub["x"], y=sub[share_col], mode="lines+markers", name=opt,
                marker=dict(size=6, color=MINT), line=dict(width=2, color=MINT),
                hovertemplate=opt + "<br>%{y:.1f}%<extra></extra>"
            ))
        y_min, bad_high = 20, False

    else:  # "%OTHER"
        thr = 10
        share_col = "OTHER_share"
        if opt == "(All)":
            for a in level_values:
                sub = agg_tipo_all[(agg_tipo_all[level_name] == a)].copy()
                sub["x"] = sub["_SEM"].map(x_map)
                sub = sub.sort_values("x")
                if sub.empty: continue
                fig.add_trace(go.Scatter(
                    x=sub["x"], y=sub[share_col], mode="lines+markers", name=a,
                    marker=dict(size=6, color=color_map[a]), line=dict(width=2, color=color_map[a]),
                    hovertemplate=a + "<br>%{y:.1f}%<extra></extra>"
                ))
        elif opt == "(TOTAL)":
            sub = total_series_builders["OTHER"].copy()
            sub["x"] = sub["_SEM"].map(x_map)
            sub = sub.sort_values("x")
            fig.add_trace(go.Scatter(
                x=sub["x"], y=sub[share_col], mode="lines+markers", name="TOTAL",
                marker=dict(size=6, color=TOTAL_SERIES_COLOR), line=dict(width=2, color=TOTAL_SERIES_COLOR),
                hovertemplate="TOTAL<br>%{y:.1f}%<extra></extra>"
            ))
        else:
            sub = agg_tipo_all[(agg_tipo_all[level_name] == opt)].copy()
            sub["x"] = sub["_SEM"].map(x_map)
            sub = sub.sort_values("x")
            fig.add_trace(go.Scatter(
                x=sub["x"], y=sub[share_col], mode="lines+markers", name=opt,
                marker=dict(size=6, color=MINT), line=dict(width=2, color=MINT),
                hovertemplate=opt + "<br>%{y:.1f}%<extra></extra>"
            ))
        y_min, bad_high = 0, True
        y_max = 40

    # Zonas de referencia
    if bad_high:
        fig.update_layout(shapes=[dict(type="rect", xref="paper", yref="y", x0=0, x1=1, y0=thr, y1=100, fillcolor="#FDE2E2", opacity=0.35, layer="below", line_width=0)])
        fig.add_hline(y=thr, line_color="#F5A3A3", line_dash="dash")
    else:
        fig.update_layout(shapes=[dict(type="rect", xref="paper", yref="y", x0=0, x1=1, y0=0, y1=thr, fillcolor="#FDE2E2", opacity=0.35, layer="below", line_width=0)])
        fig.add_hline(y=thr, line_color="red", line_dash="dash")

    if sel_x is not None:
        fig.add_vrect(x0=sel_x-0.5, x1=sel_x+0.5, fillcolor="#E8FAF7", opacity=0.5, layer="below", line_width=0)

    tickvals = list(range(len(x_labels)))
    ticktext = [str(x) for x in x_labels]
    if metric_choice == "%OTHER":
        fig.update_layout(xaxis=dict(tickmode="array", tickvals=tickvals, ticktext=ticktext), yaxis=dict(range=[y_min, y_max]))
    else:
        fig.update_layout(xaxis=dict(tickmode="array", tickvals=tickvals, ticktext=ticktext), yaxis=dict(range=[y_min, 100]))
    fig.update_xaxes(title=None)
    fig.update_yaxes(title=None)
    st.plotly_chart(fig, use_container_width=True)

    # ===== Datos para descargar (lo visible) =====
    def _series_for(level_val: str, ycol: str):
        if ycol == "P_share":
            sub = agg_ps_all[(agg_ps_all[level_name] == level_val)]
        else:
            sub = agg_tipo_all[(agg_tipo_all[level_name] == level_val)]
        m = sub.set_index("_SEM")[ycol].to_dict()
        return [m.get(x, None) for x in x_labels]

    if metric_choice == "%P":
        ycol = "P_share"
        base_cols = {}
        if opt == "(All)":
            for a in level_values:
                base_cols[a] = _series_for(a, ycol)
        elif opt == "(TOTAL)":
            sub = total_series_builders["P"].set_index("_SEM")["P_share"].to_dict()
            base_cols["TOTAL"] = [sub.get(x, None) for x in x_labels]
        else:
            base_cols[opt] = _series_for(opt, ycol)
    elif metric_choice == "%SA":
        ycol = "SA_share"
        base_cols = {}
        if opt == "(All)":
            for a in level_values:
                base_cols[a] = _series_for(a, ycol)
        elif opt == "(TOTAL)":
            sub = total_series_builders["SA"].set_index("_SEM")[ycol].to_dict()
            base_cols["TOTAL"] = [sub.get(x, None) for x in x_labels]
        else:
            base_cols[opt] = _series_for(opt, ycol)
    else:
        ycol = "OTHER_share"
        base_cols = {}
        if opt == "(All)":
            for a in level_values:
                base_cols[a] = _series_for(a, ycol)
        elif opt == "(TOTAL)":
            sub = total_series_builders["OTHER"].set_index("_SEM")[ycol].to_dict()
            base_cols["TOTAL"] = [sub.get(x, None) for x in x_labels]
        else:
            base_cols[opt] = _series_for(opt, ycol)

    export_df = pd.DataFrame({"Period": x_labels, **base_cols})
    fname = f"chart_{_slugify(fig_title)}_{_slugify(metric_choice)}_{_slugify(opt)}_{_slugify(st.session_state.get('sel_label','sel'))}.xlsx"
    _download_xlsx_button(export_df, fname, key=f"dl_hist_{_slugify(fig_title)}_{metric_choice}_{_slugify(opt)}_{_slugify(st.session_state.get('sel_label','sel'))}", label="⬇️ Datos de la gráfica (Excel)")


# ============== NORMALIZACIÓN BÁSICA EN CARTELERA ==============
col_sem = _get_any(df_car, "Semestre","Periodo","Periodo Académico","Periodo academico")
if "_SEM" not in df_car.columns and col_sem:
    df_car["_SEM"] = df_car[col_sem].astype(str).str.strip()
else:
    df_car["_SEM"] = df_car.get("_SEM", pd.Series(dtype=str))
df_car["_YEAR"] = df_car["_SEM"].map(extract_year_from_period)
df_car["_IS_INTER"] = df_car["_SEM"].str.lower().str.contains("inter", na=False)


# ================== TIMEFRAME FILTERS ==================
def mask_timeframe(series_sem: pd.Series, mode: str, selected_year: int | None, selected_sem: str | None) -> pd.Series:
    s = series_sem.astype(str)
    if mode == "Semestral" and selected_sem:
        return s.str.strip().eq(str(selected_sem))
    if mode == "Anual" and selected_year is not None:
        return s.str.startswith(str(selected_year))
    if mode == "Intersemestral" and selected_year is not None:
        return s.str.startswith(str(selected_year)) & s.str.lower().str.contains("inter")
    return pd.Series([True]*len(s), index=series_sem.index)


def filter_df_car(df: pd.DataFrame, mode: str, selected_year: int | None, selected_sem: str | None) -> pd.DataFrame:
    if "_SEM" not in df.columns:
        sc = _get_any(df, "Semestre","Periodo","Periodo Académico","Periodo academico")
        if sc:
            df = df.assign(_SEM=df[sc].astype(str).str.strip())
        else:
            return df
    m = mask_timeframe(df["_SEM"], mode, selected_year, selected_sem)
    return df[m].copy()


def filter_df_fd(df: pd.DataFrame, mode: str, selected_year: int | None, selected_sem: str | None) -> pd.DataFrame:
    semc = _get_any(df, "Semestre","Periodo","Periodo Académico","Periodo academico")
    ycol = _get_any(df, "Year","Año")
    out = df.copy()
    if semc:
        sem_series = out[semc].astype(str).str.strip()
        m = mask_timeframe(sem_series, mode, selected_year, selected_sem)
        out = out[m].copy()
    elif ycol and selected_year is not None:
        out = out[pd.to_numeric(out[ycol], errors="coerce").astype("Int64") == int(selected_year)].copy()
    return out


# ================== SIDEBAR ==================
SEMESTRAL_PERIODS = list_periods_semestral()
YEARS_ALL = list_years_from_sem()
INTER_YEARS = years_with_inter()

with st.sidebar:
    st.markdown("#### Sensitivity analysis")
    sens_mode = st.toggle(
        "Enable sensitivity mode",
        value=st.session_state.get("sens_mode", False),
        key="sens_mode",
        help="Aquí podrás hacer un análisis de sensibilidad sin modificar los datos reales."
    )
    sens_member_placeholder = st.empty()
    if "sens_ops" not in st.session_state:
        st.session_state.sens_ops = []

    if sens_mode:
        st.session_state.setdefault("sens_cat_ps", "None")
        st.session_state.setdefault("sens_cat_qual", "None")
        st.selectbox("P/S category", ["None", "P", "S"], key="sens_cat_ps")
        st.selectbox("Qualification", ["None", "SA", "PA", "SP", "IP", "OTHER"], key="sens_cat_qual")
        st.number_input("Professors", min_value=1, step=1, value=1, key="sens_count")
        st.number_input("Credits per professor", min_value=0.0, step=0.5, value=3.0, key="sens_credits")

        # ADD (suma)
        if st.button("Add", use_container_width=True, key="sens_add"):
            ops_to_add = []
            member_val = st.session_state.get("sens_member", "All")
            cnt  = int(st.session_state.get("sens_count", 1))
            cred = float(st.session_state.get("sens_credits", 3.0))
            if st.session_state.get("sens_cat_ps") and st.session_state["sens_cat_ps"] != "None":
                ops_to_add.append({"scope": "PS", "cat": st.session_state["sens_cat_ps"], "member": member_val, "credits": cred, "count": cnt})
            if st.session_state.get("sens_cat_qual") and st.session_state["sens_cat_qual"] != "None":
                ops_to_add.append({"scope": "QUAL", "cat": st.session_state["sens_cat_qual"], "member": member_val, "credits": cred, "count": cnt})
            if ops_to_add:
                st.session_state.sens_ops.extend(ops_to_add)
                st.success("Added.")

        # REMOVE (resta)
        if st.button("Remove", use_container_width=True, key="sens_remove_btn"):
            ops_to_add = []
            member_val = st.session_state.get("sens_member", "All")
            cnt  = -abs(int(st.session_state.get("sens_count", 1)))
            cred = float(st.session_state.get("sens_credits", 3.0))
            if st.session_state.get("sens_cat_ps") and st.session_state["sens_cat_ps"] != "None":
                ops_to_add.append({"scope": "PS", "cat": st.session_state["sens_cat_ps"], "member": member_val, "credits": cred, "count": cnt})
            if st.session_state.get("sens_cat_qual") and st.session_state["sens_cat_qual"] != "None":
                ops_to_add.append({"scope": "QUAL", "cat": st.session_state["sens_cat_qual"], "member": member_val, "credits": cred, "count": cnt})
            if ops_to_add:
                st.session_state.sens_ops.extend(ops_to_add)
                st.success("Removed.")

        if st.button("Reset to original", use_container_width=True, key="sens_reset"):
            st.session_state.sens_ops = []
            st.success("Reset.")

    if not sens_mode:
        st.markdown("### 📊 Go to KPI:")
        options = {
            "1 Full-time Composition": "https://facultycompositiondashboardpy-dtacyzfa3otmpbewqc5axu.streamlit.app/",
            "2 Full-time Staffing Levels": "https://facultystaffinglevelsdashboardpy-phv4t8jzbyyz5rrepqttuf.streamlit.app/",
            "3 Distribution by Academic Area": "https://facultydistributionareadashboardpy-yzwpiqdlukfdp6qcygxjhj.streamlit.app/",
            "4 Faculty Demographics": "https://facultydemographicsdashboardpy-kmsnpswxs35psbqtdtvb6y.streamlit.app/",
            "5 Full-time Faculty Questionnaire": "https://full-timefacultyactivitiespy-bbe7fmmyrxvssadnygm4fx.streamlit.app/",
            "6 Faculty Qualifications": "https://facultyqualificationspy-drvj3wpyrxvm2lrnafdwx5.streamlit.app/",
        }
        choices = list(options.keys())
        default_label = "6 Faculty Qualifications"
        default_idx = choices.index(default_label) if default_label in choices else 0
        choice = st.selectbox("Select…", choices, index=default_idx, key="kpi_nav_top")
        st.link_button("Open", options[choice], use_container_width=True)

    st.markdown('---')
    st.markdown("#### Timeframe")
    st.session_state.setdefault("time_mode", "Semestral")
    time_mode = st.radio("Timeframe", ["Semestral", "Anual", "Intersemestral"], key="time_mode", label_visibility="collapsed", horizontal=False)

    if time_mode == "Semestral":
        default_sem = SEMESTRAL_PERIODS[-1] if SEMESTRAL_PERIODS else "202510"
        st.session_state.setdefault("sel_sem", default_sem)
        sel_sem = st.selectbox("Semester", SEMESTRAL_PERIODS or [default_sem], key="sel_sem")
        sel_year = extract_year_from_period(sel_sem) or (YEARS_ALL[-1] if YEARS_ALL else 2025)
        sel_label = str(sel_sem)
    elif time_mode == "Anual":
        default_year = YEARS_ALL[-1] if YEARS_ALL else 2025
        st.session_state.setdefault("sel_year", default_year)
        sel_year = st.selectbox("Year", YEARS_ALL or [default_year], key="sel_year")
        sel_sem = None
        sel_label = f"{sel_year} (Annual)"
    else:
        default_i = INTER_YEARS[-1] if INTER_YEARS else (YEARS_ALL[-1] if YEARS_ALL else 2025)
        st.session_state.setdefault("sel_inter_year", default_i)
        sel_year = st.selectbox("Year (Intersemestral)", INTER_YEARS or YEARS_ALL or [default_i], key="sel_inter_year")
        sel_sem = None
        sel_label = f"{sel_year} Intersemestral"

    st.session_state["sel_label"] = sel_label
    st.session_state.setdefault("view_mode", "By Academic Area")
    view_mode = st.selectbox("View", ["By Program", "By Academic Area", "By Field"], key="view_mode")
    dl_bd_placeholder = st.empty()


# ================== FILTROS BASE ==================
df_car_base = df_car.copy()
base = df_fd.copy()
df_car_filt_all = filter_df_car(df_car_base, time_mode, sel_year, sel_sem)
f = filter_df_fd(df_fd, time_mode, sel_year, sel_sem)

if 'dl_bd_placeholder' in locals():
    safe = _sanitize_for_export(df_car_filt_all)
    dl_bd_placeholder.download_button(
        "Download DB (Excel)",
        data=_xlsx_bytes(safe),
        file_name=f"BD_Cartelera_{_slugify(sel_label)}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"dl_bd_{_slugify(sel_label)}"
    )

# --------- Sensitivity: “Apply to” ---------
if st.session_state.get("sens_mode", False):
    col_areaCourse = _get_any(df_car_filt_all, "Area del curso","Área del curso","Area del Curso","AREA DEL CURSO")
    col_field = _get_any(df_car_filt_all, "Field","FIELD","Campo","Área de conocimiento")
    program_col = _get_any(df_car_filt_all, "Program","PROGRAM","program","Materia")
    members = build_member_list_for_view(df_car_filt_all, view_mode, col_areaCourse, col_field, program_col)
    with st.sidebar:
        sens_member_placeholder.selectbox("Apply to", members, key="sens_member")

SENS = {"on": bool(st.session_state.get("sens_mode", False)), "ops": st.session_state.get("sens_ops", [])}

# ================== RELEVANT COLUMNS ==================
col_ps_fd   = _get_any(df_fd, "P/S", "P - S", "Participating/Supporting")
col_area_fd = _get_any(df_fd, "AREA_PROFESOR", "Area_Profesor", "Area Profesor", "Área", "Area")
col_tipo_fd = _get_any(df_fd, "TIPO", "Tipo", "Ranking", "Tipo Ranking")

col_cred       = _get_any(df_car, "Créditos", "Creditos", "Credits")
col_tipoC      = _get_any(df_car, "TIPO", "Tipo", "Tipo Ranking")
col_areaCourse = _get_any(df_car, "Area del curso","Área del curso","Area del Curso","AREA DEL CURSO")
col_prof       = _get_any(df_car, "Profesor","PROFESOR","Docente")
col_code       = _get_any(df_car, "Código Materia","Codigo Materia","CODIGO MATERIA","Código","Codigo","Course Code")
col_name       = _get_any(df_car, "Nombre largo curso","Nombre Curso","Nombre del curso","Course Name")
col_field      = _get_any(df_car, "Field","FIELD","Campo","Área de conocimiento")
col_prog       = _get_any(df_car, "Program","PROGRAM","program","Materia")
col_ps_C       = _get_any(df_car, "P/S","P - S","Participating/Supporting")

# ---------- Stylers ----------
def style_percent_tables(df_, id_col):
    sty = pd.DataFrame('', index=df_.index, columns=df_.columns)
    colP = "%P"; colSA = "%SA"; colOTHER = "%OTHER"
    p_vals = pd.to_numeric(df_[colP], errors="coerce")
    sa_vals = pd.to_numeric(df_[colSA], errors="coerce")
    other_vals = pd.to_numeric(df_[colOTHER], errors="coerce")
    is_total = df_[id_col].astype(str).str.upper().eq("TOTAL")
    sty.loc[(~is_total) & (p_vals < 60), colP] = 'background-color:#FDE2E2;'
    sty.loc[is_total & (p_vals < 75), colP] = 'background-color:#FDE2E2; font-weight:700;'
    sty.loc[sa_vals < 40, colSA] = 'background-color:#FDE2E2;'
    sty.loc[other_vals > 10, colOTHER] = 'background-color:#FDE2E2;'
    for c in sty.columns:
        sty.loc[is_total, c] = (sty.loc[is_total, c].astype(str) + 'font-weight:700;').str.replace(';;',';', regex=False)
    return sty

# ---------- util de estilo para la tabla de "Needed + Impact" ----------
def _style_needed_impact(df_, id_col):
    sty = pd.DataFrame('', index=df_.index, columns=df_.columns)
    numeric_cols = [c for c in df_.columns if c != id_col]
    # rojo claro para todo valor != 0
    for c in numeric_cols:
        vals = pd.to_numeric(df_[c], errors="coerce").fillna(0)
        sty.loc[vals != 0, c] = 'background-color:#FDE2E2;'
    return sty

# ---------- helpers de necesidades por objetivo (dos columnas) ----------
def _needed_pairs_for_obj(
    objective: str,
    scope_label: str,
    P: float, S: float, SA_: float, PA_: float, SP_: float, IP_: float, OT_: float,
    totals: dict[str,float],
    credits_each: float = 3.0
) -> tuple[int, int]:
    """
    Devuelve dos números (enteros >= 0) según el objetivo:
      - %P   -> (Need_P_more, Need_S_less)
      - %SA  -> (Need_SA_more, Need_NonSA_less)
      - %OTHER -> (Need_OTHER_less, Need_NonOTHER_more)

    Si scope="Overall", calcula con TOT (global) y limita por factibilidad del renglón cuando es "quitar".
    Nunca devuelve None; si no alcanza, devuelve el máximo posible (capped).
    """
    t_map = {"%P": (60.0, 75.0), "%SA": (40.0, 40.0), "%OTHER": (10.0, 10.0)}
    tgt_area, tgt_overall = t_map[objective]
    t = (tgt_area if scope_label == "By area" else tgt_overall) / 100.0

    # valores por fila
    TQ = SA_ + PA_ + SP_ + IP_ + OT_
    nonSA = PA_ + SP_ + IP_ + OT_
    nonOTHER = SA_ + PA_ + SP_ + IP_

    # totales
    Ptot = totals.get("P",0.0); Stot = totals.get("S",0.0)
    SAt  = totals.get("SA",0.0); PAt = totals.get("PA",0.0)
    SPt  = totals.get("SP",0.0); IPt = totals.get("IP",0.0)
    OTt  = totals.get("OTHER",0.0)
    TQt  = SAt + PAt + SPt + IPt + OTt
    nonSAt = PAt + SPt + IPt + OTt
    nonOTHERt = SAt + PAt + SPt + IPt

    # --- %P ---
    if objective == "%P":
        # Aumentar P (+3cr cada profesor)
        if scope_label == "By area":
            nP = _needed_for_pctP(P, S, tgt_area, credits_each)
        else:
            # (Ptot + c*n)/(Ptot + Stot + c*n) >= t  ->  n >= (t*(P+S) - P)/( (1-t)*c )
            den = credits_each * (1 - t)
            rhs = 0 if den <= 0 else (t*(Ptot+Stot) - Ptot) / den
            nP  = max(0, math.ceil(rhs))

        # Quitar S (-3cr)
        if scope_label == "By area":
            #  P/(P + S - c*n) >= t  ->  n >= (t*(P+S) - P)/(t*c)
            den = credits_each * t if t > 0 else float('inf')
            rhs = 0 if den == float('inf') else (t*(P+S) - P) / den
            nS_less = max(0, math.ceil(rhs))
            # factibilidad
            nmax = math.floor(S / credits_each) if credits_each > 0 else 0
            nS_less = min(nS_less, max(0, nmax))
        else:
            # overall: Ptot/(Ptot + Stot - c*n) >= t
            den = credits_each * t if t > 0 else float('inf')
            rhs = 0 if den == float('inf') else (t*(Ptot+Stot) - Ptot) / den
            nS_less = max(0, math.ceil(rhs))
            # factibilidad: solo puedo quitar del renglón actual
            nmax = math.floor(S / credits_each) if credits_each > 0 else 0
            nS_less = min(nS_less, max(0, nmax))

        return (nP, nS_less)

    # --- %SA ---
    if objective == "%SA":
        # Aumentar SA (+3cr)
        if scope_label == "By area":
            nSA = _needed_for_pctSA(SA_, nonSA, tgt_area, credits_each)
        else:
            den = credits_each * (1 - t)
            rhs = 0 if den <= 0 else (t*TQt - SAt) / den
            nSA = max(0, math.ceil(rhs))

        # Quitar No-SA (PA+SP+IP+OTHER) (-3cr)
        if scope_label == "By area":
            # SA/(SA + nonSA - c*n) >= t -> n >= (t*(SA+nonSA) - SA)/(t*c)
            den = credits_each * t if t > 0 else float('inf')
            rhs = 0 if den == float('inf') else (t*(SA_+nonSA) - SA_) / den
            nNonSA_less = max(0, math.ceil(rhs))
            nmax = math.floor(nonSA / credits_each) if credits_each > 0 else 0
            nNonSA_less = min(nNonSA_less, max(0, nmax))
        else:
            den = credits_each * t if t > 0 else float('inf')
            rhs = 0 if den == float('inf') else (t*TQt - SAt) / den
            nNonSA_less = max(0, math.ceil(rhs))
            nmax = math.floor(nonSA / credits_each) if credits_each > 0 else 0
            nNonSA_less = min(nNonSA_less, max(0, nmax))

        return (nSA, nNonSA_less)

    # --- %OTHER ---
    # Quitar OTHER (-3cr)
    if scope_label == "By area":
        # (OT - c*n)/(TQ - c*n) <= 0.10  ->  c*n >= (OT - 0.10*TQ)/0.90
        need_credits = (OT_ - 0.10*TQ) / 0.90
        nOT_less = 0 if need_credits <= 0 else math.ceil(need_credits / credits_each)
        nmax = math.floor(OT_ / credits_each) if credits_each > 0 else 0
        nOT_less = min(nOT_less, max(0, nmax))
    else:
        # overall
        need_credits = (OTt - 0.10*TQt) / 0.90
        nOT_less = 0 if need_credits <= 0 else math.ceil(need_credits / credits_each)
        nmax = math.floor(OT_ / credits_each) if credits_each > 0 else 0
        nOT_less = min(nOT_less, max(0, nmax))

    # Aumentar No-OTHER (+3cr) -> OT/(OT + nonOTHER + c*n) <= 0.10
    if scope_label == "By area":
        # n >= (0.90*OT - 0.10*nonOTHER)/(0.10*c) = (9*OT - nonOTHER)/c
        num = (9*OT_ - nonOTHER)
        den = credits_each
        nNonOT_more = 0 if num <= 0 else math.ceil(num / den)
    else:
        num = (9*OTt - nonOTHERt)
        den = credits_each
        nNonOT_more = 0 if num <= 0 else math.ceil(num / den)

    return (nOT_less, nNonOT_more)

# ---------- impacto (siempre visible) ----------
def _impact_pair(obj: str, area_vals: dict[str,float], totals: dict[str,float], scope_label: str, credits_each: float = 3.0):
    if scope_label == "By area":
        up_pp, down_pp = _impact_pp_area(obj, area_vals, credits_each)
    else:
        up_pp, down_pp = _impact_pp_overall_if_area_changes(obj, totals, credits_each)
    # devolver números (no strings)
    return round(up_pp, 2), round(down_pp, 2)

# === HEATMAP + ROJO-CLARO PARA "Needed" ===
def _style_impact_heatmap(df: pd.DataFrame, id_col: str):
    """
    - Heatmap (verde→amarillo→naranja→rojo) para columnas 'Impact +3cr (pp)' y 'Impact -3cr (pp)'.
      Se colorea por magnitud absoluta (mayor impacto = más rojo).
    - Fondo rojo claro en columnas 'Needed ...' cuando el valor != 0.
    - No toca la columna del identificador (id_col) ni otras columnas.
    """
    # DataFrame de estilos vacío
    sty = pd.DataFrame('', index=df.index, columns=df.columns)

    # --- 1) Rojo claro para "Needed ..." cuando != 0 ---
    needed_cols = [c for c in df.columns if c.startswith("Needed ")]
    for c in needed_cols:
        if c in df:
            vals = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
            sty.loc[vals != 0, c] = sty.loc[vals != 0, c].astype(str) + 'background-color:#FDE2E2;'

    # --- 2) Heatmap para columnas de Impact ---
    impact_cols = [c for c in df.columns if c.startswith("Impact ")]
    if impact_cols:
        # Usamos por defecto la magnitud del "+3cr" si existe; si no, el promedio abs de todas
        if "Impact +3cr (pp)" in df.columns:
            base_vals = pd.to_numeric(df["Impact +3cr (pp)"], errors="coerce").abs()
        else:
            base_vals = (
                df[impact_cols]
                .apply(pd.to_numeric, errors="coerce")
                .abs()
                .mean(axis=1)
            )
        base_vals = base_vals.fillna(0.0)
        vmin = float(np.nanmin(base_vals.values)) if base_vals.size else 0.0
        vmax = float(np.nanmax(base_vals.values)) if base_vals.size else 0.0
        rng = (vmax - vmin) if (vmax - vmin) > 1e-12 else 1.0  # evita división por cero

        # Paleta suave: verde → amarillo → naranja → rojo
        # (cuanto mayor el impacto, más "caliente")
        def color_for(val_abs: float) -> str:
            z = (val_abs - vmin) / rng  # 0..1  (0 = menor impacto, 1 = mayor impacto)
            if z >= 0.60:
                return "#D9F2D9"  # verde claro
            elif z >= 0.40:
                return "#FFF6B3"  # amarillo
            elif z >= 0.20:
                return "#FFD6A6"  # naranja
            else:
                return "#F5B5B5"  # rojo

        # Aplica la paleta a TODAS las columnas de impacto (según la misma escala)
        for c in impact_cols:
            col_vals = pd.to_numeric(df[c], errors="coerce").abs().fillna(0.0)
            for i, v in col_vals.items():
                sty.at[i, c] = sty.at[i, c] + f'background-color:{color_for(float(v))};'

    # Asegura que el id_col no reciba estilo accidental
    if id_col in sty.columns:
        sty[id_col] = ''

    return sty

    # Magnitudes absolutas
    vals_abs = pd.concat(
        [pd.to_numeric(df_[c], errors="coerce").abs() for c in impact_cols],
        axis=1
    )
    # vmax a partir del conjunto completo (evita escala por-columna)
    vmax = float(np.nanmax(vals_abs.values)) if vals_abs.size else 0.0
    if not np.isfinite(vmax) or vmax <= 0:
        # todos 0 o NaN: colorear todo como rojo para los impact cols
        for c in impact_cols:
            sty[c] = 'background-color:#FF0000;'
        return sty

    # Asignar color por celda
    for c in impact_cols:
        col = pd.to_numeric(df_[c], errors="coerce").abs()
        t = (col / vmax).clip(lower=0.0, upper=1.0).fillna(0.0)
        sty[c] = t.map(_interp_color).radd('background-color:')

    return sty


# ================== PRINCIPAL ==================
st.markdown("---")

# --- helpers específicos para el cabezote ---
def _guess_prof_cols(df: pd.DataFrame) -> list[str]:
    """
    Devuelve columnas candidatas para identificar un profesor.
    Prioridad: Documento/ID/Email -> nombre/profesor.
    """
    pri = []
    # IDs / correos
    for c in df.columns:
        cl = str(c).strip().lower()
        if any(k in cl for k in ["documento", "identific", "id", "correo", "email", "mail"]):
            pri.append(c)
    # Nombres / profesor
    for c in df.columns:
        cl = str(c).strip().lower()
        if any(k in cl for k in ["prof", "docent", "nombre", "name"]):
            pri.append(c)
    # Quitar duplicados preservando orden
    seen, out = set(), []
    for c in pri:
        if c not in seen:
            out.append(c); seen.add(c)
    # Fallback simple si nada matchea
    if not out:
        for cand in ["Profesor", "PROFESOR", "Docente", "Nombre", "Name", "Profesor(a)"]:
            if cand in df.columns:
                out.append(cand)
                break
    return out

def _unique_prof_count(df: pd.DataFrame, cols: list[str]) -> int:
    if df is None or df.empty:
        return 0
    # Construir una UID robusta a partir de las columnas disponibles
    use = [c for c in cols if c in df.columns]
    if not use:
        # último recurso: filas únicas por todas las columnas visibles (puede sobre-contar)
        return int(df.astype(str).drop_duplicates().shape[0])
    uid = df[use].astype(str).apply(lambda s: s.str.strip()).fillna("").agg(" | ".join, axis=1)
    return int(uid.nunique())

def _filter_fd_by_timeframe(df_fd: pd.DataFrame, time_mode: str, sel_year, sel_sem) -> pd.DataFrame:
    """
    Filtra Faculty Distribution por Semestral / Anual / Intersemestral.

    - Semestral:        == sel_sem (p.ej. '202520')
    - Anual:            empieza por sel_year (incluye 10, 20 e intersemestral)
    - Intersemestral:   == 'YYYY Intersemestral' (match tolerante a espacios y mayúsculas)
    """
    if df_fd is None or df_fd.empty:
        return df_fd.iloc[0:0]

    sem_col = _get_any(df_fd, "Semestre", "Periodo", "Periodo Académico", "Periodo academico")
    if not sem_col:
        return df_fd.iloc[0:0]

    s = df_fd[sem_col].astype(str)
    tm = (time_mode or "Semestral").strip()

    if tm == "Semestral" and sel_sem:
        m = s.str.strip().eq(str(sel_sem))
        return df_fd[m].copy()

    if tm == "Anual" and sel_year is not None:
        m = s.str.strip().str.startswith(str(sel_year))
        return df_fd[m].copy()

    if tm == "Intersemestral" and sel_year is not None:
        # Coincidencia EXACTA a "YYYY Intersemestral" (tolerante a espacios múltiples/caso)
        target = f"{sel_year} intersemestral"
        m = s.str.strip().str.casefold().eq(target)
        # tolerancia: permitir variantes con espacios extra entre año y palabra
        m = m | s.str.casefold().str.contains(rf"^\s*{sel_year}\s+intersemestral\s*$", regex=True)
        return df_fd[m].copy()

    # Fallback: sin filtro
    return df_fd.copy()

def _count_teaching_from_fd_timeaware(df_fd: pd.DataFrame, time_mode: str, sel_year, sel_sem) -> dict[str,int]:
    """
    Cuenta profesores ÚNICOS en Faculty Distribution según timeframe:
      - Full-time (FT):   PLANTA_CATEDRA == 'PLANTA'
      - Part-time (PT):   PLANTA_CATEDRA == 'CÁTEDRA' / 'CATEDRA'
      - Participating P:  P/S == 'P'
      - Supporting   S:   P/S == 'S'
    """
    if df_fd is None or df_fd.empty:
        return {"FT":0, "PT":0, "P":0, "S":0}

    dff = _filter_fd_by_timeframe(df_fd, time_mode, sel_year, sel_sem)
    if dff is None or dff.empty:
        return {"FT":0, "PT":0, "P":0, "S":0}

    prof_cols = _guess_prof_cols(dff)

    # columnas de clasificación
    pc_col = _get_any(dff, "PLANTA_CATEDRA", "Planta_Catedra", "Planta/Cátedra", "PLANTA CATEDRA", "Planta/Catedra")
    ps_col = _get_any(dff, "P/S", "P - S", "Participating/Supporting", "P S")

    # Full-time / Part-time
    ft = pt = 0
    if pc_col:
        tag = _norm_str(dff[pc_col])
        ft_df = dff[tag.eq("planta")]
        pt_df = dff[tag.isin({"catedra", "cátedra"})]
        ft = _unique_prof_count(ft_df, prof_cols)
        pt = _unique_prof_count(pt_df, prof_cols)

    # Participating / Supporting
    p_cnt = s_cnt = 0
    if ps_col:
        tps = _norm_str(dff[ps_col])
        p_df = dff[tps.eq("p")]
        s_df = dff[tps.eq("s")]
        p_cnt = _unique_prof_count(p_df, prof_cols)
        s_cnt = _unique_prof_count(s_df, prof_cols)

    return {"FT":ft, "PT":pt, "P":p_cnt, "S":s_cnt}

def compute_header_counts_teaching(df_fd: pd.DataFrame, time_mode: str, sel_year, sel_sem, sens: dict) -> dict:
    base = _count_teaching_from_fd_timeaware(df_fd, time_mode, sel_year, sel_sem)

    # Sensibilidad: +P suma a Full-time y Participating; +S suma a Part-time y Supporting
    dP = dS = 0
    if sens.get("on") and sens.get("ops"):
        for op in sens["ops"]:
            if op.get("scope") == "PS":
                if op.get("cat") == "P":
                    dP += int(op.get("count", 0))
                elif op.get("cat") == "S":
                    dS += int(op.get("count", 0))

    return {
        "Full-time":     max(0, base["FT"] + dP),
        "Part-time":     max(0, base["PT"] + dS),
        "Participating": max(0, base["P"]  + dP),
        "Supporting":    max(0, base["S"]  + dS),
    }

# === Subheader ===

st.subheader(f"Faculty Sufficiency and Qualifications — {st.session_state.get('sel_label','Selected')}")

# ====== NORMALIZACIÓN BASE PARA CARTELERA + EXCLUSIONES ======
if not all([col_cred, col_tipoC, col_areaCourse]):
    st.error("Missing columns in 'BD Cartelera 2020-2025': 'Credits', 'TIPO', and/or 'Academic Area (course)'.")
else:
    df_car_n = df_car.copy()
    df_car_n["_CRED"] = pd.to_numeric(df_car_n[col_cred], errors="coerce").fillna(0.0)
    df_car_n["_TIPO"] = _norm_str(df_car_n[col_tipoC]).map(normalize_tipo)
    if "_SEM" not in df_car_n.columns:
        sc = _get_any(df_car_n, "Semestre","Periodo","Periodo Académico","Periodo academico")
        df_car_n["_SEM"] = df_car_n[sc].astype(str).str.strip() if sc else ""
    df_car_n["_YEAR"] = df_car_n["_SEM"].map(extract_year_from_period)
    df_car_n["_AREA"] = df_car_n[col_areaCourse].astype(str).str.strip()
    col_ps_C_local = _get_any(df_car_n, "P/S","P - S","Participating/Supporting")
    df_car_n["_PS"] = _norm_str(df_car_n[col_ps_C_local]).map(normalize_ps) if col_ps_C_local else ""

    # excluir programas
    program_col0 = _get_any(df_car_n, "Program","PROGRAM","program","Materia")
    EXCLUDE_SUBJ = {"CONT", "E-IMER", "E-ENEG", "E-AFIN"}
    if program_col0:
        mask_ok = ~df_car_n[program_col0].astype(str).str.strip().str.upper().isin(EXCLUDE_SUBJ)
        df_car_global = df_car_n[mask_ok].copy()
    else:
        df_car_global = df_car_n.copy()

    # ---------- Filtro por timeframe seleccionado ----------
    sel_label = st.session_state.get("sel_label")
    time_mode = st.session_state.get("time_mode", "Semestral")
    sel_year  = st.session_state.get("sel_year")
    sel_sem   = st.session_state.get("sel_sem")

    fil = filter_df_car(df_car_global, time_mode, sel_year, sel_sem)
    df_car_filt_all = fil.copy()  # usar en expander/tabla/dona

    # ============================ VISTAS ============================
    def build_percent_table(base_idx_name, agg_tipo, agg_ps):
        den_ps = (agg_ps["P"] + agg_ps["S"]).replace(0, pd.NA)
        p_share = (agg_ps["P"] / den_ps) * 100
        s_share = 100 - p_share
        denom_q = (agg_tipo.sum(axis=1)).replace(0, pd.NA)
        dfm = pd.DataFrame({
            base_idx_name: agg_tipo.index,
            "%P": p_share,
            "%S": s_share,
            "%SA": (agg_tipo["SA"] / denom_q) * 100,
            "%OTHER": (agg_tipo["OTHER"] / denom_q) * 100,
        }).fillna(0.0)

        # TOTAL
        tot_P, tot_S = agg_ps["P"].sum(), agg_ps["S"].sum()
        tot_den_ps = tot_P + tot_S
        p_tot = (tot_P / tot_den_ps * 100) if tot_den_ps else 0.0
        s_tot = 100 - p_tot
        tipo_sums = agg_tipo[["SA","PA","SP","IP","OTHER"]].sum(axis=0)
        denom_q_tot = float(tipo_sums.sum())

        total_row = {
            base_idx_name: "TOTAL",
            "%P": round(p_tot, 1),
            "%S": round(s_tot, 1),
            "%SA": round((tipo_sums["SA"] / denom_q_tot * 100) if denom_q_tot else 0.0, 1),
            "%OTHER": round((tipo_sums["OTHER"] / denom_q_tot * 100) if denom_q_tot else 0.0, 1),
        }
        dfm[["%P","%S","%SA","%OTHER"]] = dfm[["%P","%S","%SA","%OTHER"]].round(1)
        dfm = pd.concat([dfm, pd.DataFrame([total_row])], ignore_index=True)
        return dfm[[f"{base_idx_name}", "%P", "%S", "%SA", "%OTHER"]]

    if fil.empty:
        st.info(f"No records for the selected timeframe: {sel_label}.")
    else:
        # ========== BY ACADEMIC AREA ==========
        if view_mode == "By Academic Area":
            colT, colG = st.columns([6,6], gap="large")
        
            # Agregaciones
            agg_tipo = (fil.groupby(["_AREA","_TIPO"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["SA","PA","SP","IP","OTHER"]:
                if k not in agg_tipo.columns: agg_tipo[k] = 0.0
            agg_tipo = agg_tipo[["SA","PA","SP","IP","OTHER"]]
        
            agg_ps = (fil.groupby(["_AREA","_PS"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["P","S"]:
                if k not in agg_ps.columns: agg_ps[k] = 0.0
            agg_ps = agg_ps[["P","S"]]
        
            # Sensibilidad
            base_agg_ps = agg_ps.copy()
            base_agg_tipo = agg_tipo.copy()
            if SENS["on"] and SENS["ops"]:
                mod_agg_ps, mod_agg_tipo = apply_ops_to_aggs(base_agg_ps, base_agg_tipo, SENS["ops"])
            else:
                mod_agg_ps, mod_agg_tipo = base_agg_ps, base_agg_tipo
        
            with colT:
                # Controles: solo si Sensitivity ON y toggle ON se muestra el selector; el IMPACTO ya es siempre visible
                needed_mode = False
                if SENS["on"]:
                    r1c1, r1c2, r1c3 = st.columns([1.8, 1.1, 1.6])
                    with r1c1:
                        needed_mode = st.toggle("Show necessary # of Faculty for…", value=False, key="area_needed_mode")
                    if needed_mode:
                        with r1c2:
                            objective = st.selectbox("Objective", ["%P", "%SA", "%OTHER"], key="area_objective")
                        with r1c3:
                            scope_label = st.radio("Target scope", ["By area", "Overall"], horizontal=True, key="area_scope")
                    else:
                        objective = st.session_state.get("area_objective", "%P")
                        scope_label = st.session_state.get("area_scope", "By area")
                else:
                    objective = st.session_state.get("area_objective", "%P")
                    scope_label = st.session_state.get("area_scope", "By area")
        
                if not needed_mode:
                    metrics_tbl = build_percent_table("Academic Area", mod_agg_tipo, mod_agg_ps)
                    _download_xlsx_button(metrics_tbl, f"table_ByArea_{_slugify(sel_label)}.xlsx",
                                          key=f"dl_tbl_area_{_slugify(sel_label)}", label="⬇️ Download table (Excel)")
                    styled_tbl = (
                        metrics_tbl.style
                        .format({"%P": "{:.1f}%", "%S": "{:.1f}%", "%SA": "{:.1f}%", "%OTHER": "{:.1f}%"})
                        .apply(style_percent_tables, id_col="Academic Area", axis=None)
                        .hide(axis="index")
                    )
                    st.markdown(f"<div class='scroll-wrap-400'>{styled_tbl.to_html(escape=False)}</div>", unsafe_allow_html=True)
                else:
                    # ===== Tabla: Needed (dos columnas) + Impact (siempre) SIN TOTAL =====
                    # union de índices para no perder filas
                    idx_all = sorted(set(mod_agg_ps.index.tolist()) | set(mod_agg_tipo.index.tolist()))
                    p   = mod_agg_ps["P"].reindex(idx_all, fill_value=0.0)
                    s   = mod_agg_ps["S"].reindex(idx_all, fill_value=0.0)
                    sa  = mod_agg_tipo["SA"].reindex(idx_all, fill_value=0.0)
                    pa  = mod_agg_tipo["PA"].reindex(idx_all, fill_value=0.0)
                    sp  = mod_agg_tipo["SP"].reindex(idx_all, fill_value=0.0)
                    ip  = mod_agg_tipo["IP"].reindex(idx_all, fill_value=0.0)
                    oth = mod_agg_tipo["OTHER"].reindex(idx_all, fill_value=0.0)
        
                    totals = {
                        "P": float(p.sum()), "S": float(s.sum()),
                        "SA": float(sa.sum()), "PA": float(pa.sum()),
                        "SP": float(sp.sum()), "IP": float(ip.sum()),
                        "OTHER": float(oth.sum())
                    }
        
                    # nombres de columnas según objetivo
                    if objective == "%P":
                        main_col, aux_col = "Needed P (3cr)", "Needed S less (3cr)"
                    elif objective == "%SA":
                        main_col, aux_col = "Needed SA (3cr)", "Needed Non-SA less (3cr)"
                    else:
                        main_col, aux_col = "Needed OTHER less (3cr)", "Needed Non-OTHER more (3cr)"
        
                    rows = []
                    for label in idx_all:
                        Pv, Sv = float(p.get(label,0.0)), float(s.get(label,0.0))
                        SAv, PAv = float(sa.get(label,0.0)), float(pa.get(label,0.0))
                        SPv, IPv = float(sp.get(label,0.0)), float(ip.get(label,0.0))
                        OTv      = float(oth.get(label,0.0))
        
                        need1, need2 = _needed_pairs_for_obj(
                            objective, scope_label,
                            Pv, Sv, SAv, PAv, SPv, IPv, OTv,
                            totals, credits_each=3.0
                        )
        
                        area_vals = {"P":Pv,"S":Sv,"SA":SAv,"PA":PAv,"SP":SPv,"IP":IPv,"OTHER":OTv}
                        up_pp, down_pp = _impact_pair(objective, area_vals, totals, scope_label, credits_each=3.0)
        
                        rows.append({
                            "Academic Area": label,
                            main_col: int(need1),
                            aux_col:  int(need2),
                            "Impact +3cr (pp)": up_pp,
                            "Impact -3cr (pp)": down_pp
                        })
        
                    need_tbl = pd.DataFrame(rows)
                    # formateo + heatmap (verde→amarillo→naranja→rojo) SOLO en columnas "Impact ..."
                    fmt_map = {}
                    if 'Academic Area' in need_tbl.columns:
                        fmt_map['Academic Area'] = '{}'
                    for col in ['Needed P (3cr)', 'Needed S less (3cr)', 'Needed SA (3cr)', 'Needed OTHER less (3cr)', 'Needed OTHER more (3cr)']:
                        if col in need_tbl.columns:
                            fmt_map[col] = '{:.0f}'
                    for col in ['Impact +3cr (pp)', 'Impact -3cr (pp)']:
                        if col in need_tbl.columns:
                            fmt_map[col] = '{:+.2f}'
                    
                    styled = (
                        need_tbl.style
                        .format(fmt_map)
                        .apply(_style_impact_heatmap, id_col="Academic Area", axis=None)  # HEATMAP aplicado aquí
                        .hide(axis="index")
                    )
                    _download_xlsx_button(
                        need_tbl,
                        f"needed_ByArea_{_slugify(sel_label)}_{_slugify(objective)}_{_slugify(scope_label)}.xlsx",
                        key=f"dl_need_area_{_slugify(sel_label)}_{_slugify(objective)}_{_slugify(scope_label)}",
                        label="⬇️ Descargar (Excel)"
                    )
                    st.markdown(styled.to_html(escape=False), unsafe_allow_html=True)

            # ========== Series históricas ==========
            df_hist = df_car_global.copy()
            agg_ps_all = (df_hist.groupby(["_SEM","_AREA","_PS"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["P","S"]:
                if k not in agg_ps_all.columns: agg_ps_all[k] = 0.0
            agg_ps_all["P_share"] = (agg_ps_all["P"] / (agg_ps_all["P"] + agg_ps_all["S"]).replace(0, pd.NA)) * 100
            agg_ps_all = agg_ps_all.reset_index()

            agg_tipo_all = (df_hist.groupby(["_SEM","_AREA","_TIPO"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["SA","PA","SP","IP","OTHER"]:
                if k not in agg_tipo_all.columns: agg_tipo_all[k] = 0.0
            den_all = (agg_tipo_all[["SA","PA","SP","IP","OTHER"]].sum(axis=1)).replace(0, pd.NA)
            agg_tipo_all["SA_share"] = (agg_tipo_all["SA"] / den_all) * 100
            agg_tipo_all["OTHER_share"] = (agg_tipo_all["OTHER"] / den_all) * 100
            agg_tipo_all = agg_tipo_all.reset_index()

            tot_by_sem_P = (df_hist.groupby(["_SEM","_PS"])["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["P","S"]:
                if k not in tot_by_sem_P.columns: tot_by_sem_P[k] = 0.0
            tot_by_sem_P["P_share"] = (tot_by_sem_P["P"] / (tot_by_sem_P["P"] + tot_by_sem_P["S"]).replace(0, pd.NA)) * 100
            tot_by_sem_P = tot_by_sem_P.reset_index()

            tot_by_sem_tipo = (df_hist.groupby(["_SEM","_TIPO"])["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["SA","PA","SP","IP","OTHER"]:
                if k not in tot_by_sem_tipo.columns: tot_by_sem_tipo[k] = 0.0
            den_tot = (tot_by_sem_tipo[["SA","PA","SP","IP","OTHER"]].sum(axis=1)).replace(0, pd.NA)
            tot_by_sem_tipo["SA_share"] = (tot_by_sem_tipo["SA"] / den_tot) * 100
            tot_by_sem_tipo["OTHER_share"] = (tot_by_sem_tipo["OTHER"] / den_tot) * 100
            tot_by_sem_tipo = tot_by_sem_tipo.reset_index()

            # Adaptación a modo temporal
            agg_ps_all_tm  = transform_for_time_mode_ps(agg_ps_all.rename(columns={"_AREA":"__LEVEL__"})).rename(columns={"__LEVEL__":"_AREA"})
            agg_tipo_sa_tm = transform_for_time_mode_tipo(agg_tipo_all.rename(columns={"_AREA":"__LEVEL__"}), "SA_share").rename(columns={"__LEVEL__":"_AREA"})
            agg_tipo_ot_tm = transform_for_time_mode_tipo(agg_tipo_all.rename(columns={"_AREA":"__LEVEL__"}), "OTHER_share").rename(columns={"__LEVEL__":"_AREA"})
            agg_tipo_all_tm = (
                agg_tipo_sa_tm.drop(columns=[c for c in ["OTHER_share"] if c in agg_tipo_sa_tm], errors="ignore")
                .merge(agg_tipo_ot_tm[["_SEM","_AREA","OTHER","SA","PA","SP","IP","OTHER_share"]],
                       on=["_SEM","_AREA","SA","PA","SP","IP","OTHER"], how="outer")
            )
            tot_by_sem_P_tm = transform_for_time_mode_ps(tot_by_sem_P.copy())
            tot_tipo_sa_tm  = transform_for_time_mode_tipo(tot_by_sem_tipo.copy(), "SA_share")
            tot_tipo_ot_tm  = transform_for_time_mode_tipo(tot_by_sem_tipo.copy(), "OTHER_share")
            tot_by_sem_tipo_tm = (
                tot_tipo_sa_tm.drop(columns=[c for c in ["OTHER_share"] if c in tot_tipo_sa_tm], errors="ignore")
                .merge(tot_tipo_ot_tm[["_SEM","SA","PA","SP","IP","OTHER","OTHER_share"]],
                       on=["_SEM","SA","PA","SP","IP","OTHER"], how="outer")
            )

            key_col, x_labels, x_map = build_time_axis_for_history(df_hist)
            if time_mode == "Semestral":
                sel_x = x_map.get(str(sel_sem)) if sel_sem else None
                sel_label_exact = str(sel_sem) if sel_sem else None
            elif time_mode == "Anual":
                sel_x = x_map.get(sel_year) if sel_year is not None else None
                sel_label_exact = sel_year
            else:
                inter_label = f"{sel_year} Intersemestral" if sel_year else None
                sel_x = x_map.get(inter_label) if inter_label else None
                sel_label_exact = inter_label

            if SENS["on"] and SENS["ops"] and sel_label_exact is not None:
                agg_ps_all_tm, agg_tipo_all_tm, tot_by_sem_P_tm, tot_by_sem_tipo_tm = apply_sensitivity_to_history(
                    agg_ps_all_tm, agg_tipo_all_tm, tot_by_sem_P_tm, tot_by_sem_tipo_tm,
                    level_name="_AREA",
                    sel_label_value=sel_label_exact,
                    ops=SENS["ops"],
                    member_all_label="All"
                )

            areas_all = sorted(set(agg_ps_all_tm["_AREA"].astype(str).unique()) | set(agg_tipo_all_tm["_AREA"].astype(str).unique()))
            with colG:
                draw_history(
                    "Evolution by Academic Area",
                    level_name="_AREA",
                    level_values=areas_all,
                    metric_kind="%P",
                    total_series_builders={"P": tot_by_sem_P_tm, "SA": tot_by_sem_tipo_tm, "OTHER": tot_by_sem_tipo_tm},
                    agg_ps_all=agg_ps_all_tm,
                    agg_tipo_all=agg_tipo_all_tm,
                    x_labels=x_labels, x_map=x_map, sel_x=sel_x
                )

        # -------------- BY FIELD --------------
        elif view_mode == "By Field" and col_field:
            colF_L, colF_R = st.columns([6,6], gap="large")
            fil_field = fil.copy()
            fil_field["_FIELD"] = fil_field[col_field].astype(str).str.strip()
        
            agg_tipo_f = (fil_field.groupby(["_FIELD","_TIPO"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["SA","PA","SP","IP","OTHER"]:
                if k not in agg_tipo_f.columns: agg_tipo_f[k] = 0.0
            agg_tipo_f = agg_tipo_f[["SA","PA","SP","IP","OTHER"]]
        
            agg_ps_f = (fil_field.groupby(["_FIELD","_PS"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["P","S"]:
                if k not in agg_ps_f.columns: agg_ps_f[k] = 0.0
            agg_ps_f = agg_ps_f[["P","S"]]
        
            base_agg_ps = agg_ps_f.copy()
            base_agg_tipo = agg_tipo_f.copy()
            if SENS["on"] and SENS["ops"]:
                mod_agg_ps, mod_agg_tipo = apply_ops_to_aggs(base_agg_ps, base_agg_tipo, SENS["ops"])
            else:
                mod_agg_ps, mod_agg_tipo = base_agg_ps, base_agg_tipo
        
            with colF_L:
                needed_mode_f = False
                if SENS["on"]:
                    r1c1, r1c2, r1c3 = st.columns([1.8, 1.1, 1.6])
                    with r1c1:
                        needed_mode_f = st.toggle("Show necessary # of Faculty for…", value=False, key="field_needed_mode")
                    if needed_mode_f:
                        with r1c2:
                            objective_f = st.selectbox("Objective", ["%P", "%SA", "%OTHER"], key="field_objective")
                        with r1c3:
                            scope_label_f = st.radio("Target scope", ["By area", "Overall"], horizontal=True, key="field_scope")
                    else:
                        objective_f = st.session_state.get("field_objective", "%P")
                        scope_label_f = st.session_state.get("field_scope", "By area")
                else:
                    objective_f = st.session_state.get("field_objective", "%P")
                    scope_label_f = st.session_state.get("field_scope", "By area")
        
                if not needed_mode_f:
                    metrics_tbl_f = build_percent_table("Field", mod_agg_tipo, mod_agg_ps)
                    _download_xlsx_button(metrics_tbl_f, f"table_ByField_{_slugify(sel_label)}.xlsx",
                                          key=f"dl_tbl_field_{_slugify(sel_label)}", label="⬇️ Download table (Excel)")
                    styled_tbl_f = (
                        metrics_tbl_f.style
                        .format({"%P":"{:.1f}%","%S":"{:.1f}%","%SA":"{:.1f}%","%OTHER":"{:.1f}%"})
                        .apply(style_percent_tables, id_col="Field", axis=None)
                        .hide(axis="index")
                    )
                    st.markdown(f"<div class='scroll-wrap-400'>{styled_tbl_f.to_html(escape=False)}</div>", unsafe_allow_html=True)
                else:
                    idx_all = sorted(set(mod_agg_ps.index.tolist()) | set(mod_agg_tipo.index.tolist()))
                    p   = mod_agg_ps["P"].reindex(idx_all, fill_value=0.0)
                    s   = mod_agg_ps["S"].reindex(idx_all, fill_value=0.0)
                    sa  = mod_agg_tipo["SA"].reindex(idx_all, fill_value=0.0)
                    pa  = mod_agg_tipo["PA"].reindex(idx_all, fill_value=0.0)
                    sp  = mod_agg_tipo["SP"].reindex(idx_all, fill_value=0.0)
                    ip  = mod_agg_tipo["IP"].reindex(idx_all, fill_value=0.0)
                    oth = mod_agg_tipo["OTHER"].reindex(idx_all, fill_value=0.0)
        
                    totals = {
                        "P": float(p.sum()), "S": float(s.sum()),
                        "SA": float(sa.sum()), "PA": float(pa.sum()),
                        "SP": float(sp.sum()), "IP": float(ip.sum()),
                        "OTHER": float(oth.sum())
                    }
        
                    if objective_f == "%P":
                        main_col, aux_col = "Needed P (3cr)", "Needed S less (3cr)"
                    elif objective_f == "%SA":
                        main_col, aux_col = "Needed SA (3cr)", "Needed Non-SA less (3cr)"
                    else:
                        main_col, aux_col = "Needed OTHER less (3cr)", "Needed Non-OTHER more (3cr)"
        
                    rows = []
                    for label in idx_all:
                        Pv, Sv = float(p.get(label,0.0)), float(s.get(label,0.0))
                        SAv, PAv = float(sa.get(label,0.0)), float(pa.get(label,0.0))
                        SPv, IPv = float(sp.get(label,0.0)), float(ip.get(label,0.0))
                        OTv      = float(oth.get(label,0.0))
        
                        need1, need2 = _needed_pairs_for_obj(
                            objective_f, scope_label_f,
                            Pv, Sv, SAv, PAv, SPv, IPv, OTv, totals, credits_each=3.0
                        )
                        area_vals = {"P":Pv,"S":Sv,"SA":SAv,"PA":PAv,"SP":SPv,"IP":IPv,"OTHER":OTv}
                        up_pp, down_pp = _impact_pair(objective_f, area_vals, totals, scope_label_f, credits_each=3.0)
        
                        rows.append({
                            "Field": label,
                            main_col: int(need1),
                            aux_col:  int(need2),
                            "Impact +3cr (pp)": up_pp,
                            "Impact -3cr (pp)": down_pp
                        })
        
                    need_tbl_f = pd.DataFrame(rows)

                    fmt_map_f = {}
                    if 'Field' in need_tbl_f.columns:
                        fmt_map_f['Field'] = '{}'
                    for col in ['Needed P (3cr)', 'Needed S less (3cr)', 'Needed SA (3cr)', 'Needed OTHER less (3cr)', 'Needed OTHER more (3cr)']:
                        if col in need_tbl_f.columns:
                            fmt_map_f[col] = '{:.0f}'
                    for col in ['Impact +3cr (pp)', 'Impact -3cr (pp)']:
                        if col in need_tbl_f.columns:
                            fmt_map_f[col] = '{:+.2f}'
                    
                    styled_f = (
                        need_tbl_f.style
                        .format(fmt_map_f)
                        .apply(_style_impact_heatmap, id_col="Field", axis=None)  # HEATMAP aplicado aquí
                        .hide(axis="index")
                    )
                    
                    _download_xlsx_button(
                        need_tbl_f,
                        f"needed_ByField_{_slugify(sel_label)}_{_slugify(objective_f)}_{_slugify(scope_label_f)}.xlsx",
                        key=f"dl_need_field_{_slugify(sel_label)}_{_slugify(objective_f)}_{_slugify(scope_label_f)}",
                        label="⬇️ Descargar (Excel)"
                    )
                    
                    st.markdown(styled_f.to_html(escape=False), unsafe_allow_html=True)

            # Históricos Field
            df_hist_f = df_car_global.copy()
            df_hist_f["_FIELD"] = df_hist_f[col_field].astype(str).str.strip()

            agg_ps_all_f = (df_hist_f.groupby(["_SEM","_FIELD","_PS"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["P","S"]:
                if k not in agg_ps_all_f.columns: agg_ps_all_f[k] = 0.0
            agg_ps_all_f["P_share"] = (agg_ps_all_f["P"] / (agg_ps_all_f["P"] + agg_ps_all_f["S"]).replace(0, pd.NA)) * 100
            agg_ps_all_f = agg_ps_all_f.reset_index()

            agg_tipo_all_f = (df_hist_f.groupby(["_SEM","_FIELD","_TIPO"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["SA","PA","SP","IP","OTHER"]:
                if k not in agg_tipo_all_f.columns: agg_tipo_all_f[k] = 0.0
            den_all_f = (agg_tipo_all_f[["SA","PA","SP","IP","OTHER"]].sum(axis=1)).replace(0, pd.NA)
            agg_tipo_all_f["SA_share"] = (agg_tipo_all_f["SA"] / den_all_f) * 100
            agg_tipo_all_f["OTHER_share"] = (agg_tipo_all_f["OTHER"] / den_all_f) * 100
            agg_tipo_all_f = agg_tipo_all_f.reset_index()

            tot_by_sem_tipo_f = (df_hist_f.groupby(["_SEM","_TIPO"])["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["SA","PA","SP","IP","OTHER"]:
                if k not in tot_by_sem_tipo_f.columns: tot_by_sem_tipo_f[k] = 0.0
            den_f_tot = (tot_by_sem_tipo_f[["SA","PA","SP","IP","OTHER"]].sum(axis=1)).replace(0, pd.NA)
            tot_by_sem_tipo_f["SA_share"] = (tot_by_sem_tipo_f["SA"] / den_f_tot) * 100
            tot_by_sem_tipo_f["OTHER_share"] = (tot_by_sem_tipo_f["OTHER"] / den_f_tot) * 100
            tot_by_sem_tipo_f = tot_by_sem_tipo_f.reset_index()

            tot_by_sem_f = (df_hist_f.groupby(["_SEM","_PS"])["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["P","S"]:
                if k not in tot_by_sem_f.columns: tot_by_sem_f[k] = 0.0
            tot_by_sem_f["P_share"] = (tot_by_sem_f["P"] / (tot_by_sem_f["P"] + tot_by_sem_f["S"]).replace(0, pd.NA)) * 100
            tot_by_sem_f = tot_by_sem_f.reset_index()

            agg_ps_all_tm = transform_for_time_mode_ps(agg_ps_all_f.rename(columns={"_FIELD":"__LEVEL__"})).rename(columns={"__LEVEL__":"_FIELD"})
            agg_tipo_sa_tm = transform_for_time_mode_tipo(agg_tipo_all_f.rename(columns={"_FIELD":"__LEVEL__"}), "SA_share").rename(columns={"__LEVEL__":"_FIELD"})
            agg_tipo_ot_tm = transform_for_time_mode_tipo(agg_tipo_all_f.rename(columns={"_FIELD":"__LEVEL__"}), "OTHER_share").rename(columns={"__LEVEL__":"_FIELD"})
            agg_tipo_all_tm = (
                agg_tipo_sa_tm.drop(columns=[c for c in ["OTHER_share"] if c in agg_tipo_sa_tm], errors="ignore")
                .merge(
                    agg_tipo_ot_tm[["_SEM","_FIELD","OTHER","SA","PA","SP","IP","OTHER_share"]],
                    on=["_SEM","_FIELD","SA","PA","SP","IP","OTHER"], how="outer"
                )
            )
            tot_by_sem_P_tm = transform_for_time_mode_ps(tot_by_sem_f.copy())
            tot_tipo_sa_tm  = transform_for_time_mode_tipo(tot_by_sem_tipo_f.copy(), "SA_share")
            tot_tipo_ot_tm  = transform_for_time_mode_tipo(tot_by_sem_tipo_f.copy(), "OTHER_share")
            tot_by_sem_tipo_tm = (
                tot_tipo_sa_tm.drop(columns=[c for c in ["OTHER_share"] if c in tot_tipo_sa_tm], errors="ignore")
                .merge(
                    tot_tipo_ot_tm[["_SEM","SA","PA","SP","IP","OTHER","OTHER_share"]],
                    on=["_SEM","SA","PA","SP","IP","OTHER"], how="outer"
                )
            )

            key_col, x_labels, x_map = build_time_axis_for_history(df_hist_f)
            if time_mode == "Semestral":
                sel_x = x_map.get(str(sel_sem)) if sel_sem else None
                sel_label_exact = str(sel_sem) if sel_sem else None
            elif time_mode == "Anual":
                sel_x = x_map.get(sel_year) if sel_year is not None else None
                sel_label_exact = sel_year
            else:
                inter_label = f"{sel_year} Intersemestral" if sel_year else None
                sel_x = x_map.get(inter_label) if inter_label else None
                sel_label_exact = inter_label

            if SENS["on"] and SENS["ops"] and sel_label_exact is not None:
                agg_ps_all_tm, agg_tipo_all_tm, tot_by_sem_P_tm, tot_by_sem_tipo_tm = apply_sensitivity_to_history(
                    agg_ps_all_tm, agg_tipo_all_tm, tot_by_sem_P_tm, tot_by_sem_tipo_tm,
                    level_name="_FIELD",
                    sel_label_value=sel_label_exact,
                    ops=SENS["ops"],
                    member_all_label="All"
                )

            fields_all = sorted(set(agg_ps_all_tm["_FIELD"].astype(str).unique()) | set(agg_tipo_all_tm["_FIELD"].astype(str).unique()))
            with colF_R:
                draw_history(
                    "Evolution by Academic Field",
                    level_name="_FIELD",
                    level_values=fields_all,
                    metric_kind="%P",
                    total_series_builders={"P": tot_by_sem_P_tm, "SA": tot_by_sem_tipo_tm, "OTHER": tot_by_sem_tipo_tm},
                    agg_ps_all=agg_ps_all_tm,
                    agg_tipo_all=agg_tipo_all_tm,
                    x_labels=x_labels, x_map=x_map, sel_x=sel_x
                )
        # -------------- BY PROGRAM --------------
        elif view_mode == "By Program" and col_prog:
            colP_L, colP_R = st.columns([6,6], gap="large")
            fil_prog = fil.copy()
            fil_prog["_PROG"] = fil_prog[col_prog].astype(str).str.strip()
        
            agg_tipo_p = (fil_prog.groupby(["_PROG","_TIPO"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["SA","PA","SP","IP","OTHER"]:
                if k not in agg_tipo_p.columns: agg_tipo_p[k] = 0.0
            agg_tipo_p = agg_tipo_p[["SA","PA","SP","IP","OTHER"]]
        
            agg_ps_p = (fil_prog.groupby(["_PROG","_PS"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["P","S"]:
                if k not in agg_ps_p.columns: agg_ps_p[k] = 0.0
            agg_ps_p = agg_ps_p[["P","S"]]
        
            base_agg_ps_p = agg_ps_p.copy()
            base_agg_tipo_p = agg_tipo_p.copy()
            if SENS["on"] and SENS["ops"]:
                mod_agg_ps_p, mod_agg_tipo_p = apply_ops_to_aggs(base_agg_ps_p, base_agg_tipo_p, SENS["ops"])
            else:
                mod_agg_ps_p, mod_agg_tipo_p = base_agg_ps_p, base_agg_tipo_p
        
            with colP_L:
                needed_mode_p = False
                if SENS["on"]:
                    r1c1, r1c2, r1c3 = st.columns([1.8, 1.1, 1.6])
                    with r1c1:
                        needed_mode_p = st.toggle("Show necessary # of Faculty for…", value=False, key="prog_needed_mode")
                    if needed_mode_p:
                        with r1c2:
                            objective_p = st.selectbox("Objective", ["%P", "%SA", "%OTHER"], key="prog_objective")
                        with r1c3:
                            scope_label_p = st.radio("Target scope", ["By area", "Overall"], horizontal=True, key="prog_scope")
                    else:
                        objective_p = st.session_state.get("prog_objective", "%P")
                        scope_label_p = st.session_state.get("prog_scope", "By area")
                else:
                    objective_p = st.session_state.get("prog_objective", "%P")
                    scope_label_p = st.session_state.get("prog_scope", "By area")
        
                if not needed_mode_p:
                    metrics_tbl_p = build_percent_table("Program", mod_agg_tipo_p, mod_agg_ps_p)
                    _download_xlsx_button(metrics_tbl_p, f"table_ByProgram_{_slugify(sel_label)}.xlsx",
                                          key=f"dl_tbl_prog_{_slugify(sel_label)}", label="⬇️ Download table (Excel)")
                    styled_tbl_p = (
                        metrics_tbl_p.style
                        .format({"%P":"{:.1f}%","%S":"{:.1f}%","%SA":"{:.1f}%","%OTHER":"{:.1f}%"})
                        .apply(style_percent_tables, id_col="Program", axis=None)
                        .hide(axis="index")
                    )
                    st.markdown(f"<div class='scroll-wrap-400'>{styled_tbl_p.to_html(escape=False)}</div>", unsafe_allow_html=True)
                else:
                    idx_all = sorted(set(mod_agg_ps_p.index.tolist()) | set(mod_agg_tipo_p.index.tolist()))
                    p   = mod_agg_ps_p["P"].reindex(idx_all, fill_value=0.0)
                    s   = mod_agg_ps_p["S"].reindex(idx_all, fill_value=0.0)
                    sa  = mod_agg_tipo_p["SA"].reindex(idx_all, fill_value=0.0)
                    pa  = mod_agg_tipo_p["PA"].reindex(idx_all, fill_value=0.0)
                    sp  = mod_agg_tipo_p["SP"].reindex(idx_all, fill_value=0.0)
                    ip  = mod_agg_tipo_p["IP"].reindex(idx_all, fill_value=0.0)
                    oth = mod_agg_tipo_p["OTHER"].reindex(idx_all, fill_value=0.0)
        
                    totals = {
                        "P": float(p.sum()), "S": float(s.sum()),
                        "SA": float(sa.sum()), "PA": float(pa.sum()),
                        "SP": float(sp.sum()), "IP": float(ip.sum()),
                        "OTHER": float(oth.sum())
                    }
        
                    if objective_p == "%P":
                        main_col, aux_col = "Needed P (3cr)", "Needed S less (3cr)"
                    elif objective_p == "%SA":
                        main_col, aux_col = "Needed SA (3cr)", "Needed Non-SA less (3cr)"
                    else:
                        main_col, aux_col = "Needed OTHER less (3cr)", "Needed Non-OTHER more (3cr)"
        
                    rows = []
                    for label in idx_all:
                        Pv, Sv = float(p.get(label,0.0)), float(s.get(label,0.0))
                        SAv, PAv = float(sa.get(label,0.0)), float(pa.get(label,0.0))
                        SPv, IPv = float(sp.get(label,0.0)), float(ip.get(label,0.0))
                        OTv      = float(oth.get(label,0.0))
        
                        need1, need2 = _needed_pairs_for_obj(
                            objective_p, scope_label_p,
                            Pv, Sv, SAv, PAv, SPv, IPv, OTv, totals, credits_each=3.0
                        )
                        area_vals = {"P":Pv,"S":Sv,"SA":SAv,"PA":PAv,"SP":SPv,"IP":IPv,"OTHER":OTv}
                        up_pp, down_pp = _impact_pair(objective_p, area_vals, totals, scope_label_p, credits_each=3.0)
        
                        rows.append({
                            "Program": label,
                            main_col: int(need1),
                            aux_col:  int(need2),
                            "Impact +3cr (pp)": up_pp,
                            "Impact -3cr (pp)": down_pp
                        })
        
                    need_tbl_p = pd.DataFrame(rows)
                    fmt_map_p = {}
                    if 'Program' in need_tbl_p.columns:
                        fmt_map_p['Program'] = '{}'
                    for col in ['Needed P (3cr)', 'Needed S less (3cr)', 'Needed SA (3cr)', 'Needed OTHER less (3cr)', 'Needed OTHER more (3cr)']:
                        if col in need_tbl_p.columns:
                            fmt_map_p[col] = '{:.0f}'
                    for col in ['Impact +3cr (pp)', 'Impact -3cr (pp)']:
                        if col in need_tbl_p.columns:
                            fmt_map_p[col] = '{:+.2f}'
                    
                    styled_p = (
                        need_tbl_p.style
                        .format(fmt_map_p)
                        .apply(_style_impact_heatmap, id_col="Program", axis=None)  # HEATMAP aplicado aquí
                        .hide(axis="index")
                    )
                    
                    _download_xlsx_button(
                        need_tbl_p,
                        f"needed_ByProgram_{_slugify(sel_label)}_{_slugify(objective_p)}_{_slugify(scope_label_p)}.xlsx",
                        key=f"dl_need_prog_{_slugify(sel_label)}_{_slugify(objective_p)}_{_slugify(scope_label_p)}",
                        label="⬇️ Descargar (Excel)"
                    )
                    
                    st.markdown(styled_p.to_html(escape=False), unsafe_allow_html=True)
        
            # ====== Series históricas por Program ======
            df_hist = df_car_global.copy()
            agg_ps_all_p = (df_hist.groupby(["_SEM","_PROG","_PS"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["P","S"]:
                if k not in agg_ps_all_p.columns: agg_ps_all_p[k] = 0.0
            agg_ps_all_p["P_share"] = (agg_ps_all_p["P"] / (agg_ps_all_p["P"] + agg_ps_all_p["S"]).replace(0, pd.NA)) * 100
            agg_ps_all_p = agg_ps_all_p.reset_index()
        
            agg_tipo_all_p = (df_hist.groupby(["_SEM","_PROG","_TIPO"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["SA","PA","SP","IP","OTHER"]:
                if k not in agg_tipo_all_p.columns: agg_tipo_all_p[k] = 0.0
            den_all_p = (agg_tipo_all_p[["SA","PA","SP","IP","OTHER"]].sum(axis=1)).replace(0, pd.NA)
            agg_tipo_all_p["SA_share"] = (agg_tipo_all_p["SA"] / den_all_p) * 100
            agg_tipo_all_p["OTHER_share"] = (agg_tipo_all_p["OTHER"] / den_all_p) * 100
            agg_tipo_all_p = agg_tipo_all_p.reset_index()
        
            tot_by_sem_P_p = (df_hist.groupby(["_SEM","_PS"])["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["P","S"]:
                if k not in tot_by_sem_P_p.columns: tot_by_sem_P_p[k] = 0.0
            tot_by_sem_P_p["P_share"] = (tot_by_sem_P_p["P"] / (tot_by_sem_P_p["P"] + tot_by_sem_P_p["S"]).replace(0, pd.NA)) * 100
            tot_by_sem_P_p = tot_by_sem_P_p.reset_index()
        
            tot_by_sem_tipo_p = (df_hist.groupby(["_SEM","_TIPO"])["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["SA","PA","SP","IP","OTHER"]:
                if k not in tot_by_sem_tipo_p.columns: tot_by_sem_tipo_p[k] = 0.0
            den_tot_p = (tot_by_sem_tipo_p[["SA","PA","SP","IP","OTHER"]].sum(axis=1)).replace(0, pd.NA)
            tot_by_sem_tipo_p["SA_share"] = (tot_by_sem_tipo_p["SA"] / den_tot_p) * 100
            tot_by_sem_tipo_p["OTHER_share"] = (tot_by_sem_tipo_p["OTHER"] / den_tot_p) * 100
            tot_by_sem_tipo_p = tot_by_sem_tipo_p.reset_index()
        
            # Adaptación a modo temporal
            agg_ps_all_p_tm  = transform_for_time_mode_ps(agg_ps_all_p.rename(columns={"_PROG":"__LEVEL__"})).rename(columns={"__LEVEL__":"_PROG"})
            agg_tipo_sa_p_tm = transform_for_time_mode_tipo(agg_tipo_all_p.rename(columns={"_PROG":"__LEVEL__"}), "SA_share").rename(columns={"__LEVEL__":"_PROG"})
            agg_tipo_ot_p_tm = transform_for_time_mode_tipo(agg_tipo_all_p.rename(columns={"_PROG":"__LEVEL__"}), "OTHER_share").rename(columns={"__LEVEL__":"_PROG"})
            agg_tipo_all_p_tm = (
                agg_tipo_sa_p_tm.drop(columns=[c for c in ["OTHER_share"] if c in agg_tipo_sa_p_tm], errors="ignore")
                .merge(agg_tipo_ot_p_tm[["_SEM","_PROG","OTHER","SA","PA","SP","IP","OTHER_share"]],
                       on=["_SEM","_PROG","SA","PA","SP","IP","OTHER"], how="outer")
            )
            tot_by_sem_P_p_tm = transform_for_time_mode_ps(tot_by_sem_P_p.copy())
            tot_tipo_sa_p_tm  = transform_for_time_mode_tipo(tot_by_sem_tipo_p.copy(), "SA_share")
            tot_tipo_ot_p_tm  = transform_for_time_mode_tipo(tot_by_sem_tipo_p.copy(), "OTHER_share")
            tot_by_sem_tipo_p_tm = (
                tot_tipo_sa_p_tm.drop(columns=[c for c in ["OTHER_share"] if c in tot_tipo_sa_p_tm], errors="ignore")
                .merge(tot_tipo_ot_p_tm[["_SEM","SA","PA","SP","IP","OTHER","OTHER_share"]],
                       on=["_SEM","SA","PA","SP","IP","OTHER"], how="outer")
            )
        
            key_col_p, x_labels_p, x_map_p = build_time_axis_for_history(df_hist)
            if time_mode == "Semestral":
                sel_x_p = x_map_p.get(str(sel_sem)) if sel_sem else None
                sel_label_exact_p = str(sel_sem) if sel_sem else None
            elif time_mode == "Anual":
                sel_x_p = x_map_p.get(sel_year) if sel_year is not None else None
                sel_label_exact_p = sel_year
            else:
                inter_label_p = f"{sel_year} Intersemestral" if sel_year else None
                sel_x_p = x_map_p.get(inter_label_p) if inter_label_p else None
                sel_label_exact_p = inter_label_p
        
            if SENS["on"] and SENS["ops"] and sel_label_exact_p is not None:
                agg_ps_all_p_tm, agg_tipo_all_p_tm, tot_by_sem_P_p_tm, tot_by_sem_tipo_p_tm = apply_sensitivity_to_history(
                    agg_ps_all_p_tm, agg_tipo_all_p_tm, tot_by_sem_P_p_tm, tot_by_sem_tipo_p_tm,
                    level_name="_PROG",
                    sel_label_value=sel_label_exact_p,
                    ops=SENS["ops"],
                    member_all_label="All"
                )
        
            progs_all = sorted(set(agg_ps_all_p_tm["_PROG"].astype(str).unique()) | set(agg_tipo_all_p_tm["_PROG"].astype(str).unique()))
            with colP_R:
                draw_history(
                    "Evolution by Program",
                    level_name="_PROG",
                    level_values=progs_all,
                    metric_kind="%P",
                    total_series_builders={"P": tot_by_sem_P_p_tm, "SA": tot_by_sem_tipo_p_tm, "OTHER": tot_by_sem_tipo_p_tm},
                    agg_ps_all=agg_ps_all_p_tm,
                    agg_tipo_all=agg_tipo_all_p_tm,
                    x_labels=x_labels_p, x_map=x_map_p, sel_x=sel_x_p
                )

# --------------------------
# CREDIT SUMS (EXPANDER)
# --------------------------
try:
    period_df = df_car_filt_all.copy()
    if "_CRED"  not in period_df.columns and col_cred:  period_df["_CRED"]  = pd.to_numeric(period_df[col_cred], errors="coerce").fillna(0.0)
    if "_PS"    not in period_df.columns and col_ps_C:  period_df["_PS"]    = _norm_str(period_df[col_ps_C]).map(normalize_ps)
    if "_TIPO"  not in period_df.columns and col_tipoC: period_df["_TIPO"]  = _norm_str(period_df[col_tipoC]).map(normalize_tipo)
    if "_AREA"  not in period_df.columns and col_areaCourse: period_df["_AREA"] = period_df[col_areaCourse].astype(str).str.strip()
    if "_FIELD" not in period_df.columns and col_field:      period_df["_FIELD"] = period_df[col_field].astype(str).str.strip()
    if "_PROG"  not in period_df.columns and col_prog:       period_df["_PROG"] = period_df[col_prog].astype(str).str.strip()

    view = st.session_state.view_mode if "view_mode" in st.session_state else "By Academic Area"
    if view == "By Academic Area":
        dim_col, dim_label = "_AREA", "Academic Area"
    elif view == "By Field":
        dim_col, dim_label = "_FIELD", "Field"
    else:
        dim_col, dim_label = "_PROG", "Program"

    if dim_col in period_df.columns:
        base_index = period_df.groupby(dim_col)["_CRED"].sum().sort_values(ascending=False)
        idx = base_index.index

        sum_total = base_index.rename("Credit Sum")
        sum_P  = (period_df[period_df["_PS"]   == "P"     ].groupby(dim_col)["_CRED"].sum().reindex(idx, fill_value=0.0)).rename("P Sum")
        sum_S  = (period_df[period_df["_PS"]   == "S"     ].groupby(dim_col)["_CRED"].sum().reindex(idx, fill_value=0.0)).rename("S Sum")
        sum_SA = (period_df[period_df["_TIPO"] == "SA"    ].groupby(dim_col)["_CRED"].sum().reindex(idx, fill_value=0.0)).rename("SA Sum")
        sum_PA = (period_df[period_df["_TIPO"] == "PA"    ].groupby(dim_col)["_CRED"].sum().reindex(idx, fill_value=0.0)).rename("PA Sum")
        sum_SP = (period_df[period_df["_TIPO"] == "SP"    ].groupby(dim_col)["_CRED"].sum().reindex(idx, fill_value=0.0)).rename("SP Sum")
        sum_IP = (period_df[period_df["_TIPO"] == "IP"    ].groupby(dim_col)["_CRED"].sum().reindex(idx, fill_value=0.0)).rename("IP Sum")
        sum_OT = (period_df[period_df["_TIPO"] == "OTHER" ].groupby(dim_col)["_CRED"].sum().reindex(idx, fill_value=0.0)).rename("OTHER Sum")

        tbl = pd.concat([sum_total, sum_P, sum_S, sum_SA, sum_PA, sum_SP, sum_IP, sum_OT], axis=1).fillna(0.0)

        if SENS.get("on"):
            agg_tipo = (period_df.groupby([dim_col,"_TIPO"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["SA","PA","SP","IP","OTHER"]:
                if k not in agg_tipo.columns: agg_tipo[k] = 0.0
            agg_ps = (period_df.groupby([dim_col,"_PS"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["P","S"]:
                if k not in agg_ps.columns: agg_ps[k] = 0.0
            agg_ps = agg_ps[["P","S"]]; agg_tipo = agg_tipo[["SA","PA","SP","IP","OTHER"]]

            mod_ps, mod_tipo = apply_ops_to_aggs(agg_ps, agg_tipo, SENS.get("ops", []), member_all_label="All")
            tbl["P Sum"]     = mod_ps["P"].reindex(tbl.index, fill_value=0.0)
            tbl["S Sum"]     = mod_ps["S"].reindex(tbl.index, fill_value=0.0)
            tbl["SA Sum"]    = mod_tipo["SA"].reindex(tbl.index, fill_value=0.0)
            tbl["PA Sum"]    = mod_tipo["PA"].reindex(tbl.index, fill_value=0.0)
            tbl["SP Sum"]    = mod_tipo["SP"].reindex(tbl.index, fill_value=0.0)
            tbl["IP Sum"]    = mod_tipo["IP"].reindex(tbl.index, fill_value=0.0)
            tbl["OTHER Sum"] = mod_tipo["OTHER"].reindex(tbl.index, fill_value=0.0)
            tbl["Credit Sum"]= tbl[["P Sum","S Sum"]].sum(axis=1)

        total_row = pd.DataFrame(tbl.sum(axis=0)).T
        total_row.index = ["TOTAL"]
        tbl_out = pd.concat([tbl, total_row], axis=0)

        display_label = st.session_state.get('sel_label','Selected Period')
        with st.expander(f"Credit sums by {dim_label}", expanded=False):
            export_tbl = tbl_out.reset_index().rename(columns={"index": dim_label})
            _download_xlsx_button(export_tbl,
                                  f"credit_sums_{_slugify(dim_label)}_{_slugify(display_label)}.xlsx",
                                  key=f"dl_credit_sums_{_slugify(dim_label)}_{_slugify(display_label)}",
                                  label=f"⬇️ Descargar tabla {display_label} (Excel)")
            st.dataframe(tbl_out.style.format("{:,.0f}"), use_container_width=True)

            # ===== Selector propio de dimensión (independiente del gráfico superior) =====
            # Construye lista de miembros visibles en esta tabla
            members = [str(x) for x in tbl.index.tolist() if str(x) != "TOTAL"]
            members_sorted = sorted(set(members))
            dim_options = ["(All)", "(TOTAL)"] + members_sorted
            dim_opt = st.selectbox(
                f"Select {dim_label} for the evolution lines",
                dim_options,
                index=1,  # por defecto "(TOTAL)"
                key=f"credit_dim_selector_{dim_col}"
            )
            # Nota: para evitar 5×N líneas, si eligen "(All)" mostramos TOTAL
            if dim_opt == "(All)":
                dim_opt_eff = "(TOTAL)"
            else:
                dim_opt_eff = dim_opt

            # ===== Toggle de series: Qualifications ↔ P/S =====
            mode_line = st.radio(
                "",
                ["Qualifications", "P/S"],
                horizontal=True,
                key=f"credit_line_mode_{dim_col}"
            )

            # --- histórico base normalizado ---
            df_hist = df_car_global.copy()
            if "_CRED" not in df_hist.columns and col_cred:
                df_hist["_CRED"] = pd.to_numeric(df_hist[col_cred], errors="coerce").fillna(0.0)
            if "_TIPO" not in df_hist.columns and col_tipoC:
                df_hist["_TIPO"] = _norm_str(df_hist[col_tipoC]).map(normalize_tipo)
            if "_PS" not in df_hist.columns and col_ps_C:
                df_hist["_PS"] = _norm_str(df_hist[col_ps_C]).map(normalize_ps)
            if "_SEM" not in df_hist.columns:
                sc = _get_any(df_hist, "Semestre","Periodo","Periodo Académico","Periodo academico")
                df_hist["_SEM"] = df_hist[sc].astype(str).str.strip() if sc else ""

            # columna de dimensión si falta
            if dim_col == "_AREA" and "_AREA" not in df_hist.columns and col_areaCourse:
                df_hist["_AREA"] = df_hist[col_areaCourse].astype(str).str.strip()
            if dim_col == "_FIELD" and "_FIELD" not in df_hist.columns and col_field:
                df_hist["_FIELD"] = df_hist[col_field].astype(str).str.strip()
            if dim_col == "_PROG" and "_PROG" not in df_hist.columns and col_prog:
                df_hist["_PROG"] = df_hist[col_prog].astype(str).str.strip()

            # filtro por miembro elegido en este selector
            if dim_opt_eff != "(TOTAL)" and dim_col in df_hist.columns:
                df_hist = df_hist[df_hist[dim_col].astype(str).str.strip() == str(dim_opt_eff)]

            # --- agregaciones base ---
            cats_qual = ["SA","PA","SP","IP","OTHER"]

            agg_tipo = (
                df_hist.groupby(["_SEM","_TIPO"], dropna=False)["_CRED"]
                .sum().unstack(fill_value=0.0)
            )
            for k in cats_qual:
                if k not in agg_tipo.columns:
                    agg_tipo[k] = 0.0
            agg_tipo = agg_tipo[cats_qual].reset_index()

            agg_ps = (
                df_hist.groupby(["_SEM","_PS"], dropna=False)["_CRED"]
                .sum().unstack(fill_value=0.0)
            )
            for k in ["P","S"]:
                if k not in agg_ps.columns:
                    agg_ps[k] = 0.0
            agg_ps = agg_ps[["P","S"]].reset_index()

            # --- adaptar a modo temporal (sumas) ---
            tm = st.session_state.get("time_mode", "Semestral")
            def adapt_time_sum(df_in: pd.DataFrame, value_cols: list[str]) -> pd.DataFrame:
                tmp = df_in.copy()
                tmp["_YEAR"] = tmp["_SEM"].map(extract_year_from_period)
                tmp["_INTER_LABEL"] = tmp["_SEM"].map(lambda s: f"{extract_year_from_period(s)} Intersemestral" if "inter" in str(s).lower() else None)
                if tm == "Semestral":
                    out = tmp.rename(columns={"_SEM":"_X"})
                elif tm == "Anual":
                    out = tmp.groupby("_YEAR", dropna=False)[value_cols].sum().reset_index().rename(columns={"_YEAR":"_X"})
                else:
                    inter_only = tmp[~tmp["_INTER_LABEL"].isna()].copy()
                    out = inter_only.groupby("_INTER_LABEL", dropna=False)[value_cols].sum().reset_index().rename(columns={"_INTER_LABEL":"_X"})
                return out

            plot_qual = adapt_time_sum(agg_tipo, cats_qual)
            plot_ps   = adapt_time_sum(agg_ps, ["P","S"])

            # --- eje X consistente ---
            def build_axis(df_x: pd.DataFrame) -> tuple[list, dict]:
                if tm == "Semestral":
                    x_labels = sorted(
                        {x for x in df_x["_X"].dropna().astype(str) if period_suffix(x) in {"10","20"}},
                        key=_period_sort_key
                    )
                elif tm == "Anual":
                    x_labels = sorted({int(x) for x in df_x["_X"].dropna()}, key=int)
                else:
                    x_labels = sorted(
                        df_x["_X"].dropna().astype(str).unique().tolist(),
                        key=lambda s: int(str(s).split()[0]) if str(s).split() else 0
                    )
                x_map = {lab: i for i, lab in enumerate(x_labels)}
                return x_labels, x_map

            x_labels_q, x_map_q = build_axis(plot_qual)
            x_labels_ps, x_map_ps = build_axis(plot_ps)
            plot_qual["_xi"] = plot_qual["_X"].map(x_map_q)
            plot_ps["_xi"]   = plot_ps["_X"].map(x_map_ps)
            plot_qual = plot_qual.sort_values("_xi")
            plot_ps   = plot_ps.sort_values("_xi")

            # --- sensibilidad SOLO en el período seleccionado ---
            sel_label_exact = None
            if tm == "Semestral":
                sel_sem_ = st.session_state.get("sel_sem")
                sel_label_exact = str(sel_sem_) if sel_sem_ else None
            elif tm == "Anual":
                sel_year_ = st.session_state.get("sel_year")
                sel_label_exact = sel_year_
            else:
                sel_year_ = st.session_state.get("sel_inter_year")
                sel_label_exact = f"{sel_year_} Intersemestral" if sel_year_ else None

            if SENS.get("on") and SENS.get("ops") and sel_label_exact is not None:
                # Qualifications
                sens_tipo = plot_qual[["_X"] + cats_qual].rename(columns={"_X":"_SEM"}).copy()
                dummy_ps = pd.DataFrame({"_SEM": sens_tipo["_SEM"]})
                sens_tipo2, _A, _B, _C = apply_sensitivity_to_history(
                    agg_ps_tm=dummy_ps, agg_tipo_tm=sens_tipo,
                    tot_ps_tm=dummy_ps.copy(), tot_tipo_tm=sens_tipo.copy(),
                    level_name="_SEM", sel_label_value=sel_label_exact,
                    ops=SENS["ops"], member_all_label="All"
                )
                plot_qual[cats_qual] = sens_tipo2[cats_qual].values

                # P/S
                sens_ps = plot_ps[["_X","P","S"]].rename(columns={"_X":"_SEM"}).copy()
                dummy_tipo = pd.DataFrame({"_SEM": sens_ps["_SEM"], "SA":0.0,"PA":0.0,"SP":0.0,"IP":0.0,"OTHER":0.0})
                sens_ps2, _tq, _tps, _ttq = apply_sensitivity_to_history(
                    agg_ps_tm=sens_ps, agg_tipo_tm=dummy_tipo,
                    tot_ps_tm=sens_ps.copy(), tot_tipo_tm=dummy_tipo.copy(),
                    level_name="_SEM", sel_label_value=sel_label_exact,
                    ops=SENS["ops"], member_all_label="All"
                )
                plot_ps[["P","S"]] = sens_ps2[["P","S"]].values

            # --- dibujar ---
            if mode_line == "Qualifications":
                COL_SA = "#1FA89B"  # menta verdoso
                COL_PA = "#C1C6CD"  # verde apagado
                COL_SP = "#565656"  # azul grisoso
                COL_IP = "#8F8F8F"  # gris
                COL_OT = "#A13B3B"  # rojo claro
                cmap = {"SA":COL_SA, "PA":COL_PA, "SP":COL_SP, "IP":COL_IP, "OTHER":COL_OT}

                fig = go.Figure()
                for k in ["SA","PA","SP","IP","OTHER"]:
                    fig.add_trace(go.Scatter(
                        x=plot_qual["_xi"], y=plot_qual[k],
                        mode="lines+markers",
                        name=k,
                        line=dict(width=2, color=cmap[k]),
                        marker=dict(size=6, color=cmap[k]),
                        hovertemplate=f"{k}<br>%{{y:.0f}} cr<extra></extra>"
                    ))

                sel_x = x_map_q.get(str(sel_label_exact)) if (sel_label_exact is not None) else None
                if sel_x is not None:
                    fig.add_vrect(x0=sel_x-0.5, x1=sel_x+0.5, fillcolor="#E8FAF7", opacity=0.5, layer="below", line_width=0)

                tickvals = list(range(len(x_labels_q)))
                ticktext = [str(x) for x in x_labels_q]
                fig.update_layout(
                    title=f"Evolution of Credits — Qualifications ({dim_opt_eff})",
                    margin=dict(l=10,r=10,t=40,b=60),
                    legend=dict(orientation="h", y=-0.2, yanchor="top", x=0.5, xanchor="center"),
                )
                fig.update_xaxes(title=None, tickmode="array", tickvals=tickvals, ticktext=ticktext)
                fig.update_yaxes(title="Credits", rangemode="tozero")
                st.plotly_chart(fig, use_container_width=True)

            else:
                COL_P = "#1FA89B"  # P (verde menta)
                COL_S = "#9E9E9E"  # S (gris)

                fig2 = go.Figure()
                fig2.add_trace(go.Scatter(
                    x=plot_ps["_xi"], y=plot_ps["P"],
                    mode="lines+markers",
                    name="P",
                    line=dict(width=2, color=COL_P),
                    marker=dict(size=6, color=COL_P),
                    hovertemplate="P<br>%{y:.0f} cr<extra></extra>"
                ))
                fig2.add_trace(go.Scatter(
                    x=plot_ps["_xi"], y=plot_ps["S"],
                    mode="lines+markers",
                    name="S",
                    line=dict(width=2, color=COL_S),
                    marker=dict(size=6, color=COL_S),
                    hovertemplate="S<br>%{y:.0f} cr<extra></extra>"
                ))

                sel_x2 = x_map_ps.get(str(sel_label_exact)) if (sel_label_exact is not None) else None
                if sel_x2 is not None:
                    fig2.add_vrect(x0=sel_x2-0.5, x1=sel_x2+0.5, fillcolor="#E8FAF7", opacity=0.5, layer="below", line_width=0)

                tickvals2 = list(range(len(x_labels_ps)))
                ticktext2 = [str(x) for x in x_labels_ps]
                fig2.update_layout(
                    title=f"Evolution of Credits — P/S ({dim_opt_eff})",
                    margin=dict(l=10,r=10,t=40,b=60),
                    legend=dict(orientation="h", y=-0.2, yanchor="top", x=0.5, xanchor="center"),
                )
                fig2.update_xaxes(title=None, tickmode="array", tickvals=tickvals2, ticktext=ticktext2)
                fig2.update_yaxes(title="Credits", rangemode="tozero")
                st.plotly_chart(fig2, use_container_width=True)

except Exception:
    # Evita romper la app si algo falla en este bloque
    pass

# ==========================================================
# HELPERS (únicos en este módulo, sin duplicados)
# ==========================================================
def _extract_year(s):
    m = re.search(r"(19|20)\d{2}", str(s) if s is not None else "")
    return int(m.group(0)) if m else None

def _normalize_sem_str(x: str) -> str:
    return str(x).strip().replace("\xa0", " ")

def _ensure_pid(df: pd.DataFrame) -> pd.DataFrame:
    """Agrega _PID (persona) usando ID; si no hay, usa nombre; si no, índice."""
    out = df.copy()
    idc   = _get_any(out, "ID","ID Nr.","Documento")
    namec = _get_any(out, "Profesor","PROFESOR","Docente","Nombre")
    if idc and idc in out:
        out["_PID"] = out[idc].astype(str).str.strip()
    elif namec and namec in out:
        out["_PID"] = out[namec].astype(str).str.strip().str.lower()
    else:
        out["_PID"] = out.index.astype(str)
    return out

def _norm_gender(x: str) -> str:
    v = str(x).strip().lower()
    if v in {"male","masculino","m","hombre"}:   return "Male"
    if v in {"female","femenino","f","mujer"}:   return "Female"
    return "Other"

def _is_doctoral(x: str) -> bool:
    v = str(x).strip().lower().replace(".", "")
    return ("phd" in v) or ("doctor" in v)

def _norm_ftpt(x: str) -> str:
    v = str(x).strip().upper()
    if "PLANTA"  in v: return "PLANTA"
    if "CATEDRA" in v or "CÁTEDRA" in v: return "CÁTEDRA"
    return ""

def filter_df_fd(df_fd_base: pd.DataFrame, time_mode: str, sel_year: int | None, sel_sem_code: str | int | None) -> pd.DataFrame:
    """Filtra Faculty Distribution según alcance temporal actual."""
    if df_fd_base is None or df_fd_base.empty:
        return pd.DataFrame()
    df = df_fd_base.copy()
    semc_fd = _get_any(df, "Semestre","Periodo","Periodo Académico","Periodo academico")
    if not semc_fd:
        return df
    df["_SEM_SRC"] = df[semc_fd].map(_normalize_sem_str)
    df["_YEARX"]   = df["_SEM_SRC"].map(_extract_year).astype("Int64")
    if time_mode == "Semestral" and sel_sem_code is not None:
        goal = _normalize_sem_str(sel_sem_code)
        return df[df["_SEM_SRC"].eq(goal)].copy()
    if time_mode == "Intersemestral" and sel_year is not None:
        goal = f"{int(sel_year)} Intersemestral"
        return df[df["_SEM_SRC"].str.fullmatch(re.escape(goal), case=False, na=False)].copy()
    if time_mode == "Anual" and sel_year is not None:
        return df[df["_YEARX"] == int(sel_year)].copy()
    return df

def _first_map(df_, key_col, val_col):
    if key_col not in df_ or val_col not in df_:
        return {}
    tmp = df_[[key_col, val_col]].dropna()
    return tmp.drop_duplicates(subset=[key_col]).set_index(key_col)[val_col].to_dict()

def _pick(df_, *cands):
    c = _get_any(df_, *cands)
    return df_[c] if c else pd.Series([None]*len(df_), index=df_.index)


# --------------------------
# DETAIL TABLE + DONUT + SEARCH
# (Oculto automáticamente cuando Sensitivity mode está activo)
# --------------------------
if not SENS.get("on", False):
    try:
        # ---------- Config vista ----------
        cfg = {
            "By Academic Area": {"key": "_AREA_filter",  "col": "_AREA", "label": "area",    "metric_key": "metric__AREA"},
            "By Field":         {"key": "_FIELD_filter", "col": "_FIELD","label": "campo",   "metric_key": "metric__FIELD"},
            "By Program":       {"key": "_MAT_filter",   "col": "_MAT",  "label": "programa","metric_key": "metric__MAT"},
        }
        view = st.session_state.view_mode

        if view in cfg:
            key        = cfg[view]["key"]
            col_tag    = cfg[view]["col"]
            metric_key = cfg[view]["metric_key"]
            metric_choice = st.session_state.get(metric_key, "%P")
            opt_val    = st.session_state.get(key, "(All)")

            # ---------- Base Cartelera enriquecida mínima ----------
            base = df_car_filt_all.copy()
            if "_AREA"  not in base.columns and col_areaCourse: base["_AREA"]  = base[col_areaCourse].astype(str).str.strip()
            if "_FIELD" not in base.columns and col_field:      base["_FIELD"] = base[col_field].astype(str).str.strip()
            if "_MAT"   not in base.columns and col_prog:       base["_MAT"]   = base[col_prog].astype(str).str.strip()
            if "_TIPO"  not in base.columns and col_tipoC:      base["_TIPO"]  = _norm_str(base[col_tipoC]).map(normalize_tipo)
            if "_PS"    not in base.columns and col_ps_C:       base["_PS"]    = _norm_str(base[col_ps_C]).map(normalize_ps)
            if "_CRED"  not in base.columns and col_cred:       base["_CRED"]  = pd.to_numeric(base[col_cred], errors="coerce").fillna(0.0)

            cL, cR = st.columns([7,5], gap="large")

            # ==================================================
            # IZQUIERDA: Tabla detalle + descarga
            # ==================================================
            with cL:
                # Selector de filtro de tabla según métrica
                if metric_choice == "%P":
                    table_filter = st.radio(
                        "", ["All", "Only P", "Only S"], index=0, horizontal=True,
                        key=f"table_filt_ps_{view}_{opt_val}"
                    )
                else:
                    table_filter = st.radio(
                        "", ["All", "Only SA", "Only OTHER"], index=0, horizontal=True,
                        key=f"table_filt_tipo_{view}_{opt_val}"
                    )

                base_tbl = base.copy()
                if opt_val not in {"(All)", "(TOTAL)"} and col_tag in base_tbl.columns:
                    base_tbl = base_tbl[base_tbl[col_tag] == opt_val].copy()

                if metric_choice == "%P":
                    if table_filter == "Only P":   base_tbl = base_tbl[base_tbl["_PS"] == "P"]
                    elif table_filter == "Only S": base_tbl = base_tbl[base_tbl["_PS"] == "S"]
                else:
                    if table_filter == "Only SA":        base_tbl = base_tbl[base_tbl["_TIPO"] == "SA"]
                    elif table_filter == "Only OTHER":   base_tbl = base_tbl[base_tbl["_TIPO"] == "OTHER"]

                wanted_map = {
                    "Semestre": col_sem, "Código Materia": col_code, "Créditos": col_cred,
                    "Nombre largo curso": col_name, "Program": col_prog, "Profesor": col_prof,
                    "Area del curso": col_areaCourse, "Field": col_field, "TIPO": col_tipoC, "P/S": col_ps_C,
                }
                present = {nice: col for nice, col in wanted_map.items() if col in base_tbl.columns}
                out = base_tbl[list(present.values())].rename(columns={v: k for k, v in present.items()})

                display_label = st.session_state.get('sel_label','Selected Period')
                n_courses = len(out)

                # Título dinámico
                def _title_generic(n):
                    return (f"{n} courses were taught in {display_label}"
                            if opt_val in {"(TOTAL)", "(All)"} else f"{n} courses of {opt_val} were taught in {display_label}")

                if metric_choice == "%P":
                    if   table_filter == "Only P":   title = f"{n_courses} courses taught in {display_label} by Participating Faculty"
                    elif table_filter == "Only S":   title = f"{n_courses} courses taught in {display_label} by Supporting Faculty"
                    else:                             title = _title_generic(n_courses)
                else:
                    if   table_filter == "Only SA":     title = f"{n_courses} courses taught in {display_label} by Scholarly Academics"
                    elif table_filter == "Only OTHER":  title = f"{n_courses} courses taught in {display_label} by Others"
                    else:                                title = _title_generic(n_courses)

                st.markdown(f"### {title}")
                _download_xlsx_button(
                    out,
                    f"table_detail_{_slugify(opt_val)}_{_slugify(display_label)}.xlsx",
                    key=f"dl_tbl_detail_{_slugify(opt_val)}_{_slugify(display_label)}",
                    label="⬇️ Descargar tabla (Excel)"
                )
                st.dataframe(out, use_container_width=True, hide_index=True)

            # ==================================================
            # DERECHA: Donut %P o %TIPO + descarga
            # ==================================================
            with cR:
                st.markdown("<div style='height: 110px'></div>", unsafe_allow_html=True)

                # Aggregates
                agg_tipo = (base.groupby([col_tag,"_TIPO"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0)) if col_tag in base.columns else pd.DataFrame()
                for k in ["SA","PA","SP","IP","OTHER"]:
                    if k not in agg_tipo.columns: agg_tipo[k] = 0.0
                agg_ps = (base.groupby([col_tag,"_PS"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0)) if col_tag in base.columns else pd.DataFrame()
                for k in ["P","S"]:
                    if k not in agg_ps.columns: agg_ps[k] = 0.0
                agg_ps = agg_ps[["P","S"]]
                agg_tipo = agg_tipo[["SA","PA","SP","IP","OTHER"]]

                if opt_val in {"(TOTAL)", "(All)"} or col_tag not in base.columns:
                    p_val, s_val = float(agg_ps["P"].sum() if not agg_ps.empty else 0.0), float(agg_ps["S"].sum() if not agg_ps.empty else 0.0)
                    sa = float(agg_tipo["SA"].sum() if not agg_tipo.empty else 0.0)
                    pa = float(agg_tipo["PA"].sum() if not agg_tipo.empty else 0.0)
                    sp = float(agg_tipo["SP"].sum() if not agg_tipo.empty else 0.0)
                    ip = float(agg_tipo["IP"].sum() if not agg_tipo.empty else 0.0)
                    other = float(agg_tipo["OTHER"].sum() if not agg_tipo.empty else 0.0)
                    title_suffix = "TOTAL"
                else:
                    row_ps = agg_ps.loc[[opt_val]] if opt_val in agg_ps.index else pd.DataFrame(columns=["P","S"])
                    row_q  = agg_tipo.loc[[opt_val]] if opt_val in agg_tipo.index else pd.DataFrame(columns=["SA","PA","SP","IP","OTHER"])
                    p_val, s_val = float(row_ps["P"].sum() if not row_ps.empty else 0.0), float(row_ps["S"].sum() if not row_ps.empty else 0.0)
                    sa = float(row_q["SA"].sum() if not row_q.empty else 0.0)
                    pa = float(row_q["PA"].sum() if not row_q.empty else 0.0)
                    sp = float(row_q["SP"].sum() if not row_q.empty else 0.0)
                    ip = float(row_q["IP"].sum() if not row_q.empty else 0.0)
                    other = float(row_q["OTHER"].sum() if not row_q.empty else 0.0)
                    title_suffix = opt_val

                donut_h = 360
                thrP = 75.0 if title_suffix == "TOTAL" else 60.0

                if metric_choice == "%P":
                    den = p_val + s_val
                    p_share = (p_val/den*100) if den else 0.0
                    alert = (p_share < thrP)
                    color_map = {"P": ("#F5A3A3" if alert else MINT), "S": "#B0B0B0"}
                    fig = px.pie(names=["P","S"], values=[p_val, s_val], color=["P","S"], color_discrete_map=color_map, hole=0.55)
                    fig.update_traces(textinfo="percent+label", hovertemplate="%{label}: %{percent:.1%}<extra></extra>")
                    fig.update_layout(
                        title=f"% Participating Distribution — {title_suffix}",
                        height=donut_h, margin=dict(l=10, r=10, t=40, b=10),
                        legend=dict(orientation="v", yanchor="bottom", y=0.4, xanchor="center", x=0.9)
                    )
                    st.plotly_chart(fig, use_container_width=True)

                    donut_df = pd.DataFrame({"Group": ["P","S"], "Credits": [p_val, s_val]})
                    donut_df["Percent"] = (donut_df["Credits"] / max(1e-9, donut_df["Credits"].sum()))*100
                    _download_xlsx_button(
                        donut_df,
                        f"chart_donut_PS_{_slugify(title_suffix)}_{_slugify(st.session_state.get('sel_label','sel'))}.xlsx",
                        key=f"dl_donut_ps_{_slugify(title_suffix)}_{_slugify(st.session_state.get('sel_label','sel'))}",
                        label="⬇️ Datos de la gráfica (Excel)"
                    )
                else:
                    labels_all = ["SA", "PA", "SP", "IP", "OTHER"]
                    values_all = [sa, pa, sp, ip, other]
                    filtered   = [(l, v) for l, v in zip(labels_all, values_all) if v > 0]
                    if filtered:
                        labels = [l for l, _ in filtered]; values = [v for _, v in filtered]
                        den = sum(values_all) or 1.0
                        sa_share    = sa/den*100
                        other_share = other/den*100
                        cmap = {l: "#B0B0B0" for l in labels}
                        if "SA" in labels:    cmap["SA"]    = ("#F5A3A3" if sa_share   < 40.0 else MINT)
                        if "OTHER" in labels: cmap["OTHER"] = ("#F5A3A3" if other_share > 10.0 else "#6B7280")

                        fig = px.pie(names=labels, values=values, color=labels, color_discrete_map=cmap, hole=0.55)
                        fig.update_traces(textinfo="percent+label", sort=False, hovertemplate="%{label}: %{percent:.1%}<extra></extra>")
                        title_txt = "%SA Distribution" if metric_choice == "%SA" else "%OTHER Distribution"
                        fig.update_layout(
                            title=f"{title_txt} — {title_suffix}",
                            height=donut_h, margin=dict(l=10, r=10, t=40, b=10),
                            legend=dict(orientation="v", yanchor="bottom", y=0.4, xanchor="center", x=0.9)
                        )
                        st.plotly_chart(fig, use_container_width=True)

                        donut_df = pd.DataFrame({"Type": labels_all, "Credits": values_all})
                        donut_df["Percent"] = (donut_df["Credits"] / max(1e-9, donut_df["Credits"].sum()))*100
                        _download_xlsx_button(
                            donut_df,
                            f"chart_donut_TIPO_{_slugify(title_suffix)}_{_slugify(st.session_state.get('sel_label','sel'))}.xlsx",
                            key=f"dl_donut_tipo_{_slugify(title_suffix)}_{_slugify(st.session_state.get('sel_label','sel'))}",
                            label="⬇️ Datos de la gráfica (Excel)"
                        )
                    else:
                        st.caption("No hay registros de TIPO para esta métrica en este período.")
    except Exception:
        pass


# --------------------------
# COUNTS — PIVOT / BSQ (Oculto cuando Sensitivity mode está activo)
# --------------------------
if not SENS.get("on", False):
    st.markdown("---")
    st.subheader(f"Participating vs Supporting — {st.session_state.get('sel_label','Selected')}")

    # ------- Base Faculty Distribution filtrada -------
    def _filter_fd_scope(df_fd_raw: pd.DataFrame) -> pd.DataFrame:
        if df_fd_raw.empty:
            return df_fd_raw.copy()
        out = df_fd_raw.copy()
        semc = _get_any(out, "Semestre","Periodo","Periodo Académico","Periodo academico")
        if semc:
            out["_SEM_SRC"]   = out[semc].astype(str).str.strip()
            out["_YEARX"]     = out["_SEM_SRC"].map(_extract_year).astype("Int64")
            out["_IS_INTER"]  = out["_SEM_SRC"].str.lower().str.contains("inter", na=False)
        else:
            out["_SEM_SRC"]  = ""
            out["_YEARX"]    = pd.Series(dtype="Int64")
            out["_IS_INTER"] = False

        time_mode = st.session_state.get("time_mode", "Semestral")
        sel_sem   = st.session_state.get("sel_sem")
        sel_year  = st.session_state.get("sel_year")

        if time_mode == "Semestral" and sel_sem is not None:
            return out[out["_SEM_SRC"].eq(str(sel_sem))].copy()
        if time_mode == "Intersemestral" and sel_year is not None:
            return out[(out["_YEARX"] == int(sel_year)) & (out["_IS_INTER"])].copy()
        if time_mode == "Anual" and sel_year is not None:
            return out[out["_YEARX"] == int(sel_year)].copy()
        return out

    df_fd_scope = _filter_fd_scope(df_fd)
    df_fd_f = df_fd_scope.copy()

    # -------- columnas base y extra --------
    if col_ps_fd:   df_fd_f["_PS"]   = _norm_str(df_fd_f[col_ps_fd]).map(normalize_ps)
    if col_area_fd: df_fd_f["_AREA"] = df_fd_f[col_area_fd].astype(str).str.strip()
    if col_tipo_fd: df_fd_f["_TIPO"] = _norm_str(df_fd_f[col_tipo_fd]).map(normalize_tipo)

    col_genero = _get_any(df_fd_f, "GÉNERO", "GENERO", "Genero", "Gender")
    col_degree = _get_any(df_fd_f, "Highest Degree", "HighestDegree", "DEGREE", "Grado máximo", "Grado")
    col_ftpt   = _get_any(df_fd_f, "PLANTA_CATEDRA", "Planta_Catedra", "Planta/Catedra", "Full/Part")

    # --------- Controles en una sola fila (3 botones) ----------
    pivot_mode = st.radio(
        "View",
        ["BSQ Compensation", "AREA", "Qualification Type"],
        index=0,  # BSQ por defecto y seleccionado
        horizontal=True,
        label_visibility="collapsed",
        key="counts_view_mode"
    )

    # ===================== MODO BSQ =====================
    if pivot_mode == "BSQ Compensation":
        left, right = st.columns([6,6], gap="large")

        if not all([col_genero, col_degree, col_ftpt]):
            st.error("Missing columns in 'Faculty Distribution' for BSQ tables: 'GÉNERO', 'Highest Degree', and/or 'PLANTA_CATEDRA'.")
        else:
            df_bsq = _ensure_pid(df_fd_f).assign(
                Gender     = df_fd_f[col_genero].map(_norm_gender),
                IsDoctoral = df_fd_f[col_degree].map(_is_doctoral),
                FTPT       = df_fd_f[col_ftpt].map(_norm_ftpt),
                PS         = df_fd_f["_PS"].fillna(""),
                TIPO       = df_fd_f["_TIPO"].fillna("OTHER")
            )

            def _count_by_gender(mask) -> dict:
                sub = df_bsq[mask].drop_duplicates(subset=["_PID"])
                male   = int((sub["Gender"] == "Male").sum())
                female = int((sub["Gender"] == "Female").sum())
                other  = int((sub["Gender"] == "Other").sum())
                return {"Male": male, "Female": female, "Other": other, "Total": male + female + other}

            row7a = _count_by_gender(df_bsq["PS"] == "P")
            row7b = _count_by_gender((df_bsq["PS"] == "P") & (df_bsq["IsDoctoral"]))
            row7c = _count_by_gender(df_bsq["PS"] == "S")
            row7d = _count_by_gender((df_bsq["PS"] == "S") & (df_bsq["IsDoctoral"]))

            tbl7 = pd.DataFrame([
                {"Row": "a. Total number of participating faculty members", **row7a},
                {"Row": "b. Total number of participating faculty members with doctoral degrees", **row7b},
                {"Row": "c. Total number of supporting faculty members", **row7c},
                {"Row": "d. Total number of supporting faculty members with doctoral degrees", **row7d},
            ])

            def _bold_rows_7(df_):
                sty = pd.DataFrame('', index=df_.index, columns=df_.columns)
                mask = df_["Row"].str.startswith(("b.", "d."))
                for c in df_.columns:
                    sty.loc[mask, c] = 'font-weight:700;'
                return sty

            cats = ["SA","PA","SP","IP","OTHER"]
            def _row_qual(ps_code: str, ftpt_code: str | None):
                m = (df_bsq["PS"] == ps_code)
                if ftpt_code is not None:
                    m = m & (df_bsq["FTPT"] == ftpt_code)
                sub = df_bsq[m].drop_duplicates(subset=["_PID","TIPO","PS","FTPT"])
                counts = {c: int((sub["TIPO"] == c).sum()) for c in cats}
                total = sum(counts.values())
                return {**counts, "TOTAL": total}

            r8a = _row_qual("P", "PLANTA")
            r8b = _row_qual("P", "CÁTEDRA")
            r8c = {k: r8a.get(k,0) + r8b.get(k,0) for k in cats + ["TOTAL"]}
            r8d = _row_qual("S", "PLANTA")
            r8e = _row_qual("S", "CÁTEDRA")
            r8f = {k: r8d.get(k,0) + r8e.get(k,0) for k in cats + ["TOTAL"]}

            tbl8 = pd.DataFrame([
                {"Row": "a. Full-time Participating faculty members", **r8a},
                {"Row": "b. Part-time Participating faculty members", **r8b},
                {"Row": "c. Total Participating faculty members", **r8c},
                {"Row": "d. Full-time Supporting faculty members", **r8d},
                {"Row": "e. Part-time Supporting faculty members", **r8e},
                {"Row": "f. Total Supporting faculty members", **r8f},
            ])[["Row"] + cats + ["TOTAL"]]

            def _bold_rows_8(df_):
                sty = pd.DataFrame('', index=df_.index, columns=df_.columns)
                mask = df_["Row"].str.startswith(("c.", "f."))
                for c in df_.columns:
                    sty.loc[mask, c] = 'font-weight:700;'
                return sty

            with left:
                st.markdown("**7. Participating and Supporting Faculty Counts †**")
                _download_xlsx_button(
                    tbl7, f"bsq_7_gender_counts_{_slugify(st.session_state.get('sel_label','sel'))}.xlsx",
                    key=f"dl_bsq7_{_slugify(st.session_state.get('sel_label','sel'))}",
                    label="Descargar tabla 7 (Excel)"
                )
                st.dataframe(
                    tbl7.style.apply(_bold_rows_7, axis=None).format({"Male":"{:,.0f}","Female":"{:,.0f}","Other":"{:,.0f}","Total":"{:,.0f}"}),
                    use_container_width=True, hide_index=True
                )

            with right:
                st.markdown("**8. Faculty Counts by Qualification Types †**")
                _download_xlsx_button(
                    tbl8, f"bsq_8_qual_counts_{_slugify(st.session_state.get('sel_label','sel'))}.xlsx",
                    key=f"dl_bsq8_{_slugify(st.session_state.get('sel_label','sel'))}",
                    label="Descargar tabla 8 (Excel)"
                )
                st.dataframe(
                    tbl8.style.apply(_bold_rows_8, axis=None).format({c: "{:,.0f}" for c in cats + ["TOTAL"]}),
                    use_container_width=True, hide_index=True
                )

    # ===================== MODO PIVOT ORIGINAL (AREA / TYPE) =====================
    else:
        # Define filas según modo
        if pivot_mode == "AREA":
            row_name   = "AREA"
            row_series = df_fd_f["_AREA"].astype(str).str.strip().replace({"": "N/A"})
            desired_order = None
        else:  # "Qualification Type"
            row_name   = "Type"
            row_series = df_fd_f["_TIPO"].map(lambda v: str(v).upper())
            desired_order = ["SA", "PA", "SP", "IP", "OTHER"]

        # Persona + variables para deduplicar
        df_cnt = _ensure_pid(df_fd_f)
        df_cnt[row_name] = row_series
        df_cnt["_PS2"]   = df_cnt["_PS"].fillna("")

        # DEDUP: 1 vez por persona y categoría (row_name, _PS2)
        df_cnt = df_cnt.drop_duplicates(subset=["_PID", row_name, "_PS2"])

        base = pd.DataFrame({row_name: df_cnt[row_name], "_PS": df_cnt["_PS2"]})
        table = (base.groupby([row_name, "_PS"], dropna=False)
                      .size()
                      .unstack(fill_value=0)
                      .rename(columns={"P": "Participating", "S": "Supporting"}))
        for k in ["Participating", "Supporting"]:
            if k not in table.columns: table[k] = 0
        table["__Total__"] = table["Participating"] + table["Supporting"]

        # Ajuste por sensibilidad (impacto total) — si hubiera operaciones cargadas
        if SENS["on"] and SENS.get("ops"):
            add_P = sum(op.get("count",0) for op in SENS["ops"] if op.get("scope")=="PS" and op.get("cat")=="P")
            add_S = sum(op.get("count",0) for op in SENS["ops"] if op.get("scope")=="PS" and op.get("cat")=="S")
            incs = {"Participating": int(add_P), "Supporting": int(add_S)}
        else:
            incs = {"Participating": 0, "Supporting": 0}

        df_counts = table[["Participating", "Supporting"]].astype(int).reset_index()
        total_row = pd.DataFrame([{row_name: "TOTAL",
                                   "Participating": int(df_counts["Participating"].sum()) + incs["Participating"],
                                   "Supporting":    int(df_counts["Supporting"].sum())    + incs["Supporting"]}])
        df_counts_out = pd.concat([df_counts, total_row], ignore_index=True)

        def _bold_total(df_):
            sty = pd.DataFrame('', index=df_.index, columns=df_.columns)
            mask = df_[row_name].astype(str).str.upper().eq("TOTAL")
            for c in df_.columns: sty.loc[mask, c] = 'font-weight:700;'
            return sty

        left, right = st.columns([6,6], gap="large")

        # Porcentajes y orden para gráfica
        denom = table["__Total__"].replace(0, pd.NA)
        perc_df = pd.DataFrame({
            row_name: table.index,
            "%Participating": (table["Participating"] / denom * 100).round(1).fillna(0.0),
            "%Supporting":    (table["Supporting"]    / denom * 100).round(1).fillna(0.0),
        })
        if desired_order:
            for code in desired_order:
                if code not in perc_df[row_name].tolist():
                    perc_df.loc[len(perc_df)] = [code, 0.0, 0.0]
            cat_order = desired_order
        else:
            cat_order = perc_df[row_name].tolist()

        chart_export = perc_df.melt(id_vars=row_name, value_vars=["%Participating", "%Supporting"],
                                    var_name="Group", value_name="Percent")

        with left:
            _download_xlsx_button(
                df_counts_out,
                f"ps_counts_{_slugify(row_name)}_{_slugify(st.session_state.get('sel_label','sel'))}.xlsx",
                key=f"dl_ps_counts_{_slugify(row_name)}_{_slugify(st.session_state.get('sel_label','sel'))}",
                label="Descargar tabla (Excel)"
            )
            styled_counts = (df_counts_out.style
                             .format({"Participating": "{:,.0f}", "Supporting": "{:,.0f}"})
                             .apply(_bold_total, axis=None))
            st.dataframe(styled_counts, use_container_width=True, hide_index=True)

        with right:
            fig = px.bar(
                chart_export, x=row_name, y="Percent", color="Group",
                barmode="group", text="Percent",
                color_discrete_map={"%Participating": MINT, "%Supporting": SUPPORTING},
                category_orders={row_name: cat_order}
            )
            fig.update_traces(texttemplate="%{text:.1f}%")
            fig.update_layout(
                xaxis_title=None, yaxis_title=None, height=340,
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
                legend_title_text=None, margin=dict(l=20, r=10, t=10, b=40)
            )
            st.plotly_chart(fig, use_container_width=True)
            _download_xlsx_button(
                chart_export,
                f"chart_ps_perc_{_slugify(row_name)}_{_slugify(st.session_state.get('sel_label','sel'))}.xlsx",
                key=f"dl_chart_ps_perc_{_slugify(row_name)}_{_slugify(st.session_state.get('sel_label','sel'))}",
                label="Descargar datos (Excel)"
            )


# ==========================================================
# MÓDULO FINAL: Top 5 (más/menos créditos) y FT sin cursos + Buscador
# (Oculto cuando Sensitivity mode está activo)
# ==========================================================
if not SENS.get("on", False):
    st.markdown("---")

    # ====== Título pegado a los controles ======
    head_l, head_r = st.columns([7,5], gap="large")
    with head_l:
        st.markdown("#### Faculty credit highlights (current timeframe)")
    with head_r:
        st.write("")

    @st.cache_data(ttl=0)
    def _load_planta_sheet():
        try:
            xls = pd.ExcelFile("data/Faculty/BD_Faculty.xlsx")
            dfp = pd.read_excel(xls, sheet_name="BD PLANTA 2020-2025")
            dfp.columns = dfp.columns.str.strip()
            return dfp
        except Exception:
            return pd.DataFrame()

    df_planta = _load_planta_sheet()

    # -------- helpers de columnas en Cartelera & Distribution --------
    col_prof_car = _get_any(df_car_filt_all, "Profesor","PROFESOR","Docente")
    col_cred_car = _get_any(df_car_filt_all, "Créditos","Creditos","Credits")
    col_sem_car  = _get_any(df_car_filt_all, "Semestre","Periodo","Periodo Académico","Periodo academico")
    col_code_car = _get_any(df_car_filt_all, "Código Materia","Codigo Materia","CODIGO MATERIA","Código","Codigo","Course Code")
    col_name_car = _get_any(df_car_filt_all, "Nombre largo curso","Nombre Curso","Nombre del curso","Course Name")
    col_secc_car = _get_any(df_car_filt_all, "Secc","Sección","Seccion","Section")
    col_acar_car = _get_any(df_car_filt_all, "Area del curso","Área del curso","Area del Curso","AREA DEL CURSO")
    col_field_car= _get_any(df_car_filt_all, "Field","FIELD","Campo","Área de conocimiento")
    col_prog_car = _get_any(df_car_filt_all, "Program","PROGRAM","program","Materia")
    col_campus   = _get_any(df_car_filt_all, "Campus","CAMPUS","Sede")

    if col_cred_car and "_CRED" not in df_car_filt_all.columns:
        df_car_filt_all["_CRED"] = pd.to_numeric(df_car_filt_all[col_cred_car], errors="coerce").fillna(0.0)

    # -------- Distribution (para enriquecer con ID/AREA/TIPO/P-S y FT/PT) --------
    col_id_fd_all   = _get_any(df_fd, "ID","ID Nr.","Documento")
    col_sem_fd_all  = _get_any(df_fd, "Semestre","Periodo","Periodo Académico","Periodo academico")

    time_mode    = st.session_state.get("time_mode","Semestral")
    sel_year     = st.session_state.get("sel_year")
    sel_sem_code = st.session_state.get("sel_sem")

    df_fd_sem = filter_df_fd(df_fd, time_mode, sel_year, sel_sem_code).copy()

    # --- Normalizaciones base para enrichment (después de filtrar df_fd_sem) ---
    col_prof_fd = _get_any(df_fd_sem, "Profesor","PROFESOR","Docente","Nombre")
    col_id_fd   = _get_any(df_fd_sem, "ID","ID Nr.","Documento")
    col_area_fd = _get_any(df_fd_sem, "AREA_PROFESOR","Area_Profesor","Area Profesor","Área","Area")
    col_tipo_fd = _get_any(df_fd_sem, "TIPO","Tipo","Ranking","Tipo Ranking")
    col_ps_fd   = _get_any(df_fd_sem, "P/S","P - S","Participating/Supporting")
    col_ftpt    = _get_any(df_fd_sem, "PLANTA_CATEDRA", "Planta_Catedra", "Planta/Catedra", "Full/Part")

    if col_prof_fd: df_fd_sem["_PROF_N"]    = df_fd_sem[col_prof_fd].astype(str).str.strip()
    if col_id_fd:   df_fd_sem["_ID"]        = df_fd_sem[col_id_fd].astype(str).str.strip()
    if col_area_fd: df_fd_sem["_AREA_PROF"] = df_fd_sem[col_area_fd].astype(str).str.strip()
    if col_tipo_fd: df_fd_sem["_TIPO"]      = df_fd_sem[col_tipo_fd].astype(str).str.strip()
    if col_ps_fd:   df_fd_sem["_PS"]        = _norm_str(df_fd_sem[col_ps_fd]).map(normalize_ps)

    # FT/PT directo o fallback por TIPO
    if col_ftpt:
        df_fd_sem["_FTPT"] = df_fd_sem[col_ftpt].map(_norm_ftpt)
    else:
        df_fd_sem["_FTPT"] = df_fd_sem["_TIPO"].map(_norm_ftpt) if "_TIPO" in df_fd_sem else ""

    # --- Maps por profesor (primera coincidencia) ---
    prof_to_id_map_by_name = _first_map(df_fd_sem, "_PROF_N", "_ID")
    prof_to_area_map       = _first_map(df_fd_sem, "_PROF_N", "_AREA_PROF")
    prof_to_tipo_map       = _first_map(df_fd_sem, "_PROF_N", "_TIPO")
    prof_to_ps_map         = _first_map(df_fd_sem, "_PROF_N", "_PS")

    # Conjunto de IDs PLANTA (cruce por ID)
    planta_ids = set(df_fd_sem.loc[df_fd_sem["_FTPT"] == "PLANTA", "_ID"].dropna().astype(str).unique().tolist())

    # ========= Controles de “Top / Zero” (izquierda) + Buscador (derecha) =========
    opt_highlight = st.radio(
        "Show",
        ["Top 5 most credits", "Top 5 least credits", "Full-time with 0 courses"],
        index=0, horizontal=True, label_visibility="visible", key="highlight_mode"
    )

    left, right = st.columns([7,5], gap="large")

    # ======================= PANEL IZQUIERDO =======================
    with left:
        if opt_highlight in {"Top 5 most credits", "Top 5 least credits"}:
            # switch PLANTA (por ID)
            only_ft = st.toggle("Only full-time (PLANTA)", value=False, key="top_only_ft")

            if not col_prof_car or "_CRED" not in df_car_filt_all.columns:
                st.info("Missing credits or professor column in Cartelera for this view.")
            else:
                df_top = (
                    df_car_filt_all
                    .assign(_PROF=df_car_filt_all[col_prof_car].astype(str).str.strip())
                    .groupby("_PROF")
                    .agg(Credits=("_CRED","sum"), nCourses=(col_prof_car,"count"))
                    .reset_index()
                )
                df_top["ID"] = df_top["_PROF"].map(prof_to_id_map_by_name)

                if only_ft:
                    if planta_ids:
                        df_top = df_top[df_top["ID"].astype(str).isin(planta_ids)]
                    else:
                        df_top = df_top.iloc[0:0]

                asc = (opt_highlight == "Top 5 least credits")
                df_top = df_top.sort_values("Credits", ascending=asc).head(5).copy()

                # Enriquecer
                df_top["AREA_PROFESOR"] = df_top["_PROF"].map(prof_to_area_map)
                df_top["TIPO"]          = df_top["_PROF"].map(prof_to_tipo_map)
                df_top["P/S"]           = df_top["_PROF"].map(prof_to_ps_map)

                out = (
                    df_top.rename(columns={"_PROF":"Profesor"})
                          [["Profesor","ID","AREA_PROFESOR","TIPO","P/S","Credits","nCourses"]]
                          .rename(columns={"nCourses":"#Cursos"})
                )
                title = "Top 5 professors by credits (most)" if not asc else "Top 5 professors by credits (least)"
                _download_xlsx_button(
                    out,
                    f"highlight_{_slugify(title)}_{_slugify(st.session_state.get('sel_label','sel'))}.xlsx",
                    key=f"dl_highlight_{_slugify(title)}_{_slugify(st.session_state.get('sel_label','sel'))}",
                    label="Descargar (Excel)"
                )
                st.dataframe(out.style.format({"Credits":"{:,.1f}"}), use_container_width=True, hide_index=True)

        else:
            # ========= Full-time con 0 cursos =========
            col_period_pl = _get_any(df_planta, "Periodo","PERIODO","Semestre")
            col_id_pl     = _get_any(df_planta, "ID Nr.","ID","Documento")
            if df_planta.empty or not all([col_period_pl, col_id_pl]):
                st.info("Load 'BD PLANTA 2020-2025' to compute FT with 0 courses.")
            else:
                # Alcance temporal para PLANTA y DISTRIBUTION
                if time_mode == "Semestral" and sel_sem_code is not None:
                    df_ft = df_planta[df_planta[col_period_pl].astype(str).str.strip().eq(str(sel_sem_code))].copy()
                    alcance_txt = str(sel_sem_code)
                    taught_ids = set()
                    if col_sem_fd_all and col_id_fd_all:
                        taught_ids = set(
                            df_fd.loc[df_fd[col_sem_fd_all].astype(str).str.strip().eq(str(sel_sem_code)), col_id_fd_all]
                                 .astype(str).str.strip()
                        )
                elif time_mode == "Intersemestral" and sel_year is not None:
                    goal = f"{int(sel_year)} Intersemestral"
                    mask_inter = df_planta[col_period_pl].map(_normalize_sem_str).str.fullmatch(re.escape(goal), case=False, na=False)
                    df_ft = df_planta[mask_inter].copy()
                    alcance_txt = goal
                    taught_ids = set()
                    if col_sem_fd_all and col_id_fd_all:
                        taught_ids = set(
                            df_fd.loc[
                                df_fd[col_sem_fd_all].map(_normalize_sem_str).str.fullmatch(re.escape(goal), case=False, na=False),
                                col_id_fd_all
                            ].astype(str).str.strip()
                        )
                elif time_mode == "Anual" and sel_year is not None:
                    df_ft = df_planta[df_planta[col_period_pl].astype(str).str.contains(str(sel_year), na=False)].copy()
                    alcance_txt = f"{sel_year} (annual)"
                    taught_ids = set()
                    if col_sem_fd_all and col_id_fd_all:
                        taught_ids = set(
                            df_fd.loc[df_fd[col_sem_fd_all].astype(str).str.contains(str(sel_year), na=False), col_id_fd_all]
                                 .astype(str).str.strip()
                        )
                else:
                    df_ft = pd.DataFrame()
                    taught_ids = set()
                    alcance_txt = st.session_state.get('sel_label','Selected')

                if df_ft.empty:
                    st.info(f"No full-time data found for {alcance_txt}.")
                else:
                    df_ft["_ID"] = df_ft[col_id_pl].astype(str).str.strip()
                    ft_ids = set(df_ft["_ID"])
                    ft_total = len(ft_ids)
                    ft_teaching = len(ft_ids & taught_ids)
                    st.markdown(f"**De los {ft_total} profesores de planta, {ft_teaching} están dictando en {alcance_txt}.**")

                    missing_ids = sorted(ft_ids - taught_ids)
                    sub = df_ft[df_ft["_ID"].isin(missing_ids)].copy()

                    out = pd.DataFrame({
                        "ID Nr.":        sub["_ID"],
                        "First Name":    _pick(sub, "First Name","Nombre","Nombres"),
                        "Last Name":     _pick(sub, "Last Name","Apellidos","Apellido"),
                        "Academic Area": _pick(sub, "Academic Area","Área Académica","Area Académica","AREA_PROFESOR","Área"),
                        "Faculty Ranking": _pick(sub, "Faculty Ranking","Ranking","Rango"),
                        "Faculty Qualific.": _pick(sub, "Faculty Qualific.","Qualification","Qualific.","Qualif.","Tipo Ranking","TIPO"),
                        "P/S": _pick(sub, "P/S","P - S","Participating/Supporting")
                    })

                    _download_xlsx_button(
                        out, f"ft_zero_courses_{_slugify(alcance_txt)}.xlsx",
                        key=f"dl_ft_zero_{_slugify(alcance_txt)}",
                        label="Descargar (Excel)"
                    )
                    st.dataframe(out, use_container_width=True, hide_index=True)

    # ======================= PANEL DERECHO — BUSCADOR (modo único; control pegado) =======================
    with right:
        # Base: Cartelera (alcance ya filtrado en df_car_filt_all)
        base = df_car_filt_all.copy()
        if col_prof_car: base["_PROF"] = base[col_prof_car].astype(str).str.strip()
        if col_sem_car:  base["_SEM"]  = base[col_sem_car].astype(str).str.strip()
        if col_code_car: base["_CODE"] = base[col_code_car].astype(str).str.strip()
        if col_name_car: base["_NAME"] = base[col_name_car].astype(str).str.strip()

        # Enriquecer con ID/AREA por nombre (para mostrar)
        if "_PROF" in base and prof_to_id_map_by_name:
            base["_ID"] = base["_PROF"].map(prof_to_id_map_by_name)
        if "_PROF" in base and prof_to_area_map:
            base["_AREA_PROF"] = base["_PROF"].map(prof_to_area_map)

        # Opciones autocompletar
        prof_opts = [""] + (sorted(base["_PROF"].dropna().unique().tolist()) if "_PROF" in base else [])
        if "_ID" in base:
            prof_opts = [""] + sorted(set(prof_opts[1:] + [f"ID:{v}" for v in base["_ID"].dropna().astype(str).tolist()]))

        if "_NAME" in base and base["_NAME"].notna().any():
            course_opts = sorted(base["_NAME"].dropna().unique().tolist())
        elif "_CODE" in base:
            course_opts = sorted(base["_CODE"].dropna().unique().tolist())
        else:
            course_opts = []
        course_opts = [""] + course_opts

        # Espaciador para alinear hacia abajo
        st.markdown("<div style='min-height:140px'></div>", unsafe_allow_html=True)

        # Callback: al cambiar el modo, limpiar el otro selector
        def _on_mode_change():
            mode = st.session_state.get("srch_mode_right", "Por Profesor")
            if mode == "Por Profesor":
                st.session_state["srch_course"] = ""
            else:
                st.session_state["srch_prof"] = ""

        # Selector de modo pegado al buscador
        search_mode = st.radio(
            "Search mode",
            ["Por Profesor", "Por Curso"],
            index=0, horizontal=True, key="srch_mode_right",
            on_change=_on_mode_change
        )

        # Control único a todo el ancho según modo
        if search_mode == "Por Profesor":
            st.selectbox(
                "Profesor (Nombre) o ID",
                options=prof_opts,
                index=(prof_opts.index(st.session_state.get("srch_prof",""))
                       if st.session_state.get("srch_prof","") in prof_opts else 0),
                key="srch_prof"
            )
        else:  # Por Curso
            st.selectbox(
                "Curso (Nombre)",
                options=course_opts,
                index=(course_opts.index(st.session_state.get("srch_course",""))
                       if st.session_state.get("srch_course","") in course_opts else 0),
                key="srch_course"
            )

    # ======================= RESULTADOS BUSQUEDA — FULL WIDTH =======================
    sel_prof   = st.session_state.get("srch_prof", "")
    sel_course = st.session_state.get("srch_course", "")
    search_mode = st.session_state.get("srch_mode_right", "Por Profesor")

    has_query = (search_mode == "Por Profesor" and bool(sel_prof)) or (search_mode == "Por Curso" and bool(sel_course))

    if has_query:
        base = df_car_filt_all.copy()
        if col_prof_car: base["_PROF"] = base[col_prof_car].astype(str).str.strip()
        if col_sem_car:  base["_SEM"]  = base[col_sem_car].astype(str).str.strip()
        if col_code_car: base["_CODE"] = base[col_code_car].astype(str).str.strip()
        if col_name_car: base["_NAME"] = base[col_name_car].astype(str).str.strip()
        if "_PROF" in base and prof_to_id_map_by_name:
            base["_ID"] = base["_PROF"].map(prof_to_id_map_by_name)
        if "_PROF" in base and prof_to_area_map:
            base["_AREA_PROF"] = base["_PROF"].map(prof_to_area_map)

        mask_all = pd.Series(True, index=base.index)

        if search_mode == "Por Profesor" and sel_prof:
            if sel_prof.startswith("ID:"):
                qid = sel_prof.split(":",1)[1].strip()
                m = base["_ID"].astype(str).str.fullmatch(re.escape(qid), case=False) if "_ID" in base else pd.Series(False, index=base.index)
            else:
                m = base["_PROF"].str.contains(re.escape(sel_prof), case=False, na=False) if "_PROF" in base else pd.Series(False, index=base.index)
            mask_all &= m

        if search_mode == "Por Curso" and sel_course:
            m_name = base["_NAME"].str.contains(re.escape(sel_course), case=False, na=False) if "_NAME" in base else pd.Series(False, index=base.index)
            m_code = base["_CODE"].str.contains(re.escape(sel_course), case=False, na=False) if "_CODE" in base else pd.Series(False, index=base.index)
            mask_all &= (m_name | m_code)

        res = base[mask_all].copy()

        periodo_txt = st.session_state.get('sel_label','Selected')
        if search_mode == "Por Profesor" and sel_prof:
            if "_CRED" not in res.columns and col_cred_car:
                res["_CRED"] = pd.to_numeric(res[col_cred_car], errors="coerce").fillna(0.0)
            tot_cr = float(res.get("_CRED", pd.Series([0]*len(res))).sum())
            tot_courses = int(res.shape[0])
            prof_label = sel_prof
            if sel_prof.startswith("ID:") and "_PROF" in res and not res.empty:
                profs = sorted(res["_PROF"].dropna().unique().tolist())
                if len(profs) == 1:
                    prof_label = profs[0]
            st.info(f"**El profesor {prof_label} ha dictado {tot_cr:,.1f} créditos con {tot_courses} cursos en {periodo_txt}.**")

        if search_mode == "Por Curso" and sel_course:
            profs_cnt = res["_PROF"].nunique() if "_PROF" in res else 0
            st.info(f"**El curso {sel_course} ha sido dictado por {profs_cnt} profesor(es) en {periodo_txt}.**")

        show_cols = {
            "Periodo": "_SEM" if "_SEM" in res else (col_sem_car or col_sem_fd_all),
            "Profesor": col_prof_car or col_prof_fd,
            "ID": "_ID" if "_ID" in res else col_id_fd_all,
            "AREA_PROFESOR": "_AREA_PROF" if "_AREA_PROF" in res else _get_any(df_fd, "AREA_PROFESOR","Area_Profesor","Area Profesor","Área","Area"),
            "Código Materia": col_code_car,
            "Nombre largo curso": col_name_car,
            "Secc": col_secc_car,
            "Area del curso": col_acar_car,
            "Field": col_field_car,
            "Program": col_prog_car,
            "Créditos": col_cred_car,
            "Campus": col_campus
        }

        data = {}
        out_cols = []
        for nice, col in show_cols.items():
            data[nice] = res[col] if (col in res.columns) else None
            out_cols.append(nice)

        res_out = pd.DataFrame(data, columns=out_cols).copy()
        if "Créditos" in res_out.columns:
            res_out["Créditos"] = pd.to_numeric(res_out["Créditos"], errors="coerce").fillna(0.0)

        _download_xlsx_button(
            res_out,
            f"search_results_{_slugify(st.session_state.get('sel_label','sel'))}.xlsx",
            key=f"dl_search_{_slugify(st.session_state.get('sel_label','sel'))}",
            label="Descargar resultados (Excel)"
        )
        st.dataframe(res_out, use_container_width=True, hide_index=True)












