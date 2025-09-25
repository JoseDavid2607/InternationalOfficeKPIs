# ======================= Faculty Qualifications (full app) =======================
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import re
import math
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

df_fd = load_faculty_distribution()
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
    if obj == "%P":
        return ("%P", 60.0, 75.0)
    if obj == "%SA":
        return ("%SA", 40.0, 40.0)
    # %OTHER
    return ("%OTHER", 10.0, 10.0)


# ====== NUEVOS helpers para "Overall" por área y el impacto en celdas ======
def _needed_for_overall_if_only_this_area_changes(obj: str,
                                                  totals: dict[str, float],
                                                  area_vals: dict[str, float],
                                                  target_overall: float,
                                                  credits_each: float = 3.0) -> int | None:
    """
    ¿Cuántos profesores (n, de 3cr) hay que ajustar en ESTA ÁREA para que el OVERALL alcance el target?
    Si no es posible (p.ej. no hay suficiente OTHER para reducir), regresa None.
    """
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
        # (Ptot + c*n)/(den + c*n) >= t  -> n >= (t*den - Ptot) / ((1-t)*c)
        rhs = (t*den - Ptot) / (credits_each*(1 - t))
        return max(0, math.ceil(rhs))

    if obj == "%SA":
        if TQ <= eps: return 0
        # (SA + c*n)/(TQ + c*n) >= t -> n >= (t*TQ - SA) / ((1-t)*c)
        rhs = (t*TQ - SA) / (credits_each*(1 - t))
        return max(0, math.ceil(rhs))

    # obj == "%OTHER" (semántica: disminuir OTHER)
    if TQ <= eps: return 0
    # objetivo overall: (OT - c*n)/(TQ - c*n) <= 0.10 -> c*n >= (OT - 0.10*TQ)/0.90
    need_credits = (OT - 0.10*TQ) / 0.90
    need_n = 0 if need_credits <= 0 else math.ceil(need_credits / credits_each)

    # factibilidad: verificar que en ESTA área hay suficiente OTHER para quitar
    OT_a = area_vals.get("OTHER", 0.0)
    max_remove_n = math.floor(OT_a / credits_each)  # máximo que puedo quitar en esta área
    if need_n <= max_remove_n:
        return max(0, need_n)
    return None


def _impact_pp_area(obj: str, area_vals: dict[str,float], credits_each: float = 3.0) -> tuple[float,float]:
    """Impacto en p.p. del % del ÁREA al mover ±3cr en el objetivo (P, SA u OTHER)."""
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

    # %OTHER
    if denQ <= eps: return (0.0, 0.0)
    up   = ((OT + credits_each) / (denQ + credits_each) - (OT / denQ)) * 100.0
    down = ((max(0.0, OT - credits_each)) / max(eps, denQ - credits_each) - (OT / denQ)) * 100.0 if denQ > credits_each else 0.0
    return (round(up,2), round(down,2))


def _impact_pp_overall_if_area_changes(obj: str, totals: dict[str,float], credits_each: float = 3.0) -> tuple[float,float]:
    """
    Impacto en p.p. del OVERALL si se mueve ±3cr en una sola área (el efecto en el total
    es el mismo por +3cr en cualquier área para ese objetivo).
    """
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
    sens_mode = st.toggle("Enable sensitivity mode", value=st.session_state.get("sens_mode", False), key="sens_mode")
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
        st.markdown("### Go to KPI")
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
df_fd_base = df_fd.copy()
df_car_filt_all = filter_df_car(df_car_base, time_mode, sel_year, sel_sem)
df_fd_f = filter_df_fd(df_fd_base, time_mode, sel_year, sel_sem)

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


# ================== PRINCIPAL ==================
st.markdown("---")
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
        # -------------- BY ACADEMIC AREA --------------
        if view_mode == "By Academic Area":
            colT, colG = st.columns([6,6], gap="large")

            # Agregaciones del frame seleccionado
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
                needed_mode = False
                # Controles compactos de tabla
                r1c1, r1c2, r1c3, r1c4 = st.columns([1.6, 1.1, 1.2, 1.4])
                with r1c1:
                    needed_mode = st.toggle("Show necessary # of Faculty for…", value=False, key="area_needed_mode")
                with r1c2:
                    objective = st.selectbox("Objective", ["%P", "%SA", "%OTHER"], key="area_objective")
                with r1c3:
                    scope_label = st.radio("Target scope", ["By area", "Overall"], horizontal=True, key="area_scope")
                with r1c4:
                    show_impact = st.checkbox("Show impact (±3cr)", key="area_show_impact", value=False)

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
                    # ===== Needed (una sola columna) + Impacto opcional =====
                    idx = mod_agg_ps.index
                    p = mod_agg_ps["P"].reindex(idx, fill_value=0.0)
                    s = mod_agg_ps["S"].reindex(idx, fill_value=0.0)
                    sa = mod_agg_tipo["SA"].reindex(idx, fill_value=0.0)
                    pa = mod_agg_tipo["PA"].reindex(idx, fill_value=0.0)
                    sp = mod_agg_tipo["SP"].reindex(idx, fill_value=0.0)
                    ip = mod_agg_tipo["IP"].reindex(idx, fill_value=0.0)
                    other = mod_agg_tipo["OTHER"].reindex(idx, fill_value=0.0)

                    obj_lbl, tgt_area, tgt_overall = _objective_targets(objective)
                    totals = {
                        "P": float(p.sum()), "S": float(s.sum()),
                        "SA": float(sa.sum()), "PA": float(pa.sum()),
                        "SP": float(sp.sum()), "IP": float(ip.sum()),
                        "OTHER": float(other.sum())
                    }

                    rows = []
                    colname = {"%P":"Needed P (3cr)", "%SA":"Needed SA (3cr)", "%OTHER":"Needed OTHER less (3cr)"}[objective]

                    for label in list(idx) + ["TOTAL"]:
                        if label == "TOTAL":
                            P, S = totals["P"], totals["S"]
                            SA_, PA_, SP_, IP_, OT_ = totals["SA"], totals["PA"], totals["SP"], totals["IP"], totals["OTHER"]
                        else:
                            P = float(p.get(label, 0.0)); S = float(s.get(label, 0.0))
                            SA_ = float(sa.get(label, 0.0)); PA_ = float(pa.get(label, 0.0))
                            SP_ = float(sp.get(label, 0.0)); IP_ = float(ip.get(label, 0.0))
                            OT_ = float(other.get(label, 0.0))

                        area_vals = {"P":P,"S":S,"SA":SA_,"PA":PA_,"SP":SP_,"IP":IP_,"OTHER":OT_}

                        # Needed por scope
                        if scope_label == "By area" or label == "TOTAL":
                            if objective == "%P":
                                need_val = _needed_for_pctP(P, S, tgt_area, credits_each=3.0)
                            elif objective == "%SA":
                                need_val = _needed_for_pctSA(SA_, PA_+SP_+IP_+OT_, tgt_area, credits_each=3.0)
                            else:
                                # OTHER menos para el área (≤10%)
                                TQ_area = SA_+PA_+SP_+IP_+OT_
                                need_credits = (OT_ - 0.10*TQ_area) / 0.90
                                need_val = 0 if need_credits <= 0 else math.ceil(need_credits/3.0)
                        else:
                            # scope Overall: cuántos en ESTA ÁREA para que el OVERALL llegue al target
                            need_n = _needed_for_overall_if_only_this_area_changes(objective, totals, area_vals, tgt_overall, credits_each=3.0)
                            need_val = (need_n if need_n is not None else None)

                        row = {"Academic Area": label, colname: ("" if need_val is None else int(need_val))}

                        # Impacto por celda
                        if show_impact:
                            if scope_label == "By area":
                                up_pp, down_pp = _impact_pp_area(objective, area_vals, credits_each=3.0)
                                row["Impact (±3cr)"] = f"{up_pp:+.2f} / {down_pp:+.2f} p.p. (area)"
                            else:
                                up_pp, down_pp = _impact_pp_overall_if_area_changes(objective, totals, credits_each=3.0)
                                row["Impact (±3cr)"] = f"{up_pp:+.2f} / {down_pp:+.2f} p.p. (overall)"

                        rows.append(row)

                    need_tbl = pd.DataFrame(rows)
                    _download_xlsx_button(need_tbl, f"needed_ByArea_{_slugify(sel_label)}_{_slugify(obj_lbl)}_{_slugify(scope_label)}.xlsx",
                                          key=f"dl_need_area_{_slugify(sel_label)}_{_slugify(obj_lbl)}_{_slugify(scope_label)}",
                                          label="⬇️ Descargar (Excel)")
                    st.dataframe(need_tbl, use_container_width=True, hide_index=True)

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
                r1c1, r1c2, r1c3, r1c4 = st.columns([1.6, 1.1, 1.2, 1.4])
                with r1c1:
                    needed_mode_f = st.toggle("Show necessary # of Faculty for…", value=False, key="field_needed_mode")
                with r1c2:
                    objective_f = st.selectbox("Objective", ["%P", "%SA", "%OTHER"], key="field_objective")
                with r1c3:
                    scope_label_f = st.radio("Target scope", ["By area", "Overall"], horizontal=True, key="field_scope")
                with r1c4:
                    show_impact_f = st.checkbox("Show impact (±3cr)", key="field_show_impact", value=False)

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
                    idx = mod_agg_ps.index
                    p = mod_agg_ps["P"].reindex(idx, fill_value=0.0)
                    s = mod_agg_ps["S"].reindex(idx, fill_value=0.0)
                    sa = mod_agg_tipo["SA"].reindex(idx, fill_value=0.0)
                    pa = mod_agg_tipo["PA"].reindex(idx, fill_value=0.0)
                    sp = mod_agg_tipo["SP"].reindex(idx, fill_value=0.0)
                    ip = mod_agg_tipo["IP"].reindex(idx, fill_value=0.0)
                    other = mod_agg_tipo["OTHER"].reindex(idx, fill_value=0.0)

                    obj_lbl, tgt_area, tgt_overall = _objective_targets(objective_f)
                    totals = {
                        "P": float(p.sum()), "S": float(s.sum()),
                        "SA": float(sa.sum()), "PA": float(pa.sum()),
                        "SP": float(sp.sum()), "IP": float(ip.sum()),
                        "OTHER": float(other.sum())
                    }

                    rows = []
                    colname = {"%P":"Needed P (3cr)", "%SA":"Needed SA (3cr)", "%OTHER":"Needed OTHER less (3cr)"}[objective_f]

                    for label in list(idx) + ["TOTAL"]:
                        if label == "TOTAL":
                            P, S = totals["P"], totals["S"]
                            SA_, PA_, SP_, IP_, OT_ = totals["SA"], totals["PA"], totals["SP"], totals["IP"], totals["OTHER"]
                        else:
                            P = float(p.get(label, 0.0)); S = float(s.get(label, 0.0))
                            SA_ = float(sa.get(label, 0.0)); PA_ = float(pa.get(label, 0.0))
                            SP_ = float(sp.get(label, 0.0)); IP_ = float(ip.get(label, 0.0))
                            OT_ = float(other.get(label, 0.0))

                        area_vals = {"P":P,"S":S,"SA":SA_,"PA":PA_,"SP":SP_,"IP":IP_,"OTHER":OT_}

                        if scope_label_f == "By area" or label == "TOTAL":
                            if objective_f == "%P":
                                need_val = _needed_for_pctP(P, S, tgt_area, credits_each=3.0)
                            elif objective_f == "%SA":
                                need_val = _needed_for_pctSA(SA_, PA_+SP_+IP_+OT_, tgt_area, credits_each=3.0)
                            else:
                                TQ_area = SA_+PA_+SP_+IP_+OT_
                                need_credits = (OT_ - 0.10*TQ_area) / 0.90
                                need_val = 0 if need_credits <= 0 else math.ceil(need_credits/3.0)
                        else:
                            need_n = _needed_for_overall_if_only_this_area_changes(objective_f, totals, area_vals, tgt_overall, credits_each=3.0)
                            need_val = (need_n if need_n is not None else None)

                        row = {"Field": label, colname: ("" if need_val is None else int(need_val))}

                        if show_impact_f:
                            if scope_label_f == "By area":
                                up_pp, down_pp = _impact_pp_area(objective_f, area_vals, credits_each=3.0)
                                row["Impact (±3cr)"] = f"{up_pp:+.2f} / {down_pp:+.2f} p.p. (area)"
                            else:
                                up_pp, down_pp = _impact_pp_overall_if_area_changes(objective_f, totals, credits_each=3.0)
                                row["Impact (±3cr)"] = f"{up_pp:+.2f} / {down_pp:+.2f} p.p. (overall)"

                        rows.append(row)

                    need_tbl_f = pd.DataFrame(rows)
                    _download_xlsx_button(need_tbl_f, f"needed_ByField_{_slugify(sel_label)}_{_slugify(obj_lbl)}_{_slugify(scope_label_f)}.xlsx",
                                          key=f"dl_need_field_{_slugify(sel_label)}_{_slugify(obj_lbl)}_{_slugify(scope_label_f)}",
                                          label="⬇️ Descargar (Excel)")
                    st.dataframe(need_tbl_f, use_container_width=True, hide_index=True)

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
            colM_L, colM_R = st.columns([6,6], gap="large")
            fil_mat = fil.copy()
            fil_mat["_MAT"] = fil_mat[col_prog].astype(str).str.strip()

            agg_tipo_m = (fil_mat.groupby(["_MAT","_TIPO"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["SA","PA","SP","IP","OTHER"]:
                if k not in agg_tipo_m.columns: agg_tipo_m[k] = 0.0
            agg_tipo_m = agg_tipo_m[["SA","PA","SP","IP","OTHER"]]

            agg_ps_m = (fil_mat.groupby(["_MAT","_PS"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["P","S"]:
                if k not in agg_ps_m.columns: agg_ps_m[k] = 0.0
            agg_ps_m = agg_ps_m[["P","S"]]

            base_agg_ps = agg_ps_m.copy()
            base_agg_tipo = agg_tipo_m.copy()
            if SENS["on"] and SENS["ops"]:
                mod_agg_ps, mod_agg_tipo = apply_ops_to_aggs(base_agg_ps, base_agg_tipo, SENS["ops"])
            else:
                mod_agg_ps, mod_agg_tipo = base_agg_ps, base_agg_tipo

            with colM_L:
                needed_mode_m = False
                r1, r2, r3, r4 = st.columns([1.6, 1.1, 1.2, 1.6])
                with r1:
                    needed_mode_m = st.toggle("Show necessary # of Faculty for…", value=False, key="prog_needed_mode")
                with r2:
                    objective_m = st.selectbox("Objective", ["%P", "%SA", "%OTHER"], key="prog_objective")
                with r3:
                    scope_label_m = st.radio("Target scope", ["By area", "Overall"], horizontal=True, key="prog_scope")
                with r4:
                    show_impact_m = st.checkbox("Show impact (±3cr)", key="prog_show_impact", value=False)

                if not needed_mode_m:
                    metrics_tbl_m = build_percent_table("Program", mod_agg_tipo, mod_agg_ps)
                    _download_xlsx_button(metrics_tbl_m, f"table_ByProgram_{_slugify(sel_label)}.xlsx",
                                          key=f"dl_tbl_prog_{_slugify(sel_label)}", label="⬇️ Download table (Excel)")
                    styled_tbl_m = (
                        metrics_tbl_m.style
                        .format({"%P":"{:.1f}%","%S":"{:.1f}%","%SA":"{:.1f}%","%OTHER":"{:.1f}%"})
                        .apply(style_percent_tables, id_col="Program", axis=None)
                        .hide(axis="index")
                    )
                    st.markdown(f"<div class='scroll-wrap-program'>{styled_tbl_m.to_html(escape=False)}</div>", unsafe_allow_html=True)
                else:
                    idx = mod_agg_ps.index
                    p = mod_agg_ps["P"].reindex(idx, fill_value=0.0)
                    s = mod_agg_ps["S"].reindex(idx, fill_value=0.0)
                    sa = mod_agg_tipo["SA"].reindex(idx, fill_value=0.0)
                    pa = mod_agg_tipo["PA"].reindex(idx, fill_value=0.0)
                    sp = mod_agg_tipo["SP"].reindex(idx, fill_value=0.0)
                    ip = mod_agg_tipo["IP"].reindex(idx, fill_value=0.0)
                    other = mod_agg_tipo["OTHER"].reindex(idx, fill_value=0.0)

                    obj_lbl, tgt_area, tgt_overall = _objective_targets(objective_m)
                    totals = {
                        "P": float(p.sum()), "S": float(s.sum()),
                        "SA": float(sa.sum()), "PA": float(pa.sum()),
                        "SP": float(sp.sum()), "IP": float(ip.sum()),
                        "OTHER": float(other.sum())
                    }

                    rows = []
                    colname = {"%P":"Needed P (3cr)", "%SA":"Needed SA (3cr)", "%OTHER":"Needed OTHER less (3cr)"}[objective_m]

                    for label in list(idx) + ["TOTAL"]:
                        if label == "TOTAL":
                            P, S = totals["P"], totals["S"]
                            SA_, PA_, SP_, IP_, OT_ = totals["SA"], totals["PA"], totals["SP"], totals["IP"], totals["OTHER"]
                        else:
                            P = float(p.get(label, 0.0)); S = float(s.get(label, 0.0))
                            SA_ = float(sa.get(label, 0.0)); PA_ = float(pa.get(label, 0.0))
                            SP_ = float(sp.get(label, 0.0)); IP_ = float(ip.get(label, 0.0))
                            OT_ = float(other.get(label, 0.0))

                        area_vals = {"P":P,"S":S,"SA":SA_,"PA":PA_,"SP":SP_,"IP":IP_,"OTHER":OT_}

                        if scope_label_m == "By area" or label == "TOTAL":
                            if objective_m == "%P":
                                need_val = _needed_for_pctP(P, S, tgt_area, credits_each=3.0)
                            elif objective_m == "%SA":
                                need_val = _needed_for_pctSA(SA_, PA_+SP_+IP_+OT_, tgt_area, credits_each=3.0)
                            else:
                                TQ_area = SA_+PA_+SP_+IP_+OT_
                                need_credits = (OT_ - 0.10*TQ_area) / 0.90
                                need_val = 0 if need_credits <= 0 else math.ceil(need_credits/3.0)
                        else:
                            need_n = _needed_for_overall_if_only_this_area_changes(objective_m, totals, area_vals, tgt_overall, credits_each=3.0)
                            need_val = (need_n if need_n is not None else None)

                        row = {"Program": label, colname: ("" if need_val is None else int(need_val))}

                        if show_impact_m:
                            if scope_label_m == "By area":
                                up_pp, down_pp = _impact_pp_area(objective_m, area_vals, credits_each=3.0)
                                row["Impact (±3cr)"] = f"{up_pp:+.2f} / {down_pp:+.2f} p.p. (area)"
                            else:
                                up_pp, down_pp = _impact_pp_overall_if_area_changes(objective_m, totals, credits_each=3.0)
                                row["Impact (±3cr)"] = f"{up_pp:+.2f} / {down_pp:+.2f} p.p. (overall)"

                        rows.append(row)

                    need_tbl_m = pd.DataFrame(rows)
                    _download_xlsx_button(need_tbl_m, f"needed_ByProgram_{_slugify(sel_label)}_{_slugify(obj_lbl)}_{_slugify(scope_label_m)}.xlsx",
                                          key=f"dl_need_prog_{_slugify(sel_label)}_{_slugify(obj_lbl)}_{_slugify(scope_label_m)}",
                                          label="⬇️ Descargar (Excel)")
                    st.dataframe(need_tbl_m, use_container_width=True, hide_index=True)

            # Históricos Program
            df_hist_m = df_car_global.copy()
            df_hist_m["_MAT"] = df_hist_m[col_prog].astype(str).str.strip()

            agg_ps_all_m = (df_hist_m.groupby(["_SEM","_MAT","_PS"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["P","S"]:
                if k not in agg_ps_all_m.columns: agg_ps_all_m[k] = 0.0
            agg_ps_all_m["P_share"] = (agg_ps_all_m["P"] / (agg_ps_all_m["P"] + agg_ps_all_m["S"]).replace(0, pd.NA)) * 100
            agg_ps_all_m = agg_ps_all_m.reset_index()

            agg_tipo_all_m = (df_hist_m.groupby(["_SEM","_MAT","_TIPO"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["SA","PA","SP","IP","OTHER"]:
                if k not in agg_tipo_all_m.columns: agg_tipo_all_m[k] = 0.0
            den_all_m = (agg_tipo_all_m[["SA","PA","SP","IP","OTHER"]].sum(axis=1)).replace(0, pd.NA)
            agg_tipo_all_m["SA_share"] = (agg_tipo_all_m["SA"] / den_all_m) * 100
            agg_tipo_all_m["OTHER_share"] = (agg_tipo_all_m["OTHER"] / den_all_m) * 100
            agg_tipo_all_m = agg_tipo_all_m.reset_index()

            tot_by_sem_m = (df_hist_m.groupby(["_SEM","_PS"])["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["P","S"]:
                if k not in tot_by_sem_m.columns: tot_by_sem_m[k] = 0.0
            tot_by_sem_m["P_share"] = (tot_by_sem_m["P"] / (tot_by_sem_m["P"] + tot_by_sem_m["S"]).replace(0, pd.NA)) * 100
            tot_by_sem_m = tot_by_sem_m.reset_index()

            tot_by_sem_tipo_m = (df_hist_m.groupby(["_SEM","_TIPO"])["_CRED"].sum().unstack(fill_value=0.0))
            for k in ["SA","PA","SP","IP","OTHER"]:
                if k not in tot_by_sem_tipo_m.columns: tot_by_sem_tipo_m[k] = 0.0
            den_m = (tot_by_sem_tipo_m[["SA","PA","SP","IP","OTHER"]].sum(axis=1)).replace(0, pd.NA)
            tot_by_sem_tipo_m["SA_share"] = (tot_by_sem_tipo_m["SA"] / den_m) * 100
            tot_by_sem_tipo_m["OTHER_share"] = (tot_by_sem_tipo_m["OTHER"] / den_m) * 100
            tot_by_sem_tipo_m = tot_by_sem_tipo_m.reset_index()

            agg_ps_all_tm  = transform_for_time_mode_ps(agg_ps_all_m.rename(columns={"_MAT":"__LEVEL__"})).rename(columns={"__LEVEL__":"_MAT"})
            agg_tipo_sa_tm = transform_for_time_mode_tipo(agg_tipo_all_m.rename(columns={"_MAT":"__LEVEL__"}), "SA_share").rename(columns={"__LEVEL__":"_MAT"})
            agg_tipo_ot_tm = transform_for_time_mode_tipo(agg_tipo_all_m.rename(columns={"_MAT":"__LEVEL__"}), "OTHER_share").rename(columns={"__LEVEL__":"_MAT"})
            agg_tipo_all_tm = (
                agg_tipo_sa_tm.drop(columns=[c for c in ["OTHER_share"] if c in agg_tipo_sa_tm], errors="ignore")
                .merge(
                    agg_tipo_ot_tm[["_SEM","_MAT","OTHER","SA","PA","SP","IP","OTHER_share"]],
                    on=["_SEM","_MAT","SA","PA","SP","IP","OTHER"], how="outer"
                )
            )
            tot_by_sem_P_tm = transform_for_time_mode_ps(tot_by_sem_m.copy())
            tot_tipo_sa_tm  = transform_for_time_mode_tipo(tot_by_sem_tipo_m.copy(), "SA_share")
            tot_tipo_ot_tm  = transform_for_time_mode_tipo(tot_by_sem_tipo_m.copy(), "OTHER_share")
            tot_by_sem_tipo_tm = (
                tot_tipo_sa_tm.drop(columns=[c for c in ["OTHER_share"] if c in tot_tipo_sa_tm], errors="ignore")
                .merge(
                    tot_tipo_ot_tm[["_SEM","SA","PA","SP","IP","OTHER","OTHER_share"]],
                    on=["_SEM","SA","PA","SP","IP","OTHER"], how="outer"
                )
            )

            key_col, x_labels, x_map = build_time_axis_for_history(df_hist_m)
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
                    level_name="_MAT",
                    sel_label_value=sel_label_exact,
                    ops=SENS["ops"],
                    member_all_label="All"
                )

            programs_all = sorted(set(agg_ps_all_tm["_MAT"].astype(str).unique()) | set(agg_tipo_all_tm["_MAT"].astype(str).unique()))
            with colM_R:
                draw_history(
                    "Evolution by Academic Program",
                    level_name="_MAT",
                    level_values=programs_all,
                    metric_kind="%P",
                    total_series_builders={"P": tot_by_sem_P_tm, "SA": tot_by_sem_tipo_tm, "OTHER": tot_by_sem_tipo_tm},
                    agg_ps_all=agg_ps_all_tm,
                    agg_tipo_all=agg_tipo_all_tm,
                    x_labels=x_labels, x_map=x_map, sel_x=sel_x
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
    if "_MAT"   not in period_df.columns and col_prog:       period_df["_MAT"] = period_df[col_prog].astype(str).str.strip()

    view = st.session_state.view_mode if "view_mode" in st.session_state else "By Academic Area"
    if view == "By Academic Area":
        dim_col, dim_label = "_AREA", "Academic Area"
    elif view == "By Field":
        dim_col, dim_label = "_FIELD", "Field"
    else:
        dim_col, dim_label = "_MAT", "Program"

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
        with st.expander(f"Credit sums by {dim_label} — {display_label}", expanded=False):
            export_tbl = tbl_out.reset_index().rename(columns={"index": dim_label})
            _download_xlsx_button(export_tbl,
                                  f"credit_sums_{_slugify(dim_label)}_{_slugify(display_label)}.xlsx",
                                  key=f"dl_credit_sums_{_slugify(dim_label)}_{_slugify(display_label)}",
                                  label="⬇️ Descargar tabla (Excel)")
            st.dataframe(tbl_out.style.format("{:,.0f}"), use_container_width=True)
except Exception:
    pass

# --------------------------
# DETAIL TABLE + DONUT + SEARCH
# (Oculto automáticamente cuando Sensitivity mode está activo)
# --------------------------
if not SENS.get("on", False):
    try:
        cfg = {
            "By Academic Area": {"key": "_AREA_filter",  "col": "_AREA", "label": "area",    "metric_key": "metric__AREA"},
            "By Field":         {"key": "_FIELD_filter", "col": "_FIELD","label": "campo",   "metric_key": "metric__FIELD"},
            "By Program":       {"key": "_MAT_filter",   "col": "_MAT",  "label": "programa","metric_key": "metric__MAT"},
        }
        view = st.session_state.view_mode

        if view in cfg:
            key = cfg[view]["key"]; col_tag = cfg[view]["col"]; label = cfg[view]["label"]
            metric_choice = st.session_state.get(cfg[view]["metric_key"], "%P")
            opt_val = st.session_state.get(key, "(All)")

            base = df_car_filt_all.copy()
            if "_AREA"  not in base.columns and col_areaCourse: base["_AREA"]  = base[col_areaCourse].astype(str).str.strip()
            if "_FIELD" not in base.columns and col_field:      base["_FIELD"] = base[col_field].astype(str).str.strip()
            if "_MAT"   not in base.columns and col_prog:       base["_MAT"]   = base[col_prog].astype(str).str.strip()
            if "_TIPO"  not in base.columns and col_tipoC:      base["_TIPO"]  = _norm_str(base[col_tipoC]).map(normalize_tipo)
            if "_PS"    not in base.columns and col_ps_C:       base["_PS"]    = _norm_str(base[col_ps_C]).map(normalize_ps)
            if "_CRED"  not in base.columns and col_cred:       base["_CRED"]  = pd.to_numeric(base[col_cred], errors="coerce").fillna(0.0)

            cL, cR = st.columns([7,5], gap="large")

            with cL:
                if metric_choice == "%P":
                    table_filter = st.radio("", ["All", "Only P", "Only S"], index=0, horizontal=True, key=f"table_filt_ps_{view}_{opt_val}")
                    base_tbl = base.copy()
                    if opt_val not in {"(All)", "(TOTAL)"} and col_tag in base_tbl.columns:
                        base_tbl = base_tbl[base_tbl[col_tag] == opt_val].copy()
                    if table_filter == "Only P": base_tbl = base_tbl[base_tbl["_PS"] == "P"]
                    elif table_filter == "Only S": base_tbl = base_tbl[base_tbl["_PS"] == "S"]
                else:
                    table_filter = st.radio("", ["All", "Only SA", "Only OTHER"], index=0, horizontal=True, key=f"table_filt_tipo_{view}_{opt_val}")
                    base_tbl = base.copy()
                    if opt_val not in {"(All)", "(TOTAL)"} and col_tag in base_tbl.columns:
                        base_tbl = base_tbl[base_tbl[col_tag] == opt_val].copy()
                    if table_filter == "Only SA": base_tbl = base_tbl[base_tbl["_TIPO"] == "SA"]
                    elif table_filter == "Only OTHER": base_tbl = base_tbl[base_tbl["_TIPO"] == "OTHER"]

                wanted_map = {
                    "Semestre": col_sem, "Código Materia": col_code, "Créditos": col_cred,
                    "Nombre largo curso": col_name, "Program": col_prog, "Profesor": col_prof,
                    "Area del curso": col_areaCourse, "Field": col_field, "TIPO": col_tipoC, "P/S": col_ps_C,
                }
                present_tbl = {k: v for k, v in wanted_map.items() if v in base_tbl.columns}
                out = base_tbl[list(present_tbl.values())].rename(columns={v: k for k, v in present_tbl.items()})

                display_label = st.session_state.get('sel_label','Selected Period')
                n_courses = len(out)
                if metric_choice == "%P":
                    if table_filter == "Only P":   title = f"{n_courses} courses taught in {display_label} by Participating Faculty"
                    elif table_filter == "Only S": title = f"{n_courses} courses taught in {display_label} by Supporting Faculty"
                    else:
                        title = (f"{n_courses} courses were taught in {display_label}"
                                 if opt_val in {"(TOTAL)", "(All)"} else f"{n_courses} courses of {opt_val} were taught in {display_label}")
                else:
                    if table_filter == "Only SA":       title = f"{n_courses} courses taught in {display_label} by Scholarly Academics"
                    elif table_filter == "Only OTHER":  title = f"{n_courses} courses taught in {display_label} by Others"
                    else:
                        title = (f"{n_courses} courses were taught in {display_label}"
                                 if opt_val in {"(TOTAL)", "(All)"} else f"{n_courses} courses of {opt_val} were taught in {display_label}")

                st.markdown(f"### {title}")
                _download_xlsx_button(out, f"table_detail_{_slugify(opt_val)}_{_slugify(display_label)}.xlsx",
                                      key=f"dl_tbl_detail_{_slugify(opt_val)}_{_slugify(display_label)}", label="⬇️ Descargar tabla (Excel)")
                st.dataframe(out, use_container_width=True, hide_index=True)

            with cR:
                st.markdown("<div style='height: 110px'></div>", unsafe_allow_html=True)

                agg_tipo = (base.groupby([col_tag,"_TIPO"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0)) if col_tag in base.columns else pd.DataFrame()
                for k in ["SA","PA","SP","IP","OTHER"]:
                    if k not in agg_tipo.columns: agg_tipo[k] = 0.0
                agg_ps = (base.groupby([col_tag,"_PS"], dropna=False)["_CRED"].sum().unstack(fill_value=0.0)) if col_tag in base.columns else pd.DataFrame()
                for k in ["P","S"]:
                    if k not in agg_ps.columns: agg_ps[k] = 0.0
                agg_ps = agg_ps[["P","S"]]; agg_tipo = agg_tipo[["SA","PA","SP","IP","OTHER"]]

                # (sin sensibilidad aquí porque está desactivada; si la activas, esta sección completa se oculta)

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

                donut_h   = 360
                thrP = 75.0 if title_suffix == "TOTAL" else 60.0

                if metric_choice == "%P":
                    den = p_val + s_val
                    p_share = (p_val/den*100) if den else 0.0
                    alert = (p_share < thrP)
                    color_map = {"P": ("#F5A3A3" if alert else MINT), "S": "#B0B0B0"}
                    fig = px.pie(names=["P","S"], values=[p_val, s_val], color=["P","S"], color_discrete_map=color_map, hole=0.55)
                    fig.update_traces(textinfo="percent+label", hovertemplate="%{label}: %{percent:.1%}<extra></extra>")
                    fig.update_layout(title=f"% Participating Distribution — {title_suffix}",
                                      height=donut_h, margin=dict(l=10, r=10, t=40, b=10),
                                      legend=dict(orientation="v", yanchor="bottom", y=0.4, xanchor="center", x=0.9))
                    st.plotly_chart(fig, use_container_width=True)
                    donut_df = pd.DataFrame({"Group": ["P","S"], "Credits": [p_val, s_val]})
                    donut_df["Percent"] = (donut_df["Credits"] / max(1e-9, donut_df["Credits"].sum()))*100
                    _download_xlsx_button(donut_df, f"chart_donut_PS_{_slugify(title_suffix)}_{_slugify(display_label)}.xlsx",
                                          key=f"dl_donut_ps_{_slugify(title_suffix)}_{_slugify(display_label)}", label="⬇️ Datos de la gráfica (Excel)")
                else:
                    labels_all  = ["SA", "PA", "SP", "IP", "OTHER"]
                    values_all  = [sa, pa, sp, ip, other]
                    filtered    = [(l, v) for l, v in zip(labels_all, values_all) if v > 0]
                    if filtered:
                        labels = [l for l, _ in filtered]; values = [v for _, v in filtered]
                        den = sum(values_all) or 1.0
                        sa_share    = sa/den*100
                        other_share = other/den*100
                        color_map = {}
                        for l in labels:
                            if l == "SA": color_map[l] = ("#F5A3A3" if sa_share < 40.0 else MINT)
                            elif l == "OTHER": color_map[l] = ("#F5A3A3" if other_share > 10.0 else "#6B7280")
                            else: color_map[l] = "#B0B0B0"

                        fig = px.pie(names=labels, values=values, color=labels, color_discrete_map=color_map, hole=0.55)
                        fig.update_traces(textinfo="percent+label", sort=False, hovertemplate="%{label}: %{percent:.1%}<extra></extra>")
                        title_txt = "%SA Distribution" if metric_choice == "%SA" else "%OTHER Distribution"
                        fig.update_layout(title=f"{title_txt} — {title_suffix}",
                                          height=donut_h, margin=dict(l=10, r=10, t=40, b=10),
                                          legend=dict(orientation="v", yanchor="bottom", y=0.4, xanchor="center", x=0.9))
                        st.plotly_chart(fig, use_container_width=True)
                        donut_df = pd.DataFrame({"Type": labels_all, "Credits": values_all})
                        donut_df["Percent"] = (donut_df["Credits"] / max(1e-9, donut_df["Credits"].sum()))*100
                        _download_xlsx_button(donut_df, f"chart_donut_TIPO_{_slugify(title_suffix)}_{_slugify(display_label)}.xlsx",
                                              key=f"dl_donut_tipo_{_slugify(title_suffix)}_{_slugify(display_label)}", label="⬇️ Datos de la gráfica (Excel)")
                    else:
                        st.caption("No hay registros de TIPO para esta métrica en este período.")
    except Exception:
        pass

# --------------------------
# COUNTS — PIVOT (opcional)
# --------------------------
st.markdown("---")
show_counts = st.checkbox("Show P/S counts", value=False)
if show_counts:
    st.subheader(f"Participating vs Supporting — {st.session_state.get('sel_label','Selected')} (Counts & %)")
    df_fd_f = df_fd_f.copy()
    if col_ps_fd:   df_fd_f["_PS"]   = _norm_str(df_fd_f[col_ps_fd]).map(normalize_ps)
    if col_area_fd: df_fd_f["_AREA"] = df_fd_f[col_area_fd].astype(str).str.strip()
    if col_tipo_fd: df_fd_f["_TIPO"] = _norm_str(df_fd_f[col_tipo_fd]).map(normalize_tipo)

    pivot_rows = st.radio("Pivot by", ["AREA", "Qualification Type"], index=0, horizontal=True)
    if pivot_rows == "AREA":
        row_name = "AREA"; row_series = df_fd_f["_AREA"].astype(str).str.strip().replace({"": "N/A"})
        desired_order = None
    else:
        row_name = "Type"; row_series = df_fd_f["_TIPO"].map(lambda v: str(v).upper())
        desired_order = ["SA", "PA", "SP", "IP", "OTHER"]

    base = pd.DataFrame({row_name: row_series, "_PS": df_fd_f["_PS"]})
    table = (base.groupby([row_name, "_PS"], dropna=False).size().unstack(fill_value=0)
                  .rename(columns={"P": "Participating", "S": "Supporting"}))
    for k in ["Participating", "Supporting"]:
        if k not in table.columns: table[k] = 0

    # Ajuste simple por sensibilidad (solo cuenta de profesores agregados/quitados a nivel total)
    if SENS["on"] and SENS["ops"]:
        add_P = sum(op.get("count",0) for op in SENS["ops"] if op.get("scope")=="PS" and op.get("cat")=="P")
        add_S = sum(op.get("count",0) for op in SENS["ops"] if op.get("scope")=="PS" and op.get("cat")=="S")
        table_totals_increase = {"Participating": int(add_P), "Supporting": int(add_S)}
    else:
        table_totals_increase = {"Participating": 0, "Supporting": 0}

    table["__Total__"] = table["Participating"] + table["Supporting"]

    df_counts = table[["Participating", "Supporting"]].astype(int).reset_index()
    total_row = pd.DataFrame([{row_name: "TOTAL",
                               "Participating": int(df_counts["Participating"].sum()) + table_totals_increase["Participating"],
                               "Supporting":    int(df_counts["Supporting"].sum())    + table_totals_increase["Supporting"]}])
    df_counts_out = pd.concat([df_counts, total_row], ignore_index=True)

    def _bold_total(df_):
        sty = pd.DataFrame('', index=df_.index, columns=df_.columns)
        mask = df_[row_name].astype(str).str.upper().eq("TOTAL")
        for c in df_.columns: sty.loc[mask, c] = 'font-weight:700;'
        return sty

    left, right = st.columns([6,6], gap="large")

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
        _download_xlsx_button(df_counts_out, f"ps_counts_{_slugify(row_name)}_{_slugify(st.session_state.get('sel_label','sel'))}.xlsx",
                              key=f"dl_ps_counts_{_slugify(row_name)}_{_slugify(st.session_state.get('sel_label','sel'))}",
                              label="Descargar tabla (Excel)")
        styled_counts = (df_counts_out.style
                         .format({"Participating": "{:,.0f}", "Supporting": "{:,.0f}"})
                         .apply(_bold_total, axis=None))
        st.dataframe(styled_counts, use_container_width=True, hide_index=True)

    with right:
        fig = px.bar(chart_export, x=row_name, y="Percent", color="Group",
                     barmode="group", text="Percent",
                     color_discrete_map={"%Participating": MINT, "%Supporting": SUPPORTING},
                     category_orders={row_name: cat_order})
        fig.update_traces(texttemplate="%{text:.1f}%")
        fig.update_layout(xaxis_title=None, yaxis_title=None, height=340,
                          legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
                          legend_title_text=None, margin=dict(l=20, r=10, t=10, b=40))
        st.plotly_chart(fig, use_container_width=True)
        _download_xlsx_button(chart_export, f"chart_ps_perc_{_slugify(row_name)}_{_slugify(st.session_state.get('sel_label','sel'))}.xlsx",
                              key=f"dl_chart_ps_perc_{_slugify(row_name)}_{_slugify(st.session_state.get('sel_label','sel'))}",
                              label="Descargar datos (Excel)")




