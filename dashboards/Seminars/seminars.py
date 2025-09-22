# streamlit run dashboards/Seminars/seminars_simple.py
import os
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="UASM Research Seminar Statistics — Simple", layout="wide")

# =========================
# 1) RUTAS (MISMAS DE SIEMPRE)
# =========================
def get_base() -> Path:
    """Devuelve la base usando USERPROFILE y prueba 'Estadisticas' y 'Estadísticas'."""
    up = Path(os.environ.get("USERPROFILE", ""))
    cands = [
        up / "OneDrive - Universidad de los andes" / "GESTION OFICINA INTERNACIONAL" / "Estadisticas" / "Web KPIs",
        up / "OneDrive - Universidad de los andes" / "GESTION OFICINA INTERNACIONAL" / "Estadísticas" / "Web KPIs",
    ]
    for p in cands:
        if p.exists():
            return p
    return cands[0]  # si no existe, igual lo mostramos en el error

BASE = get_base()
RUTA_XLSX = BASE / "data" / "Seminars" / "seminarios.xlsx"

# =========================
# 2) PALETA (NO SE TOCA)
# =========================
PALETA = [
    "#056D62", "#1CDFCB", "#FF7F50", "#9B59B6", "#F4A261",
    "#1B6CA8", "#0EAD69", "#E76F51", "#2C3E50", "#E9C46A"
]

# =========================
# 3) CARGA DE DATOS
# =========================
@st.cache_data(show_spinner=False)
def cargar_excel(path: Path) -> pd.DataFrame:
    if not path.exists():
        # Mensaje claro con rutas alternativas que solían aparecer
        alt1 = BASE / "Web KPIs" / "data" / "Seminars" / "seminarios.xlsx"
        alt2 = BASE / "web" / "KPIs" / "data" / "Seminars" / "seminarios.xlsx"
        st.error(
            "No se encontró el archivo Excel en:\n\n"
            f"- {path}\n- {alt1}\n- {alt2}\n\n"
            "Verifica el nombre exacto de carpetas y el archivo."
        )
        st.stop()
    try:
        return pd.read_excel(path)
    except Exception as e:
        st.error(f"No se pudo leer el Excel: {path}\n\nDetalle: {e}")
        st.stop()

df = cargar_excel(RUTA_XLSX)

# =========================
# 4) CONTROLES SÚPER SIMPLES
# =========================
st.markdown("### UASM Research Seminar Statistics")

# Elegir columnas clave sin adivinar tu esquema:
cols = list(df.columns)
if not cols:
    st.stop()

c1, c2, c3, c4 = st.columns([2, 2, 2, 2])

# Columna de periodo (opcional). Si tu archivo no tiene, deja "Ninguna".
period_col = c1.selectbox("Columna de periodo (opcional)", ["Ninguna"] + cols, index=0)
# Columna categórica para la dona (ej. Área, Serie, Departamento, etc.)
cat_col = c2.selectbox("Columna categórica", cols, index=0)
# Columna identificadora para conteo único (si no tienes, usa cualquier columna; hará count de filas)
id_col = c3.selectbox("Columna de ID (para conteo único)", ["Ninguna"] + cols, index=0)

# Si hay periodo, permite filtrar un valor
if period_col != "Ninguna":
    period_vals = sorted(df[period_col].dropna().astype(str).unique())
    periodo_sel = c4.selectbox("Valor de periodo", period_vals, index=len(period_vals) - 1)
    df_view = df[df[period_col].astype(str) == str(periodo_sel)].copy()
else:
    periodo_sel = None
    df_view = df.copy()

# =========================
# 5) DONA (Distribución por categoría)
# =========================
# Conteo único por ID si se especificó, si no, conteo de filas
if id_col != "Ninguna":
    dist = df_view.groupby(cat_col)[id_col].nunique().reset_index(name="Cuenta")
else:
    dist = df_view.groupby(cat_col).size().reset_index(name="Cuenta")

st.markdown(
    f"#### Distribución por **{cat_col}**"
    + (f" — periodo **{periodo_sel}**" if periodo_sel else "")
)

fig_donut = px.pie(
    dist, names=cat_col, values="Cuenta",
    hole=0.55, height=420, color_discrete_sequence=PALETA
)
fig_donut.update_layout(margin=dict(l=10, r=10, t=10, b=10), legend_title_text="")
st.plotly_chart(fig_donut, use_container_width=True)

# =========================
# 6) LÍNEA TEMPORAL (si hay periodo)
# =========================
if period_col != "Ninguna":
    # Serie temporal: conteo por periodo y categoría
    if id_col != "Ninguna":
        ts = df.groupby([period_col, cat_col])[id_col].nunique().reset_index(name="Cuenta")
    else:
        ts = df.groupby([period_col, cat_col]).size().reset_index(name="Cuenta")

    # Orden por el orden de aparición de los periodos (convertidos a str)
    orden = sorted(df[period_col].dropna().astype(str).unique())
    ts[period_col] = pd.Categorical(ts[period_col].astype(str), ordered=True, categories=orden)
    ts = ts.sort_values(period_col)

    st.markdown(f"#### Evolución temporal por **{cat_col}**")
    fig_line = px.line(
        ts, x=period_col, y="Cuenta", color=cat_col, markers=True,
        color_discrete_sequence=PALETA, height=460
        )
    fig_line.update_layout(margin=dict(l=10, r=10, t=10, b=10), legend_title_text="")
    st.plotly_chart(fig_line, use_container_width=True)

# =========================
# 7) TABLA BASE (vista filtrada)
# =========================
st.markdown("#### Tabla (vista actual)")
st.dataframe(df_view.reset_index(drop=True), use_container_width=True)