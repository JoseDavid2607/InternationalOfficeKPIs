import streamlit as st
import pandas as pd
import os
import plotly.express as px
from datetime import datetime

# ------------------------------------------------------
# CONFIGURACIÓN
# ------------------------------------------------------

st.set_page_config(page_title="Student Mobility Indicators", layout="wide")

PRIMARY_COLOR = "#003366"
SECONDARY_COLOR = "#0055A4"

st.markdown(
    f"""
    <style>
    .main-title {{
        font-size: 28px;
        font-weight: 700;
        color: {PRIMARY_COLOR};
    }}
    .section-title {{
        font-size: 22px;
        font-weight: 600;
        color: {SECONDARY_COLOR};
        margin-top: 35px;
    }}
    </style>
    """,
    unsafe_allow_html=True
)

st.markdown('<div class="main-title">Student Mobility Indicators</div>', unsafe_allow_html=True)

# ------------------------------------------------------
# CARGA DATOS
# ------------------------------------------------------

@st.cache_data
def load_data():
    local_path = "data/Internationalization/BD_Movilidad.xlsx"
    if os.path.exists(local_path):
        outgoing = pd.read_excel(local_path, sheet_name="Outgoing")
        incoming = pd.read_excel(local_path, sheet_name="Incoming")
    else:
        st.error("BD_Movilidad.xlsx not found.")
        st.stop()
    return outgoing, incoming

outgoing_df, incoming_df = load_data()

# ------------------------------------------------------
# ESTANDARIZACIÓN TIPO MOVILIDAD
# ------------------------------------------------------

def normalize_mobility_type(df):
    df = df.copy()
    df["Tipo de Movilidad"] = (
        df["Tipo de Movilidad"]
        .astype(str)
        .str.strip()
        .str.lower()
    )

    df["Tipo de Movilidad"] = df["Tipo de Movilidad"].replace({
        "intercambio": "Intercambio Internacional",
        "intercambio internacional": "Intercambio Internacional",
        "pasantia": "Pasantía de Investigación",
        "pasantía": "Pasantía de Investigación",
        "pasantia de investigacion": "Pasantía de Investigación",
        "pasantía de investigación": "Pasantía de Investigación"
    })

    df["Tipo de Movilidad"] = df["Tipo de Movilidad"].str.title()

    return df

outgoing_df = normalize_mobility_type(outgoing_df)
incoming_df = normalize_mobility_type(incoming_df)

# ------------------------------------------------------
# SECCIÓN UTILIZACIÓN DE CONVENIOS
# ------------------------------------------------------

st.markdown('<div class="section-title">Faculty Agreement Utilization</div>', unsafe_allow_html=True)

current_year = datetime.now().year
last_5_years = list(range(current_year - 4, current_year + 1))

mobility_types = sorted(
    set(outgoing_df["Tipo de Movilidad"].unique())
    .union(set(incoming_df["Tipo de Movilidad"].unique()))
)

def create_utilization_section(mobility_type, level_out_col, level_in_col, level_name):

    st.markdown(f"### {mobility_type} — {level_name}")

    # Filtrar por tipo y nivel
    out_filtered = outgoing_df[
        (outgoing_df["Tipo de Movilidad"] == mobility_type) &
        (outgoing_df[level_out_col].str.lower().str.contains(level_name.lower()))
    ]

    in_filtered = incoming_df[
        (incoming_df["Tipo de Movilidad"] == mobility_type) &
        (incoming_df[level_in_col].str.lower().str.contains(level_name.lower()))
    ]

    # ----------------------
    # TOP 10 ÚLTIMOS 5 AÑOS
    # ----------------------

    out_top = (
        out_filtered[out_filtered["Año de Movilidad"].isin(last_5_years)]
        .groupby("Universidad KPIs")
        .size()
        .reset_index(name="Outgoing")
    )

    in_top = (
        in_filtered[in_filtered["Año"].isin(last_5_years)]
        .groupby("Universidad KPIs")
        .size()
        .reset_index(name="Incoming")
    )

    top_merge = pd.merge(out_top, in_top, on="Universidad KPIs", how="outer").fillna(0)
    top_merge["Total"] = top_merge["Outgoing"] + top_merge["Incoming"]

    top_10 = top_merge.sort_values("Total", ascending=False).head(10)

    if not top_10.empty:
        fig = px.bar(
            top_10.sort_values("Total"),
            x="Total",
            y="Universidad KPIs",
            orientation="h",
            title="Top 10 Universities (Last 5 Years)",
            color_discrete_sequence=[PRIMARY_COLOR]
        )
        st.plotly_chart(fig, use_container_width=True)

    # ----------------------
    # MATRIZ POR AÑO
    # ----------------------

    out_group = (
        out_filtered
        .groupby(["Universidad KPIs", "País", "Año de Movilidad"])
        .size()
        .reset_index(name="Outgoing")
        .rename(columns={"Año de Movilidad": "Año"})
    )

    in_group = (
        in_filtered
        .groupby(["Universidad KPIs", "País", "Año"])
        .size()
        .reset_index(name="Incoming")
    )

    merged = pd.merge(
        out_group,
        in_group,
        on=["Universidad KPIs", "País", "Año"],
        how="outer"
    ).fillna(0)

    merged["Outgoing"] = merged["Outgoing"].astype(int)
    merged["Incoming"] = merged["Incoming"].astype(int)

    pivot = merged.pivot_table(
        index=["Universidad KPIs", "País"],
        columns="Año",
        values=["Outgoing", "Incoming"],
        fill_value=0
    )

    pivot = pivot.sort_index(axis=1, level=1)

    st.dataframe(pivot, use_container_width=True)


# ------------------------------------------------------
# GENERAR SECCIONES
# ------------------------------------------------------

for mobility in mobility_types:

    # PREGRADO
    create_utilization_section(
        mobility,
        level_out_col="Nivel",
        level_in_col="Nivel Nominación",
        level_name="Pregrado"
    )

    # POSGRADO
    create_utilization_section(
        mobility,
        level_out_col="Nivel",
        level_in_col="Nivel Nominación",
        level_name="Posgrado"
    )
