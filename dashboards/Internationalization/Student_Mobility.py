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
    if not os.path.exists(local_path):
        st.error("BD_Movilidad.xlsx not found.")
        st.stop()
    outgoing = pd.read_excel(local_path, sheet_name="Outgoing")
    incoming = pd.read_excel(local_path, sheet_name="Incoming")
    return outgoing, incoming

outgoing_df, incoming_df = load_data()

# ------------------------------------------------------
# ESTANDARIZACIÓN TIPO MOVILIDAD
# ------------------------------------------------------

def normalize_mobility(df):
    df = df.copy()
    df["Tipo de Movilidad"] = (
        df["Tipo de Movilidad"]
        .astype(str)
        .str.strip()
        .str.lower()
        .str.title()
    )
    
    df["Tipo de Movilidad"] = df["Tipo de Movilidad"].replace({
        "Intercambio": "Intercambio Internacional",
        "Intercambio Internacional": "Intercambio Internacional",
        "Pasantía": "Pasantía de Investigación",
        "Pasantia": "Pasantía de Investigación",
        "Pasantía De Investigación": "Pasantía de Investigación"
    })
    
    return df

outgoing_df = normalize_mobility(outgoing_df)
incoming_df = normalize_mobility(incoming_df)

# ------------------------------------------------------
# UNIFICAR ESTRUCTURA PARA CONVENIOS
# ------------------------------------------------------

outgoing_conv = outgoing_df.rename(columns={
    "Año de Movilidad": "Año",
    "Nivel": "Nivel Académico"
})

outgoing_conv["Flow"] = "Outgoing"

incoming_conv = incoming_df.rename(columns={
    "Nivel Nominación": "Nivel Académico"
})

incoming_conv["Flow"] = "Incoming"

combined = pd.concat([
    outgoing_conv[["Universidad KPIs", "País", "Año", "Tipo de Movilidad", "Nivel Académico", "Flow"]],
    incoming_conv[["Universidad KPIs", "País", "Año", "Tipo de Movilidad", "Nivel Académico", "Flow"]]
])

# ------------------------------------------------------
# UTILIZACIÓN DE CONVENIOS
# ------------------------------------------------------

st.markdown('<div class="section-title">Faculty Agreement Utilization</div>', unsafe_allow_html=True)

current_year = datetime.now().year
last_5_years = list(range(current_year - 4, current_year + 1))

mobility_types = sorted(combined["Tipo de Movilidad"].dropna().unique())

for mobility in mobility_types:

    st.markdown(f"### {mobility}")

    df_mob = combined[combined["Tipo de Movilidad"] == mobility]

    for level in ["Pregrado", "Posgrado"]:
        
        st.markdown(f"#### {level}")
        
        df_level = df_mob[df_mob["Nivel Académico"].str.contains(level, case=False, na=False)]

        if df_level.empty:
            st.info("No data available.")
            continue
        
        # ------------------------------
        # TOP 10 GRÁFICA (últimos 5 años)
        # ------------------------------
        
        df_last5 = df_level[df_level["Año"].isin(last_5_years)]
        
        top10 = (
            df_last5
            .groupby("Universidad KPIs")
            .size()
            .reset_index(name="Total")
            .sort_values("Total", ascending=False)
            .head(10)
        )

        if not top10.empty:
            fig = px.bar(
                top10,
                x="Total",
                y="Universidad KPIs",
                orientation="h",
                color_discrete_sequence=[PRIMARY_COLOR]
            )
            fig.update_layout(
                yaxis=dict(autorange="reversed"),
                height=400,
                margin=dict(l=20, r=20, t=30, b=20)
            )
            st.plotly_chart(fig, use_container_width=True)

        # ------------------------------
        # TABLA MATRIZ POR AÑO
        # ------------------------------

        pivot = (
            df_level
            .groupby(["Universidad KPIs", "País", "Año", "Flow"])
            .size()
            .reset_index(name="Count")
        )

        pivot_table = pivot.pivot_table(
            index=["Universidad KPIs", "País"],
            columns=["Año", "Flow"],
            values="Count",
            fill_value=0
        )

        # Ordenar columnas por año ascendente
        pivot_table = pivot_table.sort_index(axis=1, level=0)

        st.dataframe(pivot_table, use_container_width=True)

