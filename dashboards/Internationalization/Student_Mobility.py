import streamlit as st
import pandas as pd
import os
import plotly.express as px

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
        margin-top: 40px;
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
    path = "data/Internationalization/BD_Movilidad.xlsx"
    if not os.path.exists(path):
        st.error("BD_Movilidad.xlsx not found.")
        st.stop()
    return (
        pd.read_excel(path, sheet_name="Outgoing"),
        pd.read_excel(path, sheet_name="Incoming")
    )

outgoing_df, incoming_df = load_data()

# ------------------------------------------------------
# ESTANDARIZACIÓN
# ------------------------------------------------------

def normalize_mobility(df):
    df = df.copy()
    
    df["Tipo de Movilidad"] = (
        df["Tipo de Movilidad"]
        .astype(str)
        .str.strip()
        .str.title()
    )

    df["Tipo de Movilidad"] = df["Tipo de Movilidad"].replace({
        "Intercambio": "Intercambio Internacional",
        "Intercambio Internacional": "Intercambio Internacional",
        "Pasantia": "Pasantía de Investigación",
        "Pasantía": "Pasantía de Investigación",
        "Pasantia De Investigacion": "Pasantía de Investigación",
        "Pasantía De Investigación": "Pasantía de Investigación"
    })

    return df

outgoing_df = normalize_mobility(outgoing_df)
incoming_df = normalize_mobility(incoming_df)

# ------------------------------------------------------
# FUNCIÓN UTILIZACIÓN CONVENIOS
# ------------------------------------------------------

def agreement_utilization_by_type(mobility_type):

    st.markdown(f"### {mobility_type}")

    out = outgoing_df[outgoing_df["Tipo de Movilidad"] == mobility_type].copy()
    inc = incoming_df[incoming_df["Tipo de Movilidad"] == mobility_type].copy()

    # Unificar columnas año
    out["Año"] = out["Año de Movilidad"]
    inc["Año"] = inc["Año"]

    # Últimos 5 años
    max_year = max(out["Año"].max(), inc["Año"].max())
    last_5_years = list(range(int(max_year)-4, int(max_year)+1))

    # TOP 10 UNIVERSIDADES
    combined = pd.concat([
        out[["Universidad KPIs", "Año"]],
        inc[["Universidad KPIs", "Año"]]
    ])

    top10 = (
        combined[combined["Año"].isin(last_5_years)]
        .groupby("Universidad KPIs")
        .size()
        .reset_index(name="Total")
        .sort_values("Total", ascending=False)
        .head(10)
    )

    fig = px.bar(
        top10,
        x="Total",
        y="Universidad KPIs",
        orientation="h",
        color_discrete_sequence=[PRIMARY_COLOR]
    )

    fig.update_layout(
        height=500,
        yaxis=dict(autorange="reversed"),
        title="Top 10 Universities – Last 5 Years"
    )

    st.plotly_chart(fig, use_container_width=True)

    # ------------------------------------------------------
    # TABLAS POR NIVEL
    # ------------------------------------------------------

    for nivel_label, nivel_col in [
        ("Undergraduate", "Pregrado"),
        ("Graduate", "Posgrado")
    ]:

        st.markdown(f"#### {nivel_label}")

        out_nivel = out[out["Nivel"].str.contains(nivel_col, case=False, na=False)]
        inc_nivel = inc[inc["Nivel Nominación"].str.contains(nivel_col, case=False, na=False)]

        out_group = (
            out_nivel.groupby(["Universidad KPIs", "País", "Año"])
            .size()
            .reset_index(name="Outgoing")
        )

        inc_group = (
            inc_nivel.groupby(["Universidad KPIs", "País", "Año"])
            .size()
            .reset_index(name="Incoming")
        )

        merged = pd.merge(
            out_group,
            inc_group,
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

        # Orden correcto: Año → Outgoing → Incoming
        pivot = pivot.sort_index(axis=1, level=1)

        st.dataframe(pivot, use_container_width=True)


# ------------------------------------------------------
# SECCIÓN UTILIZACIÓN
# ------------------------------------------------------

st.markdown('<div class="section-title">Faculty Agreement Utilization</div>', unsafe_allow_html=True)

all_mobility_types = sorted(
    set(outgoing_df["Tipo de Movilidad"].unique())
    | set(incoming_df["Tipo de Movilidad"].unique())
)

for mobility in all_mobility_types:
    agreement_utilization_by_type(mobility)
