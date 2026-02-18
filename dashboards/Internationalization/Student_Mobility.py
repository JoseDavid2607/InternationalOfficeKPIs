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
# UTILIZACIÓN DE CONVENIOS
# ------------------------------------------------------

st.markdown('<div class="section-title">Faculty Agreement Utilization</div>', unsafe_allow_html=True)

def agreement_section(level_name, level_out_col, level_in_col):

    st.markdown(f"### {level_name}")

    out_level = outgoing_df[outgoing_df[level_out_col] == level_name].copy()
    in_level = incoming_df[incoming_df[level_in_col] == level_name].copy()

    mobility_types = sorted(
        set(out_level["Tipo de Movilidad"].dropna().unique())
        | set(in_level["Tipo de Movilidad"].dropna().unique())
    )

    for mobility in mobility_types:

        st.markdown(f"#### {mobility}")

        out_m = out_level[out_level["Tipo de Movilidad"] == mobility]
        in_m = in_level[in_level["Tipo de Movilidad"] == mobility]

        # NORMALIZAR AÑO
        out_m = out_m.rename(columns={"Año de Movilidad": "Año"})
        in_m = in_m.rename(columns={"Año": "Año"})

        # ÚLTIMOS 5 AÑOS
        max_year = max(
            out_m["Año"].max() if not out_m.empty else 0,
            in_m["Año"].max() if not in_m.empty else 0
        )
        min_year = max_year - 4

        out_5 = out_m[out_m["Año"] >= min_year]
        in_5 = in_m[in_m["Año"] >= min_year]

        # TOP 10 UNIVERSIDADES (Outgoing + Incoming)
        top_out = out_5.groupby("Universidad KPIs").size()
        top_in = in_5.groupby("Universidad KPIs").size()

        top_combined = (top_out.add(top_in, fill_value=0)
                        .sort_values(ascending=False)
                        .head(10)
                        .reset_index())

        top_combined.columns = ["Universidad", "Movilidad"]

        if not top_combined.empty:
            fig = px.bar(
                top_combined,
                x="Movilidad",
                y="Universidad",
                orientation="h",
                color_discrete_sequence=[PRIMARY_COLOR]
            )
            fig.update_layout(height=400)
            st.plotly_chart(fig, use_container_width=True)

        # TABLA MATRIZ COMPLETA
        out_ag = (
            out_m.groupby(["Universidad KPIs", "País", "Año"])
            .size()
            .reset_index(name="Outgoing")
        )

        in_ag = (
            in_m.groupby(["Universidad KPIs", "País", "Año"])
            .size()
            .reset_index(name="Incoming")
        )

        agreements = pd.merge(
            out_ag,
            in_ag,
            on=["Universidad KPIs", "País", "Año"],
            how="outer"
        ).fillna(0)

        agreements["Outgoing"] = agreements["Outgoing"].astype(int)
        agreements["Incoming"] = agreements["Incoming"].astype(int)

        years_sorted = sorted(agreements["Año"].unique())

        pivot = agreements.pivot_table(
            index=["Universidad KPIs", "País"],
            columns="Año",
            values=["Outgoing", "Incoming"],
            fill_value=0
        )

        # Reordenar columnas: Año ascendente, y dentro Outgoing primero
        pivot = pivot.sort_index(axis=1, level=1)

        st.dataframe(pivot, use_container_width=True)


# ------------------------------------------------------
# ORDEN: PREGRADO → POSGRADO
# ------------------------------------------------------

agreement_section("Pregrado", "Nivel", "Nivel Nominación")
agreement_section("Posgrado", "Nivel", "Nivel Nominación")
