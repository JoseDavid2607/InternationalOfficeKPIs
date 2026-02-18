import streamlit as st
import pandas as pd
import os

# ------------------------------------------------------
# CONFIGURACIÓN GENERAL
# ------------------------------------------------------

st.set_page_config(
    page_title="Student Mobility Indicators",
    layout="wide"
)

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
# CARGA DE DATOS
# ------------------------------------------------------

@st.cache_data
def load_data():
    try:
        # RUTA LOCAL (RECOMENDADO EN STREAMLIT CLOUD)
        local_path = "data/Internationalization/BD_Movilidad.xlsx"
        
        if os.path.exists(local_path):
            outgoing = pd.read_excel(local_path, sheet_name="Outgoing")
            incoming = pd.read_excel(local_path, sheet_name="Incoming")
        else:
            # FALLBACK GITHUB RAW (AJUSTA TU USUARIO SI ES NECESARIO)
            url = "https://raw.githubusercontent.com/<TU_USUARIO>/InternationalOfficeKPIs/main/data/Internationalization/BD_Movilidad.xlsx"
            outgoing = pd.read_excel(url, sheet_name="Outgoing")
            incoming = pd.read_excel(url, sheet_name="Incoming")
        
        return outgoing, incoming

    except Exception as e:
        st.error("Error loading BD_Movilidad.xlsx. Verify path or GitHub raw URL.")
        st.stop()

outgoing_df, incoming_df = load_data()

# ------------------------------------------------------
# SIDEBAR FILTROS
# ------------------------------------------------------

st.sidebar.header("Filters")

view_type = st.sidebar.radio(
    "View Type",
    ["Annual", "Semester"]
)

years_out = sorted(outgoing_df["Año de Movilidad"].dropna().unique())
selected_year_out = st.sidebar.multiselect(
    "Outgoing Year",
    years_out,
    default=years_out
)

years_in = sorted(incoming_df["Año"].dropna().unique())
selected_year_in = st.sidebar.multiselect(
    "Incoming Year",
    years_in,
    default=years_in
)

# ------------------------------------------------------
# SECCIÓN 1 – OUTGOING
# ------------------------------------------------------

st.markdown('<div class="section-title">Outgoing Mobility</div>', unsafe_allow_html=True)

filtered_out = outgoing_df[
    outgoing_df["Año de Movilidad"].isin(selected_year_out)
]

if view_type == "Annual":
    grouped_out = (
        filtered_out
        .groupby(["Programa Postulación", "Año de Movilidad", "Tipo de Movilidad"])
        .size()
        .reset_index(name="Students")
        .sort_values(["Programa Postulación", "Año de Movilidad"])
    )
else:
    grouped_out = (
        filtered_out
        .groupby(["Programa Postulación", "Período de Movilidad", "Tipo de Movilidad"])
        .size()
        .reset_index(name="Students")
        .sort_values(["Programa Postulación", "Período de Movilidad"])
    )

st.dataframe(grouped_out, use_container_width=True)

# ------------------------------------------------------
# SECCIÓN 2 – INCOMING
# ------------------------------------------------------

st.markdown('<div class="section-title">Incoming Mobility</div>', unsafe_allow_html=True)

filtered_in = incoming_df[
    incoming_df["Año"].isin(selected_year_in)
]

if view_type == "Annual":
    grouped_in = (
        filtered_in
        .groupby(["Programa Nominación", "Año", "Tipo de Movilidad"])
        .size()
        .reset_index(name="Students")
        .sort_values(["Programa Nominación", "Año"])
    )
else:
    grouped_in = (
        filtered_in
        .groupby(["Programa Nominación", "Semestre", "Tipo de Movilidad"])
        .size()
        .reset_index(name="Students")
        .sort_values(["Programa Nominación", "Semestre"])
    )

st.dataframe(grouped_in, use_container_width=True)

# ------------------------------------------------------
# SECCIÓN 3 – UTILIZACIÓN DE CONVENIOS
# ------------------------------------------------------

st.markdown('<div class="section-title">Faculty Agreement Utilization</div>', unsafe_allow_html=True)

# OUTGOING
out_agreements = (
    outgoing_df
    .groupby(["Universidad KPIs", "País", "Año de Movilidad"])
    .size()
    .reset_index(name="Outgoing_Count")
    .rename(columns={"Año de Movilidad": "Año"})
)

# INCOMING
in_agreements = (
    incoming_df
    .groupby(["Universidad KPIs", "País", "Año"])
    .size()
    .reset_index(name="Incoming_Count")
)

# MERGE
agreements = pd.merge(
    out_agreements,
    in_agreements,
    on=["Universidad KPIs", "País", "Año"],
    how="outer"
).fillna(0)

agreements["Outgoing_Count"] = agreements["Outgoing_Count"].astype(int)
agreements["Incoming_Count"] = agreements["Incoming_Count"].astype(int)

agreements = agreements.sort_values(["Universidad KPIs", "Año"])

with st.expander("Show Agreement Utilization Table"):
    st.dataframe(agreements, use_container_width=True)
