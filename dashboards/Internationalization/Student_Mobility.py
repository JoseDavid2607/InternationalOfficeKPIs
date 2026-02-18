import streamlit as st
import pandas as pd
import os

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
        "pasantía": "Pasantía de Investigación",
        "pasantia": "Pasantía de Investigación",
        "pasantía de investigación": "Pasantía de Investigación",
        "pasantia de investigacion": "Pasantía de Investigación"
    })
    
    return df

outgoing_df = normalize_mobility_type(outgoing_df)
incoming_df = normalize_mobility_type(incoming_df)

# ------------------------------------------------------
# FILTROS
# ------------------------------------------------------

st.sidebar.header("Filters")
view_type = st.sidebar.radio("View Type", ["Annual", "Semester"])

# ------------------------------------------------------
# FUNCIÓN TABLAS POR TIPO
# ------------------------------------------------------

def create_tables_by_mobility_type(df, program_col, year_col, semester_col=None):

    mobility_types = df["Tipo de Movilidad"].dropna().unique()
    
    for mobility in sorted(mobility_types):
        
        st.markdown(f"#### {mobility}")
        
        df_filtered = df[df["Tipo de Movilidad"] == mobility].copy()
        
        if view_type == "Annual":
            pivot = pd.pivot_table(
                df_filtered,
                index=program_col,
                columns=year_col,
                aggfunc="size",
                fill_value=0
            )
        else:
            df_filtered["Periodo"] = (
                df_filtered[year_col].astype(str) + "-" +
                df_filtered[semester_col].astype(str)
            )
            
            pivot = pd.pivot_table(
                df_filtered,
                index=program_col,
                columns="Periodo",
                aggfunc="size",
                fill_value=0
            )
        
        pivot["Total"] = pivot.sum(axis=1)
        pivot = pivot.sort_values("Total", ascending=False)
        
        st.dataframe(pivot, use_container_width=True)


# ------------------------------------------------------
# OUTGOING
# ------------------------------------------------------

st.markdown('<div class="section-title">Outgoing Mobility</div>', unsafe_allow_html=True)

create_tables_by_mobility_type(
    outgoing_df,
    program_col="Programa Postulación",
    year_col="Año de Movilidad",
    semester_col="Período de Movilidad"
)

# ------------------------------------------------------
# INCOMING
# ------------------------------------------------------

st.markdown('<div class="section-title">Incoming Mobility</div>', unsafe_allow_html=True)

create_tables_by_mobility_type(
    incoming_df,
    program_col="Programa Nominación",
    year_col="Año",
    semester_col="Semestre"
)

# ------------------------------------------------------
# UTILIZACIÓN DE CONVENIOS (MATRIZ POR AÑO)
# ------------------------------------------------------

st.markdown('<div class="section-title">Faculty Agreement Utilization</div>', unsafe_allow_html=True)

# OUTGOING
out_agreements = (
    outgoing_df
    .groupby(["Universidad KPIs", "País", "Año de Movilidad"])
    .size()
    .reset_index(name="Outgoing")
    .rename(columns={"Año de Movilidad": "Año"})
)

# INCOMING
in_agreements = (
    incoming_df
    .groupby(["Universidad KPIs", "País", "Año"])
    .size()
    .reset_index(name="Incoming")
)

# MERGE
agreements = pd.merge(
    out_agreements,
    in_agreements,
    on=["Universidad KPIs", "País", "Año"],
    how="outer"
).fillna(0)

agreements["Outgoing"] = agreements["Outgoing"].astype(int)
agreements["Incoming"] = agreements["Incoming"].astype(int)

# CREAR MATRIZ MULTI-COLUMNA
agreements_pivot = agreements.pivot_table(
    index=["Universidad KPIs", "País"],
    columns="Año",
    values=["Outgoing", "Incoming"],
    fill_value=0
)

agreements_pivot = agreements_pivot.sort_index()

with st.expander("Show Agreement Utilization Table"):
    st.dataframe(agreements_pivot, use_container_width=True)
