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
        .str.title()
    )
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
# UTILIZACIÓN DE CONVENIOS POR TIPO Y NIVEL
# ------------------------------------------------------

st.markdown('<div class="section-title">Faculty Agreement Utilization</div>', unsafe_allow_html=True)

def agreement_tables_by_type(mobility_type, level_filter, level_col_out, level_col_in):

    st.markdown(f"### {mobility_type} – {level_filter}")

    out_filtered = outgoing_df[
        (outgoing_df["Tipo de Movilidad"] == mobility_type) &
        (outgoing_df[level_col_out].str.contains(level_filter, case=False, na=False))
    ].copy()

    in_filtered = incoming_df[
        (incoming_df["Tipo de Movilidad"] == mobility_type) &
        (incoming_df[level_col_in].str.contains(level_filter, case=False, na=False))
    ].copy()

    # Obtener años combinados
    all_years = sorted(
        list(set(out_filtered["Año de Movilidad"].dropna().unique())
        .union(set(in_filtered["Año"].dropna().unique())))
    )

    if len(all_years) > 5:
        last_5_years = all_years[-5:]
    else:
        last_5_years = all_years

    # --------------------------------------------------
    # TOP 10 UNIVERSIDADES (últimos 5 años)
    # --------------------------------------------------

    out_top = (
        out_filtered[out_filtered["Año de Movilidad"].isin(last_5_years)]
        .groupby("Universidad KPIs")
        .size()
    )

    in_top = (
        in_filtered[in_filtered["Año"].isin(last_5_years)]
        .groupby("Universidad KPIs")
        .size()
    )

    top_total = (out_top + in_top).fillna(0).sort_values(ascending=False).head(10)

    if not top_total.empty:
        st.bar_chart(top_total.sort_values())

    # --------------------------------------------------
    # MATRIZ COMPLETA POR AÑO
    # --------------------------------------------------

    out_agreements = (
        out_filtered
        .groupby(["Universidad KPIs", "País", "Año de Movilidad"])
        .size()
        .reset_index(name="Outgoing")
        .rename(columns={"Año de Movilidad": "Año"})
    )

    in_agreements = (
        in_filtered
        .groupby(["Universidad KPIs", "País", "Año"])
        .size()
        .reset_index(name="Incoming")
    )

    agreements = pd.merge(
        out_agreements,
        in_agreements,
        on=["Universidad KPIs", "País", "Año"],
        how="outer"
    ).fillna(0)

    agreements["Outgoing"] = agreements["Outgoing"].astype(int)
    agreements["Incoming"] = agreements["Incoming"].astype(int)

    agreements_pivot = agreements.pivot_table(
        index=["Universidad KPIs", "País"],
        columns="Año",
        values=["Outgoing", "Incoming"],
        fill_value=0
    )

    agreements_pivot = agreements_pivot.sort_index(axis=1, level=1)

    st.dataframe(agreements_pivot, use_container_width=True)


# ------------------------------------------------------
# GENERAR TABLAS POR TIPO Y NIVEL
# ------------------------------------------------------

mobility_types = sorted(
    set(outgoing_df["Tipo de Movilidad"].unique())
    .union(set(incoming_df["Tipo de Movilidad"].unique()))
)

# Primero Pregrado
for mobility in mobility_types:
    agreement_tables_by_type(mobility, "Pregrado", "Nivel", "Nivel Nominación")

# Luego Posgrado
for mobility in mobility_types:
    agreement_tables_by_type(mobility, "Posgrado", "Nivel", "Nivel Nominación")
