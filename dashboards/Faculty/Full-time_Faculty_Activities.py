import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(
    page_title="International Mobility · Facultad de Administración",
    layout="wide"
)

st.markdown("""
<style>
.main {background-color:#F6F8FB;}
.block-container {padding-top:1rem;}
div[data-testid="stMetric"]{
    border:1px solid #E4E9F2;
    border-radius:12px;
    padding:10px;
    background:white;
}
</style>
""", unsafe_allow_html=True)

@st.cache_data
def load_data():
    # Reemplazar por lectura de Excel/CSV
    return {}

DATA = load_data()

st.title("International Mobility · Facultad de Administración")
st.caption("Migración inicial desde dashboard HTML")

sections = [
    "Data Center",
    "Faculty",
    "Visiting",
    "Research",
    "Home Campus",
    "Graduates",
    "Weeks",
    "Agreements Mobility",
    "Program Mobility",
    "PhD",
    "Agreements"
]

section = st.sidebar.radio("Navigation", sections)

if section == "Data Center":
    st.header("Data Center")
    st.info("Conectar aquí los archivos Excel que alimentan el dashboard.")

elif section == "Faculty":
    st.header("Faculty")
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Faculty", "—")
    c2.metric("International", "—")
    c3.metric("Female", "—")
    c4.metric("Areas", "—")

    st.subheader("Distribution")
    fig = px.bar(x=["Example"], y=[1])
    st.plotly_chart(fig, use_container_width=True)

elif section == "Visiting":
    st.header("Visiting")
    st.dataframe(pd.DataFrame())

elif section == "Research":
    st.header("Research")
    st.dataframe(pd.DataFrame())

elif section == "Home Campus":
    st.header("Home Campus")
    st.dataframe(pd.DataFrame())

elif section == "Graduates":
    st.header("Graduates")
    st.dataframe(pd.DataFrame())

elif section == "Weeks":
    st.header("Weeks")
    st.dataframe(pd.DataFrame())

elif section == "Agreements Mobility":
    st.header("Agreements Mobility")
    st.dataframe(pd.DataFrame())

elif section == "Program Mobility":
    st.header("Program Mobility")
    st.dataframe(pd.DataFrame())

elif section == "PhD":
    st.header("PhD")
    st.dataframe(pd.DataFrame())

elif section == "Agreements":
    st.header("Agreements")
    st.dataframe(pd.DataFrame())

st.divider()
st.caption("International Mobility · Facultad de Administración · Universidad de los Andes")
