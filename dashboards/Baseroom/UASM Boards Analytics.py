"""
Governance & Advisory Board Members Dashboard
Universidad de los Andes / UASM

Lee Governing_Bodies.xlsx (debe estar en el mismo directorio que este
archivo, o se puede subir manualmente desde la barra lateral) y muestra
un dashboard de 5 páginas -- una por cuerpo de gobierno / junta asesora:

  1) Board of Trustees
  2) Executive Committee
  3) Academic Council
  4) International Advisory Board
  5) UASM Committees

Ejecutar con:  streamlit run Governance_Dashboard.py
"""
import os
import io
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# ───────────────────────── Config & estilo ─────────────────────────
st.set_page_config(
    page_title="Governance & Advisory Boards",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded",
)

PRIMARY = "#1F6F54"      # verde institucional (Uniandes / UASM)
PRIMARY_DARK = "#14503C"
ACCENT = "#F4B400"       # amarillo del logo
SUPPORT = "#2E86AB"      # azul-verde de soporte
NEUTRAL = "#6B7280"
PALETTE = ["#1F6F54", "#2E86AB", "#F4B400", "#9D4EDD", "#E76F51",
           "#27BDAE", "#4FFF98", "#FFD166", "#118AB2", "#8AC926"]

st.markdown(f"""
<style>
    .stApp {{ background-color: #FAFBFB; }}
    [data-testid="stMetric"] {{
        background-color: #FFFFFF; border: 1px solid #E5E7EB; border-radius: 10px;
        padding: 14px 16px; box-shadow: 0 1px 2px rgba(0,0,0,0.04);
    }}
    [data-testid="stMetricLabel"] {{ color: {NEUTRAL}; font-weight: 600; }}
    [data-testid="stMetricValue"] {{ color: {PRIMARY_DARK}; }}
    h1, h2, h3 {{ color: {PRIMARY_DARK}; }}
    section[data-testid="stSidebar"] {{ background-color: #FFFFFF; border-right: 1px solid #E5E7EB; }}
</style>
""", unsafe_allow_html=True)

DEFAULT_XLSX = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Governing_Bodies.xlsx")


# ───────────────────────── Carga de datos ─────────────────────────
@st.cache_data
def load_workbook(source) -> dict:
    """Lee las 5 hojas del Excel. Cada hoja tiene un bloque de
    título/subtítulo (1 celda con contenido) seguido de una fila en
    blanco y luego el encabezado real (2+ celdas con contenido en la
    misma fila) -- se detecta automáticamente en vez de asumir una
    posición fija, para no romperse si alguna hoja no tiene subtítulo."""
    xl = pd.ExcelFile(source)
    sheets = {}
    for name in xl.sheet_names:
        raw = pd.read_excel(xl, sheet_name=name, header=None)
        header_row = 0
        for i in range(len(raw)):
            non_null = raw.iloc[i].notna().sum()
            if non_null >= 2:
                header_row = i
                break
        df = pd.read_excel(xl, sheet_name=name, skiprows=header_row)
        df = df.dropna(how="all")
        sheets[name] = df
    return sheets


def get_source():
    with st.sidebar:
        st.markdown("### 📁 Data source")
        uploaded = st.file_uploader("Governing_Bodies.xlsx", type=["xlsx"], label_visibility="collapsed")
        if uploaded is not None:
            return io.BytesIO(uploaded.getvalue())
        if os.path.exists(DEFAULT_XLSX):
            st.caption(f"Using `{os.path.basename(DEFAULT_XLSX)}` found alongside this app.")
            return DEFAULT_XLSX
        st.warning("Upload **Governing_Bodies.xlsx** to load the dashboard.")
        st.stop()


def kpi_row(items):
    cols = st.columns(len(items))
    for c, (label, value, delta) in zip(cols, items):
        with c:
            if delta:
                st.metric(label, value, delta)
            else:
                st.metric(label, value)


def render_header(title, subtitle):
    st.markdown(f"## {title}")
    st.caption(subtitle)
    st.markdown("---")


def categorize_position(pos: str) -> str:
    """Agrupa un texto de cargo libre en una categoría amplia, para poder
    graficar la composición por tipo de rol sin importar la redacción
    exacta de cada cuerpo de gobierno."""
    p = str(pos).lower()
    if "president" in p and "vice" not in p:
        return "President"
    if "vice-president" in p or "vice president" in p:
        return "Vice-President"
    if "dean" in p:
        return "Dean"
    if "director" in p:
        return "Director"
    if "coordinator" in p or "manager" in p or "officer" in p:
        return "Coordinator / Manager"
    if "professor" in p:
        return "Professor"
    if "student" in p:
        return "Student Representative"
    if "advisor" in p or "external" in p:
        return "External Advisor"
    if "representative" in p:
        return "Representative"
    return "Other / Member"


# ───────────────────────── Páginas ─────────────────────────
def page_trustees(sheets):
    df = sheets["Board of Trustees"].copy()
    render_header("Board of Trustees", "Universidad de los Andes — Numerary, Honorary and Permanent Members")

    total = len(df)
    n_num = (df["Category"] == "Numerary Members").sum()
    n_hon = (df["Category"] == "Honorary Members").sum()
    n_perm = (df["Category"] == "Permanent Members").sum()
    n_students = df["Role"].astype(str).str.contains("Student", case=False).sum()
    n_prof = df["Role"].astype(str).str.contains("Professor", case=False).sum()

    kpi_row([
        ("Total Members", total, None),
        ("Numerary", int(n_num), None),
        ("Honorary", int(n_hon), None),
        ("Permanent", int(n_perm), None),
        ("Student Reps.", int(n_students), None),
    ])

    st.markdown("")
    c1, c2 = st.columns([1, 1])
    with c1:
        counts = df["Category"].value_counts().reindex(
            ["Numerary Members", "Honorary Members", "Permanent Members"]).fillna(0)
        fig = px.bar(x=counts.index, y=counts.values, text=counts.values,
                     color=counts.index, color_discrete_sequence=PALETTE,
                     labels={"x": "", "y": "Members"}, title="Members by Category")
        fig.update_traces(textposition="outside")
        fig.update_layout(showlegend=False, margin=dict(t=50, b=10))
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        role_counts = df["Role"].value_counts()
        top_roles = role_counts[role_counts.index != "Member"]
        fig2 = go.Figure(go.Pie(
            labels=["Named role (President, VP, Student, Professor…)", "Member (no special role)"],
            values=[int(top_roles.sum()), int(total - top_roles.sum())],
            hole=0.55, marker=dict(colors=[PRIMARY, "#E5E7EB"]),
        ))
        fig2.update_layout(title="Roster Composition", margin=dict(t=50, b=10))
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown("#### Full Roster")
    cat_filter = st.multiselect("Filter by category", options=sorted(df["Category"].unique()),
                                 default=sorted(df["Category"].unique()))
    st.dataframe(df[df["Category"].isin(cat_filter)].reset_index(drop=True),
                 use_container_width=True, hide_index=True, height=420)


def page_executive(sheets):
    df = sheets["Executive Committee"].copy()
    render_header("Executive Committee", "Universidad de los Andes — Board of Trustees Executive Committee")

    kpi_row([
        ("Total Members", len(df), None),
        ("With a named role", int((df["Role"] != "Member").sum()), None),
        ("Regular Members", int((df["Role"] == "Member").sum()), None),
    ])

    st.markdown("")
    c1, c2 = st.columns([1.2, 1])
    with c1:
        st.markdown("#### Members")
        for _, row in df.iterrows():
            badge = f" — *{row['Role']}*" if row["Role"] != "Member" else ""
            st.markdown(f"- **{row['Name']}**{badge}")
    with c2:
        role_counts = df["Role"].value_counts()
        fig = px.pie(names=role_counts.index, values=role_counts.values,
                      color_discrete_sequence=PALETTE, hole=0.5, title="Roles on the Committee")
        fig.update_layout(margin=dict(t=50, b=10))
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### Full Table")
    st.dataframe(df, use_container_width=True, hide_index=True)


def page_academic_council(sheets):
    df = sheets["Academic Council"].copy()
    render_header("Academic Council", "Universidad de los Andes — Academic Council Members")

    df["Category"] = df["Position"].apply(categorize_position)
    n_deans = (df["Category"] == "Dean").sum()
    n_students = (df["Category"] == "Student Representative").sum()
    n_directors = (df["Category"] == "Director").sum()

    kpi_row([
        ("Total Members", len(df), None),
        ("Deans", int(n_deans), None),
        ("Directors", int(n_directors), None),
        ("Student Reps.", int(n_students), None),
        ("Distinct Roles", df["Category"].nunique(), None),
    ])

    st.markdown("")
    c1, c2 = st.columns([1.3, 1])
    with c1:
        counts = df["Category"].value_counts().sort_values(ascending=True)
        fig = px.bar(x=counts.values, y=counts.index, orientation="h", text=counts.values,
                     color=counts.index, color_discrete_sequence=PALETTE,
                     labels={"x": "Members", "y": ""}, title="Composition by Role")
        fig.update_traces(textposition="outside")
        fig.update_layout(showlegend=False, margin=dict(t=50, b=10),
                           xaxis=dict(range=[0, counts.max() + 2]))
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig2 = px.pie(df, names="Category", color_discrete_sequence=PALETTE,
                       title="Share by Role", hole=0.5)
        fig2.update_layout(margin=dict(t=50, b=10), showlegend=False)
        fig2.update_traces(textposition="inside", textinfo="percent+label")
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown("#### Full Roster")
    role_filter = st.multiselect("Filter by role category", options=sorted(df["Category"].unique()),
                                  default=sorted(df["Category"].unique()))
    st.dataframe(df[df["Category"].isin(role_filter)][["Name", "Position", "Category"]]
                 .reset_index(drop=True), use_container_width=True, hide_index=True, height=460)


def page_iab(sheets):
    df = sheets["International Advisory Board"].copy()
    render_header("International Advisory Board", "UASM — School of Management International Advisory Board")

    terms = sorted(df["Term"].unique())
    sel_term = st.radio("Term", ["All"] + terms, horizontal=True)
    dfv = df if sel_term == "All" else df[df["Term"] == sel_term]

    n_ext = dfv["Position in IAB"].astype(str).str.contains("external", case=False).sum()
    n_school = dfv["Position in IAB"].astype(str).str.contains("School", case=False).sum()
    n_nat = dfv["Nationality"].nunique()

    kpi_row([
        ("Total Members", len(dfv), None),
        ("External Members", int(n_ext), None),
        ("School Representatives", int(n_school), None),
        ("Nationalities", int(n_nat), None),
        ("Terms Covered", df["Term"].nunique(), None),
    ])

    st.markdown("")
    c1, c2 = st.columns([1.2, 1])
    with c1:
        fig = px.bar(dfv, x="Term", color="Position in IAB", barmode="group",
                     color_discrete_sequence=PALETTE, title="Composition by Term")
        fig.update_layout(margin=dict(t=50, b=10), xaxis_title=None, yaxis_title="Members",
                           legend=dict(orientation="h", yanchor="top", y=-0.15, x=0.5, xanchor="center", title=None))
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig2 = px.pie(dfv, names="Nationality", color_discrete_sequence=PALETTE,
                       title="Nationality Mix", hole=0.5)
        fig2.update_layout(margin=dict(t=50, b=10))
        fig2.update_traces(textposition="inside", textinfo="percent+label")
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown("#### Members returning across both terms")
    repeated = df.groupby("Members")["Term"].nunique()
    repeated = repeated[repeated > 1].index.tolist()
    if repeated:
        st.info(" · ".join(repeated))
    else:
        st.caption("No members overlap across terms.")

    st.markdown("#### Full Table")
    st.dataframe(dfv.reset_index(drop=True), use_container_width=True, hide_index=True, height=380)


def page_committees(sheets):
    df = sheets["UASM Committees"].copy()
    render_header("UASM Committees", "UASM — Committee Member Composition (December 2022)")

    n_committees = df["Committee"].nunique()
    n_slots = len(df)
    n_people = df["Name"].nunique()
    avg_size = round(n_slots / n_committees, 1)

    kpi_row([
        ("Committees", n_committees, None),
        ("Total Seats", n_slots, None),
        ("Distinct People", n_people, None),
        ("Avg. Members / Committee", avg_size, None),
    ])

    st.markdown("")
    c1, c2 = st.columns([1.3, 1])
    with c1:
        sizes = df.groupby("Committee")["Name"].count().sort_values(ascending=True)
        fig = px.bar(x=sizes.values, y=sizes.index, orientation="h", text=sizes.values,
                     color=sizes.values, color_continuous_scale=[SUPPORT, PRIMARY],
                     labels={"x": "Members", "y": ""}, title="Members per Committee")
        fig.update_traces(textposition="outside")
        fig.update_layout(coloraxis_showscale=False, margin=dict(t=50, b=10, l=10),
                           xaxis=dict(range=[0, sizes.max() + 3]), height=480)
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        top_people = df["Name"].value_counts()
        top_people = top_people[top_people > 1].sort_values(ascending=True)
        if len(top_people):
            fig2 = px.bar(x=top_people.values, y=top_people.index, orientation="h",
                          text=top_people.values, color_discrete_sequence=[ACCENT],
                          labels={"x": "Committees", "y": ""},
                          title="Most Cross-Committee Members")
            fig2.update_traces(textposition="outside", marker_color=PRIMARY)
            fig2.update_layout(margin=dict(t=50, b=10, l=10), height=480,
                                xaxis=dict(range=[0, top_people.max() + 1]))
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.info("No one currently sits on more than one committee.")

    st.markdown("#### Browse a Committee")
    sel_committee = st.selectbox("Committee", sorted(df["Committee"].unique(),
                                  key=lambda c: df[df["Committee"] == c]["Committee #"].iloc[0]))
    dfc = df[df["Committee"] == sel_committee][["Name", "Position"]].reset_index(drop=True)
    st.dataframe(dfc, use_container_width=True, hide_index=True)

    with st.expander("Show full tidy table (all committees)"):
        st.dataframe(df[["Committee", "Name", "Position"]].reset_index(drop=True),
                     use_container_width=True, hide_index=True, height=420)


# ───────────────────────── Navegación ─────────────────────────
def main():
    source = get_source()
    sheets = load_workbook(source)

    with st.sidebar:
        st.markdown(f"<h2 style='color:{PRIMARY_DARK};margin-bottom:0'>🏛️ Governance</h2>", unsafe_allow_html=True)
        st.caption("Governing Body & Advisory Board Members")
        st.markdown("---")

    pages = {
        "Board of Trustees": lambda: page_trustees(sheets),
        "Executive Committee": lambda: page_executive(sheets),
        "Academic Council": lambda: page_academic_council(sheets),
        "International Advisory Board": lambda: page_iab(sheets),
        "UASM Committees": lambda: page_committees(sheets),
    }
    icons = {
        "Board of Trustees": "🏛️", "Executive Committee": "⚖️", "Academic Council": "🎓",
        "International Advisory Board": "🌍", "UASM Committees": "🗂️",
    }

    with st.sidebar:
        st.markdown("### Navigate")
        choice = st.radio("Page", [f"{icons[p]}  {p}" for p in pages],
                           label_visibility="collapsed")
        page_name = choice.split("  ", 1)[1]

        st.markdown("---")
        total_people = pd.concat([
            sheets["Board of Trustees"]["Name"], sheets["Executive Committee"]["Name"],
            sheets["Academic Council"]["Name"], sheets["International Advisory Board"]["Members"],
            sheets["UASM Committees"]["Name"],
        ]).nunique()
        st.caption(f"**{total_people}** distinct individuals across all 5 governing bodies")

    pages[page_name]()


if __name__ == "__main__":
    main()
