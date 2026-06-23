import streamlit as st

st.set_page_config(
    page_title="Assistente de Laudos Vasculares",
    page_icon="🩺",
    layout="wide"
)

pg = st.navigation([
    st.Page("Inicio.py",                          title="Início",             icon="🏠"),
    st.Page("pages/01_Arterial_Carotidas.py",     title="Arterial Carótidas", icon="⚕️"),
    st.Page("pages/02_Venoso_MMII.py",            title="Venoso MMII",        icon="🌀"),
])
pg.run()
