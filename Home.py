import streamlit as st

st.set_page_config(
    page_title="Assistente de Laudos Vasculares",
    page_icon="🩺",
    layout="wide"
)

st.title("🩺 Assistente de Laudos Vasculares")
st.markdown("---")

col1, col2 = st.columns(2)

with col1:
    st.markdown("""
    ### ⚕️ Duplex Scan Arterial Carotídeo
    Laudo completo das artérias carótidas e vertebrais com:
    - Classificação hemodinâmica (SBC 2023 / NASCET)
    - Mapeamento de placas ateroscleróticas e Plaque-RADS
    - Suporte a tortuosidades e vasculites
    - Template personalizável com marcadores
    """)
    st.page_link("pages/01_Arterial_Carotidas.py", label="Acessar módulo arterial →", icon="⚕️")

with col2:
    st.markdown("""
    ### 🌀 Duplex Scan Venoso MMII
    Laudo completo do sistema venoso dos membros inferiores com:
    - Mapeamento da VSM, VSP e perfurantes incompetentes
    - Cartografia venosa esquemática (imagem PNG)
    - Suporte a exame unilateral e bilateral
    - Classificação por diretriz ESVS 2022
    """)
    st.page_link("pages/02_Venoso_MMII.py", label="Acessar módulo venoso →", icon="🌀")

st.markdown("---")
st.caption("Use o menu lateral para navegar entre os módulos.")
