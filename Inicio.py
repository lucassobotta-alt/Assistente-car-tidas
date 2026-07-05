import streamlit as st

# ── Cabeçalho principal ──────────────────────────────────────────────────────
st.markdown(
    """
    <div style="text-align: center; padding: 2rem 0 1rem 0;">
        <h1 style="font-size: 2.2rem; margin-bottom: 0.2rem;">
            🩺 Assistente de Laudos Vasculares
        </h1>
        <p style="color: #555; font-size: 1.05rem; margin-top: 0;">
            Sistema de apoio à elaboração de laudos de Duplex Scan vascular
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown("---")

# ── Cartão de identidade profissional ────────────────────────────────────────
st.markdown(
    """
    <div style="
        background: linear-gradient(135deg, #f0f4ff 0%, #e8f0fe 100%);
        border-left: 5px solid #1a56db;
        border-radius: 8px;
        padding: 1.6rem 2rem;
        margin-bottom: 1.8rem;
    ">
        <p style="font-size: 0.82rem; color: #666; margin-bottom: 0.3rem; text-transform: uppercase; letter-spacing: 0.07em;">
            Projeto desenvolvido por
        </p>
        <h2 style="margin: 0 0 0.6rem 0; font-size: 1.55rem; color: #1a1a2e;">
            Dr. Lucas Santos Guimarães
        </h2>
        <div style="display: flex; flex-wrap: wrap; gap: 0.5rem 1.8rem; font-size: 0.95rem; color: #333;">
            <span>🪪 <strong>CRM-SE</strong> 4061</span>
            <span>🎓 <strong>Radiologia e Diagnóstico por Imagem</strong> · RQE 4015</span>
            <span>🔬 <strong>Ecografia Vascular com Doppler</strong> · RQE 4016</span>
            <span>🏛️ Membro Titular do <strong>Colégio Brasileiro de Radiologia</strong></span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ── Módulos disponíveis ───────────────────────────────────────────────────────
st.markdown("### Módulos disponíveis")

col1, col2, col3 = st.columns(3, gap="large")

with col1:
    st.markdown(
        """
        <div style="
            border: 1px solid #dde3f0;
            border-radius: 8px;
            padding: 1.2rem 1.4rem 0.8rem 1.4rem;
            background: #fff;
            min-height: 170px;
        ">
            <h4 style="margin-top: 0;">⚕️ Duplex Scan Arterial Carotídeo</h4>
            <ul style="padding-left: 1.1rem; color: #444; font-size: 0.92rem;">
                <li>Classificação hemodinâmica (SBC 2023 / NASCET)</li>
                <li>Mapeamento de placas ateroscleróticas e Plaque-RADS</li>
                <li>Suporte a tortuosidades e vasculites</li>
                <li>Template personalizável com marcadores</li>
            </ul>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.page_link("pages/01_Arterial_Carotidas.py", label="Acessar módulo arterial →", icon="⚕️")

with col2:
    st.markdown(
        """
        <div style="
            border: 1px solid #dde3f0;
            border-radius: 8px;
            padding: 1.2rem 1.4rem 0.8rem 1.4rem;
            background: #fff;
            min-height: 170px;
        ">
            <h4 style="margin-top: 0;">🌀 Duplex Scan Venoso MMII</h4>
            <ul style="padding-left: 1.1rem; color: #444; font-size: 0.92rem;">
                <li>Mapeamento da VSM, VSP e perfurantes incompetentes</li>
                <li>Cartografia venosa esquemática (imagem PNG)</li>
                <li>Suporte a exame unilateral e bilateral</li>
                <li>Classificação por diretriz ESVS 2022</li>
            </ul>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.page_link("pages/02_Venoso_MMII.py", label="Acessar módulo venoso →", icon="🌀")

with col3:
    st.markdown(
        """
        <div style="
            border: 1px solid #dde3f0;
            border-radius: 8px;
            padding: 1.2rem 1.4rem 0.8rem 1.4rem;
            background: #fff;
            min-height: 170px;
        ">
            <h4 style="margin-top: 0;">🫀 Duplex Scan Arterial de MMII</h4>
            <ul style="padding-left: 1.1rem; color: #444; font-size: 0.92rem;">
                <li>Mapeamento arterial de AFC até A. Fibular</li>
                <li>Classificação de estenose por razão de velocidades (PVS)</li>
                <li>Propagação hemodinâmica em cascata para vasos distais</li>
                <li>Suporte a exame unilateral e bilateral</li>
            </ul>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.page_link("pages/03_Arterial_MMII.py", label="Acessar módulo arterial MMII →", icon="🫀")

st.markdown("---")
st.caption("Use o menu lateral para navegar entre os módulos.")
