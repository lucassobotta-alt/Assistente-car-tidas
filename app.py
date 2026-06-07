# app.py
import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import utils

if "reset_trigger" in st.session_state and st.session_state.reset_trigger:
    st.session_state.clear()
    st.session_state.reset_trigger = False

for key in ['lista_placas', 'lesoes_incipientes', 'calcificacoes_isoladas', 'lesoes_nao_ateromatosas']:
    if key not in st.session_state:
        st.session_state[key] = []

st.set_page_config(page_title="Laudo Neurovascular Avançado", layout="wide")
st.title("⚕️ Assistente de Laudos Vascular Completo")

with st.sidebar:
    st.markdown("## ⚙️ Painel de Controle Avançado")
    st.markdown("### 📚 Critérios Científicos")
    diretriz_selecionada = st.radio("Diretriz Hemodinâmica de Referência:", ["Diretriz SBC 2023", "Consenso Clássico NASCET"])
    ano_dislipidemia = st.selectbox("Ano da Diretriz de Dislipidemia:", ["2025", "2023"], index=0)
    st.markdown("---")
    st.markdown("### 📝 Formatação Externa (.docx)")
    fonte_doc = st.selectbox("Família da Fonte:", ["Arial", "Calibri", "Times New Roman"], index=0)
    tamanho_fonte = st.slider("Tamanho do Texto (pt):", 10, 14, 11, step=1)
    espacamento_linhas = st.slider("Espaçamento entre Linhas:", 1.0, 1.5, 1.15, step=0.05)
    quebrar_pagina_diag = st.toggle("Separar Impressão Diagnóstica em Nova Página", value=False)
    st.markdown("---")
    st.markdown("### ✍️ Identidade & Assinatura")
    nome_clinica = st.text_input("Cabeçalho / Nome da Clínica:")
    nome_medico = st.text_input("Nome do Médico:", "Lucas Santos Guimarães")
    crm_medico = st.text_input("CRM / RQE:", "4061")
    st.markdown("---")
    if st.button("🔄 Resetar Todos os Parâmetros", use_container_width=True, type="secondary"):
        st.session_state.reset_trigger = True
        st.rerun()

col_id1, col_id2 = st.columns([2, 2])
with col_id1:
    nome = st.text_input("Nome do Paciente", "Paciente Exemplo")
with col_id2:
    opcao_selecionada = st.selectbox("Condições Técnicas do Exame:", list(utils.opcoes_tecnicas.keys()))
    texto_tecnica_final = utils.opcoes_tecnicas[opcao_selecionada]

st.markdown("---")
st.markdown("### 📊 Parâmetros Hemodinâmicos")
col_hemo_dir, col_hemo_esq = st.columns(2)

with col_hemo_dir:
    st.header("LADO DIREITO")
    cmi_dir = st.number_input("CMI ACC Direita (mm)", 0.0, 5.0, 0.4, 0.1)
    estado_aci_dir = st.selectbox("Estado da ACI Direita", ["Pérvia (Calcular por Velocidade)", "Suboclusão", "Oclusão"])
    vps_aci_dir = st.number_input("VPS ACI Direita (cm/s)", 0.0, value=90.0, step=5.0)
    vcc_dir = st.number_input("VPS ACC Direita (cm/s)", 1.0, value=60.0, step=5.0)
    ace_dir = st.selectbox("ACE Direita", ["Com padrão espectral de alta resistência, compatível com perfusão de leitos musculares extracranianos.", "Alterada / Estenose hemodinâmica"])
    espectro_vert_dir = st.selectbox("Espectro AV Direita", ["Normal (Fluxo Anterógrado)", "Hipoplasia", "Roubo Latente", "Roubo Parcial (Fluxo Alternante)", "Roubo Total (Fluxo Retrógrado)"])
    vps_vert_dir = st.number_input("VPS AV Direita (cm/s)", 0.0, value=30.0, step=5.0)

with col_hemo_esq:
    st.header("LADO ESQUERDO")
    cmi_esq = st.number_input("CMI ACC Esquerda (mm)", 0.0, 5.0, 0.4, 0.1)
    estado_aci_esq = st.selectbox("Estado da ACI Esquerda", ["Pérvia (Calcular por Velocidade)", "Suboclusão", "Oclusão"])
    vps_aci_esq = st.number_input("VPS ACI Esquerda (cm/s)", 0.0, value=100.0, step=5.0)
    vcc_esq = st.number_input("VPS ACC Esquerda (cm/s)", 1.0, value=60.0, step=5.0)
    ace_esq = st.selectbox("ACE Esquerda", ["Com padrão espectral de alta resistência, compatível com perfusão de leitos musculares extracranianos.", "Alterada / Estenose hemodinâmica"])
    espectro_vert_esq = st.selectbox("Espectro AV Esquerda", ["Normal (Fluxo Anterógrado)", "Hipoplasia", "Roubo Latente", "Roubo Parcial (Fluxo Alternante)", "Roubo Total (Fluxo Retrógrado)"])
    vps_vert_esq = st.number_input("VPS AV Esquerda (cm/s)", 0.0, value=30.0, step=5.0)

with st.expander("🔄 2. Lesões Não Ateromatosas (Tortuosidades e Vasculite)"):
    col_na1, col_na2 = st.columns(2)
    with col_na1:
        vaso_na = st.selectbox("Vaso com Lesão Não Ateromatosa:", ["Artéria carótida interna", "Artéria carótida comum", "Bulbo carotídeo", "Artéria vertebral"])
        lado_na = st.selectbox("Lado do Achado:", ["Direito", "Esquerdo", "Bilateral"])
    with col_na2:
        tipo_na = st.selectbox("Selecione a Lesão Não Ateromatosa:", ["Tortuosidade / Alongamento (Alongamento simples)", "Tortuosidade / Alongamento (Kinking - Angulação aguda < 90°)", "Tortuosidade / Alongamento (Coiling - Alça completa em espiral)", "Vasculite / Arterite (Espessamento parietal concêntrico e homogêneo - Sinal do Halo)", "Vasculite / Arterite (Espessamento difuso irregular de padrão inflamatório)"])
        hemo_na = st.toggle("Gera alteração hemodinâmica?", value=False)
    if st.button("💾 Registrar Lesão Não Ateromatosa"):
        for ld in (["Direito", "Esquerdo"] if lado_na == "Bilateral" else [lado_na]):
            st.session_state.lesoes_nao_ateromatosas.append({"vaso": vaso_na, "lado": ld, "tipo": tipo_na, "hemo": hemo_na})
        st.rerun()
    for idx, na in enumerate(st.session_state.lesoes_nao_ateromatosas):
        st.write(f"• `{idx+1:02d}` **{na['vaso']} {na['lado']}**: {na['tipo']}")

with st.expander("🌱 3. Lesões Estruturais Incipientes (Alterações Precoces ≤ 1.5 mm)"):
    if st.checkbox("Incluir achados de Lesão Incipiente?"):
        c_i1, c_i2, c_i3 = st.columns(3)
        with c_i1: vaso_inc = st.selectbox("Vaso (Incipiente):", ["Artéria carótida comum direito", "Bulbo carotídeo direito", "Artéria carótida interna direito", "Artéria carótida comum esquerdo", "Bulbo carotídeo esquerdo", "Artéria carótida interna esquerdo"])
        with c_i2: local_inc = st.selectbox("Localização (Incipiente):", ["terço proximal", "terço médio", "terço distal", "segmento total"])
        with c_i3: espessura_inc = st.number_input("Espessura (mm):", 0.5, 1.5, 1.2, 0.1)
        if st.button("💾 Registrar Lesão Incipiente"):
            st.session_state.lesoes_incipientes.append({"vaso": vaso_inc, "localizacao": local_inc, "espessura": espessura_inc})
            st.rerun()
    for idx, inc in enumerate(st.session_state.lesoes_incipientes):
        st.write(f"• `{idx+1:02d}` **{inc['vaso']}**: {inc['espessura']} mm")

with st.expander("🔎 4. Mapeamento de Placas Ateroscleróticas (Consolidadas ≥ 2.0 mm)"):
    if st.checkbox("Incluir achados de placas ateromatosas?"):
        c_v, c_l = st.columns(2)
        with c_v: vaso_sel = st.selectbox("Artéria:", ["Artéria carótida comum direita", "Bulbo carotídeo direito", "Artéria carótida interna direita", "Artéria carótida comum esquerda", "Bulbo carotídeo esquerdo", "Artéria carótida interna esquerda"])
        with c_l: local_sel = st.selectbox("Segmento:", ["terço proximal", "terço médio", "terço distal", "segmento total/bifurcação"])
        culpada_hemo = st.toggle("Determina estenose significativa?", value=False)
        c1, c2, c3 = st.columns(3)
        with c1:
            composicao = st.selectbox("Composição (Plaque-RADS):", ["1. Placa calcificada", "2. Placa uniformemente ecogênica (fibrosa estável)", "3. Placa predominantemente ecogênica (fibrosa)", "4. Placa com componente anecogênico delimitado por halo hiperecogênico", "5. Placa predominantemente anecogênica", "6. Placa uniformemente anecogênica sem cápsula fibrosa definida"])
            usar_pr = st.toggle("Incluir Plaque-RADS?", value=False)
            espessura = st.number_input("Espessura Máxima (mm):", 2.0, 20.0, 2.5, 0.1)
        with c2: superficie = st.selectbox("Superfície:", ["Regular", "Irregular", "Ulcerada"])
        with c3:
            st.markdown("**Adicionais:**")
            t_tr = st.checkbox("Trombo adjacente")
            t_fl = st.checkbox("Fluxo intraplaca")
            t_ca = st.checkbox("Focos calcificação")
        if st.button("💾 Registrar Placa"):
            pr_est = utils.estimate_plaque_rads(composicao) if usar_pr else None
            st.session_state.lista_placas.append({"vaso": vaso_sel, "localizacao": local_sel, "culpada_hemo": culpada_hemo, "composicao_texto": utils.retirar_prefixo_numerico(composicao), "superficie_texto": superficie, "espessura": espessura, "plaque_rads": pr_est})
            st.rerun()
    for idx, p in enumerate(st.session_state.lista_placas):
        st.write(f"• `{idx+1:02d}` **{p['vaso']}**: {p['composicao_texto']} ({p['espessura']} mm)")

with st.expander("🪨 5. Calcificações Parietais Isoladas (Sem Formação de Placa)"):
    if st.checkbox("Incluir Calcificações Parietais Isoladas?"):
        cc1, cc2 = st.columns(2)
        with cc1: lado_c = st.selectbox("Lado:", ["Direito", "Esquerdo", "Bilateral"])
        with cc2: topo_c = st.selectbox("Anatomia:", ["bulbo carotídeo", "artéria carótida comum", "artéria carótida interna"])
        if st.button("💾 Adicionar Calcificação"):
            for ld in (["Direito", "Esquerdo"] if lado_c == "Bilateral" else [lado_c]):
                st.session_state.calcificacoes_isoladas.append({"lado": ld, "topografia": topo_c})
            st.rerun()
    for idx, c in enumerate(st.session_state.calcificacoes_isoladas):
        st.write(f"• `{idx+1:02d}` Calcificação {c['lado']} - {c['topografia']}")

st.markdown("---")
if st.button("🚀 Gerar Laudo Clínico Completo", use_container_width=True):
    p_aci_dir = any("interna direita" in x['vaso'].lower() for x in st.session_state.lista_placas)
    p_aci_esq = any("interna esquerda" in x['vaso'].lower() for x in st.session_state.lista_placas)
    
    st_dir, suf_dir = utils.obter_texto_hemo_continuo(estado_aci_dir, vps_aci_dir, vcc_dir, p_aci_dir, diretriz_selecionada)
    st_esq, suf_esq = utils.obter_texto_hemo_continuo(estado_aci_esq, vps_aci_esq, vcc_esq, p_aci_esq, diretriz_selecionada)
    
    _, txt_v_dir = utils.avaliar_vertebral(espectro_vert_dir, vps_vert_dir)
    _, txt_v_esq = utils.avaliar_vertebral(espectro_vert_esq, vps_vert_esq)

    doc = Document()
    doc.styles['Normal'].font.name = fonte_doc
    doc.styles['Normal'].font.size = Pt(tamanho_fonte)
    
    def add_p(text, bold_pre=None, align=WD_ALIGN_PARAGRAPH.LEFT, size_add=0, bold=False):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.line_spacing = espacamento_linhas
        p.paragraph_format.space_after = Pt(4)
        if bold_pre:
            r_p = p.add_run(bold_pre)
            r_p.bold = True
            r_p.font.size = Pt(tamanho_fonte + size_add)
        r = p.add_run(text)
        r.bold = bold
        r.font.size = Pt(tamanho_fonte + size_add)

    if nome_clinica:
        add_p("", bold_pre=nome_clinica.upper(), align=WD_ALIGN_PARAGRAPH.CENTER, size_add=2)
    
    add_p("", bold_pre='DUPLEX SCAN DAS ARTÉRIAS CARÓTIDAS E VERTEBRAIS', align=WD_ALIGN_PARAGRAPH.CENTER, size_add=1)
    add_p(f" {nome}", bold_pre="Paciente:")
    add_p(texto_tecnica_final, bold_pre="Técnica: ")
    add_p("", bold_pre='RELATÓRIO', size_add=1)
    
    # LADO DIREITO
    add_p("", bold_pre='LADO DIREITO', bold=True)
    txt_cc_d = f"Artéria carótida comum direita pérvia, com diâmetro e trajeto conservados, apresentando fluxo bifásico anterógrado de baixa resistência. Espessura do complexo médio-intimal: {cmi_dir} mm."
    add_p(txt_cc_d)
    
    txt_b_d = f"Bulbo carotídeo direito pérvio, com diâmetro e trajeto conservados."
    if any("bulbo carotídeo direita" in x['vaso'].lower() for x in st.session_state.lista_placas):
        p_b = [x for x in st.session_state.lista_placas if "bulbo carotídeo direita" in x['vaso'].lower()][0]
        txt_b_d += f" Apresentando na parede uma placa de ateroma {p_b['composicao_texto'].lower()}, medindo {p_b['espessura']} mm de espessura máxima."
    else:
        txt_b_d += " Sem evidências de placas ou alterações estruturais."
    add_p(txt_b_d)
    
    # 🔹 JUNÇÃO REMOVENDO A PALAVRA RÍGIDA ANTERIOR (O UTILS ADICIONA DINAMICAMENTE)
    txt_i_d = f"Artéria carótida interna direita {suf_dir}"
    add_p(txt_i_d)
    add_p(f"Artéria carótida externa direita {ace_dir.lower()}")
    add_p(txt_v_dir)
    
    # LADO ESQUERDO
    add_p("", bold_pre='LADO ESQUERDO', bold=True)
    txt_cc_e = f"Artéria carótida comum esquerda pérvia, com diâmetro e trajeto conservados, apresentando fluxo bifásico anterógrado de baixa resistência. Espessura do complexo médio-intimal: {cmi_esq} mm."
    add_p(txt_cc_e)
    
    txt_b_e = f"Bulbo carotídeo esquerdo pérvio, com diâmetro e trajeto conservados."
    if any("bulbo carotídeo esquerda" in x['vaso'].lower() for x in st.session_state.lista_placas):
        p_e = [x for x in st.session_state.lista_placas if "bulbo carotídeo esquerda" in x['vaso'].lower()][0]
        txt_b_e += f" Apresentando na parede uma placa de ateroma {p_e['composicao_texto'].lower()}, medindo {p_e['espessura']} mm de espessura máxima."
    else:
        txt_b_e += " Sem evidências de placas ou alterações estruturais."
    add_p(txt_b_e)
    
    # 🔹 JUNÇÃO REMOVENDO A PALAVRA RÍGIDA ANTERIOR
    txt_i_e = f"Artéria carótida interna esquerda {suf_esq}"
    add_p(txt_i_e)
    add_p(f"Artéria carótida externa esquerda {ace_esq.lower()}")
    add_p(txt_v_esq)
    
    # IMPRESSÃO DIAGNÓSTICA
    if quebrar_pagina_diag:
        doc.add_page_break()
    add_p("", bold_pre='IMPRESSÃO DIAGNÓSTICA', size_add=1)
    
    tem_achado = False
    if len(st.session_state.lista_placas) > 0:
        for p in st.session_state.lista_placas:
            add_p(f"– Placa de ateroma na {p['vaso'].lower()} ({p['espessura']} mm).")
            tem_achado = True
            
    if not tem_achado:
        add_p("– Artérias carótidas e vertebrais dentro dos limites da normalidade.")

    doc.add_paragraph().paragraph_format.space_before = Pt(30)
    add_p(f"{nome_medico}\nCRM {crm_medico}", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True)
    
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    
    st.success("Laudo gerado com sucesso!")
    st.download_button(label="📥 Baixar Laudo Formatado (.docx)", data=buffer, file_name="Laudo_Vascular_Completo.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
