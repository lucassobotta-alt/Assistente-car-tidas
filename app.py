# app.py
import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO

# --- DICIONÁRIOS DE CONFIGURAÇÃO ---
opcoes_tecnicas = {
    "1. Exame sem limitações técnicas": "Exame realizado em decúbito dorsal, utilizando transdutor linear de alta frequência, com avaliação bidimensional, mapeamento de fluxo a cores e Doppler pulsado, sem limitações técnicas.",
    "2. Exame com limitação por condições anatômicas desfavoráveis": "Exame realizado em decúbito dorsal, utilizando transdutor linear de alta frequência, com avaliação bidimensional, mapeamento de fluxo a cores e Doppler pulsado. Devido a condições anatômicas desfavoráveis para insonação dos vasos cervicais, foi necessária avaliação complementar com transdutor convexo, o que pode reduzir a sensibilidade para identificação de placas ateroscleróticas de pequenas dimensões.",
    "3. Exame realizado à beira do leito (UTI)": "Exame realizado à beira do leito em unidade de terapia intensiva, utilizando transdutor linear de alta frequência, com limitações técnicas inerentes às condições do exame.",
    "4. Exame à beira do leito (UTI) com curativos cervicais": "Exame realizado à beira do leito em unidade de terapia intensiva, utilizando transdutor linear de alta frequência, com limitações técnicas inerentes às condições do exame e à presença de curativos cervicais sobre acessos jugulares."
}

def obter_texto_hemo_continuo(estado, vps_aci, vcc, tem_placa=False, diretriz="Diretriz SBC 2023"):
    if estado == "Oclusão":
        return "ocludida, determinando ausência total de fluxo ao estudo Doppler pulsado e mapeamento a cores."
    if estado == "Suboclusão":
        return "subocludida, caracterizada por estreitamento luminal severo com padrão de fluxo filiforme ('trickle flow') ao estudo Doppler."
    
    relacao = round(vps_aci / vcc, 2)
    limite_vps = 140 if diretriz == "Diretriz SBC 2023" else 125
    
    if not tem_placa and vps_aci < limite_vps:
        return "com fluxo bifásico anterógrado de baixa resistência, caracterizado por diástole sustentada e velocidades dentro da normalidade, compatível com irrigação de leito encefálico de baixa impedância. Não há sinais de estenose ou turbulência."

    if diretriz == "Diretriz SBC 2023":
        if vps_aci < 140:
            return f"apresentando na parede uma placa de ateroma, determinando estenose leve (<50% pelos critérios da Diretriz SBC 2023), caracterizada por velocidade de pico sistólico de {vps_aci} cm/s."
        if vps_aci > 400 or relacao > 5.0:
            return f"determinando estenose acentuada (>90% pelos critérios da Diretriz SBC 2023), caracterizada por acentuada elevação das velocidades de fluxo com VPS de {vps_aci} cm/s e relação ACI/ACC de {relacao}."
        if 230 < vps_aci <= 400 or relacao > 4.0:
            return f"determinando estenose hemodinamicamente significativa (70-89% pelos critérios da Diretriz SBC 2023). Ao estudo Doppler, observa-se VPS de {vps_aci} cm/s e relação ACI/ACC de {relacao}."
        if 3.2 <= relacao <= 4.0:
            return f"determinando estenose moderada (60-69% pelos critérios da Diretriz SBC 2023), caracterizada por relação ACI/ACC de {relacao} e VPS de {vps_aci} cm/s."
        return f"determinando estenose moderada (50-59% pelos critérios da Diretriz SBC 2023), caracterizada por VPS de {vps_aci} cm/s e relação ACI/ACC de {relacao}."
    else:
        if vps_aci < 125:
            return f"apresentando na parede uma placa de ateroma, determinando estenose leve (<50% pelos critérios do Consenso NASCET), com VPS de {vps_aci} cm/s."
        if vps_aci >= 230 or relacao >= 4.0:
            return f"determinando estenose severa (≥70% pelos critérios do Consenso NASCET), caracterizada por VPS de {vps_aci} cm/s e relação ACI/ACC de {relacao}."
        return f"determinando estenose moderada (50-59% pelos critérios do Consenso NASCET), caracterizada por VPS de {vps_aci} cm/s e relação ACI/ACC de {relacao}."

def avaliar_vertebral(espectro, vps_vert):
    if espectro == "Normal (Fluxo Anterógrado)":
        if vps_vert >= 100:
            return f"Artéria vertebral apresentando fluxo anterógrado com acentuada elevação focal de velocidades (VPS de {vps_vert} cm/s) e turbulência local, compatível com estenose segmentar superior a 50%."
        return "Artéria vertebral pérvia, com fluxo bifásico anterógrado de baixa resistência, com diástole contínua, compatível com adequada perfusão vertebrobasilar."
    if espectro == "Hipoplasia": 
        return f"Artéria vertebral apresentando fluxo anterógrado de baixa resistência, porém exibindo calibre reduzido e velocidades proporcionalmente baixas (VPS de {vps_vert} cm/s), compatível com variante anatomofuncional (hipoplasia)."
    if espectro == "Roubo Latente": 
        return "Artéria vertebral apresentando fluxo anterógrado, porém com morfologia de onda alterada devido a uma desaceleração mesosistólica abrupta, sugerindo alteração hemodinâmica inicial por estenose da artéria subclávia proximal ipsilateral."
    if espectro == "Roubo Parcial (Fluxo Alternante)": 
        return "Artéria vertebral apresentando padrão de fluxo alternante, caracterizado por vetor sistólico retrógrado e vetor担当 diastólico anterógrado, indicando inversão parcial do fluxo por estenose acentuada da artéria subclávia proximal ipsilateral."
    if espectro == "Roubo Total (Fluxo Retrógrado)": 
        return "Artéria vertebral apresentando inversão completa e contínua do seu vetor de fluxo, confirmando o fenômeno de roubo de subclávia secundário a oclusão da artéria subclávia proximal ipsilateral."
    return f"Artéria vertebral com alterações inespecíficas do padrão de fluxo. VPS: {vps_vert} cm/s."

# --- INITIALIZATION OF STATE ---
if "reset_trigger" in st.session_state and st.session_state.reset_trigger:
    st.session_state.clear()
    st.session_state.reset_trigger = False

chaves_estado = ['lista_placas', 'lesoes_incipientes', 'calcificacoes_isoladas', 'lesoes_nao_ateromatosas']
for k in chaves_estado:
    if k not in st.session_state: st.session_state[k] = []

st.set_page_config(page_title="Laudo Neurovascular Avançado", layout="wide")
st.title("⚕️ Assistente de Laudos Vascular Completo")

with st.sidebar:
    st.markdown("## ⚙️ Painel de Controle")
    diretriz_selecionada = st.radio("Diretriz de Referência:", ["Diretriz SBC 2023", "Consenso Clássico NASCET"])
    fonte_doc = st.selectbox("Fonte:", ["Arial", "Calibri", "Times New Roman"])
    tamanho_fonte = st.slider("Tamanho Texto (pt):", 10, 14, 11)
    espacamento_linhas = st.slider("Espaçamento:", 1.0, 1.5, 1.15)
    quebrar_pagina_diag = st.toggle("Separar Diagnóstico em Nova Página", value=False)
    nome_clinica = st.text_input("Nome da Clínica:")
    nome_medico = st.text_input("Médico:", "Lucas Santos Guimarães")
    crm_medico = st.text_input("CRM / RQE:", "4061")
    if st.button("🔄 Resetar Parâmetros", use_container_width=True):
        st.session_state.reset_trigger = True
        st.rerun()

col_id1, col_id2 = st.columns(2)
with col_id1: nome = st.text_input("Nome do Paciente", "Paciente Exemplo")
with col_id2: texto_tecnica_final = opcoes_tecnicas[st.selectbox("Condições Técnicas:", list(opcoes_tecnicas.keys()))]

st.markdown("### 📊 Parâmetros Hemodinâmicos")
col_dir, col_esq = st.columns(2)

with col_dir:
    st.subheader("LADO DIREITO")
    cmi_dir = st.number_input("CMI ACC Direita (mm)", 0.0, 5.0, 0.4, 0.1)
    estado_aci_dir = st.selectbox("Estado ACI Direita", ["Pérvia (Calcular por Velocidade)", "Suboclusão", "Oclusão"])
    vps_aci_dir = st.number_input("VPS ACI Direita (cm/s)", 0.0, value=90.0, step=5.0)
    vcc_dir = st.number_input("VPS ACC Direita (cm/s)", 1.0, value=60.0, step=5.0)
    ace_dir = st.selectbox("ACE Direita", ["Com padrão espectral de alta resistência, compatível com perfusão de leitos musculares extracranianos.", "Alterada / Estenose hemodinâmica"])
    espectro_vert_dir = st.selectbox("Espectro AV Direita", ["Normal (Fluxo Anterógrado)", "Hipoplasia", "Roubo Latente", "Roubo Parcial (Fluxo Alternante)", "Roubo Total (Fluxo Retrógrado)"])
    vps_vert_dir = st.number_input("VPS AV Direita (cm/s)", 0.0, value=30.0, step=5.0)

with col_esq:
    st.subheader("LADO ESQUERDO")
    cmi_esq = st.number_input("CMI ACC Esquerda (mm)", 0.0, 5.0, 0.4, 0.1)
    estado_aci_esq = st.selectbox("Estado ACI Esquerda", ["Pérvia (Calcular por Velocidade)", "Suboclusão", "Oclusão"])
    vps_aci_esq = st.number_input("VPS ACI Esquerda (cm/s)", 0.0, value=100.0, step=5.0)
    vcc_esq = st.number_input("VPS ACC Esquerda (cm/s)", 1.0, value=60.0, step=5.0)
    ace_esq = st.selectbox("ACE Esquerda", ["Com padrão espectral de alta resistência, compatível com perfusão de leitos musculares extracranianos.", "Alterada / Estenose hemodinâmica"])
    espectro_vert_esq = st.selectbox("Espectro AV Esquerda", ["Normal (Fluxo Anterógrado)", "Hipoplasia", "Roubo Latente", "Roubo Parcial (Fluxo Alternante)", "Roubo Total (Fluxo Retrógrado)"])
    vps_vert_esq = st.number_input("VPS AV Esquerda (cm/s)", 0.0, value=30.0, step=5.0)

# --- EXPANDERS COMPLEXOS (DE VOLTA E COMPLETOS) ---
lista_arterias = [
    "Artéria carótida comum direita", "Bulbo carotídeo direito", "Artéria carótida interna direita", "Artéria carótida externa direita",
    "Artéria carótida comum esquerda", "Bulbo carotídeo esquerdo", "Artéria carótida interna esquerda", "Artéria carótida externa esquerda"
]

with st.expander("🔎 Mapeamento de Placas Ateroscleróticas (≥ 2.0 mm)"):
    col_p1, col_p2, col_p3 = st.columns(3)
    with col_p1: vaso_sel = st.selectbox("Artéria (Placa):", lista_arterias)
    with col_p2: local_sel = st.selectbox("Segmento (Placa):", ["terço proximal", "terço médio", "terço distal", "segmento total/bifurcação"])
    with col_p3: composicao = st.selectbox("Composição:", ["Placa calcificada", "Placa uniformemente ecogênica", "Placa predominantemente ecogênica", "Placa com componente anecogênico", "Placa predominantemente anecogênica", "Placa uniformemente anecogênica"])
    espessura = st.number_input("Espessura Máxima da Placa (mm):", 2.0, 20.0, 2.5, 0.1)
    if st.button("💾 Gravar Placa Ateromatosa"):
        st.session_state.lista_placas.append({"vaso": vaso_sel, "localizacao": local_sel, "composicao_texto": composicao.lower(), "espessura": espessura})
        st.rerun()
    for idx, p in enumerate(st.session_state.lista_placas):
        st.write(f"• Placa na `{p['vaso']}` ({p['localizacao']}) — {p['composicao_texto']}, {p['espessura']} mm")

with st.expander("🌱 Lesões Incipientes (< 2.0 mm)"):
    col_li1, col_li2 = st.columns(2)
    with col_li1: vaso_li = st.selectbox("Artéria (Incipiente):", lista_arterias)
    with col_li2: parede_li = st.selectbox("Parede (Incipiente):", ["posterior", "anterior", "lateral", "medial"])
    if st.button("💾 Gravar Lesão Incipiente"):
        st.session_state.lesoes_incipientes.append({"vaso": vaso_li, "parede": parede_li})
        st.rerun()
    for li in st.session_state.lesoes_incipientes:
        st.write(f"• Espessamento incipiente na parede {li['parede']} da `{li['vaso']}`.")

with st.expander("💎 Calcificações Isoladas (Sem repercussão ou estenose)"):
    vaso_calc = st.selectbox("Artéria (Calcificação):", lista_arterias)
    if st.button("💾 Gravar Calcificação Isolada"):
        st.session_state.calcificacoes_isoladas.append({"vaso": vaso_calc})
        st.rerun()
    for c in st.session_state.calcificacoes_isoladas:
        st.write(f"• Calcificação parietal isolada na `{c['vaso']}`.")

with st.expander("⚠️ Lesões Não Ateromatosas (Disfunções Estruturais)"):
    col_lna1, col_lna2 = st.columns(2)
    with col_lna1: vaso_lna = st.selectbox("Artéria afetada:", lista_arterias + ["Artéria vertebral direita", "Artéria vertebral esquerda"])
    with col_lna2: tipo_lna = st.selectbox("Tipo de Alteração:", ["Kinking (angulação severa)", "Coiling (enovelamento)", "Tortuosidade acentuada", "Sinais compatíveis com dissecção espontânea", "Sinais compatíveis com displasia fibromuscular"])
    if st.button("💾 Gravar Lesão Não Ateromatosa"):
        st.session_state.lesoes_nao_ateromatosas.append({"vaso": vaso_lna, "tipo": tipo_lna})
        st.rerun()
    for lna in st.session_state.lesoes_nao_ateromatosas:
        st.write(f"• `{lna['tipo']}` identificada na `{lna['vaso']}`.")

# --- PROCESSADOR E IMPRESSÃO DO LAUDO ---
if st.button("🚀 Gerar Laudo Clínico Completo", use_container_width=True):
    p_aci_dir = any("interna direita" in x['vaso'].lower() for x in st.session_state.lista_placas)
    p_aci_esq = any("interna esquerda" in x['vaso'].lower() for x in st.session_state.lista_placas)
    
    suf_dir = obter_texto_hemo_continuo(estado_aci_dir, vps_aci_dir, vcc_dir, p_aci_dir, diretriz_selecionada)
    suf_esq = obter_texto_hemo_continuo(estado_aci_esq, vps_aci_esq, vcc_esq, p_aci_esq, diretriz_selecionada)
    
    txt_v_dir = avaliar_vertebral(espectro_vert_dir, vps_vert_dir)
    txt_v_esq = avaliar_vertebral(espectro_vert_esq, vps_vert_esq)

    doc = Document()
    doc.styles['Normal'].font.name = fonte_doc
    doc.styles['Normal'].font.size = Pt(tamanho_fonte)
    
    def add_p(text, bold_pre=None, align=WD_ALIGN_PARAGRAPH.LEFT, size_add=0):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.line_spacing = espacamento_linhas
        p.paragraph_format.space_after = Pt(4)
        if bold_pre:
            r_p = p.add_run(bold_pre)
            r_p.bold = True
            r_p.font.size = Pt(tamanho_fonte + size_add)
        r = p.add_run(text)
        r.font.size = Pt(tamanho_fonte + size_add)

    if nome_clinica: add_p("", bold_pre=nome_clinica.upper(), align=WD_ALIGN_PARAGRAPH.CENTER, size_add=2)
    add_p("", bold_pre='DUPLEX SCAN DAS ARTÉRIAS CARÓTIDAS E VERTEBRAIS', align=WD_ALIGN_PARAGRAPH.CENTER, size_add=1)
    add_p(f" {nome}", bold_pre="Paciente:")
    add_p(f" {texto_tecnica_final}", bold_pre="Técnica:")
    add_p("", bold_pre='RELATÓRIO', size_add=1)
    
    # --- CONSTRUÇÃO DO BLOCO DE TEXTO CORRIDO ---
    def montar_descricao_vaso(nome_vaso, base_pérvia, cmi=None):
        txt = f"{nome_vaso} pérvia"
        if cmi: txt += f", com diâmetro e trajeto conservados, apresentando fluxo bifásico anterógrado de baixa resistência. Espessura do complexo médio-intimal: {cmi} mm."
        else: txt += ", com diâmetro e trajeto conservados."
        
        # Agrega achados estruturais
        placas = [x for x in st.session_state.lista_placas if x['vaso'].lower() == nome_vaso.lower()]
        incipientes = [x for x in st.session_state.lesoes_incipientes if x['vaso'].lower() == nome_vaso.lower()]
        calcs = [x for x in st.session_state.calcificacoes_isoladas if x['vaso'].lower() == nome_vaso.lower()]
        lnas = [x for x in st.session_state.lesoes_nao_ateromatosas if x['vaso'].lower() == nome_vaso.lower()]
        
        if placas:
            txt += f" Apresentando na parede uma {placas[0]['composicao_texto']} em {placas[0]['localizacao']}, medindo {placas[0]['espessura']} mm de espessura máxima."
        if incipientes:
            txt += f" Nota-se espessamento parietal incipiente em parede {incipientes[0]['parede']} (< 2.0 mm), sem repercussão hemodinâmica."
        if calcs:
            txt += " Presença de foco de calcificação parietal isolada, sem determinar estenoses."
        if lnas:
            txt += f" Observa-se {lnas[0]['tipo']}."
        if not placas and not incipientes and not calcs and not lnas and not cmi:
            txt += " Sem evidências de placas ou alterações estruturais."
        return txt

    # LADO DIREITO
    add_p("", bold_pre='LADO DIREITO')
    add_p(montar_descricao_vaso("Artéria carótida comum direita", True, cmi_dir))
    add_p(montar_descricao_vaso("Bulbo carotídeo direito", True))
    add_p(f"Artéria carótida interna direita pérvia, {suf_dir}")
    add_p(f"Artéria carótida externa direita {ace_dir.lower()}")
    add_p(txt_v_dir)
    
    # LADO ESQUERDO
    add_p("", bold_pre='LADO ESQUERDO')
    add_p(montar_descricao_vaso("Artéria carótida comum esquerda", True, cmi_esq))
    add_p(montar_descricao_vaso("Bulbo carotídeo esquerdo", True))
    add_p(f"Artéria carótida interna esquerda pérvia, {suf_esq}")
    add_p(f"Artéria carótida externa esquerda {ace_esq.lower()}")
    add_p(txt_v_esq)
    
    # CONCLUSÃO
    if quebrar_pagina_diag: doc.add_page_break()
    add_p("", bold_pre='IMPRESSÃO DIAGNÓSTICA', size_add=1)
    
    tem_achados = (st.session_state.lista_placas or st.session_state.lesoes_incipientes or 
                   st.session_state.calcificacoes_isoladas or st.session_state.lesoes_nao_ateromatosas)
    
    if not tem_achados:
        add_p("– Artérias carótidas e vertebrais dentro dos limites da normalidade.")
    else:
        for p in st.session_state.lista_placas:
            add_p(f"– Placa de ateroma na {p['vaso'].lower()} ({p['espessura']} mm).")
        for li in st.session_state.lesoes_incipientes:
            add_p(f"– Alterações ateromatosas incipientes na parede {li['parede']} da {li['vaso'].lower()}.")
        for c in st.session_state.calcificacoes_isoladas:
            add_p(f"– Calcificação parietal aterosclerótica isolada na {c['vaso'].lower()}.")
        for lna in st.session_state.lesoes_nao_ateromatosas:
            add_p(f"– {lna['tipo']} na {lna['vaso'].lower()}.")

    doc.add_paragraph().paragraph_format.space_before = Pt(25)
    add_p(f"{nome_medico}\nCRM {crm_medico}", align=WD_ALIGN_PARAGRAPH.CENTER)
    
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    st.success("Laudo completo gerado com sucesso!")
    st.download_button("📥 Baixar Laudo Formatado (.docx)", buf, "Laudo_Vascular.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
