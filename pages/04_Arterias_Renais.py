# pages/03_Arterias_Renais.py
import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO

st.title("🫘 Assistente de Laudos: Duplex Scan de Aorta e Artérias Renais")

# --- SIDEBAR ---
with st.sidebar:
    st.markdown("## ⚙️ Painel de Controle")

    st.markdown("### 📚 Referências Científicas")
    st.markdown(
        "📄 <a href='https://www.arquivosonline.com.br/2019/113/4/' target='_blank'>"
        "SBC – Vascular Ultrasound Statement (Arq Bras Cardiol, 2019)</a><br>"
        "📄 <a href='https://www.escardio.org/Guidelines/Clinical-Practice-Guidelines/Peripheral-Arterial-and-Aortic-Diseases' target='_blank'>"
        "ESC 2024 Guidelines – Peripheral Arterial and Aortic Diseases</a>",
        unsafe_allow_html=True
    )

    st.markdown("---")
    st.markdown("### 📝 Formatação Externa (.docx)")
    fonte_doc = st.selectbox("Família da Fonte:", ["Arial", "Calibri", "Times New Roman"], key="ar_fonte")
    tamanho_fonte = st.slider("Tamanho do Texto (pt):", 10, 14, 11, key="ar_tam")
    espacamento_linhas = st.slider("Espaçamento entre Linhas:", 1.0, 1.5, 1.15, step=0.05, key="ar_esp")
    quebrar_pagina_diag = st.toggle("Separar Impressão em Nova Página", value=False, key="ar_qpd")
    modo_saida = st.radio(
        "Modo de saída:",
        ["Somente DOCX", "Somente Visualização", "Visualização + DOCX"],
        index=2, key="ar_saida"
    )

    st.markdown("---")
    st.markdown("### ✍️ Identidade & Assinatura")
    nome_clinica = st.text_input("Cabeçalho / Nome da Clínica:", placeholder="Ex: Instituto de Diagnóstico Vascular", key="ar_clinica")
    nome_medico = st.text_input("Nome do Médico:", "Lucas Santos Guimarães", key="ar_medico")
    colcrm1, colcrm2 = st.columns([2, 1])
    with colcrm1:
        crm_medico = st.text_input("CRM:", "4061", key="ar_crm")
    with colcrm2:
        crm_uf = st.selectbox("UF", ["AC","AL","AP","AM","BA","CE","DF","ES","GO","MA","MT","MS","MG","PA","PB","PR","PE","PI","RJ","RN","RS","RO","RR","SC","SP","SE","TO"], index=25, key="ar_uf")
    rqe_medico = st.text_input("RQE:", "", key="ar_rqe")
    incluir_assinatura = st.toggle("Incluir assinatura / carimbo no laudo", value=True, key="ar_assin")

    st.markdown("---")
    _AR_SIDEBAR_KEYS = {
        "ar_fonte", "ar_tam", "ar_esp", "ar_qpd", "ar_saida",
        "ar_clinica", "ar_medico", "ar_crm", "ar_uf", "ar_rqe",
        "ar_assin", "ar_rodape_tog", "ar_rodape_url",
    }
    if st.button("🔄 Resetar Todos os Parâmetros", use_container_width=True, type="secondary"):
        for _k in [k for k in st.session_state if k not in _AR_SIDEBAR_KEYS]:
            del st.session_state[_k]
        st.toast("🔄 Todos os dados clínicos foram limpos!")
        st.rerun()

    st.markdown("---")
    st.markdown("### 🔗 Rodapé do Documento")
    incluir_rodape_link = st.toggle("Incluir nota de rodapé com link do sistema", value=False, key="ar_rodape_tog")
    rodape_url = ""
    if incluir_rodape_link:
        rodape_url = st.text_input("URL do sistema:", placeholder="Ex: https://seu-app.streamlit.app", key="ar_rodape_url")

# --- PACIENTE ---
nome_paciente = st.text_input("Nome do Paciente:", "", key="ar_paciente")

st.markdown("---")

# =========================================================
# 1. AORTA ABDOMINAL
# =========================================================
st.markdown("### 1. Aorta Abdominal")

col_ao1, col_ao2, col_ao3 = st.columns(3)
with col_ao1:
    aorta_calibre = st.text_input("Calibre máximo (cm):", "1.8", key="ao_calibre")
with col_ao2:
    aorta_vps = st.text_input("VPS aórtica (cm/s):", "70", key="ao_vps")
with col_ao3:
    aorta_placas = st.selectbox("Placas ateromatosas parietais:", ["Ausentes", "Presentes"], key="ao_placas")

aorta_aneurisma = st.toggle("Dilatação aneurismática / dissecção?", value=False, key="ao_aneu")
aorta_aneurisma_desc = ""
if aorta_aneurisma:
    aorta_aneurisma_desc = st.text_area(
        "Descrição da alteração (aneurisma/dissecção):",
        "Dilatação aneurismática da aorta abdominal infrarrenal, com diâmetro máximo de ___ cm.",
        key="ao_aneu_desc"
    )

# Alerta VPS aórtica fora do range de referência da RRA
try:
    _vps_ao_float = float(aorta_vps.replace(",", "."))
    if _vps_ao_float < 40 or _vps_ao_float > 100:
        st.warning(
            f"⚠️ **VPS aórtica fora do intervalo de referência ({aorta_vps} cm/s):** "
            "valores < 40 ou > 100 cm/s podem comprometer a validade da Razão Renal-Aórtica (RRA). "
            "Interprete a RRA com cautela neste exame."
        )
except ValueError:
    _vps_ao_float = None

st.markdown("---")

# =========================================================
# 2. AVALIAÇÃO MORFOLÓGICA RENAL
# =========================================================
st.markdown("### 2. Avaliação Morfológica Renal")

_ECOTEXTURA_OPTS = [
    "normais",
    "ecotextura homogênea com diferenciação corticomedular preservada",
    "ecotextura aumentada com diferenciação corticomedular reduzida",
    "ecotextura heterogênea com diferenciação corticomedular reduzida",
    "ecotextura heterogênea com diferenciação corticomedular abolida",
]

morfo_rins = {}
tabs_rins = st.tabs(["🔴 Rim Direito", "🔵 Rim Esquerdo"])

for ri, lado_rim in enumerate(["Direito", "Esquerdo"]):
    with tabs_rins[ri]:
        c1, c2 = st.columns(2)
        with c1:
            comp_rim = st.text_input(f"Comprimento do rim (cm):", "10.5", key=f"rim_comp_{lado_rim}")
        with c2:
            cortical = st.text_input(f"Espessura cortical (cm):", "1.0", key=f"rim_cort_{lado_rim}")

        ecotextura = st.selectbox(
            "Ecotextura e diferenciação corticomedular:",
            _ECOTEXTURA_OPTS,
            key=f"rim_eco_{lado_rim}"
        )

        c3, c4 = st.columns(2)
        with c3:
            calculos = st.toggle("Cálculos identificados?", key=f"rim_calc_{lado_rim}")
            calc_desc = ""
            if calculos:
                calc_desc = st.text_input("Descrição (tamanho, localização):", "cálculo de até ___ mm no polo inferior", key=f"rim_calc_desc_{lado_rim}")
        with c4:
            hidronefrose = st.toggle("Hidronefrose?", key=f"rim_hidro_{lado_rim}")
            hidro_desc = ""
            if hidronefrose:
                hidro_desc = st.selectbox(
                    "Grau de hidronefrose:",
                    ["leve (grau I)", "moderada (grau II)", "acentuada (grau III)"],
                    key=f"rim_hidro_grau_{lado_rim}"
                )

        morfo_rins[lado_rim] = {
            "comp": comp_rim,
            "cortical": cortical,
            "ecotextura": ecotextura,
            "calculos": calculos,
            "calc_desc": calc_desc,
            "hidronefrose": hidronefrose,
            "hidro_desc": hidro_desc,
        }

        # Alerta de rim pequeno
        try:
            _comp_float = float(comp_rim.replace(",", "."))
            if _comp_float < 9.0:
                st.warning(f"⚠️ Rim {lado_rim.lower()} com comprimento {comp_rim} cm < 9,0 cm — considerar redução volumétrica renal.")
            if _comp_float > 13.0:
                st.info(f"ℹ️ Rim {lado_rim.lower()} com comprimento {comp_rim} cm > 13,0 cm — considerar aumento renal.")
        except ValueError:
            pass

st.markdown("---")

# =========================================================
# 3. ARTÉRIAS RENAIS
# =========================================================
st.markdown("### 3. Artérias Renais")

_ID_ARTERY_OPTS = [
    "desde sua origem",
    "somente no segmento proximal",
    "somente no segmento distal",
    "não identificada (limitação técnica)",
]
_CALIBRE_TRAJ_OPTS = [
    "habituais",
    "calibre reduzido",
    "calibre aumentado",
    "trajeto irregular / tortuoso",
]
_FLUXO_INTRA_OPTS = [
    "baixa resistência",
    "alta resistência",
    "padrão tardus-parvus",
    "ausente (oclusão)",
]

def _safe_float(s):
    try:
        return float(str(s).replace(",", ".").strip())
    except (ValueError, TypeError):
        return None

def classificar_estenose(vps, vdf, rra, ta):
    """Retorna (grau_texto, cor) baseado nos critérios SBC/ESC."""
    if vps is None:
        return "Não calculado", "gray"
    if vps == 0:
        return "Oclusão (ausência de fluxo)", "red"
    if vps < 200:
        return "Normal (sem estenose significativa)", "green"
    # VPS ≥ 200
    if rra is not None and rra >= 3.5:
        if vdf is not None and vdf >= 150:
            if ta is not None and ta >= 70:
                return "Estenose ≥ 80% (grave)", "red"
            return "Estenose ≥ 60–80%", "orange"
        return "Estenose ≥ 60%", "orange"
    # VPS ≥ 200 mas RRA < 3.5
    return "Estenose < 60% (hemodinamicamente não significativa)", "blue"

dados_arterias = {}
tabs_art = st.tabs(["🔴 Artéria Renal Direita", "🔵 Artéria Renal Esquerda"])

for ai, lado_art in enumerate(["Direita", "Esquerda"]):
    with tabs_art[ai]:
        id_art = st.selectbox(
            "Identificação da artéria renal:",
            _ID_ARTERY_OPTS,
            key=f"art_id_{lado_art}"
        )

        nao_identificada = id_art == "não identificada (limitação técnica)"
        limitacao_desc = ""
        if nao_identificada:
            limitacao_desc = st.text_input(
                "Motivo da limitação técnica:",
                "meteorismo intestinal",
                key=f"art_lim_{lado_art}"
            )
            dados_arterias[lado_art] = {
                "id": id_art,
                "limitacao_desc": limitacao_desc,
                "nao_identificada": True,
            }
            st.info("ℹ️ A ausência de visualização isoladamente não deve ser interpretada como oclusão.")
        else:
            calibre_traj = st.selectbox(
                "Calibre e trajeto:",
                _CALIBRE_TRAJ_OPTS,
                key=f"art_cal_{lado_art}"
            )

            col_v1, col_v2, col_v3 = st.columns(3)
            with col_v1:
                vps_art = st.text_input("VPS (cm/s):", "120", key=f"art_vps_{lado_art}")
            with col_v2:
                vdf_art = st.text_input("VDF (cm/s):", "40", key=f"art_vdf_{lado_art}")
            with col_v3:
                # RRA auto-calculada — sempre recomputada a cada render
                _vps_f = _safe_float(vps_art)
                _rra_auto = ""
                if _vps_f is not None and _vps_ao_float and _vps_ao_float > 0:
                    _rra_auto = f"{_vps_f / _vps_ao_float:.2f}"
                _key_rra = f"art_rra_{lado_art}"
                st.session_state[_key_rra] = _rra_auto if _rra_auto else "—"
                rra_art = st.text_input(
                    "Razão renal-aórtica (RRA):",
                    key=_key_rra,
                    disabled=True,
                    help="Calculado automaticamente como VPS renal / VPS aórtica."
                )

            turbulencia = st.toggle(
                "Aceleração focal / turbulência pós-estenótica?",
                key=f"art_turb_{lado_art}"
            )
            turb_desc = ""
            if turbulencia:
                turb_desc = st.text_input(
                    "Localização da aceleração/turbulência:",
                    "na origem",
                    key=f"art_turb_desc_{lado_art}"
                )

            st.markdown("**Fluxo Intrarrenal**")
            col_i1, col_i2, col_i3, col_i4 = st.columns(4)
            with col_i1:
                fluxo_intra_padrao = st.selectbox("Padrão:", _FLUXO_INTRA_OPTS, key=f"art_fi_pad_{lado_art}")
            with col_i2:
                vps_intra = st.text_input("VPS intrarrenal (cm/s):", "30", key=f"art_fi_vps_{lado_art}")
            with col_i3:
                ir_intra = st.text_input("IR:", "0.65", key=f"art_fi_ir_{lado_art}")
            with col_i4:
                ta_intra = st.text_input("Tempo de aceleração (ms):", "50", key=f"art_fi_ta_{lado_art}")

            # --- Classificação automática ---
            _vps_f = _safe_float(vps_art)
            _vdf_f = _safe_float(vdf_art)
            _rra_f = _safe_float(rra_art)
            _ta_f = _safe_float(ta_intra)

            _grau, _cor = classificar_estenose(_vps_f, _vdf_f, _rra_f, _ta_f)

            _cor_map = {"green": "🟢", "blue": "🔵", "orange": "🟠", "red": "🔴", "gray": "⚪"}
            st.markdown(
                f"**{_cor_map.get(_cor, '⚪')} Classificação automática — Artéria Renal {lado_art}:** {_grau}"
            )

            if _cor in ("orange", "red"):
                st.warning(
                    f"⚠️ **Critérios compatíveis com estenose hemodinamicamente significativa — Artéria Renal {lado_art}.**  \n"
                    "Considere avaliação complementar (AngioTC, AngioRM ou arteriografia)."
                )

            dados_arterias[lado_art] = {
                "id": id_art,
                "calibre_traj": calibre_traj,
                "vps": vps_art,
                "vdf": vdf_art,
                "rra": rra_art,
                "turbulencia": turbulencia,
                "turb_desc": turb_desc,
                "fluxo_intra_padrao": fluxo_intra_padrao,
                "vps_intra": vps_intra,
                "ir_intra": ir_intra,
                "ta_intra": ta_intra,
                "grau": _grau,
                "cor": _cor,
                "nao_identificada": False,
            }

st.markdown("---")

# =========================================================
# 4. OBSERVAÇÃO TÉCNICA OPCIONAL
# =========================================================
st.markdown("### 4. Observação Técnica")
incluir_obs_tecnica = st.toggle("Incluir nota sobre limitações técnicas do exame?", value=False, key="ar_obs_tog")
obs_tecnica_texto = ""
if incluir_obs_tecnica:
    _LIMITACOES_OPTS = [
        "meteorismo intestinal",
        "obesidade",
        "respiração inadequada",
        "calcificações vasculares extensas",
        "dificuldade de demonstração das origens das artérias renais",
    ]
    limitacoes_sel = st.multiselect(
        "Fatores limitantes:",
        _LIMITACOES_OPTS,
        key="ar_lim_multi"
    )
    if limitacoes_sel:
        obs_tecnica_texto = "A avaliação foi limitada por " + ", ".join(limitacoes_sel) + ". Os segmentos com avaliação prejudicada estão indicados no relatório."

st.markdown("---")

# =========================================================
# 5. GERAÇÃO DO LAUDO
# =========================================================
def construir_laudo_renais():
    doc = Document()
    doc.styles['Normal'].font.name = fonte_doc
    doc.styles['Normal'].font.size = Pt(tamanho_fonte)

    def add_p(text, bold_pre=None, align=WD_ALIGN_PARAGRAPH.LEFT,
              space_before=0, space_after=4, bullet=False):
        p = doc.add_paragraph(style='List Bullet' if bullet else 'Normal')
        p.alignment = align
        p.paragraph_format.line_spacing = espacamento_linhas
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        if bold_pre:
            r_p = p.add_run(bold_pre)
            r_p.bold = True
        p.add_run(text)

    if nome_clinica.strip():
        p_cl = doc.add_paragraph()
        p_cl.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_cl = p_cl.add_run(nome_clinica.upper())
        r_cl.bold = True
        r_cl.font.name = fonte_doc
        r_cl.font.size = Pt(tamanho_fonte + 2)
        doc.add_paragraph().paragraph_format.space_after = Pt(8)

    add_p("DE AORTA E ARTÉRIAS RENAIS", bold_pre="DUPLEX SCAN ", space_after=12)
    if nome_paciente.strip():
        add_p(f" {nome_paciente}", bold_pre="Paciente:")
    add_p("TÉCNICA", space_before=12, space_after=6)
    add_p(
        "Exame realizado com transdutor convexo multifrequencial, utilizando recursos de modo B, "
        "Doppler colorido e Doppler espectral, com avaliação morfológica e hemodinâmica da aorta "
        "abdominal, das artérias renais e dos ramos intrarrenais.",
        space_after=12
    )

    add_p("RELATÓRIO", space_before=6, space_after=8)

    # — AORTA —
    add_p("AORTA ABDOMINAL", space_after=4)
    _placas_txt = "com placas ateromatosas parietais" if aorta_placas == "Presentes" else "sem placas ateromatosas parietais"
    add_p(f"Aorta abdominal com calibre de até {aorta_calibre} cm, {_placas_txt}.")
    add_p(f"VPS na aorta abdominal de {aorta_vps} cm/s.")
    if aorta_aneurisma and aorta_aneurisma_desc.strip():
        add_p(aorta_aneurisma_desc.strip())
    else:
        add_p("Não se observam dilatações aneurismáticas, dissecções ou outras alterações morfológicas significativas da aorta abdominal ao método.")

    # — MORFOLOGIA RENAL —
    add_p("AVALIAÇÃO MORFOLÓGICA RENAL", space_before=10, space_after=4)
    for lado_rim, mrm in morfo_rins.items():
        _calc_txt = (
            f"Identificado(s) {mrm['calc_desc']}."
            if mrm["calculos"] and mrm["calc_desc"]
            else "Não foram observados cálculos com mais de 5 mm."
        )
        _hidro_txt = (
            f"Identificada hidronefrose {mrm['hidro_desc']}."
            if mrm["hidronefrose"] and mrm["hidro_desc"]
            else "Ausência de hidronefrose."
        )
        _eco = mrm["ecotextura"]
        _eco_txt = _eco if _eco != "normais" else "normais"
        add_p(
            f"Rim {lado_rim.lower()} medindo {mrm['comp']} cm, com espessura cortical de "
            f"{mrm['cortical']} cm, apresentando ecotextura e diferenciação corticomedular "
            f"{_eco_txt}. {_calc_txt} {_hidro_txt}"
        )

    # — ARTÉRIAS RENAIS —
    conclusoes = []
    for lado_art, da in dados_arterias.items():
        add_p(f"ARTÉRIA RENAL {lado_art.upper()}", space_before=10, space_after=4)

        if da.get("nao_identificada"):
            add_p(
                f"Artéria renal {lado_art.lower()} não identificada ao método, "
                f"possivelmente em razão de {da.get('limitacao_desc', 'limitação técnica')}. "
                "A ausência de visualização isoladamente não deve ser interpretada como oclusão."
            )
            conclusoes.append(f"Artéria renal {lado_art.lower()} não avaliada por limitação técnica ({da.get('limitacao_desc','')}).")
        else:
            _id_txt = da["id"]
            _cal_txt = da["calibre_traj"]
            add_p(
                f"Artéria renal {lado_art.lower()} identificada {_id_txt}, "
                f"apresentando calibre e trajeto {_cal_txt}."
            )
            add_p(f"VPS: {da['vps']} cm/s. VDF: {da['vdf']} cm/s. Razão renal-aórtica (RRA): {da['rra']}.")

            if da["turbulencia"]:
                add_p(
                    f"Observada aceleração focal do fluxo {da['turb_desc']}, com turbulência pós-estenótica ao mapeamento colorido."
                )
            else:
                add_p("Não se observam acelerações focais significativas do fluxo, alterações espectrais ou turbulência pós-estenótica.")

            _fi_pad = da["fluxo_intra_padrao"]
            if _fi_pad == "ausente (oclusão)":
                add_p(
                    f"Fluxo intrarrenal ausente ao Doppler colorido e espectral, "
                    "sugerindo oclusão arterial."
                )
                conclusoes.append(f"Sinais de oclusão da artéria renal {lado_art.lower()}.")
            else:
                add_p(
                    f"Fluxo intrarrenal com padrão de {_fi_pad}, apresentando VPS de "
                    f"{da['vps_intra']} cm/s, IR de {da['ir_intra']} e tempo de aceleração de "
                    f"{da['ta_intra']} ms."
                )
                # Conclusão baseada na classificação
                _grau = da["grau"]
                _cor = da["cor"]
                if _cor == "green":
                    pass  # Normal — não adiciona conclusão de estenose
                elif _cor == "blue":
                    conclusoes.append(
                        f"Artéria renal {lado_art.lower()}: sinais compatíveis com estenose hemodinamicamente não significativa (< 60%)."
                    )
                elif _cor == "orange":
                    conclusoes.append(
                        f"Artéria renal {lado_art.lower()}: sinais compatíveis com estenose hemodinamicamente significativa ({_grau})."
                    )
                elif _cor == "red":
                    conclusoes.append(
                        f"Artéria renal {lado_art.lower()}: sinais de estenose grave — {_grau}."
                    )

    # — OBSERVAÇÃO TÉCNICA —
    if obs_tecnica_texto:
        add_p("OBSERVAÇÃO TÉCNICA", space_before=10, space_after=4)
        add_p(obs_tecnica_texto)

    # — IMPRESSÃO DIAGNÓSTICA —
    if quebrar_pagina_diag:
        doc.add_page_break()
    add_p("⸻", space_after=12)
    add_p("IMPRESSÃO DIAGNÓSTICA", space_after=6)

    # Aorta
    if not aorta_aneurisma:
        add_p("Aorta abdominal sem alterações hemodinamicamente significativas ao estudo Doppler.", bullet=True)
    else:
        add_p(f"Aorta abdominal com alteração morfológica: {aorta_aneurisma_desc.strip()}", bullet=True)

    # Artérias renais
    _todas_normais = all(
        (not da.get("nao_identificada")) and da.get("cor") == "green"
        for da in dados_arterias.values()
    )
    if _todas_normais:
        add_p("Ausência de sinais diretos ou indiretos de estenose hemodinamicamente significativa nas artérias renais.", bullet=True)
    else:
        for c in conclusoes:
            add_p(c, bullet=True)
        if not conclusoes:
            add_p("Ausência de sinais diretos ou indiretos de estenose hemodinamicamente significativa nas artérias renais.", bullet=True)

    # Assinatura
    if incluir_assinatura and (nome_medico or crm_medico):
        doc.add_paragraph().paragraph_format.space_before = Pt(25)
        p_assin = doc.add_paragraph()
        p_assin.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_assin.paragraph_format.line_spacing = espacamento_linhas
        if nome_medico:
            r = p_assin.add_run(f"{nome_medico}\n")
            r.bold = True
        if crm_medico:
            p_assin.add_run(f"CRM-{crm_uf} {crm_medico}\n")
        if rqe_medico.strip():
            p_assin.add_run(f"RQE {rqe_medico}")

    if incluir_rodape_link and rodape_url.strip():
        p_rod = doc.add_paragraph()
        p_rod.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_rod.paragraph_format.space_before = Pt(30)
        r_rod = p_rod.add_run(f"Laudo gerado com suporte do sistema: {rodape_url.strip()}")
        r_rod.italic = True
        r_rod.font.size = Pt(max(tamanho_fonte - 2, 8))

    return doc


if st.button("🚀 Gerar Laudo de Artérias Renais", use_container_width=True):
    doc = construir_laudo_renais()
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)

    st.success("Laudo gerado com sucesso!")
    if modo_saida in ["Somente Visualização", "Visualização + DOCX"]:
        texto_viz = "\n".join(p.text for p in doc.paragraphs if p.text.strip())
        st.markdown("## 👁️ Visualização do Laudo")
        st.text_area("Laudo Gerado", value=texto_viz, height=600)
    if modo_saida in ["Somente DOCX", "Visualização + DOCX"]:
        st.download_button(
            "📥 Baixar Laudo de Artérias Renais (.docx)",
            buf,
            "Laudo_Arterias_Renais.docx",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
