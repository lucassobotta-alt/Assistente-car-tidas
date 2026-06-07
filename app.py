# utils.py
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# 🔹 QUADRO DE CONDIÇÕES TÉCNICAS EXATAMENTE IGUAL ÀS DIRETRIZES
opcoes_tecnicas = {
    "1. Exame sem limitações técnicas": (
        "Exame realizado em decúbito dorsal, utilizando transdutor linear de alta frequência, "
        "com avaliação bidimensional, mapeamento de fluxo a cores e Doppler pulsado, sem limitações técnicas."
    ),
    "2. Exame com limitação por condições anatômicas desfavoráveis": (
        "Exame realizado em decúbito dorsal, utilizando transdutor linear de alta frequência, "
        "com avaliação bidimensional, mapeamento de fluxo a cores e Doppler pulsado. Devido a condições "
        "anatômicas desfavoráveis para insonação dos vasos cervicais, foi necessária avaliação complementar "
        "com transdutor convexo, o que pode reduzir a sensibilidade para identificação de placas "
        "ateroscleróticas de pequenas dimensões."
    ),
    "3. Exame realizado à beira do leito (UTI)": (
        "Exame realizado à beira do leito em unidade de terapia intensiva, utilizando transdutor linear de "
        "alta frequência, com limitações técnicas inerentes às condições do exame."
    ),
    "4. Exame à beira do leito (UTI) com curativos cervicais": (
        "Exame realizado à beira do leito em unidade de terapia intensiva, utilizando transdutor linear de "
        "alta frequência, com limitações técnicas inerentes às condições do exame and à presença de curativos "
        "cervicais sobre acessos jugulares."
    )
}

def retirar_prefixo_numerico(opcao_texto):
    if ". " in opcao_texto:
        return opcao_texto.split(". ", 1)[1]
    return opcao_texto

def estimate_plaque_rads(opcao_texto):
    for i in range(1, 7):
        if opcao_texto.startswith(f"{i}."):
            return f"Plaque-RADS {i}"
    return None

def obter_texto_hemo_continuo(estado, vps_aci, vcc, tem_placa=False, diretriz="Diretriz SBC 2023"):
    if estado == "Oclusão":
        return "Oclusão", "determinando oclusão completa do vaso, caracterizada por ausência total de fluxo ao estudo Doppler pulsado e mapeamento a cores."
    elif estado == "Suboclusão":
        return "Suboclusão", "determinando suboclusão do vaso, caracterizada por estreitamento luminal severo com padrão de fluxo filiforme ('trickle flow') ao estudo Doppler."
    
    relacao = round(vps_aci / vcc, 2)
    
    # 🔹 CORREÇÃO: SE NÃO HOUVER PLACA E VELOCIDADES FOREM NORMAIS, RETORNA EXATAMENTE O SEU MODELO HABITUAL
    if not tem_placa and ((diretriz == "Diretriz SBC 2023" and vps_aci < 140) or (diretriz != "Diretriz SBC 2023" and vps_aci < 125)):
        return "Normal", "pérvia, com fluxo bifásico anterógrado de baixa resistência, caracterizado por diástole sustentada e velocidades dentro da normalidade, compatível com irrigação de leito encefálico de baixa impedância. Não há sinais de estenose ou turbulência."

    # 🔹 SE HOUVER ALTERAÇÃO OU PLACA, SEGUE O CRITÉRIO SELECIONADO:
    if diretriz == "Diretriz SBC 2023":
        if vps_aci < 140:
            return "Estenose < 50%", f"determinando estenose leve (<50% pelos critérios da Diretriz SBC 2023), caracterizada por velocidade de pico sistólico na artéria carótida interna de {vps_aci} cm/s."
        if vps_aci > 400 or relacao > 5.0:
            return "Estenose > 90%", f"determinando estenose acentuada (>90% pelos critérios da Diretriz SBC 2023), caracterizada por acentuada elevação das velocidades de fluxo com VPS na artéria carótida interna de {vps_aci} cm/s e relação artéria carótida interna / artéria carótida comum de {relacao}."
        elif 230 < vps_aci <= 400 or relacao > 4.0:
            return "Estenose de 70-89%", f"determinando estenose hemodinamicamente significativa (70-89% pelos critérios da Diretriz SBC 2023), caracterizada por VPS na artéria carótida interna de {vps_aci} cm/s e relação artéria carótida interna / artéria carótida comum de {relacao}."
        elif 3.2 <= relacao <= 4.0:
            return "Estenose de 60-69%", f"determinando estenose moderada (60-69% pelos critérios da Diretriz SBC 2023), caracterizada por relação artéria carótida interna / artéria carótida comum de {relacao} and VPS na artéria carótida interna de {vps_aci} cm/s."
        else:
            return "Estenose de 50-59%", f"determinando estenose moderada (50-59% pelos critérios da Diretriz SBC 2023), caracterizada por VPS na artéria carótida interna de {vps_aci} cm/s e relação artéria carótida interna / artéria carótida comum de {relacao}."
            
    else:  # Critérios NASCET
        if vps_aci < 125:
            return "Estenose < 50%", f"determinando estenose leve (<50% pelos critérios do Consenso NASCET), com VPS na artéria carótida interna de {vps_aci} cm/s."
        if vps_aci >= 230 or relacao >= 4.0:
            return "Estenose ≥ 70%", f"determinando estenose severa (≥70% pelos critérios do Consenso NASCET), caracterizada por VPS na artéria carótida interna de {vps_aci} cm/s e relação ACI/ACC de {relacao}."
        else:
            return "Estenose de 50-69%", f"determinando estenose moderada (50-69% pelos critérios do Consenso NASCET), caracterizada por VPS na artéria carótida interna de {vps_aci} cm/s e relação ACI/ACC de {relacao}."

def avaliar_vertebral(espectro, vps_vert):
    if espectro == "Normal (Fluxo Anterógrado)":
        if vps_vert >= 100:
            return "Estenose de Vertebral (>50%)", f"Artéria vertebral apresentando fluxo anterógrado com acentuada elevação focal de velocidades (VPS de {vps_vert} cm/s) e turbulência local, compatível com estenose segmentar superior a 50%."
        else:
            return "Normal", "Artéria vertebral pérvia, com fluxo bifásico anterógrado de baixa resistência, com diástole contínua, compatível com adequada perfusão vertebrobasilar."
    elif espectro == "Hipoplasia": 
        return "Hipoplasia de Vertebral", f"Artéria vertebral apresentando fluxo anterógrado de baixa resistência, porém exibindo calibre reduzido e velocidades proporcionalmente baixas (VPS de {vps_vert} cm/s), compatível com variante anatomofuncional (hipoplasia)."
    elif espectro == "Roubo Latente": 
        return "Sinal de Roubo Latente da Subclávia", "Artéria vertebral apresentando fluxo anterógrado, porém com morfologia de onda alterada devido a uma desaceleração mesosistólica abrupta, sugerindo alteração hemodinâmica inicial por estenose da artéria subclávia proximal ipsilateral."
    elif espectro == "Roubo Parcial (Fluxo Alternante)": 
        return "Sinal de Roubo Parcial da Subclávia", "Artéria vertebral apresentando padrão de fluxo alternante, caracterizado por vetor sistólico retrógrado e vetor diastólico anterógrado, indicando inversão parcial do fluxo por estenose acentuada da artéria subclávia proximal ipsilateral."
    elif espectro == "Roubo Total (Fluxo Retrógrado)": 
        return "Sinal de Roubo Total da Subclávia", "Artéria vertebral apresentando inversão completa e contínua do seu vetor de fluxo, confirmando o fenômeno de roubo de subclávia secundário a oclusão da artéria subclávia proximal ipsilateral."
    return "Alterada", f"Artéria vertebral com alterações inespecíficas do padrão de fluxo. VPS: {vps_vert} cm/s."
