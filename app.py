def obter_texto_hemo_continuo(estado, vps_aci, vcc, tem_placa=False, diretriz="Diretriz SBC 2023"):
    """
    Retorna estritamente o sufixo hemodinâmico SEM a palavra inicial 'pérvia, '
    para evitar a duplicação no relatório gerado.
    """
    if estado == "Oclusão":
        return "ocludida, determinando ausência total de fluxo ao estudo Doppler pulsado e mapeamento a cores."
    if estado == "Suboclusão":
        return "subocludida, caracterizada por estreitamento luminal severo com padrão de fluxo filiforme ('trickle flow') ao estudo Doppler."
    
    relacao = round(vps_aci / vcc, 2)
    limite_vps = 140 if diretriz == "Diretriz SBC 2023" else 125
    
    # Caso o vaso esteja normal e sem placas
    if not tem_placa and vps_aci < limite_vps:
        return "com fluxo bifásico anterógrado de baixa resistência, caracterizado por diástole sustentada e velocidades dentro da normalidade, compatível com irrigação de leito encefálico de baixa impedância. Não há sinais de estenose ou turbulência."

    # Lógica baseada na Diretriz SBC 2023
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
    
    # Lógica baseada no Consenso Clássico NASCET
    else:
        if vps_aci < 125:
            return f"apresentando na parede uma placa de ateroma, determinando estenose leve (<50% pelos critérios do Consenso NASCET), com VPS de {vps_aci} cm/s."
        if vps_aci >= 230 or relacao >= 4.0:
            return f"determinando estenose severa (≥70% pelos critérios do Consenso NASCET), caracterizada por VPS de {vps_aci} cm/s e relação ACI/ACC de {relacao}."
        return f"determinando estenose moderada (50-69% pelos critérios do Consenso NASCET), caracterizada por VPS de {vps_aci} cm/s e relação ACI/ACC de {relacao}."
