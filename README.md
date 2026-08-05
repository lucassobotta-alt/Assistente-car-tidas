# Assistente de Laudos Vasculares

Aplicativo Streamlit para geração de laudos médicos em formato Word (.docx) para exames de ultrassonografia vascular.

## Apps disponíveis

- **Duplex Scan de Carótidas e Vertebrais** — avaliação do sistema arterial cervical
- **Duplex Scan Venoso de MMII** — mapeamento do sistema venoso superficial e profundo dos membros inferiores
- **Duplex Scan Arterial de MMII** — mapeamento arterial dos membros inferiores, da artéria femoral comum à fibular

## Como executar localmente

```bash
pip install -r requirements.txt
streamlit run venoso.py
```

## Uso offline (sem depender do Streamlit Cloud)

Como o Streamlit roda como um servidor local, o app funciona inteiramente sem internet depois de instalado — basta ter o [Python 3](https://www.python.org/downloads/) instalado no computador (no instalador do Windows, marque a opção "Add python.exe to PATH").

Baixe/clone este repositório e dê duplo clique no script correspondente ao seu sistema:

- **Windows:** `iniciar_app.bat`
- **macOS:** `iniciar_app.command`
- **Linux:** `iniciar_app.sh` (ou `./iniciar_app.sh` pelo terminal)

Na primeira execução ele instala as dependências automaticamente (só precisa de internet nessa etapa); nas próximas vezes abre direto. O app sobe em `http://localhost:8501` no seu navegador padrão. Para encerrar, feche a janela do terminal que abriu junto.

## Deploy

Hospedado no [Streamlit Community Cloud](https://share.streamlit.io).
