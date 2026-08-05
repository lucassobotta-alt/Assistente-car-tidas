#!/usr/bin/env bash
cd "$(dirname "$0")"

if ! command -v python3 &> /dev/null; then
    echo "Python 3 nao foi encontrado neste computador."
    echo "Instale em https://www.python.org/downloads/ e tente novamente."
    read -p "Pressione Enter para sair..."
    exit 1
fi

mkdir -p "$HOME/.streamlit"
if [ ! -f "$HOME/.streamlit/credentials.toml" ]; then
    printf '[general]\nemail = ""\n' > "$HOME/.streamlit/credentials.toml"
fi

echo "Verificando dependencias..."
python3 -m pip install --quiet -r requirements.txt

echo "Iniciando o Assistente de Laudos Vasculares..."
python3 -m streamlit run venoso.py --server.headless false
