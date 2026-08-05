@echo off
cd /d "%~dp0"

where python >nul 2>nul
if %errorlevel% neq 0 (
    echo Python nao foi encontrado neste computador.
    echo Instale em https://www.python.org/downloads/ ^(marque "Add python.exe to PATH" durante a instalacao^) e tente novamente.
    pause
    exit /b 1
)

if not exist "%USERPROFILE%\.streamlit" mkdir "%USERPROFILE%\.streamlit"
if not exist "%USERPROFILE%\.streamlit\credentials.toml" (
    echo [general] > "%USERPROFILE%\.streamlit\credentials.toml"
    echo email = "" >> "%USERPROFILE%\.streamlit\credentials.toml"
)

echo Verificando dependencias...
python -m pip install --quiet -r requirements.txt

echo Iniciando o Assistente de Laudos Vasculares...
python -m streamlit run venoso.py --server.headless false

pause
