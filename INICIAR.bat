@echo off
chcp 65001 >nul 2>&1
title Reconhecimento de Escrita Manual para Excel
color 1F

echo.
echo ╔══════════════════════════════════════════════════════════════╗
echo ║                                                              ║
echo ║    Reconhecimento de Escrita Manual para Excel                ║
echo ║    Iniciando a interface...                                  ║
echo ║                                                              ║
echo ╚══════════════════════════════════════════════════════════════╝
echo.

:: Verificar se a instalacao foi feita
if not exist "venv" (
    echo  ERRO: Instalacao nao encontrada!
    echo.
    echo  Voce precisa rodar o INSTALAR.bat primeiro.
    echo  De duplo clique em INSTALAR.bat e depois volte aqui.
    echo.
    pause
    exit /b 1
)

:: Ativar ambiente virtual
call venv\Scripts\activate.bat

:: Verificar se streamlit esta instalado
pip show streamlit >nul 2>&1
if %errorlevel% neq 0 (
    echo  ERRO: Bibliotecas nao instaladas!
    echo  Rode o INSTALAR.bat novamente.
    echo.
    pause
    exit /b 1
)

echo  Abrindo no navegador...
echo.
echo  ============================================================
echo   A interface vai abrir no seu navegador automaticamente.
echo.
echo   Se nao abrir, acesse: http://localhost:8501
echo.
echo   Para FECHAR a ferramenta, feche esta janela preta
echo   ou pressione Ctrl+C aqui.
echo  ============================================================
echo.

:: Iniciar Streamlit
streamlit run app_web.py --server.headless false --browser.gatherUsageStats false

pause
