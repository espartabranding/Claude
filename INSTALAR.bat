@echo off
chcp 65001 >nul 2>&1
title Instalador - Reconhecimento de Escrita Manual
color 1F

echo.
echo ╔══════════════════════════════════════════════════════════════╗
echo ║                                                              ║
echo ║    INSTALADOR - Reconhecimento de Escrita Manual p/ Excel    ║
echo ║                                                              ║
echo ║    Este instalador vai configurar tudo automaticamente.      ║
echo ║    Pode demorar alguns minutos na primeira vez.              ║
echo ║                                                              ║
echo ╚══════════════════════════════════════════════════════════════╝
echo.

:: ---------------------------------------------------------------
:: PASSO 1: Verificar se Python esta instalado
:: ---------------------------------------------------------------
echo [1/4] Verificando Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo  PYTHON NAO ENCONTRADO!
    echo.
    echo  Voce precisa instalar o Python primeiro:
    echo.
    echo  1. Acesse: https://www.python.org/downloads/
    echo  2. Clique no botao amarelo "Download Python"
    echo  3. IMPORTANTE: Na instalacao, marque a opcao
    echo     "Add Python to PATH" (checkbox embaixo)
    echo  4. Clique em "Install Now"
    echo  5. Depois de instalar, rode este INSTALAR.bat de novo
    echo.
    echo  Abrindo o site do Python para voce...
    start https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)
for /f "tokens=2" %%i in ('python --version 2^>^&1') do set PYVER=%%i
echo        Python %PYVER% encontrado!

:: ---------------------------------------------------------------
:: PASSO 2: Criar ambiente virtual (pasta isolada)
:: ---------------------------------------------------------------
echo.
echo [2/4] Criando ambiente isolado...
if not exist "venv" (
    python -m venv venv
    if %errorlevel% neq 0 (
        echo  ERRO ao criar ambiente virtual.
        echo  Tente: python -m pip install virtualenv
        pause
        exit /b 1
    )
    echo        Ambiente criado!
) else (
    echo        Ambiente ja existe, pulando...
)

:: ---------------------------------------------------------------
:: PASSO 3: Instalar dependencias Python
:: ---------------------------------------------------------------
echo.
echo [3/4] Instalando bibliotecas (pode demorar alguns minutos)...
echo        Aguarde...
call venv\Scripts\activate.bat

pip install --upgrade pip >nul 2>&1
pip install streamlit opencv-python Pillow numpy pytesseract easyocr openpyxl scikit-image imutils 2>&1 | findstr /i "successfully installed already satisfied"

if %errorlevel% neq 0 (
    echo.
    echo  Tentando instalacao alternativa...
    pip install streamlit opencv-python-headless Pillow numpy pytesseract easyocr openpyxl scikit-image imutils
)
echo.
echo        Bibliotecas instaladas!

:: ---------------------------------------------------------------
:: PASSO 4: Tesseract OCR
:: ---------------------------------------------------------------
echo.
echo [4/4] Verificando Tesseract OCR...
where tesseract >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo  ============================================================
    echo   ATENCAO: Tesseract OCR nao encontrado!
    echo  ============================================================
    echo.
    echo   Para melhor reconhecimento, instale o Tesseract:
    echo.
    echo   1. Vou abrir a pagina de download para voce
    echo   2. Baixe o arquivo .exe mais recente
    echo   3. Na instalacao, marque "Portuguese" nos idiomas
    echo   4. Mantenha o caminho padrao de instalacao
    echo.
    echo   SEM o Tesseract, a ferramenta ainda funciona usando
    echo   apenas o EasyOCR (selecione "EasyOCR" na interface).
    echo.
    echo   Deseja abrir a pagina de download? (S/N)
    set /p RESP="> "
    if /i "%RESP%"=="S" (
        start https://github.com/UB-Mannheim/tesseract/wiki
    )
) else (
    echo        Tesseract encontrado!
)

:: ---------------------------------------------------------------
:: CONCLUIDO
:: ---------------------------------------------------------------
echo.
echo ╔══════════════════════════════════════════════════════════════╗
echo ║                                                              ║
echo ║    INSTALACAO CONCLUIDA COM SUCESSO!                         ║
echo ║                                                              ║
echo ║    Para usar a ferramenta, de duplo clique em:               ║
echo ║                                                              ║
echo ║          >>> INICIAR.bat <<<                                  ║
echo ║                                                              ║
echo ╚══════════════════════════════════════════════════════════════╝
echo.
pause
