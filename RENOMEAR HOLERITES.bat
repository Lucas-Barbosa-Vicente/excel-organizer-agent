@echo off
chcp 65001 > nul
title Renomeador de Holerites
color 0A

echo.
echo ============================================
echo      RENOMEADOR DE HOLERITES
echo ============================================
echo.

:: Verifica se Python esta instalado
python --version > nul 2>&1
if errorlevel 1 (
    color 0C
    echo ERRO: Python nao encontrado no computador.
    echo.
    echo Para instalar o Python, acesse:
    echo https://www.python.org/downloads/
    echo.
    echo Marque a opcao "Add Python to PATH" durante a instalacao.
    echo.
    pause
    exit /b 1
)

:: Instala dependencias se necessario
echo Verificando dependencias...
python -c "import pandas, openpyxl" > nul 2>&1
if errorlevel 1 (
    echo Instalando dependencias necessarias, aguarde...
    pip install pandas openpyxl --quiet
    echo.
)

:: Verifica se a planilha existe
if not exist "%~dp0Holerites renomear.xlsx" (
    color 0C
    echo ERRO: Planilha nao encontrada!
    echo.
    echo Coloque o arquivo "Holerites renomear.xlsx"
    echo na mesma pasta deste programa.
    echo.
    pause
    exit /b 1
)

:: Verifica se a pasta HOLERITES existe
if not exist "%~dp0HOLERITES\" (
    color 0C
    echo ERRO: Pasta HOLERITES nao encontrada!
    echo.
    echo Crie uma pasta chamada "HOLERITES" na mesma
    echo pasta deste programa e coloque os PDFs dentro.
    echo.
    pause
    exit /b 1
)

echo Tudo pronto! Iniciando...
echo.

:: Executa o script
python "%~dp0renomear_holerites.py"

echo.
if errorlevel 1 (
    color 0C
    echo Ocorreu um erro. Verifique as mensagens acima.
) else (
    color 0A
)

echo.
pause
