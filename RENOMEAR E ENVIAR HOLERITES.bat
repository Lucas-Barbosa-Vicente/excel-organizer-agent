@echo off
chcp 65001 > nul
title Envio de Holerites
color 0A

echo.
echo ============================================
echo      ENVIO AUTOMATICO DE HOLERITES
echo ============================================
echo.

python --version > nul 2>&1
if errorlevel 1 (
    color 0C
    echo ERRO: Python nao encontrado no computador.
    echo Instale em: https://www.python.org/downloads/
    echo Marque "Add Python to PATH" durante a instalacao.
    echo.
    pause
    exit /b 1
)

echo Verificando dependencias...
python -c "import pandas, openpyxl" > nul 2>&1
if errorlevel 1 (
    echo Instalando dependencias, aguarde...
    pip install pandas openpyxl --quiet
    echo.
)

if not exist "%~dp0.env" (
    color 0C
    echo ERRO: arquivo .env nao encontrado.
    echo Crie o arquivo .env com EMAIL_USUARIO e EMAIL_SENHA.
    echo Consulte o arquivo .env.example para o formato correto.
    echo.
    pause
    exit /b 1
)

for %%F in ("%~dp0Emails funcionarios.xlsx") do (
    if not exist "%%F" (
        color 0C
        echo ERRO: "Emails funcionarios.xlsx" nao encontrado.
        echo Coloque o arquivo na mesma pasta deste programa.
        echo.
        pause
        exit /b 1
    )
)

for %%F in ("%~dp0Holerites renomear.xlsx") do (
    if not exist "%%F" (
        color 0C
        echo ERRO: "Holerites renomear.xlsx" nao encontrado.
        echo Coloque o arquivo na mesma pasta deste programa.
        echo.
        pause
        exit /b 1
    )
)

echo Iniciando pipeline...
echo.

python "%~dp0pipeline.py"

echo.
if errorlevel 1 (
    color 0C
    echo Ocorreu um erro. Verifique as mensagens acima.
) else (
    color 0A
    echo Processo concluido.
)

echo.
pause
