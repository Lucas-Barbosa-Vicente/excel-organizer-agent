@echo off
title Excel Organizer Agent

echo Iniciando Excel Organizer Agent...
echo.

:: Backend
start "Backend API" cmd /k "cd /d %~dp0backend && python -m uvicorn main:app --port 8000"

:: Aguarda 2 segundos para o backend subir
timeout /t 2 /nobreak >nul

:: Servidor do Add-in
start "Add-in Server" cmd /k "cd /d %~dp0excel-addin && node node_modules/http-server/bin/http-server src/taskpane -p 3001 --cors"

echo.
echo ============================================================
echo  Servidores iniciados!
echo.
echo  Backend:     http://localhost:8000
echo  Add-in:      http://localhost:3001
echo.
echo  Proximos passos:
echo  1. Abra o Excel
echo  2. Inserir ^> Suplementos ^> Meus Suplementos
echo  3. Carregar Suplemento ^> selecione manifest.xml
echo     Caminho: %~dp0excel-addin\manifest.xml
echo ============================================================
echo.
pause
