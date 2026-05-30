@echo off
title GS Mantos — Ferramenta de Pedidos
cd /d "%~dp0"

echo.
echo  ================================
echo   GS MANTOS - Ferramenta de Pedidos
echo  ================================
echo.
echo  Iniciando servidor...

:: Mata processos anteriores se ainda estiverem rodando
taskkill /F /IM python.exe >nul 2>&1
taskkill /F /IM ngrok.exe >nul 2>&1
timeout /t 1 >nul

:: Inicia o app Flask em background
start /B "" "C:\Users\l3ti\AppData\Local\Programs\Python\Python312\python.exe" app.py >logs_app.txt 2>&1

:: Aguarda o Flask subir
timeout /t 3 >nul

:: Inicia o ngrok em background
start /B "" "%~dp0ngrok.exe" http 5000 --log=stdout >logs_ngrok.txt 2>&1

:: Aguarda o ngrok conectar
timeout /t 10 >nul

:: Pega o link público do ngrok via API local
for /f "tokens=*" %%i in ('powershell -Command "(Invoke-RestMethod http://localhost:4040/api/tunnels).tunnels[0].public_url"') do set LINK=%%i

echo.
echo  ================================
echo   TUDO PRONTO!
echo  ================================
echo.
echo   Link para colaboradores:
echo   %LINK%
echo.
echo   Login:  gs.operacional
echo   Senha:  gs00
echo.
echo  ================================
echo.
echo  Mantenha esta janela aberta.
echo  Para encerrar, feche esta janela.
echo.
pause >nul
