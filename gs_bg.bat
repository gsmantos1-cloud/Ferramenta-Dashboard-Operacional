@echo off
cd /d "%~dp0"

:: Mata processos anteriores
taskkill /F /IM python.exe >nul 2>&1
taskkill /F /IM ngrok.exe >nul 2>&1
timeout /t 1 >nul

:: Inicia Flask em background
start /B "" "C:\Users\l3ti\AppData\Local\Programs\Python\Python312\python.exe" app.py >logs_app.txt 2>&1
timeout /t 3 >nul

:: Inicia ngrok com domínio estático
set DOMINIO=bullfight-lushness-hurler.ngrok-free.dev
start /B "" "%~dp0ngrok.exe" http --domain=%DOMINIO% 5000 --log=stdout >logs_ngrok.txt 2>&1
timeout /t 6 >nul

set LINK=https://%DOMINIO%

:: Abre o link no browser e mostra notificação
powershell -Command "Add-Type -AssemblyName System.Windows.Forms; $n = New-Object System.Windows.Forms.NotifyIcon; $n.Icon = [System.Drawing.SystemIcons]::Information; $n.Visible = $true; $n.ShowBalloonTip(8000, 'GS Mantos', 'Servidor iniciado!`nLink: https://%DOMINIO%`nLogin: gs.operacional / gs00', [System.Windows.Forms.ToolTipIcon]::None); Start-Sleep 9; $n.Dispose()"
start "" "https://%DOMINIO%"
