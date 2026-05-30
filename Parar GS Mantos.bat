@echo off
taskkill /F /IM python.exe >nul 2>&1
taskkill /F /IM ngrok.exe >nul 2>&1
powershell -Command "Add-Type -AssemblyName System.Windows.Forms; $n = New-Object System.Windows.Forms.NotifyIcon; $n.Icon = [System.Drawing.SystemIcons]::Information; $n.Visible = $true; $n.ShowBalloonTip(4000, 'GS Mantos', 'Servidor encerrado.', [System.Windows.Forms.ToolTipIcon]::None); Start-Sleep 5; $n.Dispose()"
