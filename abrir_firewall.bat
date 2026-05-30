@echo off
net session >nul 2>&1
if %errorLevel% neq 0 (
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit
)
netsh advfirewall firewall add rule name="GS Mantos Pedidos" dir=in action=allow protocol=TCP localport=5000
echo.
echo Porta 5000 liberada com sucesso!
echo Colaboradores podem acessar em: http://172.20.10.4:5000
echo.
pause
