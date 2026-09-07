@echo off
:: Adds Windows Firewall rule to allow teammates to access the PHL5 dashboard
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Requesting admin rights...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

echo Adding firewall rule for PHL5 Compliance Dashboard (port 8501)...
netsh advfirewall firewall delete rule name="PHL5 Compliance Dashboard" >nul 2>&1
netsh advfirewall firewall add rule name="PHL5 Compliance Dashboard" dir=in action=allow protocol=TCP localport=8501

echo.
echo Done! Your team can now access the dashboard at:
echo   http://10.249.194.69:8501
echo.
echo NOTE: this IP can change if this machine reconnects to VPN/WiFi.
echo If the link stops working, re-run "ipconfig" for the current IPv4
echo address and update the line above (or ask Code Puppy to do it).
echo.
pause
