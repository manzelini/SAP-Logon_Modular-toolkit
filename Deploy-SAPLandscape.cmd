@echo off
rem Lanciatore per Deploy-SAPLandscape.ps1: aggira il blocco "esecuzione script disabilitata"
rem di PowerShell senza bisogno di digitare comandi a mano. Basta fare doppio click qui.
powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Deploy-SAPLandscape.ps1"
echo.
pause
