@echo off
rem Lanciatore per Find-SAPClient.ps1: aggira il blocco "esecuzione script disabilitata"
rem di PowerShell senza bisogno di digitare comandi a mano. Basta fare doppio click qui.
rem Puoi anche passare il termine di ricerca come argomento, es:
rem     Find-SAPClient.cmd arcaplanet
powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Find-SAPClient.ps1" %*
echo.
pause
