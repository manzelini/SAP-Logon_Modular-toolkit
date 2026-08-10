@echo off
rem Lanciatore per Build-SAPGlobal.ps1: aggira il blocco "esecuzione script disabilitata"
rem di PowerShell senza bisogno di digitare comandi a mano. Basta fare doppio click qui.
rem Utile per rigenerare/validare SAPUILandscapeGlobal.xml senza fare anche il deploy.
powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Build-SAPGlobal.ps1"
echo.
pause
