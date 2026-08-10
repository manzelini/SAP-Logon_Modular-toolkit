@echo off
rem Lanciatore per Rebuild-SourcesFromMapping.ps1: aggira il blocco "esecuzione
rem script disabilitata" di PowerShell senza bisogno di digitare comandi a mano.
rem Basta fare doppio click qui. Si aspetta che nella stessa cartella ci siano
rem gia' mappa_uuid_sistemi.xlsx e la cartella SAPUILandscape_sorgenti (nomi di
rem default); per usare percorsi diversi, lancia lo script .ps1 direttamente
rem da PowerShell con i parametri -XlsxPath / -ReferenceFolder / -OutFolder.
powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Rebuild-SourcesFromMapping.ps1"
echo.
pause
