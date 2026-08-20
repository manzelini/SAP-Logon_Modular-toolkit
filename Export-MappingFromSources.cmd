@echo off
rem Lanciatore per Export-MappingFromSources.ps1: aggira il blocco "esecuzione
rem script disabilitata" di PowerShell senza bisogno di digitare comandi a mano.
rem Basta fare doppio click qui. Si aspetta che nella stessa cartella ci siano
rem gia' SAPUILandscape_sorgenti e (facoltativo, per la riconciliazione)
rem mappa_uuid_sistemi.xlsx (nomi di default); per usare percorsi diversi,
rem lancia lo script .ps1 direttamente da PowerShell con i parametri
rem -SourcesFolder / -ExistingMappingPath / -OutPath.
powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Export-MappingFromSources.ps1"
echo.
pause
