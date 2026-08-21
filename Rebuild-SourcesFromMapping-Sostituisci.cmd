@echo off
rem Variante di Rebuild-SourcesFromMapping.cmd che usa lo switch -ReplaceSources:
rem dopo aver ricostruito i sorgenti da mappa_uuid_sistemi.xlsx, se le verifiche
rem finali sono pulite lo script SOSTITUISCE DA SOLO la cartella
rem SAPUILandscape_sorgenti con quella appena ricostruita, rinominando prima la
rem vecchia con un backup a timestamp (es. SAPUILandscape_sorgenti_20260820_094037) -
rem lo stesso passaggio che altrimenti andrebbe fatto a mano (rinomina sorgenti
rem vecchia + rinomina/sposta sorgenti_REBUILT al suo posto) dopo ogni rebuild.
rem
rem Usa "Rebuild-SourcesFromMapping.cmd" (senza "-Sostituisci") se preferisci
rem invece controllare a mano il contenuto di SAPUILandscape_sorgenti_REBUILT
rem prima di sostituire i sorgenti live.
rem
rem Si aspetta che nella stessa cartella ci siano gia' mappa_uuid_sistemi.xlsx e
rem la cartella SAPUILandscape_sorgenti (nomi di default); per usare percorsi
rem diversi, lancia lo script .ps1 direttamente da PowerShell con i parametri
rem -XlsxPath / -ReferenceFolder / -OutFolder (oltre a -ReplaceSources).
powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Rebuild-SourcesFromMapping.ps1" -ReplaceSources
echo.
pause
