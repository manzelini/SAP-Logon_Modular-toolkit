<#
    Deploy-SAPLandscape.ps1
    ------------------------
    Installa/aggiorna la configurazione modulare SAP Logon (SAPUILandscape_globale)
    sulla macchina dell'utente corrente. Da eseguire UNA VOLTA per ogni utente/PC.

    Cosa fa:
      1. Rigenera SAPUILandscapeGlobal.xml da ZERO leggendo TUTTA la struttura
         corrente di "SAPUILandscape_sorgenti" (richiama Build-SAPGlobal.ps1, deve
         trovarsi nella STESSA cartella di questo script) - cosi' ogni deploy porta
         sempre l'ultima versione dei sorgenti, anche se sono stati modificati a
         mano o con New-SAPClient.ps1 senza rigenerare manualmente il globale.
      2. Fa un backup con timestamp del SAPUILandscape.xml personale esistente
         (se presente) in %appdata%\SAP\Common.
      3. Copia (sovrascrivendo) la cartella "SAPUILandscape_globale" (appena
         rigenerata al passo 1) dentro %appdata%\SAP\Common.
      4. Copia la cartella "_AVVALE_INTERNAL" (che si trova dentro "SAPUILandscape_sorgenti")
         DENTRO la cartella "SAPUILandscape_globale" appena distribuita, perche' l'<Include>
         presente in SAPUILandscapeGlobal.xml si aspetta di trovarla proprio li' accanto.
      5. Ricalcola tutti i link <Include> dentro SAPUILandscapeGlobal.xml in base al
         percorso reale dell'utente corrente (funziona con qualsiasi nome utente Windows,
         non solo "Administrator").
      6. Scrive un SAPUILandscape.xml personale "vuoto" che rimanda tutto il contenuto
         al file globale via <Include> - senza duplicare nulla.

    Uso: tasto destro sul file -> "Esegui con PowerShell".
    Se PowerShell blocca lo script (execution policy), aprire PowerShell ed eseguire:
        powershell -ExecutionPolicy Bypass -File .\Deploy-SAPLandscape.ps1
#>

$ErrorActionPreference = "Stop"

$FolderName     = "SAPUILandscape_globale"
$SourcesName    = "SAPUILandscape_sorgenti"
$InternalName   = "_AVVALE_INTERNAL"
$ScriptDir      = Split-Path -Parent $MyInvocation.MyCommand.Path
$SourceFolder   = Join-Path $ScriptDir $FolderName
$SourcesFolder  = Join-Path $ScriptDir $SourcesName
$InternalSource = Join-Path $SourcesFolder $InternalName
$BuildScript    = Join-Path $ScriptDir "Build-SAPGlobal.ps1"

$SapCommon    = Join-Path $env:APPDATA "SAP\Common"
$DestFolder   = Join-Path $SapCommon $FolderName
$InternalDest = Join-Path $DestFolder $InternalName
$PersonalFile = Join-Path $SapCommon "SAPUILandscape.xml"
$GlobalFile   = Join-Path $DestFolder "SAPUILandscapeGlobal.xml"

Write-Host "=== Deploy configurazione SAP Logon modulare ===" -ForegroundColor Cyan

if (-not (Test-Path $SourcesFolder)) {
    Write-Host "ERRORE: non trovo la cartella '$SourcesName' accanto a questo script ($ScriptDir)." -ForegroundColor Red
    Write-Host "Assicurati di aver estratto TUTTO lo zip nella stessa posizione." -ForegroundColor Red
    exit 1
}

if (-not (Test-Path $InternalSource)) {
    Write-Host "ERRORE: non trovo la cartella '$InternalName' dentro '$SourcesName' accanto a questo script ($ScriptDir)." -ForegroundColor Red
    Write-Host "Assicurati di aver estratto TUTTO lo zip (script + entrambe le cartelle) nella stessa posizione." -ForegroundColor Red
    exit 1
}

if (-not (Test-Path $BuildScript)) {
    Write-Host "ERRORE: non trovo 'Build-SAPGlobal.ps1' accanto a questo script ($ScriptDir)." -ForegroundColor Red
    exit 1
}

# 1) Rigenera SAPUILandscapeGlobal.xml da zero, leggendo tutta SAPUILandscape_sorgenti,
#    cosi' il deploy porta sempre la versione piu' aggiornata (anche se qualcuno ha
#    modificato i sorgenti a mano senza rigenerare manualmente il globale).
#    Eseguito come processo separato (non con l'operatore &): se Build-SAPGlobal.ps1
#    incontra un errore e chiama "exit", deve terminare solo se stesso, non anche
#    questo script.
Write-Host "Rigenero SAPUILandscapeGlobal.xml dai sorgenti correnti..." -ForegroundColor Cyan
$buildProc = Start-Process -FilePath "powershell.exe" `
    -ArgumentList @("-NoLogo", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$BuildScript`"") `
    -NoNewWindow -Wait -PassThru
if ($buildProc.ExitCode -ne 0) {
    Write-Host "ERRORE: la rigenerazione di SAPUILandscapeGlobal.xml e' fallita (codice $($buildProc.ExitCode)). Deploy interrotto." -ForegroundColor Red
    exit 1
}
if (-not (Test-Path $SourceFolder)) {
    Write-Host "ERRORE: la rigenerazione sembra non aver prodotto la cartella '$FolderName'. Deploy interrotto." -ForegroundColor Red
    exit 1
}

New-Item -ItemType Directory -Path $SapCommon -Force | Out-Null

# 2) Backup del file personale esistente
if (Test-Path $PersonalFile) {
    $stamp  = Get-Date -Format "yyyyMMdd_HHmmss"
    $backup = Join-Path $SapCommon "SAPUILandscape_backup_$stamp.xml"
    Copy-Item $PersonalFile $backup -Force
    Write-Host "Backup creato: $backup"
} else {
    Write-Host "Nessun SAPUILandscape.xml personale preesistente, salto il backup."
}

# 3) Copia (sovrascrivendo) la cartella globale appena rigenerata
if (Test-Path $DestFolder) {
    Remove-Item $DestFolder -Recurse -Force
}
Copy-Item $SourceFolder $DestFolder -Recurse -Force
Write-Host "Copiata la cartella '$FolderName' in: $DestFolder"

# 4) Copia "_AVVALE_INTERNAL" (dai sorgenti) DENTRO la cartella globale appena distribuita:
#    l'<Include> in SAPUILandscapeGlobal.xml punta a "<cartella globale>\_AVVALE_INTERNAL\...",
#    quindi deve trovarsi li' e non solo dentro SAPUILandscape_sorgenti.
Copy-Item $InternalSource $InternalDest -Recurse -Force
Write-Host "Copiata la cartella '$InternalName' in: $InternalDest"

# 5) Ricalcola i link <Include> nel file globale in base al percorso reale dell'utente
#    (usiamo System.Uri per ottenere una file:// URL correttamente con %20 per gli spazi, ecc.)
$NewBaseUrl = ([uri]$DestFolder).AbsoluteUri.TrimEnd('/')

$globalContent = Get-Content $GlobalFile -Raw -Encoding UTF8
# individua il prefisso attualmente usato in tutti gli url="file:///...%SAPUILandscape_globale"
if ($globalContent -match 'url="(file:///[^"]*SAPUILandscape_globale)') {
    $OldBaseUrl = $Matches[1]
    $globalContent = $globalContent.Replace($OldBaseUrl, $NewBaseUrl)
    Set-Content -Path $GlobalFile -Value $globalContent -Encoding UTF8
    Write-Host "Link <Include> ricalcolati: $OldBaseUrl  ->  $NewBaseUrl"
} else {
    Write-Host "ATTENZIONE: non ho trovato un prefisso url riconoscibile in $GlobalFile, controlla a mano." -ForegroundColor Yellow
}

# 6) Scrive il file personale "vuoto" che rimanda tutto al global
$GlobalFileUrl = ([uri]$GlobalFile).AbsoluteUri
$PersonalXml = @"
<?xml version="1.0"?>
<Landscape updated="$(Get-Date -Format s)Z" version="1" generator="SAP GUI for Windows v8000.1.17.155">
	<Workspaces>
		<Workspace uuid="bc81f604-823c-48e1-8fd6-f7a704ab9e5b" name="Local" expanded="1" hidden="0" />
	</Workspaces>
	<Services />
	<Messageservers />
	<Routers />
	<Includes>
		<Include url="$GlobalFileUrl" index="0" />
	</Includes>
</Landscape>
"@
Set-Content -Path $PersonalFile -Value $PersonalXml -Encoding UTF8
Write-Host "Scritto il file personale: $PersonalFile"

Write-Host ""
Write-Host "Fatto. Apri SAP Logon per verificare che carichi senza errori." -ForegroundColor Green
