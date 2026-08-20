<#
    Find-SAPClient.ps1
    -------------------
    Cerca rapidamente un cliente/sistema dentro SAPUILandscape_sorgenti, senza dover
    sfogliare a mano le 100+ cartelle o ricordarsi il nome esatto della cartella.

    Cerca il termine fornito (case-insensitive, substring) su:
      - nome della cartella cliente (es. "arcaplanet")
      - attributo name del <Service> (es. "ARCAPLANET - PS4 - Production")
      - attributo systemid (SID, es. "PS4")
      - attributo server (host:porta, es. "arcaplanet-ps4.")

    Mostra i risultati in tabella (Cliente | SID | Nome sistema | Server | File) e,
    se e' installato VS Code (comando "code" in PATH), permette di aprire subito il
    file XML del/dei cliente/i trovati, senza passaggi manuali.

    Uso:
      .\Find-SAPClient.ps1 "arcaplanet"
      .\Find-SAPClient.ps1 -Term "PS4"
      .\Find-SAPClient.ps1                    (modalita' interattiva: chiede il termine)

      Se lanciato da Find-SAPClient.cmd, basta doppio click e poi digitare il termine
      quando richiesto (oppure passarlo come argomento al .cmd).

    NOTA: questo script e' di sola LETTURA sui sorgenti - non modifica nulla. L'editing
    del file trovato va fatto a mano in VS Code (o editor di preferenza), poi si lancia
    Build-SAPGlobal.cmd per validare prima di ridistribuire.
#>

param(
    [Parameter(Position = 0)]
    [string]$Term
)

$ErrorActionPreference = "Stop"

$ScriptDir     = Split-Path -Parent $MyInvocation.MyCommand.Path
$SourcesFolder = Join-Path $ScriptDir "SAPUILandscape_sorgenti"

if (-not (Test-Path $SourcesFolder)) {
    Write-Host "ERRORE: non trovo 'SAPUILandscape_sorgenti' accanto a questo script ($ScriptDir)." -ForegroundColor Red
    exit 1
}

if ([string]::IsNullOrWhiteSpace($Term)) {
    $Term = Read-Host "Termine di ricerca (nome cliente, SID, o parte dell'hostname)"
}
if ([string]::IsNullOrWhiteSpace($Term)) {
    Write-Host "ERRORE: nessun termine di ricerca fornito." -ForegroundColor Red
    exit 1
}

Write-Host "=== Ricerca '$Term' in SAPUILandscape_sorgenti ===" -ForegroundColor Cyan

$clientFolders = Get-ChildItem -Path $SourcesFolder -Directory | Sort-Object Name

$results = New-Object System.Collections.Generic.List[object]

foreach ($folder in $clientFolders) {
    $file = Join-Path $folder.FullName "SAPUILandscape.xml"
    if (-not (Test-Path $file)) { continue }

    $folderMatches = $folder.Name -like "*$Term*"

    try {
        [xml]$xmlDoc = Get-Content $file -Raw -Encoding UTF8
    } catch {
        Write-Host "  ATTENZIONE: '$($folder.Name)' non e' un XML valido, saltato." -ForegroundColor Yellow
        continue
    }

    $services = @($xmlDoc.SelectNodes("//Service"))
    $anyServiceMatch = $false

    foreach ($svc in $services) {
        $name   = $svc.GetAttribute("name")
        $sid    = $svc.GetAttribute("systemid")
        $server = $svc.GetAttribute("server")

        $svcMatches = ($name -like "*$Term*") -or ($sid -like "*$Term*") -or ($server -like "*$Term*")

        if ($folderMatches -or $svcMatches) {
            $anyServiceMatch = $true
            $results.Add([PSCustomObject]@{
                Cliente = $folder.Name
                SID     = $sid
                Nome    = $name
                Server  = $server
                File    = $file
            })
        }
    }

    # Cliente il cui nome cartella matcha ma senza nessun <Service> (o nessuno ha
    # matchato sopra) - lo segnaliamo comunque, cosi' il file compare tra i risultati
    # ed e' apribile anche se non ha (ancora) sistemi definiti.
    if ($folderMatches -and -not $anyServiceMatch) {
        $results.Add([PSCustomObject]@{
            Cliente = $folder.Name
            SID     = ""
            Nome    = "(nessun sistema / solo match sul nome cartella)"
            Server  = ""
            File    = $file
        })
    }
}

if ($results.Count -eq 0) {
    Write-Host "Nessun risultato per '$Term'." -ForegroundColor Yellow
    exit 0
}

$results | Format-Table Cliente, SID, Nome, Server -AutoSize | Out-String -Width 200 | Write-Host

$distinctFiles = @($results.File | Sort-Object -Unique)

Write-Host ""
Write-Host "$($results.Count) risultato/i in $($distinctFiles.Count) file cliente diverso/i." -ForegroundColor Green

# --- apertura in VS Code, se disponibile ---
# NOTA: si usa Start-Process con una stringa argomenti quotata a mano (invece di
# "& code $path" o di ProcessStartInfo.ArgumentList) per due motivi:
#   1) il wrapper code.cmd su Windows perde la quotatura sui percorsi con spazi
#      (es. cartelle OneDrive "... S.p.A ..."), troncando il percorso;
#   2) ProcessStartInfo.ArgumentList esiste solo da PowerShell 6+ (.NET Core) ed e'
#      null in Windows PowerShell 5.1 (quella lanciata di default da "powershell.exe"),
#      causando un errore "cannot call a method on a null-valued expression".
# I percorsi Windows non possono contenere il carattere ", quindi avvolgerli tra
# virgolette e' sempre sicuro ed e' compatibile con PowerShell 5.1 e 7+.
function Open-InVSCode {
    param([string[]]$Paths)
    $argString = ($Paths | ForEach-Object { '"' + $_ + '"' }) -join ' '
    Start-Process -FilePath $codeCmd.Source -ArgumentList $argString
}

$codeCmd = Get-Command code -ErrorAction SilentlyContinue
if (-not $codeCmd) {
    Write-Host ""
    Write-Host "VS Code (comando 'code') non trovato in PATH: apri manualmente il/i file sopra." -ForegroundColor Yellow
    exit 0
}

if ($distinctFiles.Count -eq 1) {
    $answer = Read-Host "Aprire il file trovato in VS Code? (S/n)"
    if ($answer -eq "" -or $answer -match "^[sS]") {
        Open-InVSCode -Paths @($distinctFiles[0])
        Write-Host "Aperto in VS Code: $($distinctFiles[0])" -ForegroundColor Green
    }
} else {
    Write-Host ""
    Write-Host "Trovati piu' file cliente diversi. Scegli cosa aprire:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $distinctFiles.Count; $i++) {
        Write-Host "  [$($i+1)] $($distinctFiles[$i])"
    }
    Write-Host "  [A] Aprili tutti"
    Write-Host "  [Invio] Non aprire nulla"
    $choice = Read-Host "Scelta"

    if ($choice -match "^[aA]$") {
        Open-InVSCode -Paths $distinctFiles
        Write-Host "Aperti $($distinctFiles.Count) file in VS Code." -ForegroundColor Green
    } elseif ($choice -match "^\d+$") {
        $idx = [int]$choice - 1
        if ($idx -ge 0 -and $idx -lt $distinctFiles.Count) {
            Open-InVSCode -Paths @($distinctFiles[$idx])
            Write-Host "Aperto in VS Code: $($distinctFiles[$idx])" -ForegroundColor Green
        } else {
            Write-Host "Scelta non valida, nessun file aperto." -ForegroundColor Yellow
        }
    }
}