<#
    Restore-SAPLandscape.ps1
    -------------------------
    Ripristina il SAPUILandscape.xml PERSONALE (%appdata%\SAP\Common) a partire da uno
    dei backup con timestamp creati automaticamente da Deploy-SAPLandscape.ps1 al suo
    passo 2 (file "SAPUILandscape_backup_yyyyMMdd_HHmmss.xml" nella stessa cartella).

    Cosa fa:
      1. Cerca in %appdata%\SAP\Common tutti i file "SAPUILandscape_backup_*.xml".
      2. Se non specifichi un backup, te li elenca (piu' recente per primo) e te ne
         chiede uno; premendo INVIO senza digitare nulla si ripristina il piu' recente.
      3. Prima di sovrascrivere, salva un backup di sicurezza del file personale
         ATTUALE (prefisso "SAPUILandscape_prerestore_") - cosi' anche il restore e'
         reversibile.
      4. Copia il backup scelto sopra SAPUILandscape.xml.

    IMPORTANTE - cosa NON ripristina:
      Deploy-SAPLandscape.ps1 fa il backup SOLO del file personale (SAPUILandscape.xml,
      il file "Local" che rimanda tutto al globale via <Include>). La cartella
      "SAPUILandscape_globale" (SAPUILandscapeGlobal.xml + _AVVALE_INTERNAL, cioe' la
      configurazione di TUTTI i clienti) viene sovrascritta ad ogni deploy SENZA backup.
      Questo script quindi NON puo' riportare indietro il contenuto condiviso globale:
      se serve, l'unica via e' recuperare una versione precedente di
      "SAPUILandscape_sorgenti" (es. da uno zip/versione precedente) e rilanciare
      Deploy-SAPLandscape.ps1 da li'.

    Uso:
      tasto destro sul file -> "Esegui con PowerShell" (oppure il lanciatore
      Restore-SAPLandscape.cmd, che aggira il blocco execution policy).

      Parametri opzionali da riga di comando:
        -BackupFile <nome o percorso completo>   Ripristina un backup specifico.
        -Latest                                   Ripristina senza chiedere conferma
                                                    il backup piu' recente trovato.
        -ListOnly                                 Elenca solo i backup disponibili,
                                                    senza ripristinare nulla.

      Se PowerShell blocca lo script (execution policy), aprire PowerShell ed eseguire:
          powershell -ExecutionPolicy Bypass -File .\Restore-SAPLandscape.ps1
#>

[CmdletBinding()]
param(
    [string]$BackupFile,
    [switch]$Latest,
    [switch]$ListOnly
)

$ErrorActionPreference = "Stop"

$SapCommon    = Join-Path $env:APPDATA "SAP\Common"
$PersonalFile = Join-Path $SapCommon "SAPUILandscape.xml"
$BackupPrefix = "SAPUILandscape_backup_"
$PreRestorePrefix = "SAPUILandscape_prerestore_"

Write-Host "=== Ripristino SAPUILandscape.xml personale da backup ===" -ForegroundColor Cyan

if (-not (Test-Path $SapCommon)) {
    Write-Host "ERRORE: non trovo la cartella '$SapCommon'. Nessun deploy risulta mai eseguito su questo PC." -ForegroundColor Red
    exit 1
}

# --- 1) trova tutti i backup disponibili, piu' recente per primo ---
$backups = Get-ChildItem -Path $SapCommon -Filter "$BackupPrefix*.xml" -File -ErrorAction SilentlyContinue |
    Sort-Object Name -Descending

if (-not $backups -or $backups.Count -eq 0) {
    Write-Host "ERRORE: nessun backup trovato in '$SapCommon' (pattern '$BackupPrefix*.xml')." -ForegroundColor Red
    Write-Host "I backup vengono creati automaticamente da Deploy-SAPLandscape.ps1 ad ogni deploy: se non ce n'e' nessuno, significa che il deploy non e' mai stato eseguito su questo PC, oppure non esisteva ancora un file personale da salvare la prima volta." -ForegroundColor Yellow
    exit 1
}

function Format-BackupLabel($file) {
    # Nome atteso: SAPUILandscape_backup_yyyyMMdd_HHmmss.xml
    if ($file.BaseName -match '_(\d{8})_(\d{6})$') {
        $d = $Matches[1]; $t = $Matches[2]
        $stampReadable = "{0}-{1}-{2} {3}:{4}:{5}" -f $d.Substring(0,4), $d.Substring(4,2), $d.Substring(6,2), $t.Substring(0,2), $t.Substring(2,2), $t.Substring(4,2)
    } else {
        $stampReadable = $file.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
    }
    return $stampReadable
}

Write-Host "`nBackup disponibili in '$SapCommon' (dal piu' recente):" -ForegroundColor Cyan
for ($i = 0; $i -lt $backups.Count; $i++) {
    $label = Format-BackupLabel $backups[$i]
    $sizeKb = [math]::Round($backups[$i].Length / 1KB, 1)
    Write-Host ("  [{0}] {1}   ({2} KB, file: {3})" -f ($i + 1), $label, $sizeKb, $backups[$i].Name)
}

if ($ListOnly) {
    Write-Host "`n(-ListOnly: nessun ripristino eseguito.)" -ForegroundColor Yellow
    exit 0
}

# --- 2) determina quale backup ripristinare ---
$chosen = $null

if ($BackupFile) {
    $candidate = $BackupFile
    if (-not (Test-Path $candidate)) {
        $candidate = Join-Path $SapCommon $BackupFile
    }
    if (-not (Test-Path $candidate)) {
        Write-Host "`nERRORE: backup '$BackupFile' non trovato (ne' come percorso completo ne' dentro '$SapCommon')." -ForegroundColor Red
        exit 1
    }
    $chosen = Get-Item $candidate
}
elseif ($Latest) {
    $chosen = $backups[0]
    Write-Host "`n-Latest: seleziono automaticamente il piu' recente: $($chosen.Name)" -ForegroundColor Cyan
}
else {
    Write-Host ""
    $answer = Read-Host "Quale ripristinare? [1-$($backups.Count)] (INVIO = piu' recente, [1])"
    if ([string]::IsNullOrWhiteSpace($answer)) {
        $chosen = $backups[0]
    }
    else {
        $idx = 0
        if (-not [int]::TryParse($answer.Trim(), [ref]$idx) -or $idx -lt 1 -or $idx -gt $backups.Count) {
            Write-Host "ERRORE: scelta '$answer' non valida." -ForegroundColor Red
            exit 1
        }
        $chosen = $backups[$idx - 1]
    }
}

Write-Host "`nBackup selezionato: $($chosen.Name)" -ForegroundColor Cyan

# --- 3) backup di sicurezza del file personale ATTUALE, cosi' anche questo restore e' reversibile ---
if (Test-Path $PersonalFile) {
    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $preRestoreBackup = Join-Path $SapCommon "$PreRestorePrefix$stamp.xml"
    Copy-Item $PersonalFile $preRestoreBackup -Force
    Write-Host "Salvato lo stato attuale prima del ripristino in: $preRestoreBackup"
} else {
    Write-Host "Nessun SAPUILandscape.xml personale attualmente presente da salvare prima del ripristino."
}

# --- 4) ripristino vero e proprio ---
Copy-Item $chosen.FullName $PersonalFile -Force
Write-Host "`nRipristinato '$($chosen.Name)' su: $PersonalFile" -ForegroundColor Green

Write-Host "`nATTENZIONE:" -ForegroundColor Yellow
Write-Host "- Questo ripristina SOLO il file personale (workspace 'Local' + <Include> al globale)." -ForegroundColor Yellow
Write-Host "  La cartella 'SAPUILandscape_globale' condivisa (tutti i clienti) NON ha backup e NON viene toccata da questo script." -ForegroundColor Yellow
Write-Host "- Chiudi COMPLETAMENTE SAP Logon (controlla anche il Task Manager) prima di riaprirlo, altrimenti potrebbe continuare a mostrare la vecchia configurazione in memoria." -ForegroundColor Yellow

Write-Host "`nFatto." -ForegroundColor Green
