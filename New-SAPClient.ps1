<#
    New-SAPClient.ps1
    ------------------
    Crea la cartella e il file SAPUILandscape.xml per un NUOVO cliente dentro
    SAPUILandscape_sorgenti, chiedendo in sequenza sigla cliente, tipo di
    sistema, SID, SN (numero di sistema SAP) e host - e generando in automatico
    tutti gli uuid necessari (Workspace, Node, Item, Service), verificati contro
    quelli gia' usati negli altri file sorgente per escludere collisioni.

    Va eseguito nella cartella che contiene "SAPUILandscape_sorgenti" (es. quella
    estratta dallo zip di rollout, allo stesso livello di Deploy-SAPLandscape.ps1).

    Per ogni sistema chiede:
      - Tipo di sistema (es. S/4HANA, BW, ECC/R3, PI/PO, Solution Manager, CRM,
        SRM, Gateway, Fiori, MDG, PLM, DMS, SCM, TM, SLT) - usato per il nome
        della sotto-cartella "Sistemi <TIPO>". Se lasciato vuoto, il sistema
        finisce in "Altri sistemi".
      - SID (System ID, es. PRD)
      - SN  (numero di sistema SAP, 2 cifre, es. 00) -> usato per calcolare la
        porta della connessione (32NN, standard SAP GUI)
      - Host (IP o hostname del server applicativo)

    Dopo il primo sistema chiede se aggiungerne un altro per lo stesso cliente
    (utile per creare subito DEV/QAS/PRD ecc. in un solo passaggio). Se il
    cliente ha piu' di 2 sistemi E almeno due tipi diversi, i sistemi vengono
    raggruppati in sotto-cartelle per tipo, esattamente come per gli altri 89
    clienti; altrimenti restano piatti sotto il nodo cliente.

    Dopo aver creato il sorgente, richiama in automatico Build-SAPGlobal.ps1 (deve
    trovarsi accanto a questo script) per rigenerare subito SAPUILandscapeGlobal.xml
    e validarlo. NON distribuisce pero' nulla sui PC: per quello serve comunque
    lanciare Deploy-SAPLandscape.ps1/.cmd (che a sua volta rigenera anche lui il
    file globale da zero ad ogni esecuzione, quindi anche saltando questo passaggio
    automatico il risultato finale sarebbe comunque corretto).
#>

$ErrorActionPreference = "Stop"

$ScriptDir     = Split-Path -Parent $MyInvocation.MyCommand.Path
$SourcesFolder = Join-Path $ScriptDir "SAPUILandscape_sorgenti"

if (-not (Test-Path $SourcesFolder)) {
    Write-Host "ERRORE: non trovo la cartella 'SAPUILandscape_sorgenti' accanto a questo script ($ScriptDir)." -ForegroundColor Red
    exit 1
}

function ConvertTo-XmlEscaped {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return "" }
    $t = $Text -replace '&', '&amp;'
    $t = $t -replace '<', '&lt;'
    $t = $t -replace '>', '&gt;'
    $t = $t -replace '"', '&quot;'
    return $t
}

function ConvertTo-XmlCommentSafe {
    param([string]$Text)
    $t = $Text
    while ($t -match '--') { $t = $t -replace '--', '- ' }
    if ($t.EndsWith('-')) { $t += ' ' }
    return $t
}

function Get-ExistingUuids {
    # Raccoglie TUTTI gli uuid gia' usati in SAPUILandscape_sorgenti, per verificare
    # che i nuovi generati non collidano (con GUID casuali e' praticamente impossibile,
    # ma lo controlliamo comunque invece di darlo per scontato).
    $uuids = New-Object System.Collections.Generic.HashSet[string]
    Get-ChildItem -Path $SourcesFolder -Filter "SAPUILandscape.xml" -Recurse | ForEach-Object {
        [xml]$doc = Get-Content $_.FullName -Raw -Encoding UTF8
        $doc.SelectNodes("//*[@uuid]") | ForEach-Object { [void]$uuids.Add($_.uuid) }
    }
    # IMPORTANTE: "return $uuids" da solo farebbe si' che PowerShell "srotoli" l'HashSet
    # nei suoi singoli elementi (e' IEnumerable), restituendo un array di stringhe invece
    # dell'oggetto HashSet - che poi non supporta piu' .Add(). -NoEnumerate lo impedisce.
    Write-Output -NoEnumerate $uuids
}

function New-UniqueUuid {
    param($Existing)
    do { $u = [guid]::NewGuid().ToString() } while ($Existing.Contains($u))
    [void]$Existing.Add($u)
    return $u
}

Write-Host "=== Creazione nuovo cliente in SAPUILandscape_sorgenti ===" -ForegroundColor Cyan
Write-Host ""

$sigla = Read-Host "Sigla cliente (es. NUOVOCLIENTE)"
if ([string]::IsNullOrWhiteSpace($sigla)) {
    Write-Host "ERRORE: la sigla cliente e' obbligatoria." -ForegroundColor Red
    exit 1
}
$siglaSafe = ($sigla.Trim().ToUpper() -replace '[<>:"/\\|?*]', '_').Trim('.', ' ')

$clientFolder = Join-Path $SourcesFolder $siglaSafe
$clientFile   = Join-Path $clientFolder "SAPUILandscape.xml"

if (Test-Path $clientFile) {
    Write-Host "ATTENZIONE: esiste gia' un file sorgente per '$siglaSafe':" -ForegroundColor Yellow
    Write-Host "  $clientFile"
    $overwrite = Read-Host "Sovrascriverlo? (s/N)"
    if ($overwrite -ne 's' -and $overwrite -ne 'S') {
        Write-Host "Operazione annullata."
        exit 0
    }
}

$existingUuids = Get-ExistingUuids
Write-Host "($($existingUuids.Count) uuid gia' in uso nei sorgenti attuali - i nuovi verranno controllati contro questi)"
Write-Host ""

# --- raccolta sistemi (loop: puoi aggiungerne quanti vuoi per lo stesso cliente) ---
$systems = @()
$again = $true
while ($again) {
    Write-Host "--- Sistema $($systems.Count + 1) per $siglaSafe ---" -ForegroundColor Cyan

    $tipo = Read-Host "Tipo di sistema (es. S/4HANA, BW, ECC/R3, PI/PO, Solution Manager, CRM, SRM, Gateway, Fiori, MDG, PLM, DMS, SCM, TM, SLT - lascia vuoto per 'Altri sistemi')"
    if ([string]::IsNullOrWhiteSpace($tipo)) { $tipo = "Altri sistemi" } else { $tipo = $tipo.Trim() }

    $sid = Read-Host "SID (System ID, es. PRD)"
    $sid = $sid.Trim().ToUpper()
    if ([string]::IsNullOrWhiteSpace($sid)) {
        Write-Host "ERRORE: il SID e' obbligatorio, sistema scartato." -ForegroundColor Red
        continue
    }

    $sn = (Read-Host "SN (numero di sistema SAP, 2 cifre, es. 00)").Trim()
    if ($sn -notmatch '^\d{1,2}$') {
        Write-Host "ATTENZIONE: SN non valido ('$sn'), atteso un numero da 0 a 99. Uso '00'." -ForegroundColor Yellow
        $sn = "0"
    }
    $snPadded = $sn.PadLeft(2, '0')
    $port = "32$snPadded"

    $hostName = (Read-Host "Host (IP o hostname del server applicativo)").Trim()
    if ([string]::IsNullOrWhiteSpace($hostName)) {
        Write-Host "ERRORE: l'host e' obbligatorio, sistema scartato." -ForegroundColor Red
        continue
    }

    $systems += [PSCustomObject]@{
        Tipo   = $tipo
        Sid    = $sid
        Server = "$($hostName):$port"
    }

    Write-Host ""
    $ans = Read-Host "Aggiungere un altro sistema per ${siglaSafe}? (s/N)"
    $again = ($ans -eq 's' -or $ans -eq 'S')
    Write-Host ""
}

if ($systems.Count -eq 0) {
    Write-Host "ERRORE: nessun sistema valido inserito, non creo nulla." -ForegroundColor Red
    exit 1
}

# --- decide se raggruppare per tipo: SEMPRE (allineato a
#     Rebuild-SourcesFromMapping.ps1), cosi' anche un cliente nuovo con un
#     solo sistema/tipo mostra subito la sotto-cartella "Sistemi <Tipo>"
#     invece del generico "Sistemi" senza tipo aggiunto da Build-SAPGlobal.ps1 ---
$grouped     = $systems | Group-Object Tipo
$shouldGroup = $systems.Count -gt 0

$wsUuid   = New-UniqueUuid $existingUuids
$nodeUuid = New-UniqueUuid $existingUuids

$nl = "`n"
$nodesXml    = ""
$servicesXml = ""

function New-ItemBlock {
    param($Sys, $Existing, $Indent)
    $svcUuid  = New-UniqueUuid $Existing
    $itemUuid = New-UniqueUuid $Existing
    $svcNameRaw = "$siglaSafe - $($Sys.Sid)"
    $svcName    = ConvertTo-XmlEscaped $svcNameRaw
    $sidEsc     = ConvertTo-XmlEscaped $Sys.Sid
    $serverEsc  = ConvertTo-XmlEscaped $Sys.Server
    $commentTxt = ConvertTo-XmlCommentSafe " $svcNameRaw [$($Sys.Sid)] $($Sys.Server) "

    $itemXml  = "$Indent<!--$commentTxt-->$nl"
    $itemXml += "$Indent<Item uuid=`"$itemUuid`" serviceid=`"$svcUuid`" />$nl"

    $svcXml = "`t`t<Service type=`"SAPGUI`" uuid=`"$svcUuid`" name=`"$svcName`" systemid=`"$sidEsc`" mode=`"1`" server=`"$serverEsc`" sncop=`"-1`" sapcpg=`"1100`" dcpg=`"2`" />$nl"

    return @{ Item = $itemXml; Service = $svcXml }
}

if ($shouldGroup) {
    foreach ($g in $grouped) {
        $typeNodeUuid = New-UniqueUuid $existingUuids
        $typeNameRaw  = if ([string]::IsNullOrWhiteSpace($g.Name)) { "Altri sistemi" } else { "Sistemi $($g.Name)" }
        $typeNameEsc  = ConvertTo-XmlEscaped $typeNameRaw
        $itemsXml = ""
        foreach ($sys in $g.Group) {
            # 5 tabs: Workspace(2) > Node cliente(3) > Node tipo(4) > Item(5)
            $blk = New-ItemBlock -Sys $sys -Existing $existingUuids -Indent "`t`t`t`t`t"
            $itemsXml    += $blk.Item
            $servicesXml += $blk.Service
        }
        # 4 tabs: Workspace(2) > Node cliente(3) > Node tipo(4)
        $nodesXml += "`t`t`t`t<Node uuid=`"$typeNodeUuid`" name=`"$typeNameEsc`" hidden=`"0`">$nl"
        $nodesXml += $itemsXml
        $nodesXml += "`t`t`t`t</Node>$nl"
    }
} else {
    foreach ($sys in $systems) {
        # 4 tabs: Workspace(2) > Node cliente(3) > Item(4), come nei client "piatti" gia' esistenti
        $blk = New-ItemBlock -Sys $sys -Existing $existingUuids -Indent "`t`t`t`t"
        $nodesXml    += $blk.Item
        $servicesXml += $blk.Service
    }
}

$siglaEsc = ConvertTo-XmlEscaped $siglaSafe
$updated  = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss") + "Z"

$xml  = "<?xml version='1.0' encoding='utf-8'?>$nl"
$xml += "<Landscape updated=`"$updated`" version=`"1`" generator=`"SAP GUI for Windows v8000.1.17.155`">$nl"
$xml += "`t<Workspaces>$nl"
$xml += "`t`t<Workspace uuid=`"$wsUuid`" name=`"$siglaEsc`" expanded=`"0`" hidden=`"0`">$nl"
$xml += "`t`t`t<Node uuid=`"$nodeUuid`" name=`"$siglaEsc`" hidden=`"0`">$nl"
$xml += $nodesXml
$xml += "`t`t`t</Node>$nl"
$xml += "`t`t</Workspace>$nl"
$xml += "`t</Workspaces>$nl"
$xml += "`t<Services>$nl"
$xml += $servicesXml
$xml += "`t</Services>$nl"
$xml += "`t<Messageservers />$nl"
$xml += "`t<Routers />$nl"
$xml += "</Landscape>$nl"

New-Item -ItemType Directory -Path $clientFolder -Force | Out-Null
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($clientFile, $xml, $utf8NoBom)

Write-Host "Creato: $clientFile" -ForegroundColor Green

# --- verifica finale: XML ben formato + nessuna collisione di uuid nel set completo ---
try {
    [xml](Get-Content $clientFile -Raw -Encoding UTF8) | Out-Null
    Write-Host "Verifica XML: ben formato." -ForegroundColor Green
} catch {
    Write-Host "ERRORE: il file generato NON e' un XML valido: $_" -ForegroundColor Red
    exit 1
}

$allUuids = Get-ExistingUuids
$dupCount = $allUuids.Count
$totalAttrs = (Get-ChildItem -Path $SourcesFolder -Filter "SAPUILandscape.xml" -Recurse | ForEach-Object {
    ([xml](Get-Content $_.FullName -Raw -Encoding UTF8)).SelectNodes("//*[@uuid]").Count
} | Measure-Object -Sum).Sum
if ($totalAttrs -eq $dupCount) {
    Write-Host "Verifica uuid: nessuna collisione su $dupCount uuid nell'intero SAPUILandscape_sorgenti." -ForegroundColor Green
} else {
    Write-Host "ATTENZIONE: rilevate $($totalAttrs - $dupCount) collisioni di uuid nell'intero set! Controlla a mano." -ForegroundColor Red
}

Write-Host ""
Write-Host "Sistemi creati per ${siglaSafe}:" -ForegroundColor Cyan
$systems | ForEach-Object { Write-Host "  - $($_.Sid) [$($_.Tipo)] $($_.Server)" }
Write-Host ""

# --- rigenera subito il file globale, per validare il nuovo cliente senza aspettare il deploy ---
$BuildScript = Join-Path $ScriptDir "Build-SAPGlobal.ps1"
if (Test-Path $BuildScript) {
    Write-Host "=== Rigenero SAPUILandscapeGlobal.xml per includere il nuovo cliente ===" -ForegroundColor Cyan
    # Eseguito come processo separato (non con l'operatore &): se Build-SAPGlobal.ps1
    # incontra un errore e chiama "exit", deve terminare solo se stesso, non anche
    # questo script.
    $proc = Start-Process -FilePath "powershell.exe" `
        -ArgumentList @("-NoLogo", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$BuildScript`"") `
        -NoNewWindow -Wait -PassThru
    if ($proc.ExitCode -ne 0) {
        Write-Host "ATTENZIONE: la rigenerazione automatica ha restituito un errore (codice $($proc.ExitCode))." -ForegroundColor Yellow
        Write-Host "Il sorgente di ${siglaSafe} e' comunque stato creato; lancia Build-SAPGlobal.cmd a mano per i dettagli." -ForegroundColor Yellow
    }
} else {
    Write-Host "ATTENZIONE: non trovo Build-SAPGlobal.ps1 accanto a questo script." -ForegroundColor Yellow
    Write-Host "Lancialo a mano (o Deploy-SAPLandscape.ps1, che rigenera comunque il globale da solo)." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "FATTO. Ricorda: per portare il nuovo cliente su un PC devi comunque" -ForegroundColor Yellow
Write-Host "lanciare Deploy-SAPLandscape.cmd (rigenera di nuovo il globale in automatico" -ForegroundColor Yellow
Write-Host "e lo distribuisce - questo passaggio qui sopra serve solo per validare subito)." -ForegroundColor Yellow
