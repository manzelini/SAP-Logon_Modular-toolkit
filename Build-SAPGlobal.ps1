<#
    Build-SAPGlobal.ps1
    ---------------------
    Rigenera SAPUILandscape_globale\SAPUILandscapeGlobal.xml a partire da TUTTI i
    file sorgente dentro SAPUILandscape_sorgenti (un cliente = una cartella = un
    SAPUILandscape.xml), raggruppandoli in un'unica workspace "AVVALE Clients" -
    esattamente come la ricompilazione che finora veniva fatta "a mano" dopo ogni
    modifica ai sorgenti. La cartella "_AVVALE_INTERNAL" NON viene incorporata:
    resta un <Include> separato, come gia' avveniva.

    Va eseguito ogni volta che un sorgente cliente cambia (nuovo cliente creato con
    New-SAPClient.ps1, sistema aggiunto/rimosso/modificato a mano in un file
    SAPUILandscape.xml) PRIMA di ridistribuire con Deploy-SAPLandscape.ps1. Se lo
    lanci da New-SAPClient.cmd, viene richiamato in automatico alla fine.

    Uso: doppio click su Build-SAPGlobal.cmd (o esegui con PowerShell direttamente).
#>

$ErrorActionPreference = "Stop"

$ScriptDir     = Split-Path -Parent $MyInvocation.MyCommand.Path
$SourcesFolder = Join-Path $ScriptDir "SAPUILandscape_sorgenti"
$GlobalFolder  = Join-Path $ScriptDir "SAPUILandscape_globale"
$GlobalFile    = Join-Path $GlobalFolder "SAPUILandscapeGlobal.xml"
$InternalName  = "_AVVALE_INTERNAL"

# Stesso placeholder usato finora: Deploy-SAPLandscape.ps1 lo ricalcola in automatico
# in base al PC/utente reale al momento dell'installazione - non toccare a mano.
$BaseUrl = "file:///C:/Users/Administrator/AppData/Roaming/SAP/Common/SAPUILandscape_globale"

# Uuid fisso della workspace "AVVALE Clients" (lo stesso usato in tutte le rigenerazioni
# precedenti, cosi' non cambia identita' ad ogni rebuild).
$WorkspaceUuid = "4869f745-1e39-5605-8e57-ba2fa8642e48"

function Write-Section {
    param([string]$Text)
    Write-Host $Text -ForegroundColor Cyan
}

if (-not (Test-Path $SourcesFolder)) {
    Write-Host "ERRORE: non trovo 'SAPUILandscape_sorgenti' accanto a questo script ($ScriptDir)." -ForegroundColor Red
    exit 1
}

$internalFile = Join-Path (Join-Path $SourcesFolder $InternalName) "SAPUILandscape.xml"
if (-not (Test-Path $internalFile)) {
    Write-Host "ERRORE: non trovo '$InternalName\SAPUILandscape.xml' dentro '$SourcesFolder'." -ForegroundColor Red
    exit 1
}

New-Item -ItemType Directory -Path $GlobalFolder -Force | Out-Null

Write-Section "=== Rigenerazione SAPUILandscapeGlobal.xml da SAPUILandscape_sorgenti ==="

$clientFolders = Get-ChildItem -Path $SourcesFolder -Directory |
    Where-Object { $_.Name -ne $InternalName } |
    Sort-Object Name

$seenSvc = New-Object System.Collections.Generic.HashSet[string]
$seenMs  = New-Object System.Collections.Generic.HashSet[string]
$seenRt  = New-Object System.Collections.Generic.HashSet[string]

$nodesSb    = New-Object System.Text.StringBuilder
$servicesSb = New-Object System.Text.StringBuilder
$msSb       = New-Object System.Text.StringBuilder
$rtSb       = New-Object System.Text.StringBuilder

$nClients  = 0
$nSkipped  = 0
$nOrphanItems = 0

foreach ($folder in $clientFolders) {
    $file = Join-Path $folder.FullName "SAPUILandscape.xml"
    if (-not (Test-Path $file)) {
        Write-Host "  ATTENZIONE: '$($folder.Name)' non contiene SAPUILandscape.xml, saltato." -ForegroundColor Yellow
        $nSkipped++
        continue
    }

    # PreserveWhitespace=true per mantenere l'indentazione a tab gia' presente nei
    # sorgenti quando ricopiamo i blocchi cosi' come sono (niente [xml] cast qui,
    # che normalizzerebbe/perderebbe gli spazi tra i tag).
    $xmlDoc = New-Object System.Xml.XmlDocument
    $xmlDoc.PreserveWhitespace = $true
    try {
        $xmlDoc.Load($file)
    } catch {
        Write-Host "  ERRORE: '$file' non e' un XML valido, saltato: $_" -ForegroundColor Red
        $nSkipped++
        continue
    }
    $nClients++

    # --- Services: dedup per uuid, mantenendo l'eventuale commento che lo precede ---
    # NOTA: si usa SelectSingleNode/XPath invece della notazione "puntata"
    # ($xmlDoc.Landscape.Services) - quella notazione passa dall'adapter dinamico
    # di PowerShell per la navigazione XML, che in questo script ha mostrato un
    # comportamento inaffidabile (la sezione <Services> del file finale risultava
    # vuota nonostante il codice sembrasse corretto, mentre Messageservers/Routers,
    # scritti con lo stesso schema, funzionavano). SelectSingleNode passa invece
    # direttamente dalle API .NET di System.Xml, senza l'adapter di mezzo.
    # BUGFIX CONFERMATO (diagnosticato con Diagnose-Services.ps1 su dati reali):
    # PowerShell, per un XmlElement che ha un attributo "name", fa "risalire"
    # quel valore alla proprieta' .Name al posto del vero nome del tag XML - quindi
    # per <Service name="AET - DEV - Development" .../> , $node.Name restituiva
    # "AET - DEV - Development" e MAI "Service". Il controllo "$node.Name -eq
    # 'Service'" sottostante non era quindi mai vero, per nessun cliente: la
    # sezione <Services> del file finale risultava sempre vuota. Fix: non serve
    # controllare il nome del tag (dentro <Services> gli unici elementi possibili
    # sono <Service>, esattamente come Messageservers/Routers piu' sotto, che
    # infatti non hanno mai avuto questo controllo e hanno sempre funzionato).
    # NOTA: questo blocco gira PRIMA di quello sui Node/Item qui sotto, perche'
    # ci serve conoscere gli uuid dei Service DI QUESTO cliente per poter scartare
    # eventuali Item con un serviceid non risolvibile (v. $localServiceIds sotto).
    $localServiceIds = New-Object System.Collections.Generic.HashSet[string]
    $servicesNode = $xmlDoc.SelectSingleNode("/Landscape/Services")
    $pendingComment = $null
    if ($null -ne $servicesNode) {
        foreach ($node in $servicesNode.ChildNodes) {
            if ($node.NodeType -eq "Comment") {
                $pendingComment = $node
                continue
            }
            if ($node.NodeType -eq "Element") {
                $u = $node.GetAttribute("uuid")
                if ($u) { [void]$localServiceIds.Add($u) }
                if ($u -and -not $seenSvc.Contains($u)) {
                    [void]$seenSvc.Add($u)
                    if ($null -ne $pendingComment) { [void]$servicesSb.AppendLine("`t`t" + $pendingComment.OuterXml) }
                    [void]$servicesSb.AppendLine("`t`t" + $node.OuterXml)
                }
                $pendingComment = $null
            }
        }
    }

    # --- Node del cliente (di norma uno solo) copiato cosi' com'e', con i suoi commenti ---
    $ws = $xmlDoc.SelectSingleNode("/Landscape/Workspaces/Workspace")
    if ($null -eq $ws) {
        Write-Host "  ATTENZIONE: '$($folder.Name)' non ha una Workspace valida, saltato." -ForegroundColor Yellow
        $nSkipped++
        continue
    }
    foreach ($child in $ws.ChildNodes) {
        if ($child.NodeType -ne "Element") { continue }

        # Scarta gli Item con un serviceid che non risolve a nessun Service DI
        # QUESTO cliente (dati sorgente gia' rotti in partenza, es. commenti
        # "ATTENZIONE: serviceid non trovato" preesistenti in APTAR) - se non lo
        # facessimo qui, SAP Logon mostrerebbe comunque quei sistemi vuoti/rotti
        # e la verifica finale bloccherebbe l'intero deploy per colpa di dati
        # gia' orfani da prima, non per un errore introdotto da questo script.
        $orphanItems = @()
        foreach ($gc in @($child.SelectNodes(".//Item"))) {
            $sid = $gc.GetAttribute("serviceid")
            if (-not $sid -or -not $localServiceIds.Contains($sid)) { $orphanItems += $gc }
        }
        foreach ($orphan in $orphanItems) {
            $prevComment = $orphan.PreviousSibling
            if ($null -ne $prevComment -and $prevComment.NodeType -eq "Comment") { [void]$prevComment.ParentNode.RemoveChild($prevComment) }
            [void]$orphan.ParentNode.RemoveChild($orphan)
            $nOrphanItems++
        }

        # BUGFIX (confermato su AET/PIPPONE/AB HOLDING ecc.): quando il Node del
        # cliente contiene <Item> DIRETTAMENTE (client "piatto", senza sotto-Node
        # per tipo sistema), una volta annidato dentro l'unica Workspace condivisa
        # "AVVALE Clients" SAP Logon non carica quegli Item nel suo modello dati
        # (il nodo cliente compare ma resta vuoto, e nemmeno il filtro di ricerca
        # di SAP Logon li trova). I client "raggruppati per tipo" invece funzionano
        # sempre, perche' hanno un livello di Node in piu' tra il cliente e gli Item.
        # Per uniformare la struttura ed evitare il problema, se il Node del cliente
        # ha Item come figli diretti li avvolgiamo qui in un Node intermedio
        # "Sistemi", cosi' ogni cliente ha sempre almeno due livelli di Node prima
        # degli Item, esattamente come i client che gia' funzionano.
        $hasDirectItem = $false
        foreach ($gc in $child.ChildNodes) {
            if ($gc.NodeType -eq "Element" -and $gc.Name -eq "Item") { $hasDirectItem = $true; break }
        }

        if ($hasDirectItem) {
            $wrapper = $xmlDoc.CreateElement("Node")
            $wrapper.SetAttribute("uuid", [guid]::NewGuid().ToString())
            $wrapper.SetAttribute("name", "Sistemi")
            $wrapper.SetAttribute("hidden", "0")

            $existingChildren = @($child.ChildNodes)
            foreach ($n in $existingChildren) {
                [void]$child.RemoveChild($n)
                [void]$wrapper.AppendChild($xmlDoc.CreateTextNode("`n`t`t`t`t`t"))
                [void]$wrapper.AppendChild($n)
            }
            [void]$wrapper.AppendChild($xmlDoc.CreateTextNode("`n`t`t`t`t"))

            [void]$child.AppendChild($xmlDoc.CreateTextNode("`n`t`t`t`t"))
            [void]$child.AppendChild($wrapper)
            [void]$child.AppendChild($xmlDoc.CreateTextNode("`n`t`t`t"))
        }

        [void]$nodesSb.AppendLine("`t`t`t" + $child.OuterXml)
    }

    # --- Messageservers: dedup per uuid ---
    $msNode = $xmlDoc.SelectSingleNode("/Landscape/Messageservers")
    if ($null -ne $msNode) {
    foreach ($node in $msNode.ChildNodes) {
        if ($node.NodeType -ne "Element") { continue }
        $u = $node.GetAttribute("uuid")
        if ($u -and -not $seenMs.Contains($u)) {
            [void]$seenMs.Add($u)
            [void]$msSb.AppendLine("`t`t" + $node.OuterXml)
        }
    }
    }

    # --- Routers: dedup per uuid ---
    $rtNode = $xmlDoc.SelectSingleNode("/Landscape/Routers")
    if ($null -ne $rtNode) {
    foreach ($node in $rtNode.ChildNodes) {
        if ($node.NodeType -ne "Element") { continue }
        $u = $node.GetAttribute("uuid")
        if ($u -and -not $seenRt.Contains($u)) {
            [void]$seenRt.Add($u)
            [void]$rtSb.AppendLine("`t`t" + $node.OuterXml)
        }
    }
    }
}

if ($nClients -eq 0) {
    Write-Host "ERRORE: nessun cliente valido trovato in '$SourcesFolder', non rigenero nulla." -ForegroundColor Red
    exit 1
}

$updated = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss") + "Z"
$includeUrl = "$BaseUrl/$InternalName/SAPUILandscape.xml"

$out = New-Object System.Text.StringBuilder
[void]$out.AppendLine("<?xml version='1.0' encoding='utf-8'?>")
[void]$out.AppendLine("<Landscape updated=`"$updated`" version=`"1`" generator=`"SAP GUI for Windows v8000.1.17.155`">")
[void]$out.AppendLine("`t<Workspaces>")
[void]$out.AppendLine("`t`t<Workspace uuid=`"$WorkspaceUuid`" name=`"AVVALE Clients`" expanded=`"0`" hidden=`"0`">")
[void]$out.Append($nodesSb.ToString())
[void]$out.AppendLine("`t`t</Workspace>")
[void]$out.AppendLine("`t</Workspaces>")
[void]$out.AppendLine("`t<Services>")
[void]$out.Append($servicesSb.ToString())
[void]$out.AppendLine("`t</Services>")
[void]$out.AppendLine("`t<Messageservers>")
[void]$out.Append($msSb.ToString())
[void]$out.AppendLine("`t</Messageservers>")
[void]$out.AppendLine("`t<Routers>")
[void]$out.Append($rtSb.ToString())
[void]$out.AppendLine("`t</Routers>")
[void]$out.AppendLine("`t<Includes>")
[void]$out.AppendLine("`t`t<Include url=`"$includeUrl`" index=`"0`" />")
[void]$out.AppendLine("`t</Includes>")
[void]$out.AppendLine("</Landscape>")

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($GlobalFile, $out.ToString(), $utf8NoBom)

Write-Host "Scritto: $GlobalFile" -ForegroundColor Green
Write-Host "Clienti incorporati: $nClients (saltati: $nSkipped) | Service: $($seenSvc.Count) | Messageserver: $($seenMs.Count) | Router: $($seenRt.Count)"
if ($nOrphanItems -gt 0) {
    Write-Host "ATTENZIONE: scartati $nOrphanItems sistemi con collegamento gia' rotto nel sorgente (serviceid inesistente, dati preesistenti - non introdotti da questo script). Controllare i sorgenti dei clienti coinvolti quando possibile." -ForegroundColor Yellow
}

# --- verifica finale: XML ben formato + nessuna collisione di uuid con _AVVALE_INTERNAL ---
try {
    [xml](Get-Content $GlobalFile -Raw -Encoding UTF8) | Out-Null
    Write-Host "Verifica XML: ben formato." -ForegroundColor Green
} catch {
    Write-Host "ERRORE: il file rigenerato NON e' un XML valido: $_" -ForegroundColor Red
    exit 1
}

$allUuids = New-Object System.Collections.Generic.HashSet[string]
$totalAttrs = 0
foreach ($f in @($GlobalFile, $internalFile)) {
    [xml]$check = Get-Content $f -Raw -Encoding UTF8
    $nodesWithUuid = $check.SelectNodes("//*[@uuid]")
    $totalAttrs += $nodesWithUuid.Count
    foreach ($n in $nodesWithUuid) { [void]$allUuids.Add($n.uuid) }
}
if ($totalAttrs -eq $allUuids.Count) {
    Write-Host "Verifica uuid: nessuna collisione tra il file globale e _AVVALE_INTERNAL ($($allUuids.Count) uuid)." -ForegroundColor Green
} else {
    # BLOCCANTE (non solo un avviso): una collisione di uuid qui significa quasi
    # sempre una riga duplicata nella mappa Excel a cui non e' stato assegnato un
    # nuovo "UUID Item"/"UUID Service" - il sistema "nuovo" viene silenziosamente
    # scartato (il suo Item si sovrappone a quello gia' esistente, i suoi dati
    # Server/SystemID non compaiono da nessuna parte). Deploy-SAPLandscape.ps1
    # legge il codice di uscita di questo script per decidere se procedere: prima
    # di questo fix la collisione veniva solo stampata in rosso ma NON bloccava
    # il deploy, quindi un dato corrotto poteva finire distribuito inosservato.
    Write-Host "ERRORE: rilevate $($totalAttrs - $allUuids.Count) collisioni di uuid! Controlla a mano prima di distribuire (probabile riga duplicata in mappa_uuid_sistemi.xlsx senza nuovo UUID Item/UUID Service). Deploy interrotto." -ForegroundColor Red
    exit 1
}

# --- verifica aggiuntiva: ogni Item deve poter risolvere il proprio serviceid in un
# <Service> presente nel file globale (senza questo, SAP Logon mostra le cartelle
# ma nessun sistema al loro interno - e' esattamente il bug che ha causato questo
# controllo). Se anche un solo Item non risolve, e' un errore bloccante. ---
[xml]$finalCheck = Get-Content $GlobalFile -Raw -Encoding UTF8
$definedServiceIds = New-Object System.Collections.Generic.HashSet[string]
foreach ($svc in $finalCheck.SelectNodes("//Service[@uuid]")) { [void]$definedServiceIds.Add($svc.uuid) }
$allItems = $finalCheck.SelectNodes("//Item[@serviceid]")
$unresolved = 0
foreach ($it in $allItems) {
    if (-not $definedServiceIds.Contains($it.serviceid)) { $unresolved++ }
}
if ($allItems.Count -eq 0) {
    Write-Host "ERRORE: nessun <Item> trovato nel file globale, qualcosa e' andato storto." -ForegroundColor Red
    exit 1
} elseif ($unresolved -gt 0) {
    Write-Host "ERRORE: $unresolved Item su $($allItems.Count) hanno un serviceid che non corrisponde a nessun <Service> nel file globale (Service definiti: $($definedServiceIds.Count))! Il file NON va distribuito cosi'." -ForegroundColor Red
    exit 1
} else {
    Write-Host "Verifica referenze: tutti i $($allItems.Count) Item risolvono correttamente il loro <Service> ($($definedServiceIds.Count) Service definiti)." -ForegroundColor Green
}

Write-Host ""
Write-Host "FATTO. Ora puoi ridistribuire con Deploy-SAPLandscape.cmd (backup automatico incluso)." -ForegroundColor Yellow
