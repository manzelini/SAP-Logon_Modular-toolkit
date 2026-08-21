# Rebuild-SourcesFromMapping.ps1
# ================================
# Ricostruisce SAPUILandscape_sorgenti (un file SAPUILandscape.xml per
# cliente) leggendo la mappa "mappa_uuid_sistemi.xlsx" - la stessa
# esportata dallo script Python build_mapping_v2.py in direzione opposta
# (XML -> Excel). Equivalente PowerShell di rebuild_sources_from_mapping.py,
# pensato per girare su QUALSIASI PC Windows senza installare nulla: legge
# il file .xlsx direttamente come archivio zip/XML tramite le sole
# librerie .NET incluse in PowerShell (nessun modulo Excel/COM richiesto,
# nessun Python, nessuna libreria esterna).
#
# USO:
#   powershell -ExecutionPolicy Bypass -File Rebuild-SourcesFromMapping.ps1 `
#       -XlsxPath ".\mappa_uuid_sistemi.xlsx" `
#       -ReferenceFolder ".\SAPUILandscape_sorgenti" `
#       -OutFolder ".\SAPUILandscape_sorgenti_REBUILT"
#
# oppure semplicemente Rebuild-SourcesFromMapping.cmd a doppio clic se i
# tre elementi si trovano gia' con i nomi di default nella stessa cartella.
#
# CLIENTI RIMOSSI DALL'EXCEL: un cliente che nella "Mappa sistemi" non ha
# PIU' NESSUNA riga viene considerato eliminato di proposito e la sua
# cartella non viene ricopiata in $OutFolder (rebuild + deploy successivi
# lo fanno sparire davvero da SAP Logon). L'elenco di questi clienti viene
# sempre stampato per esteso (mai solo un conteggio), perche' un cliente
# con DAVVERO zero sistemi non e' rappresentabile in un foglio "un rigo
# per sistema" e finirebbe nello stesso elenco pur non essendo stato
# rimosso apposta. Se serve preservare questi casi invece di eliminarli,
# rilanciare con lo switch -PreserveClientsWithoutRows (comportamento
# storico dello script: li ricopia invariati dalla reference).
#
# DOVE FINISCONO I SORGENTI RICOSTRUITI: per default restano in $OutFolder
# (".\SAPUILandscape_sorgenti_REBUILT"), separati dalla $ReferenceFolder
# live - vanno controllati e poi sostituiti a mano alla cartella sorgenti
# vera prima di Build/Deploy. Con lo switch -ReplaceSources lo script fa
# questo passaggio da solo, ma SOLO se le verifiche finali non hanno
# rilevato errori: rinomina la $ReferenceFolder esistente aggiungendo un
# timestamp (backup, es. "SAPUILandscape_sorgenti_20260820_090633") e
# sposta $OutFolder al suo posto, cosi' diventa lei la nuova cartella
# sorgenti attiva, pronta per Build-SAPGlobal/Deploy-SAPLandscape senza
# altri passaggi manuali.
#
# Colonne attese nei fogli "Mappa sistemi" e "Avvale (interno)" (in questo
# ordine, intestazione in riga 1):
#   Cliente | Percorso cartella | Nome sistema | System ID | Server |
#   Message Server | Router | UUID Item | UUID Service |
#   Nome sistema standard | Da verificare | Tipo sistema
#
# Note tecniche importanti (lezioni gia' apprese in questo progetto, v.
# GESTIONE_MODULARE_SAP_GUI.md par. 7 - applicate qui rigorosamente):
# - MAI confrontare $node.Name su un XmlElement che ha un attributo
#   "name": PowerShell restituirebbe il VALORE dell'attributo, non il nome
#   del tag XML. Qui non serve mai un confronto simile.
# - Ogni controllo di presenza di un nodo XML usa $null -eq/-ne, mai
#   "if ($node)" (un XmlNode e' IEnumerable e la coercizione booleana e'
#   inaffidabile).
# - Ogni funzione che restituisce un ARRAY lo fa con "return ,$array" (la
#   virgola unaria) per evitare che PowerShell lo srotoli nei suoi
#   elementi quando l'array ha 0 o 1 elemento (bug gia' incontrato con una
#   HashSet restituita da funzione in Build-SAPGlobal.ps1).
# - I valori "record" con piu' campi vengono restituiti come
#   [PSCustomObject], MAI come Hashtable/array: un Hashtable puo' essere
#   enumerato dalla pipeline se attraversa piu' funzioni, un
#   PSCustomObject no.
# - L'XML e' letto/scritto sempre con XPath + XmlNamespaceManager (mai
#   dot-notation dinamica), dato che i file OOXML (.xlsx) hanno un
#   namespace di default che la dot-notation non gestisce in modo
#   affidabile.

param(
    [string]$XlsxPath = ".\mappa_uuid_sistemi.xlsx",
    [string]$ReferenceFolder = ".\SAPUILandscape_sorgenti",
    [string]$OutFolder = ".\SAPUILandscape_sorgenti_REBUILT",
    [string]$Updated = "2026-01-01T00:00:00Z",
    [string]$Version = "1",
    [string]$Generator = "SAP GUI for Windows v8000.1.17.155",
    [switch]$PreserveClientsWithoutRows,
    [switch]$ReplaceSources
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression.FileSystem

$SS_NS  = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
$REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
$RID_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

$TYPE_ORDER = @("S/4HANA", "BW", "PI/PO", "CRM", "SRM", "Solution Manager", "Gateway",
    "Fiori", "MDG", "PLM", "DMS", "SCM", "TM", "SLT", "ECC/R3", "Altri sistemi")

function Get-TypeSortIndex([string]$label) {
    $idx = [array]::IndexOf($TYPE_ORDER, $label)
    if ($idx -lt 0) { return $TYPE_ORDER.Count }
    return $idx
}

function New-SsNsManager([System.Xml.XmlDocument]$xmlDoc) {
    # NOTA: System.Xml.XmlNamespaceManager implementa IEnumerable (enumera i
    # prefissi registrati, es. "xml","s","r") - un "return $nsMgr" nudo
    # verrebbe quindi srotolato dalla pipeline in un array di STRINGHE
    # invece di restituire l'oggetto XmlNamespaceManager. Da qui l'errore
    # "Cannot find an overload for SelectSingleNode ... argument count 2":
    # il secondo argomento arrivava come string[] anziche' XmlNamespaceManager.
    # Fix: "return ,$nsMgr" (vedi stessa convenzione gia' usata nel resto
    # dello script per liste/array).
    $nsMgr = New-Object System.Xml.XmlNamespaceManager($xmlDoc.NameTable)
    $nsMgr.AddNamespace("s", $SS_NS)
    $nsMgr.AddNamespace("r", $RID_NS)
    return ,$nsMgr
}

function Get-ZipEntryXml($zip, [string]$entryName) {
    $entry = $zip.GetEntry($entryName)
    if ($null -eq $entry) { return $null }
    $stream = $entry.Open()
    try {
        $reader = New-Object System.IO.StreamReader($stream)
        $content = $reader.ReadToEnd()
        $reader.Close()
    }
    finally {
        $stream.Close()
    }
    $xmlDoc = New-Object System.Xml.XmlDocument
    $xmlDoc.LoadXml($content)
    # XmlDocument (come ogni XmlNode) implementa anch'esso IEnumerable
    # (enumera i ChildNodes): stessa cautela di New-SsNsManager sopra.
    return ,$xmlDoc
}

function Get-ColumnIndexFromRef([string]$cellRef) {
    if ($cellRef -notmatch '^([A-Z]+)\d+$') { return $null }
    $letters = $Matches[1]
    $idx = 0
    foreach ($ch in $letters.ToCharArray()) {
        $idx = $idx * 26 + ([int][char]$ch - [int][char]'A' + 1)
    }
    return $idx
}

function Get-SharedStrings($zip) {
    # se il file non usa shared strings (es. e' stato scritto con inline
    # strings, come fa openpyxl) l'entry semplicemente non esiste
    $ssXml = Get-ZipEntryXml $zip "xl/sharedStrings.xml"
    $list = New-Object System.Collections.Generic.List[string]
    if ($null -eq $ssXml) { return ,$list }
    $nsMgr = New-SsNsManager $ssXml
    foreach ($si in $ssXml.SelectNodes("//s:sst/s:si", $nsMgr)) {
        $parts = @()
        foreach ($t in $si.SelectNodes(".//s:t", $nsMgr)) { $parts += $t.InnerText }
        [void]$list.Add([string]::Join("", $parts))
    }
    return ,$list
}

function Get-CellValue($cellNode, $sharedStrings, $nsMgr) {
    if ($null -eq $cellNode) { return "" }
    $t = $cellNode.GetAttribute("t")
    if ($t -eq "inlineStr") {
        $isNode = $cellNode.SelectSingleNode("s:is", $nsMgr)
        if ($null -eq $isNode) { return "" }
        $parts = @()
        foreach ($tNode in $isNode.SelectNodes(".//s:t", $nsMgr)) { $parts += $tNode.InnerText }
        return [string]::Join("", $parts)
    }
    elseif ($t -eq "s") {
        $vNode = $cellNode.SelectSingleNode("s:v", $nsMgr)
        if ($null -eq $vNode) { return "" }
        $idx = [int]$vNode.InnerText
        if ($idx -ge 0 -and $idx -lt $sharedStrings.Count) { return $sharedStrings[$idx] }
        return ""
    }
    else {
        # numero, stringa "str" (formula) o cella vuota: il valore grezzo va bene
        $vNode = $cellNode.SelectSingleNode("s:v", $nsMgr)
        if ($null -eq $vNode) { return "" }
        return $vNode.InnerText
    }
}

function Get-SheetTarget($wbXml, $wbNsMgr, $relsXml, [string]$sheetName) {
    $sheetNode = $wbXml.SelectSingleNode("//s:sheets/s:sheet[@name='$sheetName']", $wbNsMgr)
    if ($null -eq $sheetNode) { return $null }
    $rid = $sheetNode.GetAttribute("id", $RID_NS)
    if (-not $rid) { return $null }
    $relsNsMgr = New-Object System.Xml.XmlNamespaceManager($relsXml.NameTable)
    $relsNsMgr.AddNamespace("rel", $REL_NS)
    $relNode = $relsXml.SelectSingleNode("//rel:Relationship[@Id='$rid']", $relsNsMgr)
    if ($null -eq $relNode) { return $null }
    $target = $relNode.GetAttribute("Target")
    if ($target.StartsWith("/")) { return $target.TrimStart("/") }
    return "xl/" + $target
}

$MAPPING_COLUMNS = @(
    "Cliente", "PercorsoCartella", "NomeSistema", "SystemId", "Server",
    "MessageServer", "Router", "UuidItem", "UuidService",
    "NomeSistemaStandard", "DaVerificare", "TipoSistema"
)

function Read-SheetRows($zip, [string]$sheetTarget, $sharedStrings) {
    $rows = New-Object System.Collections.Generic.List[object]
    $sheetXml = Get-ZipEntryXml $zip $sheetTarget
    if ($null -eq $sheetXml) { return ,$rows }
    $nsMgr = New-SsNsManager $sheetXml
    $numCols = $MAPPING_COLUMNS.Count

    foreach ($rowNode in $sheetXml.SelectNodes("//s:sheetData/s:row", $nsMgr)) {
        $rIdx = [int]$rowNode.GetAttribute("r")
        if ($rIdx -eq 1) { continue }  # salta intestazione

        $values = New-Object string[] $numCols
        for ($i = 0; $i -lt $numCols; $i++) { $values[$i] = "" }

        foreach ($cellNode in $rowNode.SelectNodes("s:c", $nsMgr)) {
            $ref = $cellNode.GetAttribute("r")
            $colIdx = Get-ColumnIndexFromRef $ref
            if ($null -eq $colIdx -or $colIdx -lt 1 -or $colIdx -gt $numCols) { continue }
            $values[$colIdx - 1] = Get-CellValue $cellNode $sharedStrings $nsMgr
        }

        $allEmpty = $true
        foreach ($v in $values) { if ($v -ne "") { $allEmpty = $false; break } }
        if ($allEmpty) { continue }

        $record = [ordered]@{}
        for ($i = 0; $i -lt $numCols; $i++) { $record[$MAPPING_COLUMNS[$i]] = $values[$i] }
        [void]$rows.Add([PSCustomObject]$record)
    }
    return ,$rows
}

# ============================================================
# Pool di riferimento: indicizza uuid Workspace/Node-cliente e
# gli elementi Service/Messageserver/Router gia' esistenti, per
# continuita' e per recuperare gli attributi che l'Excel non ha.
# ============================================================
function Build-ReferencePool([string]$referenceFolder) {
    $pool = [PSCustomObject]@{
        Services         = @{}   # uuid -> XmlElement <Service> (nel documento originale)
        Messageservers   = @{}
        Routers          = @{}
        ClientWsUuid     = @{}   # nome cliente -> uuid Workspace
        ClientNodeUuid   = @{}   # nome cliente -> uuid Node
        InternalWsUuid   = $null
        InternalNodeUuid = @{}   # nome sotto-nodo (es. "1.TECHEDGE") -> uuid
    }
    if (-not (Test-Path -LiteralPath $referenceFolder)) { return $pool }

    Get-ChildItem -LiteralPath $referenceFolder -Directory | ForEach-Object {
        $cname = $_.Name
        $f = Join-Path $_.FullName "SAPUILandscape.xml"
        if (-not (Test-Path -LiteralPath $f)) { return }

        $doc = New-Object System.Xml.XmlDocument
        $doc.Load($f)
        $root = $doc.DocumentElement

        foreach ($s in $root.SelectNodes("Services/Service")) {
            $u = $s.GetAttribute("uuid")
            if ($u) { $pool.Services[$u] = $s }
        }
        foreach ($m in $root.SelectNodes("Messageservers/Messageserver")) {
            $u = $m.GetAttribute("uuid")
            if ($u) { $pool.Messageservers[$u] = $m }
        }
        foreach ($r in $root.SelectNodes("Routers/Router")) {
            $u = $r.GetAttribute("uuid")
            if ($u) { $pool.Routers[$u] = $r }
        }

        $ws = $root.SelectSingleNode("Workspaces/Workspace")
        if ($null -eq $ws) { return }

        if ($cname -eq "_AVVALE_INTERNAL") {
            $pool.InternalWsUuid = $ws.GetAttribute("uuid")
            foreach ($n in $ws.SelectNodes("Node")) {
                $pool.InternalNodeUuid[$n.GetAttribute("name")] = $n.GetAttribute("uuid")
            }
        }
        else {
            $node = $ws.SelectSingleNode("Node")
            if ($null -ne $node) {
                $pool.ClientWsUuid[$cname] = $ws.GetAttribute("uuid")
                $pool.ClientNodeUuid[$cname] = $node.GetAttribute("uuid")
            }
        }
    }
    return $pool
}

function Get-XmlCommentSafeText([string]$text) {
    while ($text -match "--") { $text = $text -replace "--", "- " }
    if ($text.EndsWith("-")) { $text = $text + " " }
    return $text
}

# ricostruisce (o recupera dal pool) il <Service> per una riga, con
# l'eventuale <Messageserver>/<Router> collegato
function Resolve-ServiceForRow([System.Xml.XmlDocument]$doc, $row, $pool) {
    $svcUuidRow = $row.UuidService
    $name = $row.NomeSistemaStandard
    if (-not $name) { $name = $row.NomeSistema }
    $systemid = $row.SystemId
    $server = $row.Server

    $original = $null
    if ($svcUuidRow -and $pool.Services.ContainsKey($svcUuidRow)) {
        $original = $pool.Services[$svcUuidRow]
    }

    if ($null -ne $original) {
        $svcNode = $doc.ImportNode($original, $true)
        if ($name) { $svcNode.SetAttribute("name", $name) }
        if ($systemid) { $svcNode.SetAttribute("systemid", $systemid) }
        if ($server) { $svcNode.SetAttribute("server", $server) }

        $msNode = $null
        $msid = $svcNode.GetAttribute("msid")
        if ($msid -and $pool.Messageservers.ContainsKey($msid)) {
            $msNode = $doc.ImportNode($pool.Messageservers[$msid], $true)
        }
        $rtNode = $null
        $routerid = $svcNode.GetAttribute("routerid")
        if ($routerid -and $pool.Routers.ContainsKey($routerid)) {
            $rtNode = $doc.ImportNode($pool.Routers[$routerid], $true)
        }
        return [PSCustomObject]@{ ServiceNode = $svcNode; ServiceUuid = $svcUuidRow; MsNode = $msNode; RtNode = $rtNode }
    }

    # nessun Service originale trovato nel pool: ne creo uno minimo da zero
    $newUuid = $svcUuidRow
    if (-not $newUuid) { $newUuid = [guid]::NewGuid().ToString() }
    $svcNode = $doc.CreateElement("Service")
    $svcNode.SetAttribute("type", "SAPGUI")
    $svcNode.SetAttribute("uuid", $newUuid)
    $svcNode.SetAttribute("name", $name)
    $svcNode.SetAttribute("systemid", $systemid)
    $svcNode.SetAttribute("mode", "1")
    $svcNode.SetAttribute("server", $server)
    $svcNode.SetAttribute("sncop", "-1")
    $svcNode.SetAttribute("dcpg", "2")
    return [PSCustomObject]@{ ServiceNode = $svcNode; ServiceUuid = $newUuid; MsNode = $null; RtNode = $null }
}

function Add-ItemWithComment([System.Xml.XmlDocument]$doc, $parent, [string]$itemUuid, [string]$svcUuid, $svcNode) {
    $parts = @($svcNode.GetAttribute("name"))
    if ($svcNode.GetAttribute("systemid")) { $parts += "[" + $svcNode.GetAttribute("systemid") + "]" }
    if ($svcNode.GetAttribute("server")) { $parts += $svcNode.GetAttribute("server") }
    $commentText = Get-XmlCommentSafeText (" " + ($parts -join " ") + " ")
    [void]$parent.AppendChild($doc.CreateComment($commentText))
    $item = $doc.CreateElement("Item")
    $item.SetAttribute("uuid", $itemUuid)
    $item.SetAttribute("serviceid", $svcUuid)
    [void]$parent.AppendChild($item)
}

function New-ClientDocument([string]$clientName, $rows, $pool) {
    $doc = New-Object System.Xml.XmlDocument
    $decl = $doc.CreateXmlDeclaration("1.0", "utf-8", $null)
    [void]$doc.AppendChild($decl)

    $landscape = $doc.CreateElement("Landscape")
    $landscape.SetAttribute("updated", $Updated)
    $landscape.SetAttribute("version", $Version)
    $landscape.SetAttribute("generator", $Generator)
    [void]$doc.AppendChild($landscape)

    $wsUuid = $pool.ClientWsUuid[$clientName]
    $nodeUuid = $pool.ClientNodeUuid[$clientName]
    if (-not $wsUuid) { $wsUuid = [guid]::NewGuid().ToString() }
    if (-not $nodeUuid) { $nodeUuid = [guid]::NewGuid().ToString() }

    $workspacesEl = $doc.CreateElement("Workspaces")
    [void]$landscape.AppendChild($workspacesEl)
    $workspace = $doc.CreateElement("Workspace")
    $workspace.SetAttribute("uuid", $wsUuid)
    $workspace.SetAttribute("name", $clientName)
    $workspace.SetAttribute("expanded", "0")
    $workspace.SetAttribute("hidden", "0")
    [void]$workspacesEl.AppendChild($workspace)

    $clientNode = $doc.CreateElement("Node")
    $clientNode.SetAttribute("uuid", $nodeUuid)
    $clientNode.SetAttribute("name", $clientName)
    $clientNode.SetAttribute("hidden", "0")
    [void]$workspace.AppendChild($clientNode)

    $servicesUsed = [ordered]@{}
    $msUsed = [ordered]@{}
    $rtUsed = [ordered]@{}
    $itemEntries = New-Object System.Collections.Generic.List[object]

    foreach ($row in $rows) {
        $itemUuid = $row.UuidItem
        if (-not $itemUuid) { $itemUuid = [guid]::NewGuid().ToString() }

        $resolved = Resolve-ServiceForRow $doc $row $pool
        if (-not $servicesUsed.Contains($resolved.ServiceUuid)) {
            $servicesUsed[$resolved.ServiceUuid] = $resolved.ServiceNode
        }
        if ($null -ne $resolved.MsNode) {
            $msU = $resolved.MsNode.GetAttribute("uuid")
            if (-not $msUsed.Contains($msU)) { $msUsed[$msU] = $resolved.MsNode }
        }
        if ($null -ne $resolved.RtNode) {
            $rtU = $resolved.RtNode.GetAttribute("uuid")
            if (-not $rtUsed.Contains($rtU)) { $rtUsed[$rtU] = $resolved.RtNode }
        }

        $tipo = $row.TipoSistema
        if (-not $tipo) { $tipo = "Altri sistemi" }

        [void]$itemEntries.Add([PSCustomObject]@{
            ItemUuid = $itemUuid; ServiceUuid = $resolved.ServiceUuid; ServiceNode = $resolved.ServiceNode; Tipo = $tipo
        })
    }

    $buckets = [ordered]@{}
    foreach ($e in $itemEntries) {
        if (-not $buckets.Contains($e.Tipo)) { $buckets[$e.Tipo] = New-Object System.Collections.Generic.List[object] }
        [void]$buckets[$e.Tipo].Add($e)
    }
    # Raggruppa SEMPRE per "Tipo sistema" (anche con un solo sistema o un solo
    # tipo), cosi' una modifica alla colonna "Tipo sistema" nella mappa Excel si
    # traduce sempre in una sotto-cartella "Sistemi <Tipo>" visibile in SAP Logon
    # dopo rebuild+deploy - prima veniva creata solo con >2 sistemi E >1 tipo
    # diverso, e i client "mono-tipo" restavano piatti (poi avvolti da
    # Build-SAPGlobal.ps1 in un generico "Sistemi" che ignora il tipo).
    $shouldGroup = $itemEntries.Count -gt 0

    if ($shouldGroup) {
        $labels = @($buckets.Keys) | Sort-Object { Get-TypeSortIndex $_ }
        foreach ($label in $labels) {
            $subUuid = [guid]::NewGuid().ToString()  # nuovo uuid ad ogni rigenerazione: nessun requisito di continuita' per questi sotto-nodi
            $subName = if ($label -eq "Altri sistemi") { "Altri sistemi" } else { "Sistemi $label" }
            $subNode = $doc.CreateElement("Node")
            $subNode.SetAttribute("uuid", $subUuid)
            $subNode.SetAttribute("name", $subName)
            $subNode.SetAttribute("hidden", "0")
            [void]$clientNode.AppendChild($subNode)
            foreach ($e in $buckets[$label]) {
                Add-ItemWithComment $doc $subNode $e.ItemUuid $e.ServiceUuid $e.ServiceNode
            }
        }
    }
    else {
        foreach ($e in $itemEntries) {
            Add-ItemWithComment $doc $clientNode $e.ItemUuid $e.ServiceUuid $e.ServiceNode
        }
    }

    $svcContainer = $doc.CreateElement("Services")
    [void]$landscape.AppendChild($svcContainer)
    foreach ($k in $servicesUsed.Keys) { [void]$svcContainer.AppendChild($servicesUsed[$k]) }

    $msContainer = $doc.CreateElement("Messageservers")
    [void]$landscape.AppendChild($msContainer)
    foreach ($k in $msUsed.Keys) { [void]$msContainer.AppendChild($msUsed[$k]) }

    $rtContainer = $doc.CreateElement("Routers")
    [void]$landscape.AppendChild($rtContainer)
    foreach ($k in $rtUsed.Keys) { [void]$rtContainer.AppendChild($rtUsed[$k]) }

    # stessa cautela IEnumerable di Get-ZipEntryXml/New-SsNsManager
    return ,$doc
}

function New-InternalDocument($rows, $pool) {
    # Foglio "Avvale (interno)": la Workspace "AVVALE (interno)" contiene
    # DIRETTAMENTE piu' Node fratelli (uno per ogni valore distinto di
    # "Percorso cartella", es. "1.TECHEDGE", "2.AVVALE") - NON un unico
    # Node-cliente che li avvolge, a differenza di tutti gli altri client.
    $doc = New-Object System.Xml.XmlDocument
    $decl = $doc.CreateXmlDeclaration("1.0", "utf-8", $null)
    [void]$doc.AppendChild($decl)

    $landscape = $doc.CreateElement("Landscape")
    $landscape.SetAttribute("updated", $Updated)
    $landscape.SetAttribute("version", $Version)
    $landscape.SetAttribute("generator", $Generator)
    [void]$doc.AppendChild($landscape)

    $wsUuid = $pool.InternalWsUuid
    if (-not $wsUuid) { $wsUuid = [guid]::NewGuid().ToString() }

    $workspacesEl = $doc.CreateElement("Workspaces")
    [void]$landscape.AppendChild($workspacesEl)
    $workspace = $doc.CreateElement("Workspace")
    $workspace.SetAttribute("uuid", $wsUuid)
    $workspace.SetAttribute("name", "AVVALE (interno)")
    $workspace.SetAttribute("expanded", "0")
    $workspace.SetAttribute("hidden", "0")
    [void]$workspacesEl.AppendChild($workspace)

    $byPath = [ordered]@{}
    foreach ($row in $rows) {
        $path = $row.PercorsoCartella
        if (-not $path) { $path = "AVVALE" }
        if (-not $byPath.Contains($path)) { $byPath[$path] = New-Object System.Collections.Generic.List[object] }
        [void]$byPath[$path].Add($row)
    }

    $servicesUsed = [ordered]@{}
    $msUsed = [ordered]@{}
    $rtUsed = [ordered]@{}

    foreach ($path in $byPath.Keys) {
        $nodeUuid = $pool.InternalNodeUuid[$path]
        if (-not $nodeUuid) { $nodeUuid = [guid]::NewGuid().ToString() }
        $node = $doc.CreateElement("Node")
        $node.SetAttribute("uuid", $nodeUuid)
        $node.SetAttribute("name", $path)
        $node.SetAttribute("hidden", "0")
        [void]$workspace.AppendChild($node)

        foreach ($row in $byPath[$path]) {
            $itemUuid = $row.UuidItem
            if (-not $itemUuid) { $itemUuid = [guid]::NewGuid().ToString() }
            $resolved = Resolve-ServiceForRow $doc $row $pool
            if (-not $servicesUsed.Contains($resolved.ServiceUuid)) {
                $servicesUsed[$resolved.ServiceUuid] = $resolved.ServiceNode
            }
            if ($null -ne $resolved.MsNode) {
                $msU = $resolved.MsNode.GetAttribute("uuid")
                if (-not $msUsed.Contains($msU)) { $msUsed[$msU] = $resolved.MsNode }
            }
            if ($null -ne $resolved.RtNode) {
                $rtU = $resolved.RtNode.GetAttribute("uuid")
                if (-not $rtUsed.Contains($rtU)) { $rtUsed[$rtU] = $resolved.RtNode }
            }
            Add-ItemWithComment $doc $node $itemUuid $resolved.ServiceUuid $resolved.ServiceNode
        }
    }

    $svcContainer = $doc.CreateElement("Services")
    [void]$landscape.AppendChild($svcContainer)
    foreach ($k in $servicesUsed.Keys) { [void]$svcContainer.AppendChild($servicesUsed[$k]) }
    $msContainer = $doc.CreateElement("Messageservers")
    [void]$landscape.AppendChild($msContainer)
    foreach ($k in $msUsed.Keys) { [void]$msContainer.AppendChild($msUsed[$k]) }
    $rtContainer = $doc.CreateElement("Routers")
    [void]$landscape.AppendChild($rtContainer)
    foreach ($k in $rtUsed.Keys) { [void]$rtContainer.AppendChild($rtUsed[$k]) }

    return ,$doc
}

function Save-ClientDocument([System.Xml.XmlDocument]$doc, [string]$path) {
    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Indent = $true
    $settings.IndentChars = "`t"
    $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
    $settings.OmitXmlDeclaration = $false
    $writer = [System.Xml.XmlWriter]::Create($path, $settings)
    try { $doc.Save($writer) } finally { $writer.Close() }
}

# ============================================================
# MAIN
# ============================================================
if (-not (Test-Path -LiteralPath $XlsxPath)) {
    Write-Host "ERRORE: file Excel non trovato: $XlsxPath" -ForegroundColor Red
    exit 1
}

Write-Host "Apro $XlsxPath ..."
$zip = [System.IO.Compression.ZipFile]::OpenRead((Resolve-Path $XlsxPath))
try {
    $wbXml = Get-ZipEntryXml $zip "xl/workbook.xml"
    if ($null -eq $wbXml) { throw "xl/workbook.xml non trovato: il file non sembra un .xlsx valido." }
    $wbNsMgr = New-SsNsManager $wbXml
    $relsXml = Get-ZipEntryXml $zip "xl/_rels/workbook.xml.rels"

    $mappaTarget = Get-SheetTarget $wbXml $wbNsMgr $relsXml "Mappa sistemi"
    if ($null -eq $mappaTarget) { throw "Foglio 'Mappa sistemi' non trovato nel file Excel." }
    $internoTarget = Get-SheetTarget $wbXml $wbNsMgr $relsXml "Avvale (interno)"

    $sharedStrings = Get-SharedStrings $zip

    Write-Host "Leggo il foglio 'Mappa sistemi' ..."
    $clientRows = Read-SheetRows $zip $mappaTarget $sharedStrings
    Write-Host "  righe lette: $($clientRows.Count)"

    $internalRows = New-Object System.Collections.Generic.List[object]
    if ($null -ne $internoTarget) {
        Write-Host "Leggo il foglio 'Avvale (interno)' ..."
        $internalRows = Read-SheetRows $zip $internoTarget $sharedStrings
        Write-Host "  righe lette: $($internalRows.Count)"
    }
}
finally {
    $zip.Dispose()
}

if ($clientRows.Count -gt 0) {
    Write-Host "`nEsempio prima riga letta (controllo di coerenza):"
    $first = $clientRows[0]
    Write-Host "  Cliente=$($first.Cliente) | Nome sistema=$($first.NomeSistema) | System ID=$($first.SystemId) | Server=$($first.Server) | Tipo sistema=$($first.TipoSistema)"
}

Write-Host "`nIndicizzo la cartella di riferimento: $ReferenceFolder ..."
$pool = Build-ReferencePool $ReferenceFolder
Write-Host "  clienti indicizzati: $($pool.ClientWsUuid.Keys.Count) | Service noti: $($pool.Services.Keys.Count)"
if ($pool.ClientWsUuid.Keys.Count -eq 0) {
    Write-Host "ATTENZIONE: nessun cliente trovato nella cartella di riferimento ($ReferenceFolder)." -ForegroundColor Yellow
    Write-Host "  Controlla il percorso: di default lo script cerca '.\SAPUILandscape_sorgenti'" -ForegroundColor Yellow
    Write-Host "  nella cartella da cui viene lanciato. Se SAPUILandscape_sorgenti si trova" -ForegroundColor Yellow
    Write-Host "  altrove (es. una cartella piu' in alto), rilancia con -ReferenceFolder <percorso>." -ForegroundColor Yellow
    Write-Host "  Senza reference, gli uuid di continuita' e gli attributi Service non vengono" -ForegroundColor Yellow
    Write-Host "  recuperati: i sorgenti verrebbero ricostruiti da zero per ogni sistema." -ForegroundColor Yellow
}

if (Test-Path -LiteralPath $OutFolder) {
    Remove-Item -LiteralPath $OutFolder -Recurse -Force
}
New-Item -ItemType Directory -Path $OutFolder | Out-Null

$byClient = [ordered]@{}
foreach ($row in $clientRows) {
    $cname = $row.Cliente
    if (-not $cname) { continue }
    if (-not $byClient.Contains($cname)) { $byClient[$cname] = New-Object System.Collections.Generic.List[object] }
    [void]$byClient[$cname].Add($row)
}

Write-Host "`nRicostruisco i sorgenti per cliente ..."
$nClients = 0
# NOTA IMPORTANTE: i nomi delle variabili in PowerShell NON distinguono
# maiuscole/minuscole - "$outFolder" e "$OutFolder" sono la STESSA
# variabile. Usare qui una variabile locale chiamata "$outFolder" andrebbe
# quindi a SOVRASCRIVERE silenziosamente il parametro "$OutFolder" ad ogni
# iterazione, facendo crescere il percorso base ad ogni ciclo (bug reale
# gia' riscontrato: cartelle annidate con tutti i nomi cliente concatenati,
# fino a superare il limite di lunghezza percorso di Windows). Per questo
# la variabile locale si chiama "$clientDestDir", nome che non collide.
foreach ($cname in $byClient.Keys) {
    $doc = New-ClientDocument $cname $byClient[$cname] $pool
    $clientDestDir = Join-Path $OutFolder $cname
    New-Item -ItemType Directory -Path $clientDestDir -Force | Out-Null
    Save-ClientDocument $doc (Join-Path $clientDestDir "SAPUILandscape.xml")
    $nClients++
}

if ($internalRows.Count -gt 0) {
    $doc = New-InternalDocument $internalRows $pool
    $internalDestDir = Join-Path $OutFolder "_AVVALE_INTERNAL"
    New-Item -ItemType Directory -Path $internalDestDir -Force | Out-Null
    Save-ClientDocument $doc (Join-Path $internalDestDir "SAPUILandscape.xml")
}

# --- clienti presenti nella reference ma senza NESSUNA riga in Excel ---
# Di default sono considerati RIMOSSI DI PROPOSITO (l'utente li ha tolti
# dall'Excel) e la loro cartella NON viene ricopiata in $OutFolder: e'
# cosi' che cancellare le righe di un cliente dall'Excel si traduce
# davvero nella sua sparizione dopo rebuild + deploy, invece di farlo
# ricomparire identico ad ogni rebuild successivo (bug osservato: un
# cliente rimane fisicamente nella reference con lo stesso Copy-Item, che
# ne preserva pure il timestamp originale, quindi il rebuild successivo lo
# ritrova ancora la' e lo ricopia di nuovo all'infinito).
# Caso limite: un cliente con DAVVERO zero sistemi non e' rappresentabile
# in un foglio "un rigo per sistema" e finirebbe nello stesso elenco pur
# non essendo stato rimosso apposta - per questo l'elenco dei clienti
# esclusi viene SEMPRE stampato per esteso (mai un semplice conteggio), e
# chi ha bisogno del vecchio comportamento (ricopiarli invariati) puo'
# rilanciare lo script con -PreserveClientsWithoutRows.
$missingClients = New-Object System.Collections.Generic.List[string]
if (Test-Path -LiteralPath $ReferenceFolder) {
    foreach ($cname in $pool.ClientWsUuid.Keys) {
        if ($byClient.Contains($cname)) { continue }
        $src = Join-Path $ReferenceFolder $cname
        if (Test-Path -LiteralPath $src) { [void]$missingClients.Add($cname) }
    }
}

if ($missingClients.Count -gt 0) {
    if ($PreserveClientsWithoutRows) {
        Write-Host "`nClienti senza righe in Excel, ricopiati invariati dalla reference (-PreserveClientsWithoutRows attivo):" -ForegroundColor Yellow
        foreach ($cname in $missingClients) {
            $src = Join-Path $ReferenceFolder $cname
            $dst = Join-Path $OutFolder $cname
            Copy-Item -LiteralPath $src -Destination $dst -Recurse
            $nClients++
            Write-Host "  - $cname"
        }
    }
    else {
        Write-Host "`nClienti senza righe in Excel, considerati RIMOSSI (cartella NON ricopiata in $OutFolder):" -ForegroundColor Yellow
        foreach ($cname in $missingClients) { Write-Host "  - $cname" }
        Write-Host "Se qualcuno di questi ha invece DAVVERO zero sistemi ed e' da tenere, rilancia con -PreserveClientsWithoutRows." -ForegroundColor Yellow
    }
}

Write-Host "`nClienti ricostruiti: $nClients (+ _AVVALE_INTERNAL: $(if ($internalRows.Count -gt 0) { 'si' } else { 'no' }))"

# --- verifiche finali: XML ben formato + nessuna collisione di uuid + nessun item orfano ---
Write-Host "`nVerifiche finali:"
$allUuids = @{}
$collisions = 0
$orphans = 0
$badXml = 0
# NOTA: uso deliberatamente "foreach" (parola chiave di linguaggio, stesso
# scope del chiamante) e NON "Get-ChildItem | ForEach-Object { }" (il cui
# scriptblock gira in un NUOVO scope figlio): dentro il blocco incremento
# variabili scalari come $badXml/$collisions/$orphans con "++", e dentro
# uno scriptblock di ForEach-Object quell'incremento aggiornerebbe solo una
# copia locale, lasciando silenziosamente i contatori esterni sempre a 0.
foreach ($dirInfo in (Get-ChildItem -LiteralPath $OutFolder -Directory)) {
    $cname = $dirInfo.Name
    $f = Join-Path $dirInfo.FullName "SAPUILandscape.xml"
    if (-not (Test-Path -LiteralPath $f)) { continue }
    $doc = New-Object System.Xml.XmlDocument
    try {
        $doc.Load($f)
    }
    catch {
        Write-Host "  ERRORE XML malformato in $cname : $($_.Exception.Message)" -ForegroundColor Red
        $badXml++
        continue
    }
    $root = $doc.DocumentElement
    $svcIds = New-Object System.Collections.Generic.HashSet[string]
    foreach ($s in $root.SelectNodes("Services/Service")) {
        $u = $s.GetAttribute("uuid")
        if ($u) { [void]$svcIds.Add($u) }
    }
    foreach ($it in $root.SelectNodes(".//Item")) {
        $sid = $it.GetAttribute("serviceid")
        if (-not $sid -or -not $svcIds.Contains($sid)) { $orphans++ }
    }
    foreach ($el in $root.SelectNodes("//*[@uuid]")) {
        $u = $el.GetAttribute("uuid")
        if ($allUuids.ContainsKey($u) -and $allUuids[$u] -ne $cname) { $collisions++ }
        $allUuids[$u] = $cname
    }
}
Write-Host "  file XML malformati: $badXml"
Write-Host "  uuid totali indicizzati: $($allUuids.Keys.Count) | collisioni: $collisions"
Write-Host "  item con serviceid non risolvibile: $orphans"

$rebuildOk = ($badXml -eq 0 -and $collisions -eq 0 -and $orphans -eq 0)
if ($rebuildOk) {
    Write-Host "`nTutto ok. Sorgenti pronti in: $OutFolder" -ForegroundColor Green
}
else {
    Write-Host "`nATTENZIONE: controlla gli avvisi sopra prima di usare questi sorgenti." -ForegroundColor Yellow
}

# --- -ReplaceSources: sostituisce da sola la cartella sorgenti live con
# quella appena ricostruita, dopo averne fatto un backup con timestamp.
# Va fatto SOLO se le verifiche sopra sono pulite: altrimenti si rischia
# di sostituire i sorgenti veri con qualcosa di rotto.
if ($ReplaceSources) {
    if (-not $rebuildOk) {
        Write-Host "`n-ReplaceSources richiesto ma le verifiche hanno rilevato problemi (vedi sopra): NON sostituisco $ReferenceFolder." -ForegroundColor Red
        Write-Host "I sorgenti ricostruiti restano solo in $OutFolder per un controllo manuale." -ForegroundColor Red
    }
    elseif (-not (Test-Path -LiteralPath $ReferenceFolder)) {
        Write-Host "`n-ReplaceSources richiesto ma $ReferenceFolder non esiste: niente da sostituire." -ForegroundColor Yellow
    }
    else {
        $refFull = (Resolve-Path -LiteralPath $ReferenceFolder).ProviderPath
        $refLeaf = Split-Path -Leaf $refFull
        $refParent = Split-Path -Parent $refFull
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $backupName = "${refLeaf}_$timestamp"
        $backupPath = Join-Path $refParent $backupName

        if (Test-Path -LiteralPath $backupPath) {
            Write-Host "`nATTENZIONE: esiste gia' una cartella $backupPath - -ReplaceSources annullato per sicurezza (rilancia tra qualche secondo)." -ForegroundColor Red
        }
        else {
            Write-Host "`nSostituzione in corso (-ReplaceSources) ..."
            Rename-Item -LiteralPath $refFull -NewName $backupName
            Write-Host "  backup del vecchio '$refLeaf' -> '$backupName'"
            Move-Item -LiteralPath $OutFolder -Destination $refFull
            Write-Host "  '$OutFolder' -> '$refFull' (nuova cartella sorgenti attiva)"
            Write-Host "`nFatto. Ora puoi lanciare Build-SAPGlobal.cmd e Deploy-SAPLandscape.cmd." -ForegroundColor Green
        }
    }
}
