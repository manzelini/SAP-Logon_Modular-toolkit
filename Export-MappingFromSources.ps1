# Export-MappingFromSources.ps1
# ================================
# Ricostruisce mappa_uuid_sistemi.xlsx leggendo SAPUILandscape_sorgenti (le
# fonti XML per cliente, che sono il vero "stato attuale" del landscape) -
# la direzione OPPOSTA di Rebuild-SourcesFromMapping.ps1 (che va invece da
# Excel verso XML). Utile per ricostruire/riallineare la mappa se va persa o
# corrotta, o se si sospetta sia rimasta indietro rispetto a modifiche fatte
# a mano sui sorgenti (§2.6.2 del manuale).
#
# USO:
#   powershell -ExecutionPolicy Bypass -File Export-MappingFromSources.ps1 `
#       -SourcesFolder ".\SAPUILandscape_sorgenti" `
#       -ExistingMappingPath ".\mappa_uuid_sistemi.xlsx" `
#       -OutPath ".\mappa_uuid_sistemi_REBUILT_20260814_101500.xlsx"
#
# oppure semplicemente Export-MappingFromSources.cmd a doppio clic se i tre
# elementi si trovano gia' con i nomi di default nella stessa cartella.
#
# IMPORTANTE - stesso principio di sicurezza di Rebuild-SourcesFromMapping.ps1:
# questo script NON sovrascrive mai mappa_uuid_sistemi.xlsx direttamente.
# Scrive sempre in un file NUOVO, con un timestamp nel nome per non
# sovrascrivere nemmeno le esecuzioni precedenti (default
# mappa_uuid_sistemi_REBUILT_yyyyMMdd_HHmmss.xlsx, es.
# mappa_uuid_sistemi_REBUILT_20260814_101500.xlsx): controllane il
# contenuto, poi sostituisci tu stesso il file originale quando sei sicuro
# che vada bene.
#
# COSA VIENE RICOSTRUITO DA ZERO (autoritativo, preso sempre dall'XML):
#   Cliente, Percorso cartella, Nome sistema, System ID, Server,
#   Message Server, Router, UUID Item, UUID Service.
#
# COSA VIENE RICONCILIATO CON LA MAPPA ESISTENTE (se fornita, tramite
# -ExistingMappingPath) invece di essere rigenerato da zero:
#   Nome sistema standard, Da verificare, Tipo sistema (per _AVVALE_INTERNAL,
#   dove non esiste raggruppamento per tipo nell'XML).
#   La mappa XML infatti registra un solo nome per sistema (quello attuale):
#   non puo' da sola ricostruire una proposta di rinomina gia' elaborata a
#   mano in passato per un sistema non ancora verificato. Per questo, se un
#   UUID Service esiste gia' nella mappa esistente E il nome corrente in XML
#   non e' cambiato, il "Nome sistema standard" e il flag "Da verificare"
#   vengono riportati invariati dalla mappa esistente. Solo per i sistemi
#   NUOVI o con nome cambiato rispetto alla mappa esistente, il flag viene
#   ricalcolato: se il nome corrente rispetta gia' esattamente lo schema
#   "CLIENTE - SID - Ambiente" (Ambiente tra Production/Quality/Development/
#   Test/Sandbox) allora "Da verificare"="NO" e lo standard coincide col
#   nome corrente; altrimenti "Da verificare"="SI" e lo standard resta
#   provvisoriamente uguale al nome corrente, in attesa di verifica manuale
#   (stessa convenzione gia' in uso nella mappa esistente per le righe non
#   ancora risolte: NON si inventa mai un ambiente dalla sola iniziale del
#   SID, indicatore gia' verificato come inaffidabile - v. GESTIONE_MODULARE
#   _SAP_GUI.md).
#
# "Tipo sistema" per i client normali viene invece SEMPRE letto in modo
# affidabile dalla struttura XML stessa: dal nome del Node-tipo che contiene
# l'Item ("Sistemi <Tipo>" o "Altri sistemi"), che Rebuild-SourcesFromMapping
# .ps1/New-SAPClient.ps1 creano sempre (raggruppamento per tipo sempre attivo
# - v. §2.1 del manuale). Solo per _AVVALE_INTERNAL (che non ha
# raggruppamento per tipo nell'XML) e per eventuali sorgenti rimasti "piatti"
# (mai passati da Rebuild/New-SAPClient) si ricade sulla mappa esistente, o
# su "Altri sistemi" di default se non c'e' nessuna informazione pregressa.
#
# FOGLI E SEARCH HELP:
#   Ordine schede: "Riepilogo per cliente", "Mappa sistemi", "Avvale
#   (interno)". La colonna "Tipo sistema" di "Mappa sistemi" e "Avvale
#   (interno)" ha una search help (data validation a elenco, come una F4 SAP)
#   con i valori di "Tipo sistema" gia' presenti nei dati, in ordine
#   alfabetico e senza duplicati: l'elenco vive nel foglio di supporto
#   nascosto "Liste" (puoi farlo comparire con tasto destro su una scheda ->
#   Mostra... se vuoi controllarlo).
#
# Stesse convenzioni tecniche gia' in uso in Rebuild-SourcesFromMapping.ps1:
# - Ogni controllo di presenza di un nodo XML usa $null -eq/-ne, mai
#   "if ($node)".
# - Ogni funzione che restituisce un ARRAY lo fa con "return ,$array".
# - I "record" con piu' campi sono [PSCustomObject], mai Hashtable/array.
# - L'XML (lettura E scrittura) usa sempre le API di System.Xml
#   (CreateElement/SetAttribute/InnerText), mai concatenazione di stringhe:
#   l'escaping dei caratteri speciali (&, <, >, ", ') e' cosi' sempre
#   corretto e automatico.

param(
    [string]$SourcesFolder      = ".\SAPUILandscape_sorgenti",
    [string]$ExistingMappingPath = ".\mappa_uuid_sistemi.xlsx",
    [string]$OutPath            = (".\mappa_uuid_sistemi_REBUILT_{0}.xlsx" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression

$SS_NS  = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
$REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
$RID_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# "Solution Manager" e' incluso perche' e' un terzo segmento consolidato e
# sempre accettato per questi sistemi (nessuna eccezione osservata sui dati
# reali: 24/24 istanze con questo suffisso hanno "Da verificare"="NO"), pur
# non essendo un "Ambiente" in senso stretto. Altre varianti storiche non
# ancora standardizzate (es. "01 Sviluppo ECC", nomi in italiano, numerati)
# vengono deliberatamente NON riconosciute qui: per un sistema nuovo o il cui
# nome e' cambiato rispetto alla mappa precedente, verranno marcate "Da
# verificare"="SI" invece di essere accettate automaticamente, cosi' restano
# sempre visibili per una verifica umana anziche' sparire silenziosamente.
$ENV_CANONICI = @("Production", "Quality", "Development", "Test", "Sandbox", "Solution Manager")

# ============================================================
# Lettura di un .xlsx esistente (stesse funzioni, stesso approccio gia'
# testato in Rebuild-SourcesFromMapping.ps1 - qui servono solo per la
# RICONCILIAZIONE con la mappa esistente, se fornita).
# ============================================================
function New-SsNsManager([System.Xml.XmlDocument]$xmlDoc) {
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

# stesso ordine colonne 1..12 usato in Rebuild-SourcesFromMapping.ps1
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
        if ($rIdx -eq 1) { continue }

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

# Indicizza la mappa esistente per UUID Service, per la riconciliazione.
function Read-ExistingMapping([string]$path) {
    $byService = @{}
    if (-not (Test-Path -LiteralPath $path)) { return $byService }
    Write-Host "Leggo la mappa esistente per riconciliazione: $path ..."
    $zip = [System.IO.Compression.ZipFile]::OpenRead((Resolve-Path $path))
    try {
        $wbXml = Get-ZipEntryXml $zip "xl/workbook.xml"
        if ($null -eq $wbXml) { throw "xl/workbook.xml non trovato: il file non sembra un .xlsx valido." }
        $wbNsMgr = New-SsNsManager $wbXml
        $relsXml = Get-ZipEntryXml $zip "xl/_rels/workbook.xml.rels"
        $sharedStrings = Get-SharedStrings $zip

        foreach ($sheetName in @("Mappa sistemi", "Avvale (interno)")) {
            $target = Get-SheetTarget $wbXml $wbNsMgr $relsXml $sheetName
            if ($null -eq $target) { continue }
            $rows = Read-SheetRows $zip $target $sharedStrings
            foreach ($r in $rows) {
                if ($r.UuidService) { $byService[$r.UuidService] = $r }
            }
        }
    }
    finally {
        $zip.Dispose()
    }
    Write-Host "  righe indicizzate per riconciliazione: $($byService.Keys.Count)"
    return $byService
}

# ============================================================
# Lettura dei sorgenti XML (fonte autoritativa)
# ============================================================
function Get-XmlAttr($node, [string]$name) {
    if ($null -eq $node) { return "" }
    $v = $node.GetAttribute($name)
    if ($null -eq $v) { return "" }
    return $v
}

function Resolve-MessageServerValue($doc, [string]$msid) {
    if (-not $msid) { return "" }
    $ms = $doc.SelectSingleNode("/Landscape/Messageservers/Messageserver[@uuid='$msid']")
    if ($null -eq $ms) { return "" }
    $host_ = Get-XmlAttr $ms "host"
    $port  = Get-XmlAttr $ms "port"
    if ($host_ -and $port) { return "${host_}:$port" }
    if ($host_) { return $host_ }
    return Get-XmlAttr $ms "name"
}

function Resolve-RouterValue($doc, [string]$routerid) {
    if (-not $routerid) { return "" }
    $rt = $doc.SelectSingleNode("/Landscape/Routers/Router[@uuid='$routerid']")
    if ($null -eq $rt) { return "" }
    $r = Get-XmlAttr $rt "router"
    if ($r) { return $r }
    return Get-XmlAttr $rt "name"
}

# Costruisce una riga [PSCustomObject] per un <Item>, dato il documento del
# cliente, il nome del Node-tipo che lo contiene direttamente (o $null se e'
# figlio diretto del Node cliente/path), Cliente e PercorsoCartella.
function New-RowFromItem($doc, $itemNode, [string]$tipoNodeName, [string]$cliente, [string]$percorso, $existingByService) {
    $itemUuid = Get-XmlAttr $itemNode "uuid"
    $svcUuid  = Get-XmlAttr $itemNode "serviceid"
    if (-not $svcUuid) { return $null }
    $svc = $doc.SelectSingleNode("/Landscape/Services/Service[@uuid='$svcUuid']")
    if ($null -eq $svc) { return $null }

    $nome     = Get-XmlAttr $svc "name"
    $systemId = Get-XmlAttr $svc "systemid"
    $server   = Get-XmlAttr $svc "server"
    $msValue  = Resolve-MessageServerValue $doc (Get-XmlAttr $svc "msid")
    $rtValue  = Resolve-RouterValue $doc (Get-XmlAttr $svc "routerid")

    $existing = $null
    if ($existingByService.ContainsKey($svcUuid)) { $existing = $existingByService[$svcUuid] }

    # --- Tipo sistema ---
    if ($tipoNodeName) {
        if ($tipoNodeName -eq "Altri sistemi") { $tipo = "Altri sistemi" }
        elseif ($tipoNodeName.StartsWith("Sistemi ")) { $tipo = $tipoNodeName.Substring(8) }
        else { $tipo = $tipoNodeName }
    }
    elseif ($null -ne $existing -and $existing.TipoSistema) {
        $tipo = $existing.TipoSistema
    }
    else {
        $tipo = "Altri sistemi"
    }

    # --- Nome sistema standard / Da verificare ---
    if ($null -ne $existing -and $existing.NomeSistema -eq $nome -and $existing.NomeSistemaStandard) {
        $standard = $existing.NomeSistemaStandard
        $daVerificare = if ($existing.DaVerificare) { $existing.DaVerificare } else { "SI" }
    }
    else {
        $clienteEsc = [regex]::Escape($cliente)
        $sidEsc     = [regex]::Escape($systemId)
        $envAlt     = [string]::Join("|", $ENV_CANONICI)
        $pattern    = "^$clienteEsc - $sidEsc - ($envAlt)$"
        if ($nome -match $pattern) {
            $standard = $nome
            $daVerificare = "NO"
        }
        else {
            $standard = $nome
            $daVerificare = "SI"
        }
    }

    return [PSCustomObject]@{
        Cliente              = $cliente
        PercorsoCartella     = $percorso
        NomeSistema          = $nome
        SystemId             = $systemId
        Server               = $server
        MessageServer        = $msValue
        Router               = $rtValue
        UuidItem             = $itemUuid
        UuidService          = $svcUuid
        NomeSistemaStandard  = $standard
        DaVerificare         = $daVerificare
        TipoSistema          = $tipo
    }
}

function Read-ClientRows([string]$clientDir, [string]$clientName, $existingByService) {
    $rows = New-Object System.Collections.Generic.List[object]
    $f = Join-Path $clientDir "SAPUILandscape.xml"
    if (-not (Test-Path -LiteralPath $f)) { return ,$rows }
    $doc = New-Object System.Xml.XmlDocument
    $doc.Load($f)
    $clientNode = $doc.SelectSingleNode("/Landscape/Workspaces/Workspace/Node")
    if ($null -eq $clientNode) { return ,$rows }

    foreach ($child in $clientNode.ChildNodes) {
        if ($child.NodeType -ne "Element") { continue }
        # NOTA: si usa .LocalName, MAI .Name, per il nome del tag XML. Per un
        # XmlElement <Node name="..."> PowerShell restituirebbe da $child.Name
        # il VALORE dell'attributo "name" (es. "Altri sistemi"), non il nome
        # del tag "Node" - un confronto "$child.Name -eq 'Node'" risulterebbe
        # sempre falso e nessun sistema verrebbe mai letto (lezione gia'
        # documentata in Rebuild-SourcesFromMapping.ps1 e Build-SAPGlobal.ps1,
        # v. GESTIONE_MODULARE_SAP_GUI.md par. 7 - qui inizialmente non
        # applicata per errore).
        if ($child.LocalName -eq "Item") {
            # sorgente ancora "piatto" (non passato da Rebuild/New-SAPClient)
            $row = New-RowFromItem $doc $child $null $clientName $clientName $existingByService
            if ($null -ne $row) { [void]$rows.Add($row) }
        }
        elseif ($child.LocalName -eq "Node") {
            $tipoName = Get-XmlAttr $child "name"
            foreach ($itemNode in $child.SelectNodes("Item")) {
                $row = New-RowFromItem $doc $itemNode $tipoName $clientName $clientName $existingByService
                if ($null -ne $row) { [void]$rows.Add($row) }
            }
        }
    }
    return ,$rows
}

function Read-InternalRows([string]$internalDir, $existingByService) {
    $rows = New-Object System.Collections.Generic.List[object]
    $f = Join-Path $internalDir "SAPUILandscape.xml"
    if (-not (Test-Path -LiteralPath $f)) { return ,$rows }
    $doc = New-Object System.Xml.XmlDocument
    $doc.Load($f)
    $ws = $doc.SelectSingleNode("/Landscape/Workspaces/Workspace")
    if ($null -eq $ws) { return ,$rows }
    $clienteLabel = Get-XmlAttr $ws "name"
    if (-not $clienteLabel) { $clienteLabel = "AVVALE (interno)" }

    foreach ($pathNode in $ws.SelectNodes("Node")) {
        $percorso = Get-XmlAttr $pathNode "name"
        foreach ($itemNode in $pathNode.SelectNodes("Item")) {
            $row = New-RowFromItem $doc $itemNode $null $clienteLabel $percorso $existingByService
            if ($null -ne $row) { [void]$rows.Add($row) }
        }
    }
    return ,$rows
}

# ============================================================
# Scrittura xlsx (OOXML minimale, via System.Xml - nessuna libreria esterna)
# ============================================================
function New-SheetXml([string[]]$headers, [object[]]$rows, [string[]]$fields, [int]$headerStyle, [int]$dataStyle, [double[]]$colWidths, [string]$validationSqref = $null, [string]$validationFormula1 = $null) {
    $doc = New-Object System.Xml.XmlDocument
    $decl = $doc.CreateXmlDeclaration("1.0", "UTF-8", "yes")
    [void]$doc.AppendChild($decl)

    $worksheet = $doc.CreateElement("worksheet", $SS_NS)
    [void]$doc.AppendChild($worksheet)

    $nCols = $headers.Count
    $nRowsTotal = $rows.Count + 1
    $lastColLetter = [char](64 + $nCols)  # ok fino a 26 colonne, qui ne servono 12
    $dimension = $doc.CreateElement("dimension", $SS_NS)
    $dimension.SetAttribute("ref", "A1:$lastColLetter$nRowsTotal")
    [void]$worksheet.AppendChild($dimension)

    $sheetViews = $doc.CreateElement("sheetViews", $SS_NS)
    $sheetView = $doc.CreateElement("sheetView", $SS_NS)
    $sheetView.SetAttribute("workbookViewId", "0")
    $pane = $doc.CreateElement("pane", $SS_NS)
    $pane.SetAttribute("ySplit", "1")
    $pane.SetAttribute("topLeftCell", "A2")
    $pane.SetAttribute("activePane", "bottomLeft")
    $pane.SetAttribute("state", "frozen")
    [void]$sheetView.AppendChild($pane)
    $selection = $doc.CreateElement("selection", $SS_NS)
    $selection.SetAttribute("pane", "bottomLeft")
    $selection.SetAttribute("activeCell", "A2")
    $selection.SetAttribute("sqref", "A2")
    [void]$sheetView.AppendChild($selection)
    [void]$sheetViews.AppendChild($sheetView)
    [void]$worksheet.AppendChild($sheetViews)

    if ($null -ne $colWidths -and $colWidths.Count -gt 0) {
        $cols = $doc.CreateElement("cols", $SS_NS)
        for ($i = 0; $i -lt $colWidths.Count; $i++) {
            $col = $doc.CreateElement("col", $SS_NS)
            $col.SetAttribute("min", [string]($i + 1))
            $col.SetAttribute("max", [string]($i + 1))
            $col.SetAttribute("width", [string]$colWidths[$i])
            $col.SetAttribute("customWidth", "1")
            [void]$cols.AppendChild($col)
        }
        [void]$worksheet.AppendChild($cols)
    }

    $sheetData = $doc.CreateElement("sheetData", $SS_NS)
    [void]$worksheet.AppendChild($sheetData)

    function Add-Row($rowNum, [string[]]$values, [int]$style, [bool]$numeric0Based) {
        $rowEl = $doc.CreateElement("row", $SS_NS)
        $rowEl.SetAttribute("r", [string]$rowNum)
        for ($c = 0; $c -lt $values.Count; $c++) {
            $colLetter = [char](65 + $c)
            $cellEl = $doc.CreateElement("c", $SS_NS)
            $cellEl.SetAttribute("r", "$colLetter$rowNum")
            $cellEl.SetAttribute("s", [string]$style)
            $val = $values[$c]
            if ($numeric0Based -and $c -eq ($values.Count - 1) -and $val -match '^\d+$') {
                $vEl = $doc.CreateElement("v", $SS_NS)
                $vEl.InnerText = $val
                [void]$cellEl.AppendChild($vEl)
            }
            else {
                $cellEl.SetAttribute("t", "inlineStr")
                $isEl = $doc.CreateElement("is", $SS_NS)
                $tEl = $doc.CreateElement("t", $SS_NS)
                $tEl.InnerText = $val
                if ($val -ne $val.Trim()) { $tEl.SetAttribute("xml:space", "preserve") }
                [void]$isEl.AppendChild($tEl)
                [void]$cellEl.AppendChild($isEl)
            }
            [void]$rowEl.AppendChild($cellEl)
        }
        [void]$sheetData.AppendChild($rowEl)
    }

    Add-Row 1 $headers $headerStyle $false

    $r = 2
    foreach ($row in $rows) {
        $values = New-Object string[] $fields.Count
        for ($i = 0; $i -lt $fields.Count; $i++) { $values[$i] = [string]$row.($fields[$i]) }
        Add-Row $r $values $dataStyle $true
        $r++
    }

    $autoFilter = $doc.CreateElement("autoFilter", $SS_NS)
    $autoFilter.SetAttribute("ref", "A1:$lastColLetter$nRowsTotal")
    [void]$worksheet.AppendChild($autoFilter)

    # Search help (data validation a elenco) opzionale, es. per far scegliere
    # "Tipo sistema" da un elenco predefinito invece di digitarlo a mano -
    # va aggiunta subito dopo autoFilter per rispettare l'ordine richiesto
    # dallo schema OOXML del worksheet.
    if ($validationSqref -and $validationFormula1) {
        $dataValidations = $doc.CreateElement("dataValidations", $SS_NS)
        $dataValidations.SetAttribute("count", "1")
        $dv = $doc.CreateElement("dataValidation", $SS_NS)
        $dv.SetAttribute("type", "list")
        $dv.SetAttribute("allowBlank", "1")
        $dv.SetAttribute("showInputMessage", "1")
        $dv.SetAttribute("showErrorMessage", "1")
        $dv.SetAttribute("errorTitle", "Tipo sistema non valido")
        $dv.SetAttribute("error", "Scegli un valore dall'elenco (search help), oppure aggiungilo prima alla colonna Tipo sistema di un'altra riga e rigenera la mappa.")
        $dv.SetAttribute("sqref", $validationSqref)
        $f1 = $doc.CreateElement("formula1", $SS_NS)
        $f1.InnerText = $validationFormula1
        [void]$dv.AppendChild($f1)
        [void]$dataValidations.AppendChild($dv)
        [void]$worksheet.AppendChild($dataValidations)
    }

    return ,$doc
}

function Get-StylesXml() {
    $xml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="$SS_NS">
  <fonts count="3">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><sz val="10"/><name val="Arial"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Arial"/></font>
  </fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF2F5597"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF808080"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="4">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
  </cellXfs>
</styleSheet>
"@
    return $xml
}

function Get-ContentTypesXml() {
    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet4.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"@
}

function Get-RootRelsXml() {
    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="$REL_NS">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"@
}

function Get-WorkbookRelsXml() {
    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="$REL_NS">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet4.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"@
}

function Get-WorkbookXml([int]$tipoListLastRow) {
    # Ordine schede: Riepilogo per cliente prima, poi Mappa sistemi, poi
    # Avvale (interno) - i file fisici sheet1/2/3.xml restano invariati
    # (mappati da rId1/rId2/rId3), cambia solo l'ordine di comparsa nelle
    # linguette, che dipende unicamente dall'ordine degli elementi <sheet>
    # qui sotto. "Liste" (sheet4, r:id rId4) e' il foglio di supporto
    # nascosto con l'elenco univoco di "Tipo sistema" per il search help:
    # il foglio fisico viene sempre scritto (anche vuoto), ma il nome
    # definito (e quindi la search help stessa) viene aggiunto solo se
    # $tipoListLastRow -gt 0, cioe' se esiste almeno un valore da elencare.
    $definedNamesXml = ""
    if ($tipoListLastRow -gt 0) {
        $definedNamesXml = "`n  <definedNames>`n    <definedName name=`"ListaTipoSistema`">'Liste'!`$A`$2:`$A`$$tipoListLastRow</definedName>`n  </definedNames>"
    }
    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="$SS_NS" xmlns:r="$RID_NS">
  <sheets>
    <sheet name="Riepilogo per cliente" sheetId="1" r:id="rId2"/>
    <sheet name="Mappa sistemi" sheetId="2" r:id="rId1"/>
    <sheet name="Avvale (interno)" sheetId="3" r:id="rId3"/>
    <sheet name="Liste" sheetId="4" state="hidden" r:id="rId4"/>
  </sheets>$definedNamesXml
</workbook>
"@
}

function Get-CoreXml() {
    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>mappa_uuid_sistemi</dc:title>
  <dc:creator>Export-MappingFromSources.ps1</dc:creator>
</cp:coreProperties>
"@
}

function Get-AppXml() {
    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Export-MappingFromSources.ps1</Application>
</Properties>
"@
}

function Save-XmlDoc([System.Xml.XmlDocument]$doc, $entry) {
    $stream = $entry.Open()
    try {
        $settings = New-Object System.Xml.XmlWriterSettings
        $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
        $settings.OmitXmlDeclaration = $false
        $writer = [System.Xml.XmlWriter]::Create($stream, $settings)
        try { $doc.Save($writer) } finally { $writer.Close() }
    }
    finally {
        $stream.Close()
    }
}

function Add-TextEntry($zip, [string]$name, [string]$content) {
    $entry = $zip.CreateEntry($name)
    $stream = $entry.Open()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($content)
        $stream.Write($bytes, 0, $bytes.Length)
    }
    finally {
        $stream.Close()
    }
}

# ============================================================
# MAIN
# ============================================================
if (-not (Test-Path -LiteralPath $SourcesFolder)) {
    Write-Host "ERRORE: cartella sorgenti non trovata: $SourcesFolder" -ForegroundColor Red
    exit 1
}

$existingByService = Read-ExistingMapping $ExistingMappingPath

Write-Host "`nLeggo i sorgenti da: $SourcesFolder ..."
$allClientRows = New-Object System.Collections.Generic.List[object]
$internalRows  = New-Object System.Collections.Generic.List[object]
$nClients = 0

foreach ($dirInfo in (Get-ChildItem -LiteralPath $SourcesFolder -Directory | Sort-Object Name)) {
    $cname = $dirInfo.Name
    if ($cname -eq "_AVVALE_INTERNAL") {
        $rows = Read-InternalRows $dirInfo.FullName $existingByService
        foreach ($r in $rows) { [void]$internalRows.Add($r) }
        continue
    }
    $rows = Read-ClientRows $dirInfo.FullName $cname $existingByService
    if ($rows.Count -gt 0) { $nClients++ }
    foreach ($r in $rows) { [void]$allClientRows.Add($r) }
}

Write-Host "  clienti letti: $nClients | sistemi (Mappa sistemi): $($allClientRows.Count) | sistemi (Avvale interno): $($internalRows.Count)"

# ordina per Cliente (stabile, mantiene l'ordine di lettura all'interno di
# ciascun cliente, che segue gia' il raggruppamento per tipo dell'XML)
$allClientRows = @($allClientRows | Sort-Object Cliente)

# Riepilogo per cliente: conteggio righe per Cliente, ordine alfabetico
$riepilogo = @($allClientRows | Group-Object Cliente | Sort-Object Name | ForEach-Object {
    [PSCustomObject]@{ Cliente = $_.Name; NSistemi = $_.Count }
})

$daVerificareCount = @($allClientRows + $internalRows | Where-Object { $_.DaVerificare -eq "SI" }).Count
Write-Host "  sistemi con 'Da verificare = SI': $daVerificareCount"

# Elenco univoco (ordine alfabetico, nessun duplicato) dei valori di "Tipo
# sistema" gia' presenti nei dati letti, usato come search help per la
# colonna "Tipo sistema" nei fogli "Mappa sistemi" e "Avvale (interno)".
$tipiUnici = @(($allClientRows + $internalRows) | ForEach-Object { $_.TipoSistema } | Where-Object { $_ } | Sort-Object -Unique)
$listeRows = @($tipiUnici | ForEach-Object { [PSCustomObject]@{ TipoSistema = $_ } })
Write-Host "  valori univoci 'Tipo sistema' (search help): $($tipiUnici.Count)"

# ============================================================
# Scrittura del file xlsx
# ============================================================
$fullOutPath = [System.IO.Path]::GetFullPath($OutPath)
if (Test-Path -LiteralPath $fullOutPath) { Remove-Item -LiteralPath $fullOutPath -Force }

$zipStream = [System.IO.File]::Open($fullOutPath, [System.IO.FileMode]::CreateNew)
$zip = New-Object System.IO.Compression.ZipArchive($zipStream, [System.IO.Compression.ZipArchiveMode]::Create)
try {
    $tipoListLastRow = if ($tipiUnici.Count -gt 0) { $tipiUnici.Count + 1 } else { 0 }

    Add-TextEntry $zip "[Content_Types].xml" (Get-ContentTypesXml)
    Add-TextEntry $zip "_rels/.rels" (Get-RootRelsXml)
    Add-TextEntry $zip "xl/workbook.xml" (Get-WorkbookXml $tipoListLastRow)
    Add-TextEntry $zip "xl/_rels/workbook.xml.rels" (Get-WorkbookRelsXml)
    Add-TextEntry $zip "xl/styles.xml" (Get-StylesXml)
    Add-TextEntry $zip "docProps/core.xml" (Get-CoreXml)
    Add-TextEntry $zip "docProps/app.xml" (Get-AppXml)

    $headers = @("Cliente", "Percorso cartella", "Nome sistema", "System ID", "Server", "Message Server", "Router", "UUID Item", "UUID Service", "Nome sistema standard", "Da verificare", "Tipo sistema")
    $fields  = @("Cliente", "PercorsoCartella", "NomeSistema", "SystemId", "Server", "MessageServer", "Router", "UuidItem", "UuidService", "NomeSistemaStandard", "DaVerificare", "TipoSistema")
    $colWidths = @(20, 20, 36, 10, 30, 20, 20, 38, 38, 36, 13, 18)

    # search help sulla colonna L ("Tipo sistema") di "Mappa sistemi" e
    # "Avvale (interno)", solo sulle righe dati effettivamente presenti e
    # solo se esiste almeno un valore da proporre.
    $validFormula = if ($tipoListLastRow -gt 0) { "ListaTipoSistema" } else { $null }
    $sheet1ValidSqref = if ($tipoListLastRow -gt 0 -and $allClientRows.Count -gt 0) { "L2:L$($allClientRows.Count + 1)" } else { $null }
    $sheet3ValidSqref = if ($tipoListLastRow -gt 0 -and $internalRows.Count -gt 0) { "L2:L$($internalRows.Count + 1)" } else { $null }

    $sheet1 = New-SheetXml $headers $allClientRows $fields 2 1 $colWidths $sheet1ValidSqref $validFormula
    $entry1 = $zip.CreateEntry("xl/worksheets/sheet1.xml")
    Save-XmlDoc $sheet1 $entry1

    $riepHeaders = @("Cliente", "N. sistemi")
    $riepFields  = @("Cliente", "NSistemi")
    $riepWidths  = @(28, 12)
    $sheet2 = New-SheetXml $riepHeaders $riepilogo $riepFields 2 1 $riepWidths
    $entry2 = $zip.CreateEntry("xl/worksheets/sheet2.xml")
    Save-XmlDoc $sheet2 $entry2

    $sheet3 = New-SheetXml $headers $internalRows $fields 3 1 $colWidths $sheet3ValidSqref $validFormula
    $entry3 = $zip.CreateEntry("xl/worksheets/sheet3.xml")
    Save-XmlDoc $sheet3 $entry3

    $listeHeaders = @("Tipo sistema")
    $listeFields  = @("TipoSistema")
    $listeWidths  = @(30)
    $sheet4 = New-SheetXml $listeHeaders $listeRows $listeFields 2 1 $listeWidths
    $entry4 = $zip.CreateEntry("xl/worksheets/sheet4.xml")
    Save-XmlDoc $sheet4 $entry4
}
finally {
    $zip.Dispose()
    $zipStream.Close()
}

Write-Host "`nFATTO. Mappa ricostruita in: $fullOutPath" -ForegroundColor Green
Write-Host "Controlla il contenuto, poi sostituisci tu stesso mappa_uuid_sistemi.xlsx quando sei sicuro che vada bene." -ForegroundColor Yellow
if ($daVerificareCount -gt 0) {
    Write-Host "$daVerificareCount sistemi risultano con 'Da verificare = SI': confrontali con la mappa precedente prima di sostituirla, per non perdere eventuali proposte di rinomina gia' elaborate a mano." -ForegroundColor Yellow
}
