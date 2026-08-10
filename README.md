# SAP Logon - Gestione modulare landscape

Toolkit PowerShell per gestire in modo centralizzato e modulare le connessioni SAP GUI
(SAP Logon) di Avvale: un file XML sorgente per cliente, compilato automaticamente in un
unico file globale distribuito a tutti i colleghi.

Per la descrizione completa dell'architettura, dei flussi operativi e della cronologia dei
problemi risolti, vedi [`LEGGIMI.txt`](./LEGGIMI.txt).

## Script inclusi

| Script | Cosa fa |
|---|---|
| `Build-SAPGlobal.ps1` | Rigenera/valida `SAPUILandscapeGlobal.xml` dai sorgenti, senza distribuire nulla. |
| `Deploy-SAPLandscape.ps1` | Installa/aggiorna la configurazione sul PC dell'utente corrente (richiama Build, fa backup, distribuisce). |
| `New-SAPClient.ps1` | Crea la cartella e il file sorgente per un nuovo cliente, in modo guidato. |
| `Rebuild-SourcesFromMapping.ps1` | Ricostruisce/corregge i sorgenti a partire dalla mappa `mappa_uuid_sistemi.xlsx` (non inclusa in questo repo). |
| `Restore-SAPLandscape.ps1` | Ripristina il file personale da uno dei backup con timestamp creati da Deploy. |

Ogni script ha un lanciatore `.cmd` gemello (stesso nome) che aggira il blocco
"esecuzione script disabilitata" di PowerShell: basta fare doppio click su quello invece
del `.ps1`.

## Nota sui dati

Questo repository contiene SOLO il codice. I dati operativi del landscape
(`mappa_uuid_sistemi.xlsx`, le cartelle `SAPUILandscape_sorgenti` e
`SAPUILandscape_globale`, con IP/hostname/nomi dei sistemi SAP dei clienti) restano
esclusivamente su OneDrive e sono esclusi via `.gitignore`, cosi' da evitare che finiscano
per errore nella cronologia git.
