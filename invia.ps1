# CON.FI - Invio dei dati Danea al server CON.FI (SOLA LETTURA sull'archivio).
# Gira ogni 15 minuti come attivita' pianificata "CON.FI Danea". Registro: C:\ProgramData\ConfiDanea\invia.log
param([switch]$Interattivo)

. (Join-Path $PSScriptRoot 'comune.ps1')

$Log = Join-Path $Cartella 'invia.log'
function Registra([string]$testo, $colore = 'White') {
    $riga = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '  ' + $testo
    Add-Content -Path $Log -Value $riga -Encoding UTF8
    if ($Interattivo) { Write-Host $testo -ForegroundColor $colore }
}

New-Item -ItemType Directory -Force -Path $Cartella | Out-Null
# il registro tiene solo le ultime 2000 righe
if ((Test-Path $Log) -and (Get-Item $Log).Length -gt 400KB) {
    $ultime = Get-Content $Log -Tail 2000; Set-Content -Path $Log -Value $ultime -Encoding UTF8
}

$esito = 0
$con = $null
try {
    $cfg = Leggi-Config
    $inizio = Get-Date
    Carica-LibreriaFirebird
    $arch = Trova-Archivio
    $con = Apri-Danea $arch.FullName
    $tr = Transazione-SolaLettura $con

    $nDoc = 0; $nScad = 0; $nAnag = 0
    # anche le altre date dei documenti (per le fatture dei fornitori "Data" e' la data di registrazione:
    # la data della fattura del fornitore e' in un'altra colonna) e gli altri numeri documento
    $altre = @()
    try {
        $sqlCol = 'select trim(rf.RDB$FIELD_NAME), f.RDB$FIELD_TYPE from RDB$RELATION_FIELDS rf join RDB$FIELDS f on f.RDB$FIELD_NAME = rf.RDB$FIELD_SOURCE where rf.RDB$RELATION_NAME = ''TDocTestate'''
        $cmdCol = [FirebirdSql.Data.FirebirdClient.FbCommand]::new($sqlCol, $con, $tr)
        $rc = $cmdCol.ExecuteReader()
        while ($rc.Read()) {
            $nome = [string]$rc.GetValue(0); $tipo = [int]$rc.GetValue(1)
            if ($nome -in @('Data', 'NumDoc')) { continue }
            if ($tipo -in @(12, 35) -or ($nome -match '^(Num|Rif)' -and $tipo -in @(14, 37))) { $altre += '"' + $nome + '"' }
        }
        $rc.Dispose(); $cmdCol.Dispose()
    } catch { Registra "Avviso: colonne aggiuntive non lette ($($_.Exception.Message))" 'Yellow'; $altre = @() }
    $extra = if ($altre.Count) { ',' + ($altre -join ',') } else { '' }
    $base = 'select "IDDoc","TipoDoc","Data","NumDoc","IDAnagr","Anagr_Nome","TotNetto","TotIva","TotDoc","Pagam_ImportoDaSaldare","Pagam_Saldato"'
    try { $documenti = Tabella-Json $con $tr ($base + $extra + ' from "TDocTestate"') ([ref]$nDoc) }
    catch {
        if (-not $extra) { throw }
        Registra "Avviso: colonne aggiuntive non leggibili, invio senza ($($_.Exception.Message))" 'Yellow'
        $documenti = Tabella-Json $con $tr ($base + ' from "TDocTestate"') ([ref]$nDoc)
    }
    $scadenze  = Tabella-Json $con $tr 'select "IDPrimaNota","IDDoc","IDAnagr","DataScad","DataPagam","Importo","Saldato","Risorsa","CategPagamento" from "TPrimaNota"' ([ref]$nScad)
    $anagrafica = Tabella-Json $con $tr 'select "IDAnagr","CodAnagr","Nome","Cliente","Fornitore","Citta","Prov","Nazione","PartitaIva" from "TAnagrafica"' ([ref]$nAnag)
    $tr.Commit()
    $con.Close(); $con = $null

    $archivio = '{"percorso":' + (ValoreJson $arch.FullName) + ',"mb":' + [int]($arch.Length / 1MB) + ',"modificato":' + (ValoreJson $arch.LastWriteTime) + '}'
    $json = '{"versione":1,"generato":' + (ValoreJson (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss'Z'")) +
            ',"pc":' + (ValoreJson $env:COMPUTERNAME) + ',"archivio":' + $archivio +
            ',"documenti":' + $documenti + ',"scadenze":' + $scadenze + ',"anagrafica":' + $anagrafica + '}'

    $bytes = [Text.Encoding]::UTF8.GetBytes($json)
    $ms = New-Object IO.MemoryStream
    $gz = New-Object IO.Compression.GZipStream($ms, [IO.Compression.CompressionMode]::Compress)
    $gz.Write($bytes, 0, $bytes.Length); $gz.Close()
    $corpo = $ms.ToArray()

    $url = ($cfg.server.TrimEnd('/')) + '/api/ingest'
    $risp = Invoke-WebRequest -UseBasicParsing -Uri $url -Method Post -Body $corpo -ContentType 'application/octet-stream' `
            -Headers @{ Authorization = "Bearer $($cfg.codice)" } -TimeoutSec 180
    $sec = [int]((Get-Date) - $inizio).TotalSeconds
    Registra ("OK: inviati $nDoc documenti, $nScad scadenze, $nAnag anagrafiche (" + [int]($corpo.Length / 1KB) + " KB, $sec s)") 'Green'
}
catch {
    $esito = 1
    $msg = $_.Exception.Message
    if ($_.Exception.Response) {
        try { $msg += ' | ' + (New-Object IO.StreamReader($_.Exception.Response.GetResponseStream())).ReadToEnd() } catch { }
    }
    Registra "ERRORE: $msg" 'Red'
}
finally {
    if ($con) { try { $con.Close() } catch { } }
}
exit $esito
