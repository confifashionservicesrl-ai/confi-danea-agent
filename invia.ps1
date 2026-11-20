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
    # righe dei documenti (descrizione, quantita', importi) di fatture, note e fatture dei fornitori dal 2025:
    # servono a dividere ricavi e costi per reparto e per commessa. Se non riesce si invia senza.
    $nRighe = 0; $righe = $null
    try {
        $cmdT = [FirebirdSql.Data.FirebirdClient.FbCommand]::new('select first 1 trim(RDB$RELATION_NAME) from RDB$RELATIONS where RDB$RELATION_NAME starting with ''TDocRig''', $con, $tr)
        $tab = [string]$cmdT.ExecuteScalar(); $cmdT.Dispose()
        if (-not $tab) { throw 'tabella righe non trovata' }
        $sqlR = 'select trim(rf.RDB$FIELD_NAME), f.RDB$FIELD_TYPE, f.RDB$FIELD_SUB_TYPE from RDB$RELATION_FIELDS rf join RDB$FIELDS f on f.RDB$FIELD_NAME = rf.RDB$FIELD_SOURCE where rf.RDB$RELATION_NAME = ''' + $tab + ''' order by rf.RDB$FIELD_POSITION'
        $cmdR = [FirebirdSql.Data.FirebirdClient.FbCommand]::new($sqlR, $con, $tr)
        $rr = $cmdR.ExecuteReader(); $colR = @()
        while ($rr.Read()) {
            $nome = [string]$rr.GetValue(0); $tipo = [int]$rr.GetValue(1)
            $sub = if ($rr.IsDBNull(2)) { 0 } else { [int]$rr.GetValue(2) }
            if ($tipo -eq 261 -and $sub -ne 1) { continue }   # blob binari (immagini): non servono
            $colR += 'r."' + $nome + '"'
        }
        $rr.Dispose(); $cmdR.Dispose()
        $sqlRighe = 'select ' + ($colR -join ',') + ' from "' + $tab + '" r where r."IDDoc" in (select t."IDDoc" from "TDocTestate" t where t."Data" >= ''2025-01-01'' and t."TipoDoc" in (''I'',''F'',''J'',''N'',''O'',''R'',''U'',''V'',''W''))'
        $righe = Tabella-Json $con $tr $sqlRighe ([ref]$nRighe)
    } catch { Registra "Avviso: righe dei documenti non lette ($($_.Exception.Message))" 'Yellow'; $righe = $null; $nRighe = 0 }
    # fatture elettroniche ricevute dai fornitori (dal 2026): l'XML originale, che ha le righe con commessa,
    # capi e prezzo (in Danea la fattura del fornitore e' registrata con una riga sola). Si cerca, tra le
    # tabelle legate ai documenti, la colonna binaria che contiene l'XML. Se non c'e' si invia senza.
    $nFe = 0; $fe = $null; $feInfo = @()
    try {
        $sqlB = 'select trim(rf.RDB$RELATION_NAME), trim(rf.RDB$FIELD_NAME) from RDB$RELATION_FIELDS rf join RDB$FIELDS f on f.RDB$FIELD_NAME = rf.RDB$FIELD_SOURCE ' +
                'where f.RDB$FIELD_TYPE = 261 and rf.RDB$RELATION_NAME not starting with ''RDB$'' and rf.RDB$RELATION_NAME in ' +
                '(select r2.RDB$RELATION_NAME from RDB$RELATION_FIELDS r2 where trim(r2.RDB$FIELD_NAME) = ''IDDoc'')'
        $cmdB = [FirebirdSql.Data.FirebirdClient.FbCommand]::new($sqlB, $con, $tr)
        $rb = $cmdB.ExecuteReader(); $coppie = @()
        while ($rb.Read()) { $coppie += , @([string]$rb.GetValue(0), [string]$rb.GetValue(1)) }
        $rb.Dispose(); $cmdB.Dispose()
        $docForn = 'select t."IDDoc" from "TDocTestate" t where t."Data" >= ''2026-01-01'' and t."TipoDoc" in (''U'',''V'',''W'')'
        $lat = [Text.Encoding]::GetEncoding(28591)
        $scelta = $null
        foreach ($c in $coppie) {
            $trovati = 0; $letti = 0
            try {
                $cmdP = [FirebirdSql.Data.FirebirdClient.FbCommand]::new('select first 5 x."' + $c[1] + '" from "' + $c[0] + '" x where x."IDDoc" in (' + $docForn + ') and x."' + $c[1] + '" is not null', $con, $tr)
                $rp = $cmdP.ExecuteReader()
                while ($rp.Read()) {
                    $v = $rp.GetValue(0); $letti++
                    $testo = if ($v -is [byte[]]) { $lat.GetString($v) } else { [string]$v }
                    if ($testo -match 'FatturaElettronica') { $trovati++ }
                }
                $rp.Dispose(); $cmdP.Dispose()
            } catch { }
            $feInfo += ($c[0] + '.' + $c[1] + ':' + $letti + '/' + $trovati)
            if ($trovati -gt 0 -and -not $scelta) { $scelta = $c }
        }
        # se l'XML non e' nell'archivio: le colonne FE_ della testata (es. nome o percorso del file) di una fattura fornitore
        if (-not $scelta) {
            try {
                $cmdN = [FirebirdSql.Data.FirebirdClient.FbCommand]::new('select trim(rf.RDB$FIELD_NAME) from RDB$RELATION_FIELDS rf join RDB$FIELDS f on f.RDB$FIELD_NAME = rf.RDB$FIELD_SOURCE where rf.RDB$RELATION_NAME = ''TDocTestate'' and rf.RDB$FIELD_NAME starting with ''FE'' and f.RDB$FIELD_TYPE in (14, 37)', $con, $tr)
                $rn = $cmdN.ExecuteReader(); $colFe = @(); while ($rn.Read()) { $colFe += [string]$rn.GetValue(0) }; $rn.Dispose(); $cmdN.Dispose()
                if ($colFe.Count) {
                    $cmdS = [FirebirdSql.Data.FirebirdClient.FbCommand]::new('select first 1 ' + (($colFe | ForEach-Object { 't."' + $_ + '"' }) -join ',') + ' from "TDocTestate" t where t."Data" >= ''2026-01-01'' and t."TipoDoc" in (''U'',''V'',''W'') order by t."Data" desc', $con, $tr)
                    $rs = $cmdS.ExecuteReader()
                    if ($rs.Read()) { for ($i = 0; $i -lt $colFe.Count; $i++) { $feInfo += ($colFe[$i] + '=' + [string]$rs.GetValue($i)) } }
                    $rs.Dispose(); $cmdS.Dispose()
                }
            } catch { $feInfo += ('colonne FE: ' + $_.Exception.Message) }
        }
        if ($scelta) {
            $cmdF = [FirebirdSql.Data.FirebirdClient.FbCommand]::new('select x."IDDoc", x."' + $scelta[1] + '" from "' + $scelta[0] + '" x where x."IDDoc" in (' + $docForn + ') and x."' + $scelta[1] + '" is not null', $con, $tr)
            $rf = $cmdF.ExecuteReader(); $sbF = [System.Text.StringBuilder]::new(); [void]$sbF.Append('[')
            while ($rf.Read()) {
                $v = $rf.GetValue(1)
                $b = if ($v -is [byte[]]) { $v } else { [Text.Encoding]::UTF8.GetBytes([string]$v) }
                if ($b.Length -gt 3MB) { continue }
                if ($lat.GetString($b) -notmatch 'FatturaElettronica') { continue }
                if ($nFe) { [void]$sbF.Append(',') }
                [void]$sbF.Append('{"IDDoc":').Append((ValoreJson $rf.GetValue(0))).Append(',"b64":"').Append([Convert]::ToBase64String($b)).Append('"}')
                $nFe++
            }
            [void]$sbF.Append(']'); $rf.Dispose(); $cmdF.Dispose()
            $fe = $sbF.ToString()
        }
    } catch { $feInfo += ('errore: ' + $_.Exception.Message) }
    $tr.Commit()
    $con.Close(); $con = $null

    $archivio = '{"percorso":' + (ValoreJson $arch.FullName) + ',"mb":' + [int]($arch.Length / 1MB) + ',"modificato":' + (ValoreJson $arch.LastWriteTime) + '}'
    $json = '{"versione":1,"generato":' + (ValoreJson (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss'Z'")) +
            ',"pc":' + (ValoreJson $env:COMPUTERNAME) + ',"archivio":' + $archivio +
            ',"documenti":' + $documenti + ',"scadenze":' + $scadenze + ',"anagrafica":' + $anagrafica +
            $(if ($righe) { ',"righe":' + $righe } else { '' }) +
            $(if ($fe) { ',"fe":' + $fe } else { '' }) + ',"fe_info":' + (ValoreJson ($feInfo -join ' | ')) + '}'

    $bytes = [Text.Encoding]::UTF8.GetBytes($json)
    $ms = New-Object IO.MemoryStream
    $gz = New-Object IO.Compression.GZipStream($ms, [IO.Compression.CompressionMode]::Compress)
    $gz.Write($bytes, 0, $bytes.Length); $gz.Close()
    $corpo = $ms.ToArray()

    $url = ($cfg.server.TrimEnd('/')) + '/api/ingest'
    $risp = Invoke-WebRequest -UseBasicParsing -Uri $url -Method Post -Body $corpo -ContentType 'application/octet-stream' `
            -Headers @{ Authorization = "Bearer $($cfg.codice)" } -TimeoutSec 180
    $sec = [int]((Get-Date) - $inizio).TotalSeconds
    Registra ("OK: inviati $nDoc documenti, $nRighe righe, $nFe fatture elettroniche ricevute, $nScad scadenze, $nAnag anagrafiche (" + [int]($corpo.Length / 1KB) + " KB, $sec s)") 'Green'
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
