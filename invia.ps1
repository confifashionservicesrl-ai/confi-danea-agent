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
    # capi e prezzo (in Danea la fattura del fornitore e' registrata con una riga sola). Si cerca in tutte le
    # tabelle dell'archivio la colonna binaria che contiene l'XML e il modo di collegarla al documento
    # (colonna IDDoc, oppure il codice del servizio fatture elettroniche FE_IDAgyo/FE_IDAgyoReg della testata).
    $nFe = 0; $fe = $null; $feInfo = @()
    try {
        $sqlB = 'select trim(rf.RDB$RELATION_NAME), trim(rf.RDB$FIELD_NAME) from RDB$RELATION_FIELDS rf join RDB$FIELDS f on f.RDB$FIELD_NAME = rf.RDB$FIELD_SOURCE ' +
                'join RDB$RELATIONS r on r.RDB$RELATION_NAME = rf.RDB$RELATION_NAME where f.RDB$FIELD_TYPE = 261 and coalesce(r.RDB$SYSTEM_FLAG, 0) = 0 and r.RDB$VIEW_BLR is null'
        $cmdB = [FirebirdSql.Data.FirebirdClient.FbCommand]::new($sqlB, $con, $tr)
        $rb = $cmdB.ExecuteReader(); $coppie = @()
        while ($rb.Read()) { $coppie += , @([string]$rb.GetValue(0), [string]$rb.GetValue(1)) }
        $rb.Dispose(); $cmdB.Dispose()
        $lat = [Text.Encoding]::GetEncoding(28591)
        $docForn = 'select t."IDDoc" from "TDocTestate" t where t."Data" >= ''2026-01-01'' and t."TipoDoc" in (''U'',''V'',''W'')'
        $scelta = $null; $sqlScelta = $null
        foreach ($c in $coppie) {
            $trovati = 0; $letti = 0
            try {
                $cmdP = [FirebirdSql.Data.FirebirdClient.FbCommand]::new('select first 3 x."' + $c[1] + '" from "' + $c[0] + '" x where x."' + $c[1] + '" is not null', $con, $tr)
                $cmdP.CommandTimeout = 60
                $rp = $cmdP.ExecuteReader()
                while ($rp.Read()) {
                    $v = $rp.GetValue(0); $letti++
                    $testo = if ($v -is [byte[]]) { $lat.GetString($v) } else { [string]$v }
                    if ($testo -match 'FatturaElettronica') { $trovati++ }
                }
                $rp.Dispose(); $cmdP.Dispose()
            } catch { }
            if ($trovati -eq 0) { continue }
            # colonne della tabella che contiene l'XML
            $sqlK = 'select trim(rf.RDB$FIELD_NAME), f.RDB$FIELD_TYPE from RDB$RELATION_FIELDS rf join RDB$FIELDS f on f.RDB$FIELD_NAME = rf.RDB$FIELD_SOURCE where rf.RDB$RELATION_NAME = ''' + $c[0] + ''' order by rf.RDB$FIELD_POSITION'
            $cmdK = [FirebirdSql.Data.FirebirdClient.FbCommand]::new($sqlK, $con, $tr)
            $rk = $cmdK.ExecuteReader(); $cols = @(); while ($rk.Read()) { $cols += , @([string]$rk.GetValue(0), [int]$rk.GetValue(1)) }; $rk.Dispose(); $cmdK.Dispose()
            $feInfo += ('XML in ' + $c[0] + '.' + $c[1] + ' (' + $trovati + '/' + $letti + '); colonne: ' + (($cols | ForEach-Object { $_[0] }) -join ','))
            if ($scelta) { continue }
            # collegamento al documento: IDDoc, oppure una colonna di testo uguale a FE_IDAgyo / FE_IDAgyoReg della testata
            $prove = @()
            if ($cols | Where-Object { $_[0] -eq 'IDDoc' }) { $prove += 'x."IDDoc" = t."IDDoc"' }
            foreach ($k in ($cols | Where-Object { $_[1] -in @(14, 37) })) { foreach ($fk in 'FE_IDAgyoReg', 'FE_IDAgyo') { $prove += ('x."' + $k[0] + '" = t."' + $fk + '"') } }
            foreach ($cond in $prove) {
                try {
                    $q = 'select count(*) from "' + $c[0] + '" x join "TDocTestate" t on ' + $cond + ' where t."IDDoc" in (' + $docForn + ') and x."' + $c[1] + '" is not null'
                    $cmdC = [FirebirdSql.Data.FirebirdClient.FbCommand]::new($q, $con, $tr); $cmdC.CommandTimeout = 60
                    $n = [int]$cmdC.ExecuteScalar(); $cmdC.Dispose()
                    $feInfo += ('  collegamento ' + $cond + ': ' + $n)
                    if ($n -gt 0) { $scelta = $c; $sqlScelta = 'select t."IDDoc", x."' + $c[1] + '" from "' + $c[0] + '" x join "TDocTestate" t on ' + $cond + ' where t."IDDoc" in (' + $docForn + ') and x."' + $c[1] + '" is not null'; break }
                } catch { $feInfo += ('  collegamento ' + $cond + ': errore') }
            }
        }
        if (-not $coppie.Count) { $feInfo += 'nessuna colonna binaria nell''archivio' }
        elseif (-not ($feInfo | Where-Object { $_ -like 'XML in*' })) { $feInfo += ('nessun XML nelle ' + $coppie.Count + ' colonne binarie: ' + (($coppie | Select-Object -First 40 | ForEach-Object { $_[0] + '.' + $_[1] }) -join ',')) }
        if ($sqlScelta) {
            $cmdF = [FirebirdSql.Data.FirebirdClient.FbCommand]::new($sqlScelta, $con, $tr)
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
