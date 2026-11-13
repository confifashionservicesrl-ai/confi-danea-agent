# CON.FI - Test di collegamento all'archivio Danea Easyfatt (SOLA LETTURA)
# Non modifica nulla: apre una transazione di sola lettura, conta i documenti e chiude.

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$Cartella   = Join-Path $env:ProgramData 'ConfiDanea'
$Log        = Join-Path $Cartella 'test-connessione.log'
$DllVersion = '7.10.1'
$Utente     = 'SYSDBA'
$Password   = 'masterkey'   # password standard di Firebird; se Danea l'ha cambiata il test lo segnala

New-Item -ItemType Directory -Force -Path $Cartella | Out-Null
Start-Transcript -Path $Log -Force | Out-Null

function Scrivi($testo, $colore = 'White') { Write-Host $testo -ForegroundColor $colore }

Scrivi ''
Scrivi '=== CON.FI - Test collegamento a Danea (sola lettura) ===' 'Cyan'
Scrivi ''

try {
    # 1) Libreria Firebird per .NET (scaricata una sola volta da nuget.org)
    $DllDir = Join-Path $Cartella "fbclient-$DllVersion"
    $Dll = Get-ChildItem -Path $DllDir -Recurse -Filter 'FirebirdSql.Data.FirebirdClient.dll' -ErrorAction SilentlyContinue |
           Where-Object { $_.FullName -match '\\lib\\net4' } | Select-Object -First 1
    if (-not $Dll) {
        Scrivi '1/4 Scarico la libreria Firebird (una sola volta)...'
        $Pkg = Join-Path $Cartella "fbclient-$DllVersion.zip"
        Invoke-WebRequest -UseBasicParsing -Uri "https://www.nuget.org/api/v2/package/FirebirdSql.Data.FirebirdClient/$DllVersion" -OutFile $Pkg
        Expand-Archive -Path $Pkg -DestinationPath $DllDir -Force
        Remove-Item $Pkg -Force
        $Dll = Get-ChildItem -Path $DllDir -Recurse -Filter 'FirebirdSql.Data.FirebirdClient.dll' |
               Where-Object { $_.FullName -match '\\lib\\net4' } | Select-Object -First 1
    } else {
        Scrivi '1/4 Libreria Firebird gia presente.'
    }
    if (-not $Dll) { throw 'Libreria Firebird non trovata dopo il download.' }
    Add-Type -Path $Dll.FullName
    Scrivi "    OK ($($Dll.FullName))" 'Green'

    # 2) Archivio Danea
    Scrivi '2/4 Cerco l''archivio Danea...'
    $Archivio = 'D:\Danea\CON FI FASHION SERVICE.eft'
    if (-not (Test-Path $Archivio)) {
        $trovato = Get-ChildItem 'D:\Danea\*.eft' -ErrorAction SilentlyContinue | Sort-Object Length -Descending | Select-Object -First 1
        if (-not $trovato) { throw 'Archivio Danea non trovato in D:\Danea' }
        $Archivio = $trovato.FullName
    }
    $info = Get-Item $Archivio
    Scrivi "    OK: $Archivio  ($([int]($info.Length/1MB)) MB, modificato $($info.LastWriteTime))" 'Green'

    # 3) Porta di Firebird: prima la standard 3050, poi quelle su cui ascolta fbserver
    Scrivi '3/4 Cerco la porta di Firebird...'
    $Porte = @(3050)
    try {
        $fb = Get-Process -Name fbserver -ErrorAction Stop
        $Porte += (Get-NetTCPConnection -State Listen -OwningProcess $fb.Id -ErrorAction Stop).LocalPort
    } catch { }
    $Porte = $Porte | Select-Object -Unique
    Scrivi "    Porte da provare: $($Porte -join ', ')"

    # 4) Collegamento in sola lettura
    Scrivi '4/4 Mi collego in sola lettura...'
    $Collegato = $null
    $UltimoErrore = $null
    foreach ($p in $Porte) {
        [string]$cs = "DataSource=localhost;Port=$([int]$p);Database=$Archivio;User=$Utente;Password=$Password;Charset=WIN1252;Dialect=3;Pooling=false"
        try {
            $con = [FirebirdSql.Data.FirebirdClient.FbConnection]::new($cs)
            $con.Open(); $Collegato = $con; $PortaOk = $p; break
        }
        catch {
            $UltimoErrore = "porta ${p}: " + $_.Exception.Message
            Scrivi "    $UltimoErrore" 'Yellow'
            if ($con) { $con.Dispose() }
        }
    }
    if (-not $Collegato) { throw "Collegamento non riuscito. Ultimo errore: $UltimoErrore" }
    Scrivi "    OK, collegato sulla porta $PortaOk" 'Green'

    $opz = [FirebirdSql.Data.FirebirdClient.FbTransactionOptions]::new()
    $opz.TransactionBehavior = [FirebirdSql.Data.FirebirdClient.FbTransactionBehavior]'Read,ReadCommitted,RecVersion,NoWait'
    $tr = $Collegato.BeginTransaction($opz)

    function Valore([string]$sql) {
        $cmd = [FirebirdSql.Data.FirebirdClient.FbCommand]::new($sql, $Collegato, $tr)
        try { return $cmd.ExecuteScalar() } finally { $cmd.Dispose() }
    }

    $nDoc   = Valore 'select count(*) from "TDocTestate"'
    $ultima = Valore 'select max("Data") from "TDocTestate"'
    $nFatt  = Valore 'select count(*) from "TDocTestate" where "TipoDoc" = ''I'' and extract(year from "Data") = extract(year from current_date)'
    $nAnag  = Valore 'select count(*) from "TAnagrafica"'

    $tr.Commit()
    $Collegato.Close()

    Scrivi ''
    Scrivi '=== TUTTO OK ===' 'Green'
    Scrivi "Documenti in archivio:        $nDoc"
    Scrivi "Fatture di quest'anno:        $nFatt"
    Scrivi "Clienti e fornitori:          $nAnag"
    Scrivi "Data dell'ultimo documento:   $ultima"
    Scrivi ''
    Scrivi 'Fai una foto a questa schermata e mandala a Claude.' 'Cyan'
}
catch {
    Scrivi ''
    Scrivi '=== ERRORE ===' 'Red'
    Scrivi $_.Exception.Message 'Red'
    Scrivi ("Riga " + $_.InvocationInfo.ScriptLineNumber + ": " + $_.InvocationInfo.Line.Trim()) 'DarkGray'
    if ($_.Exception.Message -match 'password|user name') {
        Scrivi 'La password di Firebird non e quella standard: Danea l''ha cambiata.' 'Yellow'
    }
    Scrivi ''
    Scrivi 'Fai una foto a questa schermata e mandala a Claude.' 'Cyan'
}
finally {
    Stop-Transcript | Out-Null
}
