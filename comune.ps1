# CON.FI - funzioni comuni dell'agent Danea (sola lettura). Usato da invia.ps1 e installa.ps1.

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$Cartella    = Join-Path $env:ProgramData 'ConfiDanea'
$ConfigFile  = Join-Path $Cartella 'config.json'
$DllVersion  = '7.10.1'
$FbUtente    = 'SYSDBA'
$FbPassword  = 'masterkey'   # password standard del Firebird installato da Danea

function Carica-LibreriaFirebird {
    New-Item -ItemType Directory -Force -Path $Cartella | Out-Null
    $dir = Join-Path $Cartella "fbclient-$DllVersion"
    $trova = { Get-ChildItem -Path $dir -Recurse -Filter 'FirebirdSql.Data.FirebirdClient.dll' -ErrorAction SilentlyContinue |
               Where-Object { $_.FullName -match '\\lib\\net4' } | Select-Object -First 1 }
    $dll = & $trova
    if (-not $dll) {
        $pkg = Join-Path $Cartella "fbclient-$DllVersion.zip"
        Invoke-WebRequest -UseBasicParsing -Uri "https://www.nuget.org/api/v2/package/FirebirdSql.Data.FirebirdClient/$DllVersion" -OutFile $pkg
        Expand-Archive -Path $pkg -DestinationPath $dir -Force
        Remove-Item $pkg -Force
        $dll = & $trova
    }
    if (-not $dll) { throw 'Libreria Firebird non trovata.' }
    Add-Type -Path $dll.FullName
    Add-Type -AssemblyName System.Web
}

function Trova-Archivio {
    $a = 'D:\Danea\CON FI FASHION SERVICE.eft'
    if (Test-Path $a) { return (Get-Item $a) }
    $t = Get-ChildItem 'D:\Danea\*.eft' -ErrorAction SilentlyContinue | Sort-Object Length -Descending | Select-Object -First 1
    if (-not $t) { throw 'Archivio Danea non trovato in D:\Danea' }
    return $t
}

function Apri-Danea([string]$archivio) {
    $porte = @(3050)
    try {
        $fb = Get-Process -Name fbserver -ErrorAction Stop
        $porte += (Get-NetTCPConnection -State Listen -OwningProcess $fb.Id -ErrorAction Stop).LocalPort
    } catch { }
    $errori = @()
    foreach ($p in ($porte | Select-Object -Unique)) {
        [string]$cs = "DataSource=localhost;Port=$([int]$p);Database=$archivio;User=$FbUtente;Password=$FbPassword;Charset=WIN1252;Dialect=3;Pooling=false"
        try {
            $con = [FirebirdSql.Data.FirebirdClient.FbConnection]::new($cs)
            $con.Open()
            return $con
        } catch { $errori += "porta ${p}: $($_.Exception.Message)" }
    }
    throw ("Collegamento a Danea non riuscito. " + ($errori -join ' | '))
}

function Transazione-SolaLettura($con) {
    $opz = [FirebirdSql.Data.FirebirdClient.FbTransactionOptions]::new()
    $opz.TransactionBehavior = [FirebirdSql.Data.FirebirdClient.FbTransactionBehavior]'Read,ReadCommitted,RecVersion,NoWait'
    return $con.BeginTransaction($opz)
}

$Inv = [Globalization.CultureInfo]::InvariantCulture

function ValoreJson($v) {
    if ($v -is [DBNull] -or $null -eq $v) { return 'null' }
    if ($v -is [string])   { return '"' + [System.Web.HttpUtility]::JavaScriptStringEncode($v) + '"' }
    if ($v -is [datetime]) {
        if ($v.TimeOfDay.Ticks -eq 0) { return '"' + $v.ToString('yyyy-MM-dd', $Inv) + '"' }
        return '"' + $v.ToString("yyyy-MM-dd'T'HH:mm:ss", $Inv) + '"'
    }
    if ($v -is [bool])     { if ($v) { return 'true' } else { return 'false' } }
    if ($v -is [decimal] -or $v -is [double] -or $v -is [single] -or $v -is [int16] -or $v -is [int32] -or $v -is [int64] -or $v -is [byte]) {
        return ([System.IConvertible]$v).ToString($Inv)
    }
    if ($v -is [byte[]])   { return 'null' }   # campi binari (immagini/blob): non servono
    return '"' + [System.Web.HttpUtility]::JavaScriptStringEncode([string]$v) + '"'
}

# Esegue una SELECT e restituisce {"colonne":[...],"righe":[[...],...]} come testo JSON
function Tabella-Json($con, $tr, [string]$sql, [ref]$conteggio) {
    $cmd = [FirebirdSql.Data.FirebirdClient.FbCommand]::new($sql, $con, $tr)
    $r = $cmd.ExecuteReader()
    try {
        $sb = [System.Text.StringBuilder]::new()
        $n = $r.FieldCount
        $nomi = for ($i = 0; $i -lt $n; $i++) { '"' + $r.GetName($i) + '"' }
        [void]$sb.Append('{"colonne":[').Append(($nomi -join ',')).Append('],"righe":[')
        $prima = $true; $righe = 0
        $vals = New-Object 'object[]' $n
        while ($r.Read()) {
            [void]$r.GetValues($vals)
            if (-not $prima) { [void]$sb.Append(',') }
            $prima = $false
            [void]$sb.Append('[')
            for ($i = 0; $i -lt $n; $i++) {
                if ($i) { [void]$sb.Append(',') }
                [void]$sb.Append((ValoreJson $vals[$i]))
            }
            [void]$sb.Append(']')
            $righe++
        }
        [void]$sb.Append(']}')
        $conteggio.Value = $righe
        return $sb.ToString()
    } finally { $r.Dispose(); $cmd.Dispose() }
}

function Leggi-Config {
    if (-not (Test-Path $ConfigFile)) { throw "Configurazione mancante ($ConfigFile): esegui installa.cmd" }
    return (Get-Content $ConfigFile -Raw | ConvertFrom-Json)
}
