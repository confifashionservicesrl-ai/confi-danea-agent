# CON.FI - Installazione dell'agent Danea (sola lettura) come attivita' pianificata ogni 15 minuti.
. (Join-Path $PSScriptRoot 'comune.ps1')

function Scrivi($t, $c = 'White') { Write-Host $t -ForegroundColor $c }
$NomeAttivita = 'CON.FI Danea'

try {
    Scrivi ''
    Scrivi '=== CON.FI - Installazione invio dati Danea (sola lettura) ===' 'Cyan'
    Scrivi ''

    # 1) copia dei programmi in C:\ProgramData\ConfiDanea
    New-Item -ItemType Directory -Force -Path $Cartella | Out-Null
    foreach ($f in 'comune.ps1', 'invia.ps1', 'test-connessione.ps1') {
        Copy-Item -Path (Join-Path $PSScriptRoot $f) -Destination (Join-Path $Cartella $f) -Force
    }
    Get-ChildItem $Cartella -Filter *.ps1 | Unblock-File
    Scrivi "1/4 Programmi copiati in $Cartella" 'Green'

    # 2) indirizzo del server e codice
    $vecchia = $null
    if (Test-Path $ConfigFile) { try { $vecchia = Get-Content $ConfigFile -Raw | ConvertFrom-Json } catch { } }
    $predefinito = if ($vecchia) { $vecchia.server } else { 'https://confi-danea-production.up.railway.app' }
    Scrivi ''
    Scrivi "Indirizzo del server (premi Invio per lasciare: $predefinito)"
    $server = Read-Host 'Indirizzo'
    if ([string]::IsNullOrWhiteSpace($server)) { $server = $predefinito }
    $server = $server.Trim().TrimEnd('/')
    if ($server -notmatch '^https?://') { $server = 'https://' + $server }

    $codice = $null
    if ($vecchia -and $vecchia.codice) {
        Scrivi 'Codice di invio: premi Invio per tenere quello gia salvato, oppure scrivine uno nuovo'
    } else {
        Scrivi 'Codice di invio (quello che ti ha dato Claude, con i trattini):'
    }
    $codice = (Read-Host 'Codice').Trim().ToUpper()
    if ([string]::IsNullOrWhiteSpace($codice)) {
        if ($vecchia -and $vecchia.codice) { $codice = $vecchia.codice } else { throw 'Codice di invio mancante.' }
    }
    [pscustomobject]@{ server = $server; codice = $codice } | ConvertTo-Json | Set-Content -Path $ConfigFile -Encoding UTF8
    # il file con il codice lo leggono solo SYSTEM e gli amministratori
    & icacls.exe $ConfigFile /inheritance:r /grant:r '*S-1-5-18:F' '*S-1-5-32-544:F' | Out-Null
    Scrivi '2/4 Configurazione salvata' 'Green'

    # 3) primo invio, a schermo
    Scrivi ''
    Scrivi '3/4 Primo invio dei dati (puo richiedere un minuto)...'
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $Cartella 'invia.ps1') -Interattivo
    $primoOk = ($LASTEXITCODE -eq 0)

    # 4) attivita' pianificata: ogni 15 minuti e all'avvio di Windows, come SYSTEM
    $azione = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument ('-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + (Join-Path $Cartella 'invia.ps1') + '"')
    $ogni15 = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(2) -RepetitionInterval (New-TimeSpan -Minutes 15) -RepetitionDuration (New-TimeSpan -Days 3650)
    $avvio  = New-ScheduledTaskTrigger -AtStartup
    $chi    = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
    $regole = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 10) -MultipleInstances IgnoreNew -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
    Register-ScheduledTask -TaskName $NomeAttivita -Action $azione -Trigger $ogni15, $avvio -Principal $chi -Settings $regole -Force | Out-Null
    Scrivi "4/4 Attivita pianificata '$NomeAttivita' creata: invio ogni 15 minuti" 'Green'

    Scrivi ''
    if ($primoOk) { Scrivi '=== INSTALLAZIONE COMPLETATA ===' 'Green' }
    else { Scrivi '=== INSTALLATO, MA IL PRIMO INVIO NON E RIUSCITO (vedi errore in rosso sopra) ===' 'Yellow' }
    Scrivi "Registro degli invii: $(Join-Path $Cartella 'invia.log')"
}
catch {
    Scrivi ''
    Scrivi '=== ERRORE ===' 'Red'
    Scrivi $_.Exception.Message 'Red'
    Scrivi ("Riga " + $_.InvocationInfo.ScriptLineNumber + ": " + $_.InvocationInfo.Line.Trim()) 'DarkGray'
}
Scrivi ''
Scrivi 'Fai una foto a questa schermata e mandala a Claude.' 'Cyan'
