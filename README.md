# confi-danea-agent

Programma di **sola lettura** dell'archivio Danea Easyfatt di CON.FI FASHION SERVICE.

Non contiene dati aziendali né password: gira dentro il PC Windows dove si trova Danea
e legge l'archivio tramite il Firebird di Danea, senza modificarlo.

## Versione attuale: test di collegamento

1. **Code → Download ZIP**, poi estrai lo ZIP.
2. Doppio clic su **`installa.cmd`**. Se Windows avvisa "Windows ha protetto il PC",
   clicca **Ulteriori informazioni → Esegui comunque**.
3. Il programma scarica la libreria Firebird per .NET (da nuget.org, solo la prima volta),
   si collega in sola lettura e mostra quanti documenti trova.

Il registro dell'ultimo test è in `C:\ProgramData\ConfiDanea\test-connessione.log`.
