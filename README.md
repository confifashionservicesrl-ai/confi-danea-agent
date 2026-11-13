# confi-danea-agent

Programma di **sola lettura** dell'archivio Danea Easyfatt di CON.FI FASHION SERVICE.

Gira dentro il Windows dove si trova Danea, legge l'archivio tramite il Firebird di Danea con una
transazione di sola lettura (non modifica nulla) e ogni 15 minuti manda una fotografia dei dati
(testate dei documenti, scadenze, clienti/fornitori) al server privato CON.FI.

Qui non ci sono dati aziendali né password: il codice di invio si scrive durante l'installazione
e resta solo su quel PC (`C:\ProgramData\ConfiDanea\config.json`, leggibile solo dagli amministratori).

## Installazione
1. **Code → Download ZIP**, poi tasto destro sullo ZIP → **Estrai tutto**.
2. Doppio clic su **`installa.cmd`** e rispondi **Sì** alla richiesta di Windows (servono i permessi
   da amministratore per creare l'attività pianificata). Se compare "Windows ha protetto il PC":
   **Ulteriori informazioni → Esegui comunque**.
3. Premi Invio per l'indirizzo proposto (o scrivi quello giusto) e scrivi il **codice di invio**.
4. Fa subito un primo invio e crea l'attività pianificata **"CON.FI Danea"** (ogni 15 minuti e all'avvio).

Per reinstallare o cambiare indirizzo/codice basta rilanciare `installa.cmd`.

## File
- `invia.ps1` — legge Danea e invia (registro: `C:\ProgramData\ConfiDanea\invia.log`)
- `comune.ps1` — collegamento a Firebird in sola lettura, conversione in JSON
- `test-connessione.cmd` — solo test: si collega e conta i documenti, non invia niente

Per togliere tutto: Utilità di pianificazione → elimina "CON.FI Danea", poi cancella `C:\ProgramData\ConfiDanea`.
