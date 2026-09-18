# Caricamento resiliente della dashboard Serie A

La modifica riduce i casi di caricamento infinito e `Load failed` causati da
interruzioni temporanee fra il browser mobile e Google Apps Script.

## Cosa cambia

- `serieAGetDashboard` restituisce partite e classifica in una sola richiesta.
  Le vecchie azioni restano disponibili per compatibilità.
- Il frontend ripete una volta soltanto le letture interrotte, dopo una breve
  attesa. Login e salvataggi non vengono ripetuti automaticamente.
- L'ultima dashboard caricata correttamente viene conservata nel browser per
  un massimo di 7 giorni, separatamente per giocatore. Se Google non risponde,
  quei dati restano visibili con un avviso e un pulsante Retry.
- Quando vengono mostrati dati salvati, le partite il cui kickoff è già passato
  vengono bloccate localmente. Il backend continua comunque a essere l'autorità
  definitiva per ogni salvataggio.
- Il logout e la scadenza della sessione cancellano la copia locale del relativo
  giocatore. La cache può non essere disponibile nella navigazione privata.

Non cambiano PIN, sessioni, regole di punteggio, Golden, pronostici, risultati,
promemoria o trigger.

## Ordine sicuro di pubblicazione

1. Nel progetto Apps Script **Serie A Predictor**, sostituire tutto il contenuto
   dell'esistente `SerieAApi.gs` con `SerieAApi_DASHBOARD_FIX.txt` e salvare.
   Non creare un secondo file con le stesse funzioni.
2. Aggiornare il deployment web esistente: **Esegui il deployment → Gestisci
   deployment → matita → Nuova versione → Esegui il deployment**. Conservare lo
   stesso deployment e lo stesso URL `/exec` usato dal sito.
3. Verificare il nuovo endpoint prima di pubblicare il frontend. Una richiesta
   autenticata `serieAGetDashboard` deve restituire sia `matches` sia
   `leaderboard`. Il frontend include comunque un fallback temporaneo per un
   backend vecchio, ma quel fallback usa ancora due richieste.
4. Solo dopo il deployment Apps Script, pubblicare `serie-a/index.html` su
   GitHub Pages.

Non serve modificare `Code.gs`: il suo router inoltra già le azioni Serie A a
`handleSerieARequest_`. Non servono nuovi trigger o proprietà script.

## Verifica

- Primo caricamento riuscito: compare brevemente `Updated just now`.
- Rete interrotta dopo almeno un caricamento riuscito: la pagina mostra i dati
  salvati e `Refresh failed · saved data remains available`, senza svuotarsi.
- Rete ripristinata: Retry aggiorna la dashboard.
- Logout: la copia locale del giocatore viene cancellata.
- Il salvataggio di un pronostico mantiene un solo tentativo di rete.

Test offline:

```text
node --test tests/serie-a-dashboard-resilience.test.cjs
```

I test simulano l'errore mobile `Load failed`, il retry, la cache separata per
giocatore e il nuovo endpoint senza contattare Google o modificare dati reali.
