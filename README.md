# Cantina

App web per gestire la cantina: giacenze per location, trasferimenti (anche in transito), consumi, acquisti e consegne, verifiche, export in Excel.

- Dati: database Supabase (progetto `cantina`), protetti da login. Solo le email nell'elenco "Persone autorizzate" vedono i dati.
- App: pagina statica pubblicata con GitHub Pages da questo repository.
- `.github/workflows/keepalive.yml`: ping ogni 3 giorni per evitare la pausa del piano gratuito di Supabase.

Per ricostruire `app.js` dopo una modifica a `src/`: `npm install && npm run build`.
