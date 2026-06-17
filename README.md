# SG Arese Waterpolo Analytics

Dashboard Streamlit interattiva per la stagione SG Arese Waterpolo 2025-26.

## File inclusi

- `app.py`: applicazione Streamlit
- `SG_Arese_2025_26_Analisi.xlsx`: workbook dati finale letto dall'app
- `prima_squadra_25_26.jpg`: immagine squadra usata nella sidebar
- `requirements.txt`: dipendenze minime per deploy
- `.python-version` e `runtime.txt`: runtime Python consigliato per Streamlit Cloud

## Avvio locale

Dalla cartella della repo deploy:

```bash
streamlit run app.py
```

Se usi un virtual environment, attivalo prima oppure richiama direttamente il binario `streamlit` del tuo ambiente locale.

## Deploy su Streamlit Community Cloud

1. Vai su `https://share.streamlit.io`.
2. Seleziona repository e branch.
3. Imposta `app.py` come main file path.
4. Avvia il deploy.
5. Se il deploy usa una versione Python non compatibile, imposta Python `3.11` dalla dashboard Streamlit Cloud.

## Note sui dati

- L'app legge direttamente `SG_Arese_2025_26_Analisi.xlsx` dalla root della repo.
- Le statistiche giocatore sono basate sulle partite con video/statistiche disponibili.
- Le righe con tutte le statistiche evento a `-` sono escluse dagli aggregati giocatore.
- `% tiro media` e `% tiro stagionale` sono metriche distinte.
- Questa e' la versione completa della dashboard, inclusiva di metriche di valutazione ed errore soggettivo.

## Aggiornare i dati della stagione 2025-26

Per aggiornare la dashboard dopo una modifica alle statistiche, normalmente basta aggiornare il file origine statistiche e rigenerare il workbook analisi.

Nei comandi sotto si usano questi placeholder:

- `<cartella-analisi>`: cartella locale privata dove tieni file origine, calendario, script di generazione e virtual environment.
- `<cartella-deploy>`: cartella locale della repo GitHub/Streamlit.

### 1. Modifica file origine

Apri e aggiorna il file statistiche origine nella tua `<cartella-analisi>`, ad esempio:

```text
<cartella-analisi>/SG Arese 2025-26 _ Statistiche.xlsx
```

Poi salva e chiudi Excel.

Note importanti:

- Se un giocatore non ha giocato una partita, inserisci `-` in tutte le colonne evento `B:N` della sua riga partita.
- Le colonne percentuali `O:P` possono contenere formule: lo script considera assente il giocatore guardando solo le colonne evento `B:N`.
- Se modifichi solo statistiche, gol, assist, tiri, errori, parate, ecc., non devi toccare il codice.

### 2. Rigenera workbook analisi

Dalla `<cartella-analisi>`:

```bash
python generate_arese_analysis.py
```

Output atteso:

```text
<cartella-analisi>/SG_Arese_2025_26_Analisi.xlsx
```

Se usi un virtual environment, attivalo prima oppure richiama direttamente il binario `python` del tuo ambiente locale.

### 3. Copia il workbook nella repo deploy

Da qualunque cartella:

```bash
cp "<cartella-analisi>/SG_Arese_2025_26_Analisi.xlsx" "<cartella-deploy>/SG_Arese_2025_26_Analisi.xlsx"
```

Oppure, se sei gia' dentro `<cartella-deploy>`:

```bash
cp "<cartella-analisi>/SG_Arese_2025_26_Analisi.xlsx" ./SG_Arese_2025_26_Analisi.xlsx
```

### 4. Testa localmente l'app

Dalla `<cartella-deploy>`:

```bash
streamlit run app.py
```

Controlli consigliati:

- `Overview` carica senza errori.
- `Giocatori` mostra percentuali in formato `%`.
- `Scheda Giocatore` funziona per un giocatore di movimento e per un portiere.
- `Scatter Lab` non mostra errori.

### 5. Pubblica aggiornamento su GitHub/Streamlit

Dalla `<cartella-deploy>`:

```bash
git status
git add SG_Arese_2025_26_Analisi.xlsx
git commit -m "Update analytics data"
git push
```

Streamlit Cloud aggiornera' automaticamente l'app dopo il push. Se non succede, usa `Reboot app` o `Clear cache and rerun` dalla dashboard Streamlit Cloud.

## Quando serve modificare il codice

Non basta aggiornare il file Excel origine se cambiano struttura o anagrafica.

Modifica `generate_arese_analysis.py` se cambiano:

- nomi giocatori
- ruoli
- mano dominante
- jolly
- portieri
- pesi della valutazione
- nomi dei fogli partita
- avversari
- calendario
- formule o colonne del workbook statistiche

Punti principali nello script:

- `STATS_XLSX`: nome file statistiche origine
- `CALENDAR_HTML`: calendario HTML usato per risultati/classifica
- `OUTPUT_XLSX`: nome workbook finale generato
- `WORKBOOK_OPPONENT_MAP`: mapping tra nomi foglio e avversari
- `PLAYER_PROFILES`: anagrafica giocatori, ruoli, mano, jolly
- `MOTION_WEIGHTS` e `GOALKEEPER_WEIGHTS`: pesi valutazione

## Preparare una prossima stagione

Workflow consigliato per una nuova stagione.

### 1. Duplica il progetto di analisi

Crea una nuova cartella locale privata, ad esempio:

```text
<cartella-analisi-nuova-stagione>
```

Copiaci dentro almeno:

- `generate_arese_analysis.py`
- nuovo file statistiche stagione
- nuovo calendario stagione
- eventuale foto squadra aggiornata

### 2. Aggiorna nomi file nello script

Nel nuovo `generate_arese_analysis.py`, aggiorna:

```python
STATS_XLSX = BASE_DIR / "..."
CALENDAR_HTML = BASE_DIR / "..."
OUTPUT_XLSX = BASE_DIR / "..."
```

Esempio output nuova stagione:

```python
OUTPUT_XLSX = BASE_DIR / "SG_Arese_2026_27_Analisi.xlsx"
```

### 3. Aggiorna avversari e calendario

Aggiorna:

```python
WORKBOOK_OPPONENT_MAP
```

Il formato previsto dei fogli partita e':

```text
A|AVVERSARIO
R|AVVERSARIO
```

dove `A` indica andata e `R` ritorno.

### 4. Aggiorna anagrafica giocatori

Aggiorna:

```python
PLAYER_PROFILES
```

Per ogni giocatore definire:

- `ruolo_macro`: `movimento` o `portiere`
- `mano`: `D` o `S`
- `ruolo_base`: es. `1-2`, `3`, `4-5`, `6`, `GK`
- `ruolo_label`: descrizione ruolo
- `jolly`: `SI` o `NO`
- `ruoli_jolly`: eventuali ruoli multipli

### 5. Rigenera e verifica

Dalla cartella della nuova analisi:

```bash
python generate_arese_analysis.py
```

Verifica nel workbook generato:

- foglio `Partite`
- foglio `Giocatori_Stagione`
- foglio `Portieri`
- foglio `Confronto_Giocatori`
- foglio `Metodo`

### 6. Aggiorna app e deploy

Nella repo deploy:

- sostituisci il file Excel finale
- se il nome file cambia, aggiorna `WORKBOOK_PATH` in `app.py`
- sostituisci foto squadra se necessario
- aggiorna testi/titoli stagione in `app.py` e `README.md`

Poi:

```bash
git add .
git commit -m "Prepare analytics app for new season"
git push
```

## Checklist rapida aggiornamento dati

```bash
cd <cartella-analisi>
python generate_arese_analysis.py
cp "./SG_Arese_2025_26_Analisi.xlsx" "<cartella-deploy>/SG_Arese_2025_26_Analisi.xlsx"
cd <cartella-deploy>
git status
git add SG_Arese_2025_26_Analisi.xlsx
git commit -m "Update analytics data"
git push
```
