# SG Arese Waterpolo Analytics

Dashboard Streamlit interattiva per la stagione SG Arese Waterpolo 2025-26.

## File inclusi

- `app.py`: applicazione Streamlit
- `SG_Arese_2025_26_Analisi.xlsx`: workbook dati finale
- `prima_squadra_25_26.jpg`: immagine squadra usata nella sidebar
- `requirements.txt`: dipendenze minime per deploy

## Avvio locale

```bash
streamlit run app.py
```

## Deploy su Streamlit Community Cloud

1. Crea una repository GitHub con questi file.
2. Vai su `https://share.streamlit.io`.
3. Seleziona repository e branch.
4. Imposta `app.py` come main file path.
5. Avvia il deploy.

## Note sui dati

- L'app legge direttamente `SG_Arese_2025_26_Analisi.xlsx`.
- Le statistiche giocatore sono basate sulle partite con video/statistiche disponibili.
- Le righe con tutte le statistiche evento a `-` sono escluse dagli aggregati giocatore.
- `% tiro media` e `% tiro stagionale` sono metriche distinte.
- Questa e' la versione completa della dashboard, inclusiva di metriche di valutazione ed errore soggettivo.
