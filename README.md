# ⏱️ SAP Ore Analyzer

App Streamlit per analizzare le ore esportate da SAP.

## Setup rapido

```bash
# Crea virtual environment (opzionale ma consigliato)
python -m venv venv
source venv/bin/activate  # Linux/Mac
# oppure: venv\Scripts\activate  # Windows

# Installa dipendenze
pip install -r requirements.txt

# Avvia l'app
streamlit run app.py
```

L'app si aprirà nel browser su `http://localhost:8501`.

## Come usarla

1. **Carica** il file Excel esportato da SAP (sidebar sinistra)
2. **Filtra** per Reparto → WBS → Persone (filtri a cascata)
3. **Imposta un Target** ore per il confronto visivo
4. **Esporta** il report Excel completo con grafici

## Struttura file SAP attesa

| Colonna | Campo SAP                              | Uso                 |
|---------|----------------------------------------|---------------------|
| A       | Testo breve                            | Codice WBS          |
| C       | Testo contab.                          | Reparto (UTE, UTES) |
| G       | Data                                   | Data giornata       |
| H       | Nome del dipendente o del candidato    | Persona             |
| K       | Numero (un. di mis.)                   | Ore                 |
| N       | Descrizione attività                   | Descrizione         |

Le righe di subtotale SAP (senza WBS) vengono automaticamente scartate.

## Funzionalità

- **KPI Cards**: ore totali, persone, WBS attive, media giornaliera
- **Target vs Reale**: barra con status colorato (verde/giallo/rosso)
- **Heatmap**: matrice Persone × WBS con ore
- **Ore per Persona**: classifica orizzontale
- **Export Excel**: report multi-foglio con grafici embedded
