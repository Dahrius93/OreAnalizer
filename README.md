# SAP Ore Analyzer

App Streamlit per analizzare e visualizzare le ore registrate su SAP, esportate in formato Excel.  
Permette di filtrare per reparto, WBS, tipo attività e persona, confrontare le ore reali con un target e scaricare un report Excel completo con grafici embedded.

---

## Struttura del progetto

```
OreAnalizer/
├── App.py          # Entry point: configurazione pagina, CSS, UI sidebar e main
├── config.py       # Costanti: nomi colonne SAP, ACTIVITY_TYPES, TYPE_COLORS
├── data.py         # Caricamento, pulizia, parsing attività, filtri
├── charts.py       # Generazione grafici Plotly (heatmap, donut, barre)
├── report.py       # Generazione report Excel multi-foglio con grafici
└── requirements.txt
```

---

## Requisiti

Python 3.9+

| Pacchetto     | Versione minima | Uso                                   |
|---------------|-----------------|---------------------------------------|
| streamlit     | 1.30.0          | Framework UI                          |
| pandas        | 2.0.0           | Manipolazione dati                    |
| plotly        | 5.18.0          | Grafici interattivi                   |
| openpyxl      | 3.1.0           | Lettura file Excel SAP                |
| xlsxwriter    | 3.1.0           | Scrittura report Excel                |
| kaleido       | 0.2.1           | Esportazione grafici Plotly come PNG  |

---

## Installazione e avvio

```bash
# 1. Crea e attiva un virtual environment
python -m venv .venv
.venv\Scripts\activate        # Windows
# source .venv/bin/activate   # Linux / macOS

# 2. Installa le dipendenze
pip install -r requirements.txt

# 3. Avvia l'app
streamlit run App.py
```

L'app si apre automaticamente nel browser su `http://localhost:8501`.

---

## Utilizzo

1. **Carica** il file Excel esportato da SAP tramite il pannello laterale sinistro
2. **Filtra** i dati a cascata: Reparto → WBS → Tipo Attività → Persone
3. **Imposta un Target** ore (opzionale) per il confronto visivo
4. **Esplora** i grafici nella pagina principale
5. **Scarica** il report Excel completo con grafici embedded

---

## Struttura attesa del file SAP

Il file Excel deve essere l'export standard SAP delle ore giornaliere.  
Le righe di subtotale (senza WBS) vengono scartate automaticamente al caricamento.

| Colonna | Campo SAP                              | Utilizzo nell'app          |
|---------|----------------------------------------|----------------------------|
| A       | Testo breve                            | Codice WBS                 |
| B       | Creato il                              | (non usato direttamente)   |
| C       | Testo contab.                          | Reparto (es. UTE, UTES)    |
| G       | Data                                   | Data giornata lavorativa   |
| H       | Nome del dipendente o del candidato    | Persona                    |
| I       | Periodo                                | (non usato direttamente)   |
| J       | Esercizio                              | (non usato direttamente)   |
| K       | Numero (un. di mis.)                   | Ore registrate             |
| M       | Network                                | Codice network SAP         |
| N       | Descrizione attività                   | Tipo attività (parsing)    |

---

## Tipi di attività

Il tipo attività viene estratto automaticamente dal prefisso della colonna **N** (Descrizione attività).  
Sono supportati i seguenti formati: `SW - testo`, `SW–testo`, `SW testo`, `SW` (solo prefisso).

| Prefisso | Descrizione                                      |
|----------|--------------------------------------------------|
| HW       | Progettazione HW / Fornitori elettrici           |
| SW       | Progettazione SW                                 |
| REV      | Revisione / Collaudo commesse precedenti         |
| RI       | Riunioni / Coordinamento UTE                     |
| PS       | Test macchina / FAT                              |
| C        | Cantiere / Sviluppo in loco                      |
| NC       | Gestione NC elettriche                           |
| APRE     | Assistenza cantiere (montaggio/collaudo)         |
| APOST    | Assistenza post vendita                          |
| ITEC     | Studi prevendita / Offerta                       |
| RD       | R&D                                              |
| V        | Attività varie (archivio, DB, nuovi prodotti)    |
| GEST     | Gestione (fornitori, offerte, DOC, manualistica) |

Descrizioni non riconosciute vengono classificate come **N/D**.

---

## Funzionalità principali

### KPI Cards
Pannello con 5 indicatori: ore totali, numero persone, WBS attive, tipi attività, media ore/giorno.

### Target vs Reale
Barra di confronto con status colorato:
- Verde — utilizzo ≤ 80% del target
- Giallo — utilizzo tra 80% e 100%
- Rosso — target superato

### Grafici
| Grafico | Descrizione |
|---------|-------------|
| Donut Tipo Attività | Distribuzione percentuale delle ore per tipo |
| Stacked Bar Persona × Tipo | Ore per persona ripartite per tipo attività |
| Heatmap Persone × WBS | Matrice ore con scala colore |
| Heatmap Persone × Tipo Attività | Matrice ore per tipo attività |
| Bar Ore per Persona | Classifica ore totali per persona |

### Report Excel
File multi-foglio scaricabile contenente:
- **Riepilogo** — parametri filtro, ore totali, target, periodo, data generazione
- **Ore per Persona** — ore totali, giorni lavorati, media ore/giorno
- **Ore per WBS** — ore totali e numero registrazioni per WBS
- **Ore per Tipo Attività** — ore, registrazioni e percentuale per ogni tipo
- **Dettaglio** — tutte le righe filtrate con colonne principali
- **Grafici** — immagini PNG dei grafici embedded nel foglio

> Per l'esportazione dei grafici come immagini è necessaria la libreria **kaleido**.
