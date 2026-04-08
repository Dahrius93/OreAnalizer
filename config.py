# ──────────────────────────────────────────────
# COLUMN MAPPING (SAP export structure)
# Nomi esatti delle colonne nell'Excel SAP.
# I commenti indicano la lettera di colonna nell'export standard.
# ──────────────────────────────────────────────
COL_WBS = "Testo breve"                             # A — codice WBS commessa
COL_CREATED = "Creato il"                            # B — data creazione riga
COL_REPARTO = "Testo contab."                        # C — codice reparto (es. UTE, UTES)
COL_DATE = "Data"                                    # G — data giornata lavorativa
COL_PERSON = "Nome del dipendente o del candidato"   # H — nome dipendente
COL_PERIOD = "Periodo"                               # I — periodo contabile
COL_YEAR = "Esercizio"                               # J — anno esercizio
COL_HOURS = "Numero (un. di mis.)"                   # K — ore registrate
COL_NETWORK = "Network"                              # M — codice network SAP
COL_DESC = "Descrizione attività"                    # N — testo libero con prefisso tipo attività

# Colonne virtuali aggiunte da load_and_clean() al momento del caricamento.
# Non esistono nel file SAP originale.
COL_ACT_TYPE = "_tipo_attivita"       # prefisso estratto da COL_DESC (es. "SW", "HW")
COL_ACT_LABEL = "_tipo_attivita_label"  # etichetta estesa: "SW — Progettazione SW"
COL_ACT_DETAIL = "_desc_dettaglio"    # testo rimanente dopo il prefisso

# ──────────────────────────────────────────────
# ACTIVITY TYPE MAPPING
# Dizionario prefisso → descrizione estesa.
# Usato per il parsing della colonna N e per le etichette nei grafici.
# ──────────────────────────────────────────────
ACTIVITY_TYPES = {
    "HW":    "Progettazione HW / Fornitori elettrici",
    "SW":    "Progettazione SW",
    "REV":   "Revisione / Collaudo commesse precedenti",
    "RI":    "Riunioni / Coordinamento UTE",
    "PS":    "Test macchina / FAT",
    "C":     "Cantiere / Sviluppo in loco",
    "NC":    "Gestione NC elettriche",
    "APRE":  "Assistenza cantiere (montaggio/collaudo)",
    "APOST": "Assistenza post vendita",
    "ITEC":  "Studi prevendita / Offerta",
    "RD":    "R&D",
    "V":     "Attività varie (archivio, DB, nuovi prodotti)",
    "HW-GES": "Gestione fornitori HW",
    "AHW":   "Assistenza hardware",
    "$POST": "Assistenza pagata",
}

# Alias: prefissi legacy → chiave canonica in ACTIVITY_TYPES.
# Usato in data.py per normalizzare prefissi vecchi verso quelli nuovi.
ACTIVITY_ALIASES = {
    "GES":  "HW-GES",
    "GEST": "HW-GES",
}

# Mappa prefisso → colore esadecimale usato nei grafici Plotly.
# "N/D" è il colore di fallback per attività non classificate.
TYPE_COLORS = {
    "HW":    "#3b82f6",
    "SW":    "#8b5cf6",
    "REV":   "#06b6d4",
    "RI":    "#f59e0b",
    "PS":    "#10b981",
    "C":     "#ef4444",
    "NC":    "#f97316",
    "APRE":  "#ec4899",
    "APOST": "#14b8a6",
    "ITEC":  "#a855f7",
    "RD":    "#6366f1",
    "V":     "#64748b",
    "HW-GES": "#84cc16",
    "AHW":   "#0ea5e9",
    "$POST": "#f43f5e",
    "N/D":   "#334155",
}
