# Analisi Politiche - Project Context

## Panoramica Progetto

**Analisi Politiche** è un'applicazione web Streamlit per l'analisi e il conteggio delle azioni da file Excel relative alle politiche. L'applicazione permette di elaborare dati, generare report analitici e exportare risultati in formato Excel con grafici professionali.

---

## Architettura del Progetto

### Struttura Directory

```
analisiPolitiche/
├── app.py                    # Entry point applicazione Streamlit
├── config.json               # Configurazione (tariffe, filtri, output)
├── requirements.txt          # Dipendenze Python
├── README.md                 # Documentazione
└── src/
    ├── __init__.py           # Esportazione moduli
    ├── config.py             # Gestione configurazione
    ├── data_loader.py        # Caricamento e preprocessing dati Excel
    ├── analysis.py           # Logica di conteggio e analisi
    └── excel_export.py       # Export e formattazione risultati Excel
```

---

## Componenti Principali

### 1. **ConfigManager** (`src/config.py`)

Gestisce la configurazione dell'applicazione da file JSON.

**Responsabilità:**
- Carica/salva configurazione da `config.json`
- Gestisce tariffe (codici azione → importo €)
- Gestisce lista eventi da escludere
- Mantiene valori di default se file non esiste

**Principali Proprietà:**
- `tariffe`: Dict[str, float] - Codici azione e relative tariffe
- `escludi_eventi`: List[str] - Eventi da filtrare
- `prefisso_output`: str - Prefisso file export
- `codici_validi`: List[str] - Lista codici configurati

**Metodi Chiave:**
- `aggiungi_tariffa(codice, valore)` - Aggiunge/aggiorna tariffa
- `rimuovi_tariffa(codice)` - Rimuove tariffa
- `save()` - Persiste config su file
- `reload()` - Ricarica da file

**Configurazione Default:**
```json
{
    "tariffe": {
        "A03": 37.14, "A06": 35.57,
        "B03": 37.14, "B04": 37.14,
        "C06": 499.88
    },
    "filtri": {
        "escludi_eventi": ["Annullamento (prima dell'inizio)", "Proposta"]
    },
    "output": {
        "prefisso_nome": "export_analisi"
    }
}
```

---

### 2. **DataLoader** (`src/data_loader.py`)

Carica e preprocessa i dati da file Excel.

**Responsabilità:**
- Legge file Excel (con pandas)
- Estrae codici azione e tipi (A, B, C)
- Filtra righe con eventi da escludere
- Filtra righe con codici non configurati
- Traccia righe scartate e motivi di esclusione

**Attributi:**
- `tariffe`: Dict[str, float] - Tariffe per validazione
- `escludi_eventi`: List[str] - Eventi da filtrare
- `codici_validi`: List[str] - Codici ammessi

**Metodo Principale:**
- `load(file_source)` → Tuple[pd.DataFrame, pd.DataFrame]
  - Restituisce: (df_valido, df_scartato)

**Campi Elaborati:**
- `Codice`: Estratto da "Attività" (es. A03)
- `Tipo`: Prima lettera del codice (A, B, C)
- `Data Riferimento`: Data Proposta per C06, Data Fine per altri
- `Anno-Mese`: Per aggregazioni per mese
- `_indice_originale`: Numero riga Excel originale
- `_motivo_esclusione`: Ragione dello scarto (se applicabile)

**Metodi Secondari:**
- `get_statistiche_base(df)` - Statsistiche sul dataset
- `prepara_scartate_per_export(df_scartate)` - Formatta righe scartate

---

### 3. **AnalisiPolitiche** (`src/analysis.py`)

Esegue analisi quantitative sui dati.

**Responsabilità:**
- Conteggi per persona, tipo, mese
- Calcoli ricavi con tariffe
- Analisi per operatore
- Andamenti mensili

**Principale Input:**
- `df`: DataFrame preprocessato
- `tariffe`: Dict[str, float]

**Metodi di Analisi Principale:**

| Metodo | Output | Definizione |
|--------|--------|------------|
| `conteggio_per_persona_tipo()` | DataFrame | Azioni per persona con suddivisione per tipo (A, B, C) e codici specifici |
| `conteggio_per_persona_tipo_mese()` | DataFrame | Come sopra ma per ogni mese |
| `conteggio_totale_tipo()` | DataFrame | Totali aggregati per tipo e codice |
| `conteggio_totale_tipo_mese()` | DataFrame | Totali per tipo per ogni mese |
| `conteggio_per_operatore()` | DataFrame | Azioni raggruppate per operatore |
| `conteggio_per_operatore_mese()` | DataFrame | Come sopra per ogni mese |
| `calcolo_ricavi_per_mese()` | DataFrame | Conteggi e ricavi per codice e mese |
| `riepilogo_ricavi()` | DataFrame | Ricavi totali per codice |
| `ricavi_totali()` | float | Ricavo complessivo totale |
| `top_persone(n)` | DataFrame | Prime N persone per azioni |
| `utenti_per_operatore()` | DataFrame | Numero utenti unici per operatore |
| `andamento_mensile()` | DataFrame | Azioni mensili per tipo |
| `ricavi_per_codice()` | DataFrame | Ricavi per ogni codice |

**Metodi Helper:**
- `_get_operatore_per_persona()` - Associa operatore più recente a ogni persona

---

### 4. **ExcelExporter** (`src/excel_export.py`)

Esporta dati in format Excel con formattazione e grafici.

**Responsabilità:**
- Crea workbook Excel multi-foglio
- Applica stili professionali (header, bordi, colori)
- Genera grafici (linee, barre, torta)
- Formatta numeri e date

**Fogli Generati:**
1. **Riepilogo** - Statistiche generali e summari
2. **Per Persona** - Conteggi per persona/tipo/mese
3. **Per Operatore** - Conteggi per operatore
4. **Ricavi** - Analisi ricavi per codice e mese
5. **Righe Scartate** - Dettaglio righe filtrate
6. **Codici Non Validi** - Codici esclusi per mancanza tariffa

**Stili e Formattazione:**
- Header: Blu (#4472C4) con testo bianco, grassetto
- Bordi: Sottile su tutte le celle
- Grafici: Automaticamente dimensionati e colorati

**Metodi Chiave:**
- `export()` → bytes - Genera file Excel completo
- `_write_dataframe()` - Scrive DataFrame con stili
- `_create_line_chart()` - Grafico andamenti mensili
- `_create_bar_chart()` - Grafico ricavi per codice
- `_create_pie_chart()` - Grafico utenti per operatore

---

### 5. **app.py** - Interfaccia Streamlit

Entry point dell'applicazione web.

**Sezioni Principali:**

1. **Sidebar - Configurazione Dinamica**
   - Gestione tariffe (add/edit/remove)
   - Gestione eventi esclusi
   - Ripristino defaults
   - Salvataggio configurazione

2. **Upload File Excel**
   - Caricamento file da utente
   - Visualizzazione anteprima dati

3. **Tab di Navigazione:**
   - **Analisi**: Conteggi per persona e tipo
   - **Per Mese**: Andamento mensile
   - **Operatori**: Analisi per operatore
   - **Ricavi**: Calcoli economici
   - **Scartate**: Righe filtrate con motivo
   - **Grafici**: Visualizzazioni

4. **Export**
   - Pulsante download per file Excel formattato

**Dipendenze Streamlit:**
- `streamlit` - Framework web
- `plotly` - Grafici interattivi
- `pandas` - Manipolazione dati

---

## Flusso Dati

```
File Excel Upload
        ↓
    DataLoader
    ├─→ Estrae codici/tipi
    ├─→ Valida contro tariffe
    ├─→ Filtra per eventi
    └─→ Separa: df_valido | df_scartato
        ↓
    AnalisiPolitiche
    ├─→ Conteggi per persona/tipo
    ├─→ Conteggi per operatore
    ├─→ Conteggi per mese
    └─→ Calcoli ricavi
        ↓
    ExcelExporter
    ├─→ Crea workbook
    ├─→ Popola fogli con dati
    ├─→ Aggiunge grafici
    └─→ Salva bytes
        ↓
    Download Excel
```

---

## Colonne Dati Attese da Excel

Il file Excel di input deve contenere le seguenti colonne:
- `Destinatario` - Nome persona che riceve l'azione
- `Operatore` - Nome operatore che esegue l'azione
- `Attività` - Descrizione con codice (es. "A03 - Riunione")
- `Evento` - Tipo di evento (es. "Realizzazione", "Proposta")
- `Data Fine` - Data di conclusione (gg/mm/yyyy)
- `Data Proposta` - Data proposta (gg/mm/yyyy, usata per C06)

---

## Configurazione

### Modifica Tariffe
1. Via UI Streamlit (sidebar)
2. Direttamente in `config.json`

### Modifica Eventi Esclusi
1. Via UI Streamlit (sidebar)
2. Direttamente in `config.json` sezione `filtri.escludi_eventi`

### Salvataggio Configurazione
- UI salva automaticamente in `config.json`
- Persistenza tra sessioni

---

## Dipendenze Python

```
streamlit>=1.28.0      # Framework web interattivo
pandas>=2.0.0          # Manipolazione dati
openpyxl>=3.1.0        # Lettura/scrittura Excel
xlrd>=2.0.1            # Supporto file Excel
plotly>=5.18.0         # Grafici interattivi
```

---

## Funzionalità Chiave

### ✅ Conteggi Granulari
- Per persona (destinatario)
- Per tipo azione (A, B, C)
- Per codice specifico (A03, B04, C06, etc.)
- Per operatore
- Per mese/anno
- Combinazioni: persona-tipo-mese, operatore-mese

### ✅ Analisi Economica
- Calcolo ricavi con tariffe configurabili
- Ricavi per codice e mese
- Riepilogo ricavi aggregato
- Ricavo totale

### ✅ Controllo Qualità Dati
- Tracciamento righe scartate
- Motivo esclusione dettagliato
- Visualizzazione codici non riconosciuti

### ✅ Reportistica
- Export Excel multi-foglio
- Formattazione professionale
- Grafici automatici:
  - Andamento mensile (linee)
  - Ricavi per codice (barre)
  - Distribuzione utenti (torta)

### ✅ Configurabilità
- Tariffe modificabili da UI
- Filtri dinamici
- Nessena ricompilazione necessaria

---

## Utilizzo Tipico

1. **Setup Iniziale**
   - Clonare repository / aprire workspace
   - `pip install -r requirements.txt`
   - Opzionalmente: modificare `config.json`

2. **Avvio Applicazione**
   - `streamlit run app.py`
   - Si apre browser a `http://localhost:8501`

3. **Caricamento Dati**
   - Upload file Excel
   - Sistema valida automaticamente

4. **Analisi Interattiva**
   - Navigare tra tab per diverse View
   - Modificare tariffe nella sidebar se necessario
   - Visualizzare grafici

5. **Export Risultati**
   - Click su "Download Report Excel"
   - File con tutte le analisi e grafici

---

## Punti Tecnici Importanti

### Estrazione Codice Azione
- Regex: `r'^([A-Z]\d+)'` su campo "Attività"
- Estrae prime lettere e cifre (es. "A03" da "A03 - Nome Attività")

### Calcolo Data Riferimento
- **C06**: Usa `Data Proposta`
- **Altre**: Usa `Data Fine`
- Necessario per corretta aggregazione mensile

### Deduplica Operatore
- Per ogni persona: prende operatore dell'azione più recente
- Applicato in conteggi per garantire coerenza

### Filtraggio Righe
Due step:
1. Righe con evento in `escludi_eventi`
2. Righe con codice non in tariffe

Tutte tracciate in `df_scartate` con motivo

### Calcolo Ricavi
- Formula semplice: `Conteggio_Codice × Tariffa_Codice`
- Nessun arrotondamento forzato (pandas gestisce precisione)

---

## Estensibilità Futura

### Possibili Miglioramenti
1. **Database**: Sostituire JSON con DB per grandi volumi
2. **API**: Aggiungere endpoint FastAPI per accesso programmatico
3. **Grafici Avanzati**: Heatmap, scatter plot per relazioni
4. **Notifiche**: Alert per anomalie nei dati
5. **Audit**: Logging delle modifiche configurazione
6. **Validazione Schema**: Verifica struttura Excel rigida
7. **Scheduling**: Export automatico periodico
8. **Filtri Dinamici**: UI per definire criteri ricerca custom

---

## Diagnostic Utilities

### Debugging
- Stampare `df.info()` per schema
- Controllare `df_scartate` per capire filtraggi
- Verificare `config.tariffe` per coerenza

### Testing
- Mockare DataLoader con dati di test
- Verificare calcoli ricavi manualmente
- Controllare formattazione Excel

---

## Note sulla Configurazione

- **config.json** viene creato automaticamente se non esiste
- Le modifiche via UI vengono persistite immediatamente
- Valori default hard-coded in `ConfigManager.DEFAULT_CONFIG`
- Merge intelligente: se config.json manca campi, vengono aggiunti dai default

---

**Data Update Documento**: 16 Febbraio 2026
**Versione Progetto**: 1.0 (Analisi Base + Export Excel)
