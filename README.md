# 📊 Analisi Politiche

Applicazione web per l'analisi e il conteggio delle azioni da file Excel delle politiche.

## Funzionalità

- ✅ Conteggio azioni per persona e tipo (A, B, C) con dettaglio codici
- ✅ Conteggio azioni per persona, tipo e mese
- ✅ Totali complessivi per tipo e per mese
- ✅ Conteggio per operatore
- ✅ Calcolo ricavi con tariffe configurabili
- ✅ Export in Excel formattato con più fogli
- ✅ Visualizzazione righe scartate con motivo di esclusione
- ✅ Configurazione dinamica tariffe ed eventi da escludere

## Struttura Progetto

```
analisi-politiche/
├── app.py                 # Entry point applicazione Streamlit
├── config.json            # File di configurazione
├── requirements.txt       # Dipendenze Python
├── README.md              # Questo file
├── .gitignore             # File da ignorare in Git
└── src/
    ├── __init__.py        # Init package
    ├── config.py          # Gestione configurazione
    ├── data_loader.py     # Caricamento e preprocessing dati
    ├── analysis.py        # Funzioni di analisi/conteggio
    └── excel_export.py    # Formattazione e export Excel
```

## Configurazione

Il file `config.json` contiene le impostazioni dell'applicazione:

```json
{
    "tariffe": {
        "A03": 37.14,
        "A06": 35.57,
        "B03": 37.14,
        "B04": 37.14,
        "C06": 499.88
    },
    "filtri": {
        "escludi_eventi": [
            "Annullamento (prima dell'inizio)",
            "Proposta"
        ]
    },
    "output": {
        "prefisso_nome": "export_analisi"
    }
}
```

### Parametri

| Sezione | Parametro | Descrizione |
|---------|-----------|-------------|
| `tariffe.*` | Codice: Valore | Tariffa in euro per ogni codice azione |
| `filtri.escludi_eventi` | Lista | Eventi da escludere dall'analisi (con eccezioni per C06) |
| `output.prefisso_nome` | String | Prefisso per i file di output |

**Nota importante sui filtri**: L'evento "Proposta" viene escluso per tutte le azioni TRANNE C06, che deve mantenere le proposte dato che utilizza la Data Proposta come riferimento.

Le tariffe e gli eventi possono essere modificati anche dalla sidebar dell'applicazione.

## Installazione Locale

### Prerequisiti

- Python 3.8 o superiore
- pip

### Setup

```bash
# Clona il repository
git clone https://github.com/TUO_USERNAME/analisi-politiche.git
cd analisi-politiche

# Crea virtual environment (opzionale ma consigliato)
python -m venv venv
source venv/bin/activate  # Linux/macOS
venv\Scripts\activate     # Windows

# Installa dipendenze
pip install -r requirements.txt
```

### Esecuzione

```bash
streamlit run app.py
```

L'applicazione si aprirà automaticamente nel browser all'indirizzo `http://localhost:8501`

## Deploy su Streamlit Cloud

1. **Crea repository GitHub** con tutti i file del progetto
2. **Vai su** [share.streamlit.io](https://share.streamlit.io)
3. **Accedi con GitHub**
4. **New app** → seleziona il repository → `app.py` → **Deploy!**

Riceverai un URL pubblico per accedere all'applicazione.

## Formato File Input

Il file Excel deve contenere le seguenti colonne:

| Colonna | Descrizione |
|---------|-------------|
| `Destinatario` | Nome della persona |
| `Operatore` | Nome dell'operatore |
| `Attività` | Codice attività (es. A03, B04, C06) |
| `Evento` | Stato dell'attività |
| `Data Fine` | Data di completamento |
| `Data Proposta` | Data proposta (usata per C06) |

## Output Excel

Il report generato contiene 9 fogli:

1. **Riepilogo** - Statistiche generali, totali e top 10 persone
2. **Per Persona-Tipo** - Conteggio dettagliato per persona
3. **Per Persona-Tipo-Mese** - Conteggio per persona e mese
4. **Totali per Tipo** - Riepilogo globale
5. **Totali per Tipo-Mese** - Distribuzione mensile
6. **Per Operatore** - Conteggio per operatore
7. **Per Operatore-Mese** - Conteggio per operatore e mese
8. **Ricavi per Mese** - Dettaglio ricavi mensili
9. **Righe Scartate** - Righe escluse con motivo

## Licenza

Uso interno - Btinkeeng s.r.l.
