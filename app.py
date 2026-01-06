"""
Analisi Politiche - Web App
Streamlit app per l'analisi e il conteggio delle azioni da file Excel.
"""

import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

# Configurazione pagina
st.set_page_config(
    page_title="Analisi Politiche",
    page_icon="📊",
    layout="wide"
)

# Configurazione di default (modificabile nella sidebar)
DEFAULT_CONFIG = {
    "tariffe": {
        "A03": 37.14,
        "A06": 35.57,
        "B03": 37.14,
        "B04": 37.14,
        "C06": 499.88
    },
    "filtri": {
        "escludi_eventi": ["Annullamento (prima dell'inizio)", "Proposta"]
    }
}


def load_data(uploaded_file, config: dict) -> pd.DataFrame:
    """Carica il file Excel e prepara i dati."""
    escludi_eventi = config.get('filtri', {}).get('escludi_eventi', [])
    tariffe = config.get('tariffe', DEFAULT_CONFIG['tariffe'])
    codici_validi = list(tariffe.keys())
    
    df = pd.read_excel(uploaded_file)
    
    info_filtri = []
    
    # Escludi righe con eventi da escludere
    righe_iniziali = len(df)
    for evento in escludi_eventi:
        df = df[df['Evento'] != evento]
    righe_escluse = righe_iniziali - len(df)
    if righe_escluse > 0:
        info_filtri.append(f"Escluse {righe_escluse} righe per eventi filtrati")
    
    # Crea una copia per evitare warning pandas
    df = df.copy()
    
    # Estrai codice azione (es. A03, B04, C06)
    df['Codice'] = df['Attività'].str.extract(r'^([A-Z]\d+)')
    
    # Filtra solo le azioni con codici presenti nelle tariffe
    righe_prima_filtro = len(df)
    df = df[df['Codice'].isin(codici_validi)].copy()
    righe_filtrate = righe_prima_filtro - len(df)
    if righe_filtrate > 0:
        info_filtri.append(f"Escluse {righe_filtrate} righe con codici non in tariffe")
    
    # Estrai tipo (prima lettera: A, B, C)
    df['Tipo'] = df['Codice'].str[0]
    
    # Determina la data da usare
    df['Data Riferimento'] = df.apply(
        lambda row: row['Data Proposta'] if row['Codice'] == 'C06' else row['Data Fine'],
        axis=1
    )
    
    # Converti in datetime
    df['Data Riferimento'] = pd.to_datetime(df['Data Riferimento'], format='%d/%m/%Y', errors='coerce')
    
    # Estrai anno-mese per aggregazione
    df['Anno-Mese'] = df['Data Riferimento'].dt.to_period('M')
    
    return df, info_filtri


def get_operatore_per_persona(df: pd.DataFrame) -> pd.DataFrame:
    """Trova l'operatore più recente per ogni persona."""
    df_sorted = df.sort_values('Data Riferimento', ascending=False)
    operatori = df_sorted.groupby('Destinatario').first()['Operatore'].reset_index()
    operatori.columns = ['Destinatario', 'Operatore']
    return operatori


def conteggio_per_persona_tipo(df: pd.DataFrame) -> pd.DataFrame:
    """Conteggio azioni per persona e per tipo."""
    pivot_tipo = pd.pivot_table(
        df, index='Destinatario', columns='Tipo',
        aggfunc='size', fill_value=0
    ).reset_index()
    
    pivot_codice = pd.pivot_table(
        df, index='Destinatario', columns='Codice',
        aggfunc='size', fill_value=0
    ).reset_index()
    
    codici_dettaglio = ['A03', 'A06', 'B03', 'B04', 'C06']
    codici_presenti = [c for c in codici_dettaglio if c in pivot_codice.columns]
    
    pivot = pivot_tipo.merge(
        pivot_codice[['Destinatario'] + codici_presenti],
        on='Destinatario', how='left'
    )
    
    for col in ['A', 'B', 'C', 'A03', 'A06', 'B03', 'B04', 'C06']:
        if col not in pivot.columns:
            pivot[col] = 0
    
    operatori = get_operatore_per_persona(df)
    pivot = pivot.merge(operatori, on='Destinatario', how='left')
    
    pivot = pivot[['Destinatario', 'Operatore', 'A', 'A03', 'A06', 'B', 'B03', 'B04', 'C', 'C06']]
    pivot['Totale'] = pivot['A'] + pivot['B'] + pivot['C']
    pivot = pivot.sort_values('Destinatario')
    
    return pivot


def conteggio_per_persona_tipo_mese(df: pd.DataFrame) -> pd.DataFrame:
    """Conteggio azioni per persona, tipo e mese."""
    df_temp = df.copy()
    df_temp['Tipo-Mese'] = df_temp['Tipo'] + '_' + df_temp['Anno-Mese'].astype(str)
    
    pivot = pd.pivot_table(
        df_temp, index='Destinatario', columns='Tipo-Mese',
        aggfunc='size', fill_value=0
    ).reset_index()
    
    operatori = get_operatore_per_persona(df)
    pivot = pivot.merge(operatori, on='Destinatario', how='left')
    
    cols = pivot.columns.tolist()
    cols.remove('Operatore')
    cols.insert(1, 'Operatore')
    pivot = pivot[cols]
    
    return pivot.sort_values('Destinatario')


def conteggio_totale_tipo(df: pd.DataFrame) -> pd.DataFrame:
    """Conteggio totale per tipo."""
    conteggio_tipo = df.groupby('Tipo').size().reset_index(name='Conteggio')
    conteggio_codice = df.groupby('Codice').size().reset_index(name='Conteggio')
    conteggio_codice.columns = ['Tipo', 'Conteggio']
    
    risultato = pd.concat([conteggio_tipo, conteggio_codice], ignore_index=True)
    risultato = risultato.sort_values('Tipo')
    
    totale = pd.DataFrame([{'Tipo': 'TOTALE', 'Conteggio': len(df)}])
    risultato = pd.concat([risultato, totale], ignore_index=True)
    
    return risultato


def conteggio_totale_tipo_mese(df: pd.DataFrame) -> pd.DataFrame:
    """Conteggio totale per tipo e mese."""
    pivot = pd.pivot_table(
        df, index='Tipo', columns='Anno-Mese',
        aggfunc='size', fill_value=0
    ).reset_index()
    
    pivot.columns = [str(col) for col in pivot.columns]
    
    mesi_cols = [c for c in pivot.columns if c != 'Tipo']
    totali = pivot[mesi_cols].sum()
    totale_row = pd.DataFrame([['TOTALE'] + totali.tolist()], columns=pivot.columns)
    pivot = pd.concat([pivot, totale_row], ignore_index=True)
    
    return pivot


def conteggio_per_operatore(df: pd.DataFrame) -> pd.DataFrame:
    """Conteggio azioni per operatore."""
    pivot = pd.pivot_table(
        df, index='Operatore', columns='Codice',
        aggfunc='size', fill_value=0
    ).reset_index()
    
    for codice in ['A03', 'A06', 'B03', 'B04', 'C06']:
        if codice not in pivot.columns:
            pivot[codice] = 0
    
    cols = ['Operatore', 'A03', 'A06', 'B03', 'B04', 'C06']
    pivot = pivot[cols]
    pivot['Totale'] = pivot['A03'] + pivot['A06'] + pivot['B03'] + pivot['B04'] + pivot['C06']
    
    return pivot.sort_values('Operatore')


def conteggio_per_operatore_mese_tipo(df: pd.DataFrame) -> pd.DataFrame:
    """Conteggio azioni per operatore, mese e tipo."""
    df_temp = df.copy()
    df_temp['Mese'] = df_temp['Anno-Mese'].astype(str)
    
    pivot = pd.pivot_table(
        df_temp, index=['Operatore', 'Mese'], columns='Codice',
        aggfunc='size', fill_value=0
    ).reset_index()
    
    for codice in ['A03', 'A06', 'B03', 'B04', 'C06']:
        if codice not in pivot.columns:
            pivot[codice] = 0
    
    cols = ['Operatore', 'Mese', 'A03', 'A06', 'B03', 'B04', 'C06']
    pivot = pivot[cols]
    pivot['Totale'] = pivot['A03'] + pivot['A06'] + pivot['B03'] + pivot['B04'] + pivot['C06']
    
    return pivot.sort_values(['Operatore', 'Mese'])


def calcolo_ricavi(df: pd.DataFrame, tariffe: dict) -> pd.DataFrame:
    """Calcola ricavi per tipologia e mese."""
    df_temp = df.copy()
    df_temp['Mese'] = df_temp['Anno-Mese'].astype(str)
    
    pivot = pd.pivot_table(
        df_temp, index='Mese', columns='Codice',
        aggfunc='size', fill_value=0
    ).reset_index()
    
    for codice in tariffe.keys():
        if codice not in pivot.columns:
            pivot[codice] = 0
    
    cols_codice = list(tariffe.keys())
    pivot = pivot[['Mese'] + cols_codice]
    
    for codice, tariffa in tariffe.items():
        pivot[f'{codice}_ricavo'] = pivot[codice] * tariffa
    
    cols_ricavo = [f'{c}_ricavo' for c in cols_codice]
    pivot['Totale_Conteggio'] = pivot[cols_codice].sum(axis=1)
    pivot['Totale_Ricavo'] = pivot[cols_ricavo].sum(axis=1)
    
    return pivot.sort_values('Mese')


def riepilogo_ricavi_per_tipo(df: pd.DataFrame, tariffe: dict) -> pd.DataFrame:
    """Riepilogo ricavi aggregato per tipo."""
    conteggio = df.groupby('Codice').size().reset_index(name='Conteggio')
    
    conteggio['Tariffa (€)'] = conteggio['Codice'].map(tariffe)
    conteggio['Ricavo (€)'] = conteggio['Conteggio'] * conteggio['Tariffa (€)']
    
    totale = pd.DataFrame([{
        'Codice': 'TOTALE',
        'Tariffa (€)': None,
        'Conteggio': conteggio['Conteggio'].sum(),
        'Ricavo (€)': conteggio['Ricavo (€)'].sum()
    }])
    
    risultato = pd.concat([conteggio, totale], ignore_index=True)
    return risultato[['Codice', 'Tariffa (€)', 'Conteggio', 'Ricavo (€)']]


def to_excel(dfs: dict) -> bytes:
    """Converte i DataFrame in un file Excel."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for nome_foglio, df in dfs.items():
            df.to_excel(writer, sheet_name=nome_foglio, index=False)
    return output.getvalue()


# ============================================================
# INTERFACCIA STREAMLIT
# ============================================================

st.title("📊 Analisi Politiche")
st.markdown("Carica un file Excel per analizzare conteggi e ricavi delle azioni.")

# Sidebar per configurazione
with st.sidebar:
    st.header("⚙️ Configurazione")
    
    st.subheader("Tariffe (€)")
    tariffe = {}
    for codice, default_value in DEFAULT_CONFIG['tariffe'].items():
        tariffe[codice] = st.number_input(
            f"Tariffa {codice}",
            value=default_value,
            step=0.01,
            format="%.2f",
            key=f"tariffa_{codice}"
        )
    
    st.subheader("Eventi da Escludere")
    eventi_default = ", ".join(DEFAULT_CONFIG['filtri']['escludi_eventi'])
    eventi_input = st.text_area(
        "Un evento per riga",
        value="\n".join(DEFAULT_CONFIG['filtri']['escludi_eventi']),
        height=100
    )
    escludi_eventi = [e.strip() for e in eventi_input.split("\n") if e.strip()]

config = {
    "tariffe": tariffe,
    "filtri": {"escludi_eventi": escludi_eventi}
}

# Upload file
uploaded_file = st.file_uploader(
    "Carica file Excel (.xls o .xlsx)",
    type=['xls', 'xlsx']
)

if uploaded_file is not None:
    try:
        with st.spinner("Analisi in corso..."):
            df, info_filtri = load_data(uploaded_file, config)
        
        # Mostra info filtri
        if info_filtri:
            for info in info_filtri:
                st.info(info)
        
        # Metriche principali
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Righe Analizzate", len(df))
        with col2:
            st.metric("Persone Uniche", df['Destinatario'].nunique())
        with col3:
            st.metric("Operatori", df['Operatore'].nunique())
        with col4:
            ricavi_totali = (df.groupby('Codice').size() * df['Codice'].map(tariffe)).sum()
            st.metric("Ricavi Totali", f"€ {ricavi_totali:,.2f}")
        
        # Tabs per le diverse viste
        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
            "📋 Per Persona",
            "📅 Per Persona/Mese",
            "📊 Totali",
            "👤 Per Operatore",
            "💰 Ricavi",
            "📈 Ricavi/Mese"
        ])
        
        with tab1:
            st.subheader("Conteggio per Persona e Tipo")
            df_persona = conteggio_per_persona_tipo(df)
            st.dataframe(df_persona, use_container_width=True, hide_index=True)
        
        with tab2:
            st.subheader("Conteggio per Persona, Tipo e Mese")
            df_persona_mese = conteggio_per_persona_tipo_mese(df)
            st.dataframe(df_persona_mese, use_container_width=True, hide_index=True)
        
        with tab3:
            col_a, col_b = st.columns(2)
            with col_a:
                st.subheader("Totali per Tipo")
                st.dataframe(conteggio_totale_tipo(df), use_container_width=True, hide_index=True)
            with col_b:
                st.subheader("Totali per Tipo e Mese")
                st.dataframe(conteggio_totale_tipo_mese(df), use_container_width=True, hide_index=True)
        
        with tab4:
            col_a, col_b = st.columns(2)
            with col_a:
                st.subheader("Per Operatore")
                st.dataframe(conteggio_per_operatore(df), use_container_width=True, hide_index=True)
            with col_b:
                st.subheader("Per Operatore e Mese")
                st.dataframe(conteggio_per_operatore_mese_tipo(df), use_container_width=True, hide_index=True)
        
        with tab5:
            st.subheader("Riepilogo Ricavi per Tipo")
            st.dataframe(riepilogo_ricavi_per_tipo(df, tariffe), use_container_width=True, hide_index=True)
        
        with tab6:
            st.subheader("Ricavi per Mese")
            st.dataframe(calcolo_ricavi(df, tariffe), use_container_width=True, hide_index=True)
        
        # Download Excel
        st.markdown("---")
        
        # Prepara tutti i fogli per l'export
        excel_data = to_excel({
            "Per Persona-Tipo": conteggio_per_persona_tipo(df),
            "Per Persona-Tipo-Mese": conteggio_per_persona_tipo_mese(df),
            "Totali per Tipo": conteggio_totale_tipo(df),
            "Totali per Tipo-Mese": conteggio_totale_tipo_mese(df),
            "Per Operatore": conteggio_per_operatore(df),
            "Per Operatore-Mese": conteggio_per_operatore_mese_tipo(df),
            "Riepilogo Ricavi": riepilogo_ricavi_per_tipo(df, tariffe),
            "Ricavi per Mese": calcolo_ricavi(df, tariffe)
        })
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        filename = f"{timestamp}_export_analisi.xlsx"
        
        st.download_button(
            label="📥 Scarica Report Excel",
            data=excel_data,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"Errore durante l'analisi: {str(e)}")
        st.exception(e)

else:
    st.info("👆 Carica un file Excel per iniziare l'analisi")
    
    with st.expander("ℹ️ Informazioni sul formato del file"):
        st.markdown("""
        Il file Excel deve contenere le seguenti colonne:
        - **Destinatario**: Nome della persona
        - **Operatore**: Nome dell'operatore
        - **Attività**: Codice attività (es. A03, B04, C06)
        - **Evento**: Stato dell'attività
        - **Data Fine**: Data di completamento
        - **Data Proposta**: Data proposta (usata per C06)
        """)