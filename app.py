"""
Analisi Politiche - Applicazione Web
Entry point dell'applicazione Streamlit.
"""

import streamlit as st
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go

# Import moduli locali
from src import ConfigManager, DataLoader, AnalisiPolitiche, ExcelExporter

# Configurazione pagina
st.set_page_config(
    page_title="Analisi Politiche",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)


def init_session_state():
    """Inizializza lo stato della sessione."""
    if 'config_manager' not in st.session_state:
        st.session_state.config_manager = ConfigManager()
    
    if 'tariffe' not in st.session_state:
        st.session_state.tariffe = st.session_state.config_manager.tariffe.copy()
    
    if 'escludi_eventi' not in st.session_state:
        st.session_state.escludi_eventi = st.session_state.config_manager.escludi_eventi.copy()


def render_sidebar():
    """Renderizza la sidebar con la configurazione."""
    with st.sidebar:
        st.header("⚙️ Configurazione")
        
        # --- TARIFFE DINAMICHE ---
        st.subheader("💰 Tariffe (€)")
        
        # Form per aggiungere nuova tariffa
        with st.expander("➕ Aggiungi nuova tariffa"):
            col1, col2 = st.columns(2)
            with col1:
                nuovo_codice = st.text_input(
                    "Codice (es. A07)", 
                    key="nuovo_codice", 
                    max_chars=5
                )
            with col2:
                nuova_tariffa = st.number_input(
                    "Tariffa €", 
                    value=0.0, 
                    step=0.01, 
                    key="nuova_tariffa"
                )
            
            if st.button("Aggiungi", key="btn_aggiungi"):
                if nuovo_codice and nuovo_codice.strip():
                    codice = nuovo_codice.strip().upper()
                    st.session_state.tariffe[codice] = nuova_tariffa
                    st.success(f"Aggiunta tariffa {codice}: €{nuova_tariffa:.2f}")
                    st.rerun()
        
        # Mostra tariffe esistenti
        st.markdown("**Tariffe attuali:**")
        tariffe_da_rimuovere = []
        
        for codice in sorted(st.session_state.tariffe.keys()):
            col1, col2, col3 = st.columns([2, 2, 1])
            with col1:
                st.text(codice)
            with col2:
                nuova_val = st.number_input(
                    f"tariffa_{codice}",
                    value=st.session_state.tariffe[codice],
                    step=0.01,
                    format="%.2f",
                    key=f"tariffa_{codice}",
                    label_visibility="collapsed"
                )
                st.session_state.tariffe[codice] = nuova_val
            with col3:
                if st.button("🗑️", key=f"rimuovi_{codice}"):
                    tariffe_da_rimuovere.append(codice)
        
        # Rimuovi tariffe marcate
        for codice in tariffe_da_rimuovere:
            del st.session_state.tariffe[codice]
            st.rerun()
        
        # Pulsanti gestione tariffe
        col1, col2 = st.columns(2)
        with col1:
            if st.button("🔄 Ripristina"):
                st.session_state.tariffe = st.session_state.config_manager.tariffe.copy()
                st.rerun()
        with col2:
            if st.button("💾 Salva"):
                st.session_state.config_manager.tariffe = st.session_state.tariffe.copy()
                if st.session_state.config_manager.save():
                    st.success("Salvato!")
                else:
                    st.error("Errore nel salvataggio")
        
        st.divider()
        
        # --- EVENTI DA ESCLUDERE ---
        st.subheader("🚫 Eventi da Escludere")
        eventi_input = st.text_area(
            "Un evento per riga",
            value="\n".join(st.session_state.escludi_eventi),
            height=100,
            key="eventi_esclusi_input"
        )
        st.session_state.escludi_eventi = [
            e.strip() for e in eventi_input.split("\n") if e.strip()
        ]
        
        # Pulsante salva eventi
        if st.button("💾 Salva Eventi"):
            st.session_state.config_manager.escludi_eventi = st.session_state.escludi_eventi.copy()
            if st.session_state.config_manager.save():
                st.success("Eventi salvati!")
            else:
                st.error("Errore nel salvataggio")


def render_results(df, df_scartate, analisi):
    """Renderizza i risultati dell'analisi."""
    
    # Metriche principali
    st.subheader("📈 Riepilogo")
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        st.metric("Righe Analizzate", len(df))
    with col2:
        st.metric("Persone Uniche", df['Destinatario'].nunique())
    with col3:
        st.metric("Operatori", df['Operatore'].nunique())
    with col4:
        st.metric("Righe Scartate", len(df_scartate))
    with col5:
        ricavi_totali = analisi.ricavi_totali()
        st.metric("Ricavi Totali", f"€ {ricavi_totali:,.2f}")
    
    # --- AZIONI SCARTATE PER CODICE ---
    if len(df_scartate) > 0:
        loader = DataLoader(st.session_state.tariffe, st.session_state.escludi_eventi)
        df_scartate_codice = loader.conteggio_scartate_per_codice(df_scartate)
        
        if len(df_scartate_codice) > 1:  # Più di solo la riga TOTALE
            st.markdown("---")
            st.subheader("⚠️ Azioni Scartate per Codice (non in tariffe)")
            
            col_tabella, col_info = st.columns([1, 2])
            with col_tabella:
                st.dataframe(df_scartate_codice, use_container_width=True, hide_index=True)
            with col_info:
                st.info(
                    "Questi codici sono stati trovati nel file ma non sono presenti "
                    "nella configurazione delle tariffe. Per includerli nell'analisi, "
                    "aggiungi i codici mancanti nella sezione Tariffe della barra laterale."
                )
    
    # --- GRAFICI ---
    st.markdown("---")
    
    col_grafico1, col_grafico2 = st.columns(2)
    
    with col_grafico1:
        st.subheader("📈 Andamento Mensile Azioni")
        df_andamento = analisi.andamento_mensile()
        
        # Prepara dati per il grafico a linee
        tipi_presenti = [c for c in df_andamento.columns if c not in ['Mese', 'Totale']]
        
        fig_linee = go.Figure()
        
        # Aggiungi una linea per ogni tipo
        colori = {'A': '#1f77b4', 'B': '#ff7f0e', 'C': '#2ca02c'}
        for tipo in tipi_presenti:
            fig_linee.add_trace(go.Scatter(
                x=df_andamento['Mese'],
                y=df_andamento[tipo],
                mode='lines+markers',
                name=f'Tipo {tipo}',
                line=dict(color=colori.get(tipo, '#666666'), width=2),
                marker=dict(size=8)
            ))
        
        # Aggiungi linea totale
        fig_linee.add_trace(go.Scatter(
            x=df_andamento['Mese'],
            y=df_andamento['Totale'],
            mode='lines+markers',
            name='Totale',
            line=dict(color='#d62728', width=3, dash='dash'),
            marker=dict(size=10)
        ))
        
        fig_linee.update_layout(
            xaxis_title="Mese",
            yaxis_title="Quantità",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            height=400,
            margin=dict(l=20, r=20, t=40, b=20)
        )
        
        st.plotly_chart(fig_linee, use_container_width=True)
    
    with col_grafico2:
        st.subheader("💰 Ricavi per Tipo Azione")
        df_ricavi = analisi.ricavi_per_codice()
        
        # Grafico a barre orizzontali
        fig_barre = go.Figure(go.Bar(
            x=df_ricavi['Ricavo'],
            y=df_ricavi['Codice'],
            orientation='h',
            text=df_ricavi['Ricavo'].apply(lambda x: f'€ {x:,.2f}'),
            textposition='outside',
            marker_color=['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd'][:len(df_ricavi)]
        ))
        
        fig_barre.update_layout(
            xaxis_title="Ricavo (€)",
            yaxis_title="Codice Azione",
            height=400,
            margin=dict(l=20, r=80, t=40, b=20),
            xaxis=dict(tickformat='€,.0f')
        )
        
        st.plotly_chart(fig_barre, use_container_width=True)
    
    # --- UTENTI PER OPERATORE ---
    st.markdown("---")
    st.subheader("👥 Utenti per Operatore")
    
    col_tabella, col_grafico = st.columns([1, 2])
    
    with col_tabella:
        df_utenti = analisi.utenti_per_operatore()
        st.dataframe(df_utenti, use_container_width=True, hide_index=True)
    
    with col_grafico:
        # Escludi la riga TOTALE per il grafico
        df_utenti_grafico = df_utenti[df_utenti['Operatore'] != 'TOTALE']
        
        fig_utenti = px.pie(
            df_utenti_grafico, 
            values='Numero Utenti', 
            names='Operatore',
            hole=0.4,  # Grafico a ciambella
            color_discrete_sequence=px.colors.qualitative.Set2
        )
        fig_utenti.update_traces(textposition='inside', textinfo='percent+value')
        fig_utenti.update_layout(
            height=350,
            margin=dict(l=20, r=20, t=20, b=20),
            legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5)
        )
        st.plotly_chart(fig_utenti, use_container_width=True)
    
    st.markdown("---")
    
    # Tabs per le diverse viste
    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
        "📋 Per Persona",
        "📅 Per Persona/Mese",
        "📊 Totali",
        "👤 Per Operatore",
        "💰 Ricavi",
        "📈 Ricavi/Mese",
        "🚫 Righe Scartate"
    ])
    
    with tab1:
        st.subheader("Conteggio per Persona e Tipo")
        st.dataframe(
            analisi.conteggio_per_persona_tipo(), 
            use_container_width=True, 
            hide_index=True
        )
    
    with tab2:
        st.subheader("Conteggio per Persona, Tipo e Mese")
        st.dataframe(
            analisi.conteggio_per_persona_tipo_mese(), 
            use_container_width=True, 
            hide_index=True
        )
    
    with tab3:
        col_a, col_b = st.columns(2)
        with col_a:
            st.subheader("Totali per Tipo")
            st.dataframe(
                analisi.conteggio_totale_tipo(), 
                use_container_width=True, 
                hide_index=True
            )
        with col_b:
            st.subheader("Totali per Tipo e Mese")
            st.dataframe(
                analisi.conteggio_totale_tipo_mese(), 
                use_container_width=True, 
                hide_index=True
            )
    
    with tab4:
        col_a, col_b = st.columns(2)
        with col_a:
            st.subheader("Per Operatore")
            st.dataframe(
                analisi.conteggio_per_operatore(), 
                use_container_width=True, 
                hide_index=True
            )
        with col_b:
            st.subheader("Per Operatore e Mese")
            st.dataframe(
                analisi.conteggio_per_operatore_mese(), 
                use_container_width=True, 
                hide_index=True
            )
    
    with tab5:
        st.subheader("Riepilogo Ricavi per Tipo")
        st.dataframe(
            analisi.riepilogo_ricavi(), 
            use_container_width=True, 
            hide_index=True
        )
    
    with tab6:
        st.subheader("Ricavi per Mese")
        st.dataframe(
            analisi.calcolo_ricavi_per_mese(), 
            use_container_width=True, 
            hide_index=True
        )
    
    with tab7:
        st.subheader("Righe Scartate")
        if len(df_scartate) > 0:
            # Raggruppa per motivo
            loader = DataLoader(st.session_state.tariffe, st.session_state.escludi_eventi)
            motivi = loader.riepilogo_scartate(df_scartate)
            
            st.markdown("**Riepilogo motivi di esclusione:**")
            for motivo, count in motivi.items():
                st.write(f"- {motivo}: **{count}** righe")
            
            st.divider()
            
            # Mostra dettaglio
            df_scartate_display = loader.prepara_scartate_per_export(df_scartate)
            st.dataframe(df_scartate_display, use_container_width=True, hide_index=True)
        else:
            st.success("✅ Nessuna riga scartata!")


def main():
    """Funzione principale dell'applicazione."""
    
    # Inizializza stato
    init_session_state()
    
    # Titolo
    st.title("📊 Analisi Politiche")
    st.markdown("Carica un file Excel per analizzare conteggi e ricavi delle azioni.")
    
    # Sidebar
    render_sidebar()
    
    # Upload file
    uploaded_file = st.file_uploader(
        "Carica file Excel (.xls o .xlsx)",
        type=['xls', 'xlsx']
    )
    
    if uploaded_file is not None:
        try:
            with st.spinner("Analisi in corso..."):
                # Carica dati
                loader = DataLoader(
                    st.session_state.tariffe, 
                    st.session_state.escludi_eventi
                )
                df, df_scartate = loader.load(uploaded_file)
                
                # Analizza
                analisi = AnalisiPolitiche(df, st.session_state.tariffe)
            
            # Mostra risultati
            render_results(df, df_scartate, analisi)
            
            # Download Excel
            st.markdown("---")
            
            exporter = ExcelExporter(analisi, df_scartate)
            excel_data = exporter.export()
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M")
            prefisso = st.session_state.config_manager.prefisso_output
            filename = f"{timestamp}_{prefisso}.xlsx"
            
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
        
        with st.expander("📋 Tariffe configurate"):
            st.markdown("Tariffe attualmente configurate (modificabili nella barra laterale):")
            for codice, tariffa in sorted(st.session_state.tariffe.items()):
                st.write(f"- **{codice}**: €{tariffa:.2f}")


if __name__ == "__main__":
    main()
