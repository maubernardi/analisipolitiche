"""
Caricamento e preprocessing dei dati Excel.
"""

import pandas as pd
from typing import Tuple, List, Dict, Any, Union
from io import BytesIO


class DataLoader:
    """Carica e preprocessa i dati da file Excel."""
    
    def __init__(self, tariffe: Dict[str, float], escludi_eventi: List[str]):
        """
        Inizializza il loader.
        
        Args:
            tariffe: Dizionario codice -> tariffa (es. {"A03": 37.14})
            escludi_eventi: Lista di eventi da escludere
        """
        self.tariffe = tariffe
        self.escludi_eventi = escludi_eventi
        self.codici_validi = list(tariffe.keys())
    
    def load(self, file_source: Union[str, BytesIO]) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """
        Carica e preprocessa i dati.
        
        Args:
            file_source: Percorso file o BytesIO (per upload Streamlit)
        
        Returns:
            Tuple (df_valido, df_scartato) con i dati processati
        """
        df_raw = pd.read_excel(file_source)
        
        # DataFrame per le righe scartate
        righe_scartate = []
        
        # Crea una copia per lavorare
        df = df_raw.copy()
        
        # Aggiungi colonna per tracciare l'indice originale (riga Excel)
        df['_indice_originale'] = df.index + 2  # +2 per header Excel
        
        # 1. Escludi righe con eventi da escludere
        for evento in self.escludi_eventi:
            mask = df['Evento'] == evento
            scartate = df[mask].copy()
            if len(scartate) > 0:
                scartate['_motivo_esclusione'] = f"Evento escluso: {evento}"
                righe_scartate.append(scartate)
            df = df[~mask]
        
        # Crea una copia per evitare warning pandas
        df = df.copy()
        
        # Estrai codice azione (es. A03, B04, C06)
        df['Codice'] = df['Attività'].str.extract(r'^([A-Z]\d+)')
        
        # 2. Escludi righe con codici non validi (non in tariffe)
        mask_codici_invalidi = ~df['Codice'].isin(self.codici_validi)
        scartate_codici = df[mask_codici_invalidi].copy()
        if len(scartate_codici) > 0:
            scartate_codici['_motivo_esclusione'] = scartate_codici['Codice'].apply(
                lambda x: f"Codice non in tariffe: {x}" if pd.notna(x) else "Codice non riconosciuto"
            )
            righe_scartate.append(scartate_codici)
        
        df = df[~mask_codici_invalidi].copy()
        
        # Estrai tipo (prima lettera: A, B, C)
        df['Tipo'] = df['Codice'].str[0]
        
        # Determina la data da usare (C06 usa Data Proposta, altre usano Data Fine)
        df['Data Riferimento'] = df.apply(
            lambda row: row['Data Proposta'] if row['Codice'] == 'C06' else row['Data Fine'],
            axis=1
        )
        
        # Converti in datetime
        df['Data Riferimento'] = pd.to_datetime(
            df['Data Riferimento'], 
            format='%d/%m/%Y', 
            errors='coerce'
        )
        
        # Estrai anno-mese per aggregazione
        df['Anno-Mese'] = df['Data Riferimento'].dt.to_period('M')
        
        # Combina tutte le righe scartate
        if righe_scartate:
            df_scartate = pd.concat(righe_scartate, ignore_index=True)
        else:
            df_scartate = pd.DataFrame()
        
        return df, df_scartate
    
    def get_statistiche_base(self, df: pd.DataFrame) -> Dict[str, Any]:
        """
        Restituisce statistiche base sul DataFrame.
        
        Args:
            df: DataFrame processato
        
        Returns:
            Dizionario con statistiche
        """
        return {
            'totale_righe': len(df),
            'persone_uniche': df['Destinatario'].nunique(),
            'operatori_unici': df['Operatore'].nunique(),
            'tipi_presenti': sorted(df['Tipo'].unique().tolist()),
            'codici_presenti': sorted(df['Codice'].unique().tolist()),
            'periodo': {
                'inizio': df['Data Riferimento'].min(),
                'fine': df['Data Riferimento'].max()
            }
        }
    
    def prepara_scartate_per_export(self, df_scartate: pd.DataFrame) -> pd.DataFrame:
        """
        Prepara il DataFrame delle righe scartate per la visualizzazione/export.
        
        Args:
            df_scartate: DataFrame con le righe scartate
        
        Returns:
            DataFrame formattato per export
        """
        if len(df_scartate) == 0:
            return pd.DataFrame(columns=[
                'Riga Excel', 'Destinatario', 'Operatore', 
                'Attività', 'Evento', 'Motivo Esclusione'
            ])
        
        cols_to_show = [
            '_indice_originale', 'Destinatario', 'Operatore', 
            'Attività', 'Evento', '_motivo_esclusione'
        ]
        cols_available = [c for c in cols_to_show if c in df_scartate.columns]
        
        df_export = df_scartate[cols_available].copy()
        df_export.columns = [
            'Riga Excel', 'Destinatario', 'Operatore', 
            'Attività', 'Evento', 'Motivo Esclusione'
        ][:len(cols_available)]
        
        return df_export
    
    def riepilogo_scartate(self, df_scartate: pd.DataFrame) -> Dict[str, int]:
        """
        Restituisce un riepilogo dei motivi di esclusione.
        
        Args:
            df_scartate: DataFrame con le righe scartate
        
        Returns:
            Dizionario motivo -> conteggio
        """
        if len(df_scartate) == 0 or '_motivo_esclusione' not in df_scartate.columns:
            return {}
        
        return df_scartate['_motivo_esclusione'].value_counts().to_dict()
