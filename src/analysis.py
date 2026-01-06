"""
Funzioni di analisi e conteggio per le politiche.
"""

import pandas as pd
from typing import Dict, List


class AnalisiPolitiche:
    """Classe per l'analisi dei dati delle politiche."""
    
    def __init__(self, df: pd.DataFrame, tariffe: Dict[str, float]):
        """
        Inizializza l'analisi.
        
        Args:
            df: DataFrame preprocessato con i dati
            tariffe: Dizionario codice -> tariffa
        """
        self.df = df
        self.tariffe = tariffe
        self.codici = list(tariffe.keys())
    
    def _get_operatore_per_persona(self) -> pd.DataFrame:
        """Trova l'operatore più recente per ogni persona."""
        df_sorted = self.df.sort_values('Data Riferimento', ascending=False)
        operatori = df_sorted.groupby('Destinatario').first()['Operatore'].reset_index()
        operatori.columns = ['Destinatario', 'Operatore']
        return operatori
    
    def conteggio_per_persona_tipo(self) -> pd.DataFrame:
        """Conteggio azioni per persona e per tipo."""
        # Pivot per tipo (A, B, C)
        pivot_tipo = pd.pivot_table(
            self.df, index='Destinatario', columns='Tipo',
            aggfunc='size', fill_value=0
        ).reset_index()
        
        # Pivot per codice (A03, A06, B03, etc.)
        pivot_codice = pd.pivot_table(
            self.df, index='Destinatario', columns='Codice',
            aggfunc='size', fill_value=0
        ).reset_index()
        
        # Codici presenti nel dataset
        codici_presenti = [c for c in self.codici if c in pivot_codice.columns]
        
        # Merge
        pivot = pivot_tipo.merge(
            pivot_codice[['Destinatario'] + codici_presenti],
            on='Destinatario', how='left'
        )
        
        # Assicura che esistano tutte le colonne necessarie
        tipi_unici = sorted(self.df['Tipo'].unique())
        for col in tipi_unici + self.codici:
            if col not in pivot.columns:
                pivot[col] = 0
        
        # Aggiungi operatore
        operatori = self._get_operatore_per_persona()
        pivot = pivot.merge(operatori, on='Destinatario', how='left')
        
        # Costruisci ordine colonne dinamicamente
        cols_ordinate = ['Destinatario', 'Operatore']
        for tipo in sorted(tipi_unici):
            cols_ordinate.append(tipo)
            codici_tipo = sorted([c for c in self.codici if c.startswith(tipo)])
            cols_ordinate.extend(codici_tipo)
        
        # Filtra solo colonne esistenti
        cols_ordinate = [c for c in cols_ordinate if c in pivot.columns]
        pivot = pivot[cols_ordinate]
        
        # Calcola totale
        pivot['Totale'] = pivot[[c for c in tipi_unici if c in pivot.columns]].sum(axis=1)
        pivot = pivot.sort_values('Destinatario')
        
        return pivot
    
    def conteggio_per_persona_tipo_mese(self) -> pd.DataFrame:
        """Conteggio azioni per persona, tipo e mese."""
        df_temp = self.df.copy()
        df_temp['Tipo-Mese'] = df_temp['Tipo'] + '_' + df_temp['Anno-Mese'].astype(str)
        
        pivot = pd.pivot_table(
            df_temp, index='Destinatario', columns='Tipo-Mese',
            aggfunc='size', fill_value=0
        ).reset_index()
        
        operatori = self._get_operatore_per_persona()
        pivot = pivot.merge(operatori, on='Destinatario', how='left')
        
        # Riordina colonne
        cols = pivot.columns.tolist()
        cols.remove('Operatore')
        cols.insert(1, 'Operatore')
        pivot = pivot[cols]
        
        return pivot.sort_values('Destinatario')
    
    def conteggio_totale_tipo(self) -> pd.DataFrame:
        """Conteggio totale per tipo."""
        conteggio_tipo = self.df.groupby('Tipo').size().reset_index(name='Conteggio')
        conteggio_codice = self.df.groupby('Codice').size().reset_index(name='Conteggio')
        conteggio_codice.columns = ['Tipo', 'Conteggio']
        
        risultato = pd.concat([conteggio_tipo, conteggio_codice], ignore_index=True)
        risultato = risultato.sort_values('Tipo')
        
        totale = pd.DataFrame([{'Tipo': 'TOTALE', 'Conteggio': len(self.df)}])
        risultato = pd.concat([risultato, totale], ignore_index=True)
        
        return risultato
    
    def conteggio_totale_tipo_mese(self) -> pd.DataFrame:
        """Conteggio totale per tipo e mese."""
        pivot = pd.pivot_table(
            self.df, index='Tipo', columns='Anno-Mese',
            aggfunc='size', fill_value=0
        ).reset_index()
        
        pivot.columns = [str(col) for col in pivot.columns]
        
        mesi_cols = [c for c in pivot.columns if c != 'Tipo']
        totali = pivot[mesi_cols].sum()
        totale_row = pd.DataFrame([['TOTALE'] + totali.tolist()], columns=pivot.columns)
        pivot = pd.concat([pivot, totale_row], ignore_index=True)
        
        return pivot
    
    def conteggio_per_operatore(self) -> pd.DataFrame:
        """Conteggio azioni per operatore."""
        pivot = pd.pivot_table(
            self.df, index='Operatore', columns='Codice',
            aggfunc='size', fill_value=0
        ).reset_index()
        
        for codice in self.codici:
            if codice not in pivot.columns:
                pivot[codice] = 0
        
        cols = ['Operatore'] + sorted([c for c in self.codici if c in pivot.columns])
        pivot = pivot[cols]
        pivot['Totale'] = pivot[[c for c in self.codici if c in pivot.columns]].sum(axis=1)
        
        return pivot.sort_values('Operatore')
    
    def conteggio_per_operatore_mese(self) -> pd.DataFrame:
        """Conteggio azioni per operatore, mese e tipo."""
        df_temp = self.df.copy()
        df_temp['Mese'] = df_temp['Anno-Mese'].astype(str)
        
        pivot = pd.pivot_table(
            df_temp, index=['Operatore', 'Mese'], columns='Codice',
            aggfunc='size', fill_value=0
        ).reset_index()
        
        for codice in self.codici:
            if codice not in pivot.columns:
                pivot[codice] = 0
        
        cols = ['Operatore', 'Mese'] + sorted([c for c in self.codici if c in pivot.columns])
        pivot = pivot[cols]
        pivot['Totale'] = pivot[[c for c in self.codici if c in pivot.columns]].sum(axis=1)
        
        return pivot.sort_values(['Operatore', 'Mese'])
    
    def calcolo_ricavi_per_mese(self) -> pd.DataFrame:
        """Calcola ricavi per tipologia e mese."""
        df_temp = self.df.copy()
        df_temp['Mese'] = df_temp['Anno-Mese'].astype(str)
        
        pivot = pd.pivot_table(
            df_temp, index='Mese', columns='Codice',
            aggfunc='size', fill_value=0
        ).reset_index()
        
        for codice in self.tariffe.keys():
            if codice not in pivot.columns:
                pivot[codice] = 0
        
        cols_codice = sorted(self.tariffe.keys())
        pivot = pivot[['Mese'] + cols_codice]
        
        # Calcola ricavi
        for codice, tariffa in self.tariffe.items():
            pivot[f'{codice}_ricavo'] = pivot[codice] * tariffa
        
        cols_ricavo = [f'{c}_ricavo' for c in cols_codice]
        pivot['Totale_Conteggio'] = pivot[cols_codice].sum(axis=1)
        pivot['Totale_Ricavo'] = pivot[cols_ricavo].sum(axis=1)
        
        return pivot.sort_values('Mese')
    
    def riepilogo_ricavi(self) -> pd.DataFrame:
        """Riepilogo ricavi aggregato per tipo."""
        conteggio = self.df.groupby('Codice').size().reset_index(name='Conteggio')
        
        conteggio['Tariffa (€)'] = conteggio['Codice'].map(self.tariffe)
        conteggio['Ricavo (€)'] = conteggio['Conteggio'] * conteggio['Tariffa (€)']
        
        totale = pd.DataFrame([{
            'Codice': 'TOTALE',
            'Tariffa (€)': None,
            'Conteggio': conteggio['Conteggio'].sum(),
            'Ricavo (€)': conteggio['Ricavo (€)'].sum()
        }])
        
        risultato = pd.concat([conteggio, totale], ignore_index=True)
        return risultato[['Codice', 'Tariffa (€)', 'Conteggio', 'Ricavo (€)']]
    
    def ricavi_totali(self) -> float:
        """Calcola il totale dei ricavi."""
        conteggio = self.df.groupby('Codice').size()
        ricavi = sum(
            conteggio.get(codice, 0) * tariffa 
            for codice, tariffa in self.tariffe.items()
        )
        return ricavi
    
    def top_persone(self, n: int = 10) -> pd.DataFrame:
        """Restituisce le prime N persone per numero di azioni."""
        return self.conteggio_per_persona_tipo().nlargest(n, 'Totale')
    
    def utenti_per_operatore(self) -> pd.DataFrame:
        """Conta il numero di utenti unici gestiti da ogni operatore."""
        utenti = self.df.groupby('Operatore')['Destinatario'].nunique().reset_index()
        utenti.columns = ['Operatore', 'Numero Utenti']
        utenti = utenti.sort_values('Numero Utenti', ascending=False)
        
        # Aggiungi riga totale
        totale = pd.DataFrame([{
            'Operatore': 'TOTALE',
            'Numero Utenti': self.df['Destinatario'].nunique()
        }])
        utenti = pd.concat([utenti, totale], ignore_index=True)
        
        return utenti
    
    def andamento_mensile(self) -> pd.DataFrame:
        """Restituisce l'andamento mensile delle azioni per tipo."""
        df_temp = self.df.copy()
        df_temp['Mese'] = df_temp['Anno-Mese'].astype(str)
        
        pivot = pd.pivot_table(
            df_temp, index='Mese', columns='Tipo',
            aggfunc='size', fill_value=0
        ).reset_index()
        
        # Aggiungi totale
        tipi = [c for c in pivot.columns if c != 'Mese']
        pivot['Totale'] = pivot[tipi].sum(axis=1)
        
        return pivot.sort_values('Mese')
    
    def ricavi_per_codice(self) -> pd.DataFrame:
        """Calcola i ricavi totali per ogni codice azione."""
        conteggio = self.df.groupby('Codice').size().reset_index(name='Conteggio')
        conteggio['Tariffa'] = conteggio['Codice'].map(self.tariffe)
        conteggio['Ricavo'] = conteggio['Conteggio'] * conteggio['Tariffa']
        return conteggio.sort_values('Ricavo', ascending=True)  # Ascending per grafico orizzontale
