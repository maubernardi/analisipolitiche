"""
Gestione della configurazione dell'applicazione.
Carica e salva le impostazioni da/su file config.json.
"""

import json
from pathlib import Path
from typing import Dict, List, Any


class ConfigManager:
    """Gestisce la configurazione dell'applicazione."""
    
    DEFAULT_CONFIG = {
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
    
    def __init__(self, config_path: str = None):
        """
        Inizializza il gestore configurazione.
        
        Args:
            config_path: Percorso del file config.json. Se None, usa il percorso di default.
        """
        if config_path is None:
            # Cerca config.json nella directory root del progetto
            self.config_path = Path(__file__).parent.parent / "config.json"
        else:
            self.config_path = Path(config_path)
        
        self._config = self._load_config()
    
    def _load_config(self) -> Dict[str, Any]:
        """Carica la configurazione dal file."""
        if self.config_path.exists():
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    # Merge con default per eventuali chiavi mancanti
                    return self._merge_with_defaults(config)
            except (json.JSONDecodeError, IOError) as e:
                print(f"Errore nel caricamento config: {e}. Uso valori di default.")
                return self.DEFAULT_CONFIG.copy()
        else:
            # Crea file config con valori di default
            self._save_config(self.DEFAULT_CONFIG)
            return self.DEFAULT_CONFIG.copy()
    
    def _merge_with_defaults(self, config: Dict) -> Dict:
        """Merge configurazione caricata con valori di default."""
        merged = self.DEFAULT_CONFIG.copy()
        
        if "tariffe" in config:
            merged["tariffe"] = config["tariffe"]
        if "filtri" in config:
            merged["filtri"] = config["filtri"]
        if "output" in config:
            merged["output"] = {**merged["output"], **config["output"]}
        
        return merged
    
    def _save_config(self, config: Dict[str, Any] = None) -> bool:
        """
        Salva la configurazione su file.
        
        Args:
            config: Configurazione da salvare. Se None, salva quella corrente.
        
        Returns:
            True se il salvataggio è riuscito, False altrimenti.
        """
        if config is None:
            config = self._config
        
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
            return True
        except IOError as e:
            print(f"Errore nel salvataggio config: {e}")
            return False
    
    def save(self) -> bool:
        """Salva la configurazione corrente su file."""
        return self._save_config()
    
    def reload(self) -> None:
        """Ricarica la configurazione dal file."""
        self._config = self._load_config()
    
    # --- Proprietà per accesso alla configurazione ---
    
    @property
    def tariffe(self) -> Dict[str, float]:
        """Restituisce le tariffe configurate."""
        return self._config.get("tariffe", {})
    
    @tariffe.setter
    def tariffe(self, value: Dict[str, float]) -> None:
        """Imposta le tariffe."""
        self._config["tariffe"] = value
    
    @property
    def escludi_eventi(self) -> List[str]:
        """Restituisce la lista degli eventi da escludere."""
        return self._config.get("filtri", {}).get("escludi_eventi", [])
    
    @escludi_eventi.setter
    def escludi_eventi(self, value: List[str]) -> None:
        """Imposta gli eventi da escludere."""
        if "filtri" not in self._config:
            self._config["filtri"] = {}
        self._config["filtri"]["escludi_eventi"] = value
    
    @property
    def prefisso_output(self) -> str:
        """Restituisce il prefisso per i file di output."""
        return self._config.get("output", {}).get("prefisso_nome", "export_analisi")
    
    @property
    def codici_validi(self) -> List[str]:
        """Restituisce la lista dei codici validi (chiavi delle tariffe)."""
        return list(self.tariffe.keys())
    
    # --- Metodi per modifica tariffe ---
    
    def aggiungi_tariffa(self, codice: str, valore: float) -> None:
        """Aggiunge o aggiorna una tariffa."""
        self._config["tariffe"][codice.upper()] = valore
    
    def rimuovi_tariffa(self, codice: str) -> bool:
        """
        Rimuove una tariffa.
        
        Returns:
            True se la tariffa è stata rimossa, False se non esisteva.
        """
        if codice.upper() in self._config["tariffe"]:
            del self._config["tariffe"][codice.upper()]
            return True
        return False
    
    def reset_tariffe(self) -> None:
        """Ripristina le tariffe ai valori di default."""
        self._config["tariffe"] = self.DEFAULT_CONFIG["tariffe"].copy()
    
    def reset_all(self) -> None:
        """Ripristina tutta la configurazione ai valori di default."""
        self._config = self.DEFAULT_CONFIG.copy()
    
    # --- Export/Import ---
    
    def to_dict(self) -> Dict[str, Any]:
        """Restituisce la configurazione come dizionario."""
        return self._config.copy()
    
    def from_dict(self, config: Dict[str, Any]) -> None:
        """Carica la configurazione da un dizionario."""
        self._config = self._merge_with_defaults(config)
