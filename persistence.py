# persistence.py — salvataggio/ripristino del file SAP e dei filtri su disco locale.
# Espone: save_uploaded_file, load_saved_file, save_filters, load_filters, clear_saved_data.

import json
import os
from pathlib import Path

_DATA_DIR = Path.home() / ".oreanalizer"
_FILE_PATH = _DATA_DIR / "last_upload.xlsx"
_FILTERS_PATH = _DATA_DIR / "last_filters.json"


def _ensure_dir():
    _DATA_DIR.mkdir(parents=True, exist_ok=True)


def save_uploaded_file(uploaded_file) -> None:
    """Salva su disco il file caricato via st.file_uploader."""
    _ensure_dir()
    uploaded_file.seek(0)
    _FILE_PATH.write_bytes(uploaded_file.read())
    uploaded_file.seek(0)


def load_saved_file():
    """Restituisce i byte del file salvato, o None se non esiste."""
    if _FILE_PATH.exists():
        return _FILE_PATH.read_bytes()
    return None


def saved_file_name() -> str | None:
    """Restituisce il nome del file salvato (solo il filename), o None."""
    if _FILE_PATH.exists():
        return _FILE_PATH.name
    return None


def save_filters(filters: dict) -> None:
    """Salva i filtri attivi come JSON."""
    _ensure_dir()
    _FILTERS_PATH.write_text(json.dumps(filters, default=str, ensure_ascii=False), encoding="utf-8")


def load_filters() -> dict:
    """Carica i filtri salvati. Restituisce {} se non esistono."""
    if _FILTERS_PATH.exists():
        try:
            return json.loads(_FILTERS_PATH.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def clear_saved_data() -> None:
    """Elimina file e filtri salvati."""
    if _FILE_PATH.exists():
        _FILE_PATH.unlink()
    if _FILTERS_PATH.exists():
        _FILTERS_PATH.unlink()
