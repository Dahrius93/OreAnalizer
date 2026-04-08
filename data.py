# data.py — caricamento, pulizia e filtraggio del dataset SAP.
# Espone: parse_activity_type, load_and_clean, apply_filters, _get_date_range.

import re
import pandas as pd
import streamlit as st

from config import (
    COL_WBS, COL_REPARTO, COL_DATE, COL_PERSON, COL_HOURS,
    COL_DESC, COL_ACT_TYPE, COL_ACT_LABEL, COL_ACT_DETAIL,
    ACTIVITY_TYPES, ACTIVITY_ALIASES,
)


def parse_activity_type(desc: str) -> tuple:
    """
    Parse activity prefix from description.
    Handles: 'SW - Software QG1', 'SW-Software', 'SW Software', 'APOST - qualcosa'
    Returns (prefix, detail) e.g. ('SW', 'Software QG1')
    """
    if not isinstance(desc, str) or not desc.strip():
        return ("N/D", "")

    desc = desc.strip()

    # Known prefixes sorted longest-first to avoid partial matches.
    # Include also legacy aliases so they get recognized before normalization.
    known = sorted(set(ACTIVITY_TYPES.keys()) | set(ACTIVITY_ALIASES.keys()), key=len, reverse=True)

    for prefix in known:
        # Match: PREFIX followed by separator (- or space) or end of string
        pattern = rf"^{re.escape(prefix)}\s*[-–—]\s*(.*)$"
        m = re.match(pattern, desc, re.IGNORECASE)
        if m:
            return (prefix.upper(), m.group(1).strip())

        # Also match PREFIX + space + text (no dash)
        pattern2 = rf"^{re.escape(prefix)}\s+(.+)$"
        m2 = re.match(pattern2, desc, re.IGNORECASE)
        if m2:
            return (prefix.upper(), m2.group(1).strip())

        # Exact match (just the prefix, nothing else)
        if desc.upper() == prefix.upper():
            return (prefix.upper(), "")

    return ("N/D", desc)


def normalize_activity_type(act_type: str) -> str:
    """Normalizza prefissi legacy verso la chiave canonica (es. GES → HW-GES)."""
    return ACTIVITY_ALIASES.get(act_type.upper(), act_type)


@st.cache_data
def load_and_clean(uploaded_file) -> pd.DataFrame:
    """Load SAP Excel export, drop summary rows, parse activity types."""
    df = pd.read_excel(uploaded_file)

    df = df.dropna(subset=[COL_WBS])
    df = df.dropna(subset=[COL_REPARTO, COL_PERSON])

    df[COL_HOURS] = pd.to_numeric(df[COL_HOURS], errors="coerce").fillna(0)
    df[COL_DATE] = pd.to_datetime(df[COL_DATE], errors="coerce")
    df[COL_WBS] = df[COL_WBS].astype(str).str.strip()
    df[COL_REPARTO] = df[COL_REPARTO].astype(str).str.strip().str.upper()
    df[COL_PERSON] = df[COL_PERSON].astype(str).str.strip().str.title()
    df[COL_DESC] = df[COL_DESC].astype(str).str.strip() if COL_DESC in df.columns else ""

    # Parse activity type from description, then normalize legacy aliases
    parsed = df[COL_DESC].apply(parse_activity_type)
    df[COL_ACT_TYPE] = parsed.apply(lambda x: normalize_activity_type(x[0]))
    df[COL_ACT_DETAIL] = parsed.apply(lambda x: x[1])
    df[COL_ACT_LABEL] = df[COL_ACT_TYPE].map(
        lambda t: f"{t} — {ACTIVITY_TYPES.get(t, 'Non classificato')}"
    )

    return df


def apply_filters(df, reparti, wbs_list, persone, act_types, date_start=None, date_end=None):
    """Applica i filtri a cascata al DataFrame. Ogni lista vuota significa "tutti"."""
    filtered = df.copy()
    if reparti:
        filtered = filtered[filtered[COL_REPARTO].isin(reparti)]
    if wbs_list:
        filtered = filtered[filtered[COL_WBS].isin(wbs_list)]
    if persone:
        filtered = filtered[filtered[COL_PERSON].isin(persone)]
    if act_types:
        filtered = filtered[filtered[COL_ACT_TYPE].isin(act_types)]
    if date_start is not None:
        try:
            filtered = filtered[filtered[COL_DATE] >= pd.Timestamp(date_start)]
        except Exception:
            pass
    if date_end is not None:
        try:
            filtered = filtered[filtered[COL_DATE] <= pd.Timestamp(date_end)]
        except Exception:
            pass
    return filtered


def _get_date_range(df):
    """Restituisce il range di date del DataFrame come stringa 'dd/mm/yyyy → dd/mm/yyyy'."""
    dates = df[COL_DATE].dropna()
    if dates.empty:
        return "N/A"
    return f"{dates.min().strftime('%d/%m/%Y')} → {dates.max().strftime('%d/%m/%Y')}"
