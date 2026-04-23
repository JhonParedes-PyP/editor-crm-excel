# utils.py - Funciones reutilizables para procesamiento de datos

import json
import os
import pandas as pd
from config import CRM_COLUMNS, NUMERIC_COLS, PROFILES_FILE, NONE_LABEL

def load_profiles() -> dict:
    if os.path.exists(PROFILES_FILE):
        try:
            with open(PROFILES_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return {}
    return {}

def save_profiles(profiles: dict) -> None:
    with open(PROFILES_FILE, "w", encoding="utf-8") as f:
        json.dump(profiles, f, ensure_ascii=False, indent=2)

def auto_map(source_cols: list) -> dict:
    """Mapeo automático por nombre idéntico (insensible a mayúsculas/espacios)."""
    lookup = {c.upper().strip(): c for c in source_cols}
    return {crm: lookup[crm] for crm in CRM_COLUMNS if crm in lookup}

def is_non_numeric(val) -> bool:
    """True si el valor no puede convertirse a número."""
    if pd.isna(val) or str(val).strip() in ("", "nan", "None"):
        return False
    try:
        float(str(val).replace(",", ".").replace(" ", ""))
        return False
    except ValueError:
        return True

def apply_mapping(df: pd.DataFrame, mapping: dict) -> pd.DataFrame:
    """Retorna un DataFrame con exactamente las 26 columnas CRM en orden fijo."""
    n = len(df)
    data = {}
    for col in CRM_COLUMNS:
        src = mapping.get(col, "")
        data[col] = df[src].tolist() if src and src in df.columns else [""] * n
    return pd.DataFrame(data)

def get_alerts(df: pd.DataFrame) -> list:
    """Valida DOC_DNI_RUC vacío y montos no numéricos. Retorna lista de (nivel, mensaje)."""
    alerts = []

    empty_mask = df["DOC_DNI_RUC"].isna() | df["DOC_DNI_RUC"].astype(str).str.strip().eq("")
    n_empty = int(empty_mask.sum())
    if n_empty:
        alerts.append(("warning", f"**DOC_DNI_RUC**: {n_empty} registro(s) vacíos"))

    for col in NUMERIC_COLS:
        bad = df[col].apply(is_non_numeric)
        if bad.any():
            examples = df.loc[bad, col].astype(str).unique()[:3].tolist()
            alerts.append(("error", f"**{col}**: valores no numéricos detectados → {examples}"))

    return alerts
