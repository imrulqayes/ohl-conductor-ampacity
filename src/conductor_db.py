"""
Conductor Database Module.

Reads conductor electrical and thermal properties from the Excel database file
located in the 'data/' folder.

Expected Excel columns (names are matched case-insensitively):
  - Conductor Name
  - Diameter (mm)
  - R75 (ohm/m)   — AC resistance at 75 °C
  - R25 (ohm/m)   — AC resistance at 25 °C
  - Conductor heat capacity (J/m·°C)

The database can be extended with more conductors at any time; no code changes are needed.

Author       : Imrul Qayes
Email        : imrul27@gmail.com
Date Created : 03/2025
Last Modified: 04/2026
Version      : 1.0.0
AI Assistant : Developed with the assistance of Claude (Anthropic) — https://claude.ai
"""

# Licensed under the MIT License. See LICENSE file in the project root for details.

import re
import warnings
from pathlib import Path

import pandas as pd

DATA_DIR = Path(__file__).parent.parent / "data"

# Flexible column-name matching patterns (lowercase keywords → config key)
_COLUMN_PATTERNS = {
    "name":          ["conductor name", "conductor", "name", "designation", "type", "code"],
    "diameter_mm":   ["diameter", "od", "overall diameter", "outer diameter", "outside diameter", "o.d"],
    "temp_low":      ["temp_low", "t_low", "temp low", "low temp", "temperature low"],
    "temp_high":     ["temp_high", "t_high", "temp high", "high temp", "temperature high"],
    "r_low":         ["r25", "r20", "r_low", "r low", "resistance low", "ac resistance low",
                      "rac low", "r @ 25", "r @ 20"],
    "r_high":        ["r75", "r_high", "r high", "resistance high", "ac resistance high",
                      "rac high", "r @ 75"],
    "heat_capacity": ["conductor heat capacity", "heat capacity", "mcp", "thermal capacity",
                      "heat cap", "specific heat capacity"],
    "max_temp":      ["max allowable temp", "max temp", "maximum allowable temp",
                      "maximum temperature", "max allowable temperature", "t_max",
                      "max operating temp", "max continuous temp"],
}


def _normalize_col(text: str) -> str:
    """Lowercase, strip, and remove unit annotations like (mm), (Ω/m), etc."""
    text = str(text).lower().strip()
    text = re.sub(r"\(.*?\)", "", text)   # remove (…)
    text = re.sub(r"\[.*?\]", "", text)   # remove […]
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _find_column(df_columns, patterns: list) -> str | None:
    """Return the first DataFrame column whose normalised name matches any pattern."""
    normalised = {_normalize_col(c): c for c in df_columns}
    for pattern in patterns:
        if pattern in normalised:
            return normalised[pattern]
    return None


def _find_excel_file() -> Path:
    """Locate the conductor Excel database in the data directory."""
    files = sorted(DATA_DIR.glob("*.xlsx")) + sorted(DATA_DIR.glob("*.xls"))
    if not files:
        raise FileNotFoundError(
            f"No Excel file found in '{DATA_DIR}'. "
            "Please place your conductor database there."
        )
    if len(files) > 1:
        # Prefer a file whose name contains 'conductor'
        preferred = [f for f in files if "conductor" in f.name.lower()]
        if preferred:
            return preferred[0]
    return files[0]


def _read_conductor_sheet(excel_file: Path) -> pd.DataFrame:
    """
    Read the conductor data sheet from an Excel file.

    Looks for a sheet named 'data' (case-insensitive) first.
    If not found, tries each sheet in order and picks the first one
    that contains a recognisable conductor name column.
    """
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        xl = pd.ExcelFile(excel_file)

        # 1. Prefer a sheet literally named 'data'
        for sheet in xl.sheet_names:
            if sheet.strip().lower() == "data":
                return xl.parse(sheet)

        # 2. Fall back: find the first sheet with a name column
        for sheet in xl.sheet_names:
            df = xl.parse(sheet)
            if _find_column(df.columns, _COLUMN_PATTERNS["name"]) is not None:
                return df

        # 3. Last resort: first sheet
        return xl.parse(0)


def list_conductors() -> list[str]:
    """Return a sorted list of all conductor names in the database."""
    excel_file = _find_excel_file()
    df = _read_conductor_sheet(excel_file)

    name_col = _find_column(df.columns, _COLUMN_PATTERNS["name"])
    diam_col = _find_column(df.columns, _COLUMN_PATTERNS["diameter_mm"])
    if name_col is None:
        return []

    # Keep only rows that have a valid numeric diameter — this filters out
    # section headers (e.g. "Temperature Coefficients…") and metadata rows.
    if diam_col:
        valid = pd.to_numeric(df[diam_col], errors="coerce").notna()
        df = df[valid]

    names = df[name_col].dropna().astype(str).str.strip().tolist()
    names = [n for n in names if n and not n.lower().startswith("nan")]
    return sorted(set(names))


def get_conductor_data(name: str) -> dict:
    """
    Look up conductor properties by name from the Excel database.

    The match is case-insensitive but requires an exact spelling.

    Parameters:
        name : Conductor name (e.g., 'Grosbeak', 'drake', 'HAWK')

    Returns:
        dict with keys:
            name           – matched conductor name as stored in database
            diameter_mm    – outside diameter (mm)
            temp_low       – low reference temperature for resistance (°C)
            temp_high      – high reference temperature for resistance (°C)
            r_low          – AC resistance at low reference temperature (Ω/m)
            r_high         – AC resistance at high reference temperature (Ω/m)
            heat_capacity  – conductor heat capacity (J/m·°C); None if not found

    Raises:
        FileNotFoundError : If no Excel file is found in the data directory
        ValueError        : If the conductor name is not found or required columns are missing
    """
    excel_file = _find_excel_file()
    df = _read_conductor_sheet(excel_file)

    # ---- Find name column ------------------------------------------------
    name_col = _find_column(df.columns, _COLUMN_PATTERNS["name"])
    if name_col is None:
        raise ValueError(
            f"Could not find a 'Conductor Name' column in '{excel_file.name}'.\n"
            f"Available columns: {list(df.columns)}"
        )

    # ---- Case-insensitive exact name match --------------------------------
    mask = df[name_col].astype(str).str.strip().str.lower() == name.strip().lower()
    matches = df[mask].dropna(subset=[name_col])

    if matches.empty:
        available = [n for n in df[name_col].dropna().astype(str).tolist()
                     if not n.lower().startswith("temperature")]
        raise ValueError(
            f"Conductor '{name}' not found in the database.\n"
            f"Available conductors: {available}\n"
            f"Note: spelling must match exactly (case-insensitive)."
        )

    if len(matches) > 1:
        warnings.warn(f"Multiple entries found for '{name}'; using the first match.")

    row = matches.iloc[0]

    # ---- Extract required values -----------------------------------------
    def _get(key: str, required: bool = True):
        col = _find_column(df.columns, _COLUMN_PATTERNS[key])
        if col is None:
            if required:
                raise ValueError(
                    f"Required column for '{key}' not found in '{excel_file.name}'.\n"
                    f"Available columns: {list(df.columns)}"
                )
            return None
        try:
            return float(row[col])
        except (ValueError, TypeError):
            if required:
                raise ValueError(
                    f"Non-numeric value in column '{col}' for conductor '{name}': {row[col]}"
                )
            return None

    diameter_mm   = _get("diameter_mm")
    r_low_raw     = _get("r_low")
    r_high_raw    = _get("r_high")
    heat_capacity = _get("heat_capacity", required=False)
    max_temp      = _get("max_temp",      required=False)

    # Determine resistance units: ohm/m vs ohm/km
    # Typical ACSR/ACSS values: ohm/m ~ 5e-5 to 5e-4; ohm/km ~ 0.05 to 0.5
    if r_low_raw > 0.01:
        r_low  = r_low_raw  / 1000.0
        r_high = r_high_raw / 1000.0
    else:
        r_low  = r_low_raw
        r_high = r_high_raw

    # Reference temperatures — try to read from database; fall back to IEEE 738 table defaults
    temp_low  = _get("temp_low",  required=False)
    temp_high = _get("temp_high", required=False)

    # Infer reference temperatures from column name if not explicit columns present
    if temp_low is None:
        r_low_col = _find_column(df.columns, _COLUMN_PATTERNS["r_low"])
        if r_low_col:
            digits = re.findall(r"\d+", r_low_col)
            temp_low = float(digits[0]) if digits else 25.0
        else:
            temp_low = 25.0

    if temp_high is None:
        r_high_col = _find_column(df.columns, _COLUMN_PATTERNS["r_high"])
        if r_high_col:
            digits = re.findall(r"\d+", r_high_col)
            temp_high = float(digits[0]) if digits else 75.0
        else:
            temp_high = 75.0

    return {
        "name":          str(row[name_col]).strip(),
        "diameter_mm":   diameter_mm,
        "temp_low":      temp_low,
        "temp_high":     temp_high,
        "r_low":         r_low,
        "r_high":        r_high,
        "heat_capacity": heat_capacity,
        "max_temp":      max_temp,
    }
