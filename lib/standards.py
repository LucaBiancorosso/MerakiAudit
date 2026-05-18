from __future__ import annotations

from pathlib import Path
from typing import Any

import openpyxl


# Fields that are identity keys, not audited config fields
_KEY_FIELDS = {"org_id", "ssid_name"}

# Sentinel for a blank standard cell
NOT_DEFINED = "NOT_DEFINED"


def _cell_to_str(value: Any) -> str:
    """
    Robustly convert an Excel cell value to a clean string, handling all
    the ways Excel and openpyxl can mangle data:

      - float with no fractional part (e.g. 123456.0 -> '123456')
        Happens when a numeric org_id column is not formatted as Text in Excel.
      - bool stored as Python bool (e.g. True -> 'True', False -> 'False')
        Happens when the user types TRUE/FALSE and Excel stores as boolean.
      - scientific notation floats (e.g. 1.23457e+11)
        Happens for large org IDs that Excel auto-converts to scientific notation.
      - plain int / str: converted with str() as usual.
    """
    if value is None:
        return ""
    if isinstance(value, bool):
        # Must come before int check — bool is a subclass of int in Python
        return str(value)           # True -> 'True', False -> 'False'
    if isinstance(value, float):
        if value == int(value):
            return str(int(value))  # 123456.0 -> '123456'
        return str(value)
    return str(value).strip()


def _normalise(value: Any) -> str:
    """Return a clean lowercase string for comparison."""
    return _cell_to_str(value).strip().lower()


def load_standards(path: Path) -> dict[tuple[str, str], dict[str, Any]]:
    """
    Read the Standards sheet and return a lookup keyed by (org_id, ssid_name).

    Each value is a dict of {field: expected_value}.
    Blank cells are stored as NOT_DEFINED sentinel so callers can distinguish
    'not configured' from 'empty string expected'.
    """
    wb = openpyxl.load_workbook(path, data_only=True)

    if "Standards" not in wb.sheetnames:
        raise ValueError(f"'Standards' sheet not found in {path}")

    ws = wb["Standards"]
    rows = list(ws.iter_rows(values_only=True))

    if not rows:
        raise ValueError("Standards sheet is empty")

    headers = [_cell_to_str(h).strip() for h in rows[0]]

    standards: dict[tuple[str, str], dict[str, Any]] = {}

    for row in rows[1:]:
        row_dict = dict(zip(headers, row))

        # org_id: normalise robustly (handles int, float, str from Excel)
        org_id = _normalise(row_dict.get("org_id"))
        ssid_name = _cell_to_str(row_dict.get("ssid_name")).strip()

        # Skip blank rows or the hint row in the template
        if not org_id or not ssid_name or ssid_name.lower() in ("ssid_name", "← required"):
            continue

        # Key uses lowercase ssid_name so lookup is case-insensitive
        key = (org_id, ssid_name.lower())
        fields: dict[str, Any] = {}

        for header, value in row_dict.items():
            if header in _KEY_FIELDS or not header:
                continue
            cell_str = _cell_to_str(value).strip()
            fields[header] = NOT_DEFINED if cell_str == "" else cell_str

        standards[key] = fields

    return standards


def load_ignored_fields(path: Path) -> set[str]:
    """
    Read the Ignore sheet and return a set of field names to skip globally.
    Returns empty set if sheet is missing or empty.
    """
    wb = openpyxl.load_workbook(path, data_only=True)

    if "Ignore" not in wb.sheetnames:
        return set()

    ws = wb["Ignore"]
    ignored: set[str] = set()

    for row in ws.iter_rows(min_row=2, values_only=True):
        cell_str = _cell_to_str(row[0]).strip()
        if cell_str:
            ignored.add(cell_str)

    return ignored


def compare_field(
    field: str,
    expected: str,
    actual: Any,
    ignored_fields: set[str],
) -> dict[str, str]:
    """
    Compare a single field and return a result dict with keys:
        field, expected, actual, result
    """
    if field in ignored_fields:
        return {
            "field":    field,
            "expected": expected,
            "actual":   _cell_to_str(actual),
            "result":   "IGNORED",
        }

    if expected == NOT_DEFINED:
        return {
            "field":    field,
            "expected": "",
            "actual":   _cell_to_str(actual),
            "result":   NOT_DEFINED,
        }

    actual_str   = _normalise(actual)
    expected_str = _normalise(expected)

    # Special handling for radiusHosts: compare as unordered comma-separated sets
    if field == "radiusHosts":
        actual_set   = {h.strip() for h in actual_str.split(",") if h.strip()}
        expected_set = {h.strip() for h in expected_str.split(",") if h.strip()}
        result = "PASS" if actual_set == expected_set else "FAIL"
    else:
        result = "PASS" if actual_str == expected_str else "FAIL"

    return {
        "field":    field,
        "expected": expected,
        "actual":   _cell_to_str(actual),
        "result":   result,
    }
