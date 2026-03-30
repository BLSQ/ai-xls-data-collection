"""Stateless text utilities for canonical name generation and cell reference parsing.

These functions are used across multiple classes within the schema_validation
pipeline.  They have no side effects and no dependencies beyond the standard
library and openpyxl.
"""

from __future__ import annotations

import datetime
import re
import unicodedata

from openpyxl.utils import column_index_from_string, get_column_letter

# Ligature expansion table (French text commonly uses œ/æ).
_LIGATURES = str.maketrans({"œ": "oe", "Œ": "OE", "æ": "ae", "Æ": "AE"})


def to_canonical(text: str | None) -> str:
    """Convert free text to a deterministic snake_case canonical name.

    Steps: strip → NFKD normalise → drop combining marks → expand
    ligatures → replace non-alphanumeric with ``_`` → collapse → lower.

    Args:
        text: The raw text to canonicalise.

    Returns:
        A lowercase, ASCII-only, underscore-separated identifier.
        Returns ``"unnamed"`` for empty/None input.
    """
    if not text:
        return "unnamed"
    text = str(text).strip()
    nfkd = unicodedata.normalize("NFKD", text)
    text = "".join(c for c in nfkd if not unicodedata.combining(c))
    text = text.translate(_LIGATURES)
    text = re.sub(r"[^a-zA-Z0-9]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_").lower()
    return text or "unnamed"


def to_column_canonical(text: str | None) -> str:
    """Generate a canonical name for a column header.

    Strips parenthetical descriptions first so that
    ``"Niveau central (Directions, Programmes, …)"`` becomes
    ``"niveau_central"`` rather than an unwieldy long slug.

    Args:
        text: The raw column header text.

    Returns:
        A canonical column name.
    """
    text = re.sub(r"\s*\([^)]*\)", "", str(text or "")).strip()
    return to_canonical(text)


def column_index(letter: str) -> int:
    """Convert a column letter (e.g. ``'C'``) to a 1-based integer index.

    Args:
        letter: One or more uppercase letters identifying the column.

    Returns:
        The 1-based column index.
    """
    return column_index_from_string(letter)


def column_letter(index: int) -> str:
    """Convert a 1-based column index to a letter (e.g. ``3`` → ``'C'``).

    Args:
        index: The 1-based column index.

    Returns:
        The uppercase column letter(s).
    """
    return get_column_letter(index)


def cell_reference_to_column(cell_ref: str) -> str:
    """Extract the column letter(s) from a cell reference like ``'C7'`` or ``'AB12'``.

    Args:
        cell_ref: An A1-notation cell reference.

    Returns:
        The alphabetical column part.
    """
    return "".join(c for c in cell_ref if c.isalpha())


def cell_reference_to_row(cell_ref: str) -> int:
    """Extract the row number from a cell reference like ``'C7'``.

    Args:
        cell_ref: An A1-notation cell reference.

    Returns:
        The numeric row part.
    """
    return int("".join(c for c in cell_ref if c.isdigit()))


def format_example(value) -> str:
    """Format a cell value as a human-readable example string.

    Dates become ``YYYY-MM-DD``, integer-valued floats drop the decimal,
    and everything else is converted to a stripped string.

    Args:
        value: A raw cell value (string, number, datetime, or None).

    Returns:
        A formatted example string (empty string for None).
    """
    if value is None:
        return ""
    if isinstance(value, datetime.datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    return str(value).strip()


def infer_type(value) -> str:
    """Infer the schema type (``string``, ``number``, ``date``) from a Python value.

    Args:
        value: A raw cell value.

    Returns:
        One of ``"string"``, ``"number"``, or ``"date"``.
    """
    if value is None:
        return "string"
    if isinstance(value, datetime.datetime):
        return "date"
    if isinstance(value, (int, float)):
        return "number"
    text = str(value).strip()
    if re.match(r"^\d{4}-\d{2}-\d{2}", text):
        return "date"
    try:
        float(text.replace(",", "").replace(" ", ""))
        return "number"
    except (ValueError, AttributeError):
        pass
    return "string"
