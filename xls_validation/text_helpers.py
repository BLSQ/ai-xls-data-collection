"""Stateless text normalization and matching utilities.

Provides configurable text normalization (whitespace, accents, case),
deterministic canonical name generation, and fuzzy text matching used
by the section validators.

Pure functions — no side effects, no OpenHEXA dependency.
"""

from __future__ import annotations

import re
import unicodedata

# Ligature expansion table (French text commonly uses œ/æ).
_LIGATURES = str.maketrans({"œ": "oe", "Œ": "OE", "æ": "ae", "Æ": "AE"})


def normalize_text(text: str, config: dict) -> str:
    """Normalize a cell value for comparison according to validation_config.

    Applies whitespace collapsing, trailing-colon stripping, case folding,
    and accent removal based on the flags in *config*.

    Args:
        text: The raw text to normalize (may be None).
        config: A dict with optional boolean keys ``strip_whitespace``,
            ``strip_trailing_colon``, ``case_sensitive``, ``normalize_accents``.

    Returns:
        The normalized text string (empty string for None input).
    """
    if text is None:
        return ""
    normalized = str(text)
    if config.get("strip_whitespace", True):
        normalized = normalized.strip()
        normalized = re.sub(r"\s+", " ", normalized)
    if config.get("strip_trailing_colon", True):
        normalized = normalized.rstrip(":")
    if not config.get("case_sensitive"):
        normalized = normalized.lower()
    if config.get("normalize_accents", True):
        normalized = _remove_accents(normalized)
    return normalized


def _remove_accents(text: str) -> str:
    """Strip combining diacritical marks from NFKD-decomposed text."""
    nfkd = unicodedata.normalize("NFKD", text)
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def to_canonical(text: str) -> str:
    """Convert free text to a deterministic snake_case canonical name.

    Must match the logic in ``schema_validation/text_helpers.to_canonical``
    so that names generated during schema creation can be compared at
    validation time.

    Args:
        text: The raw text to canonicalize.

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


def text_matches(
    actual: str,
    expected: str,
    variants: list[str],
    config: dict,
) -> str:
    """Check whether *actual* matches *expected* or any known variant.

    Args:
        actual: The text found in the cell.
        expected: The canonical expected text from the schema.
        variants: A list of known alternative spellings.
        config: Normalization config forwarded to :func:`normalize_text`.

    Returns:
        ``"exact"`` if the normalized texts match, ``"variant"`` if a
        variant matched, or ``""`` (empty string) if no match.
    """
    normalized_actual = normalize_text(actual, config)
    if normalized_actual == normalize_text(expected, config):
        return "exact"
    for variant in variants:
        if normalized_actual == normalize_text(variant, config):
            return "variant"
    return ""
