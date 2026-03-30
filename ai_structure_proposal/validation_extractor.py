"""Deterministic data-validation extraction from Excel workbooks.

Reads openpyxl data-validation objects and resolves dropdown list values
where possible (cell ranges, named ranges, inline lists).  INDIRECT
formulas are flagged but not resolved here — the schema-generation
pipeline handles them with full cell-value context.
"""


class ValidationExtractor:
    """Extracts and resolves Excel data-validation rules deterministically.

    Works with the ``data_only=False`` workbook so that formulas and
    named-range definitions are available.

    Args:
        worksheet_full: Worksheet opened with ``data_only=False``.
        workbook_full:  Workbook opened with ``data_only=False``.
    """

    def __init__(self, worksheet_full, workbook_full):
        """Store references to the full (formula-mode) worksheet and workbook.

        Args:
            worksheet_full: The target sheet (data_only=False).
            workbook_full:  The parent workbook (data_only=False).
        """
        self._worksheet = worksheet_full
        self._workbook = workbook_full

    def extract(self) -> list[dict]:
        """Extract all list-type data validations with resolved values.

        Returns:
            A list of dicts, each describing one dropdown validation::

                {
                    "cells":            "B17:B100",
                    "type":             "list",
                    "formula":          "=Categories!$A$2:$A$20",
                    "resolved_values":  ["Cat A", "Cat B", ...] or None,
                    "allow_blank":      True
                }
        """
        results: list[dict] = []

        if (
            not hasattr(self._worksheet, "data_validations")
            or self._worksheet.data_validations is None
        ):
            return results

        for data_validation in self._worksheet.data_validations.dataValidation:
            if data_validation.type != "list":
                continue

            formula = str(data_validation.formula1 or "")
            resolved_values = self._try_resolve_formula(formula)

            results.append(
                {
                    "cells": str(data_validation.sqref),
                    "type": "list",
                    "formula": formula,
                    "resolved_values": resolved_values if resolved_values else None,
                    "allow_blank": bool(data_validation.allow_blank),
                }
            )

        return results

    # ------------------------------------------------------------------
    # Formula resolution helpers
    # ------------------------------------------------------------------

    def _try_resolve_formula(self, formula_string: str) -> list[str] | None:
        """Best-effort resolution of a data-validation formula.

        Handles cell-range references, named ranges, and inline
        semicolon/comma-separated lists.  Returns ``None`` for INDIRECT
        formulas (which require cell-value context to resolve).

        Args:
            formula_string: The raw formula from the data validation.

        Returns:
            A list of resolved string values, or ``None`` if resolution
            is not possible.
        """
        formula_string = formula_string.strip('"')
        if not formula_string:
            return None

        # INDIRECT — skip (needs cell values to resolve).
        if formula_string.upper().startswith("INDIRECT("):
            return None

        # Cell range reference (contains $, :, or !).
        if "$" in formula_string or ":" in formula_string or "!" in formula_string:
            return self._resolve_range(formula_string)

        # Named range.
        if hasattr(self._workbook, "defined_names"):
            definition = self._workbook.defined_names.get(formula_string)
            if definition:
                return self._resolve_named_range(definition)

        # Inline list (semicolons or commas).
        if ";" in formula_string:
            return [v.strip() for v in formula_string.split(";") if v.strip()]
        if "," in formula_string:
            return [v.strip() for v in formula_string.split(",") if v.strip()]

        return [formula_string] if formula_string else None

    def _resolve_range(self, formula_string: str) -> list[str] | None:
        """Resolve a cell-range reference (e.g. ``Sheet1!$A$2:$A$20``) to values.

        Args:
            formula_string: A cell-range formula, optionally prefixed with
                a sheet name.

        Returns:
            A list of non-empty string values from the range, or ``None``
            on failure.
        """
        try:
            if "!" in formula_string:
                sheet_part, range_part = formula_string.rsplit("!", 1)
                target_sheet_name = sheet_part.strip("'\"")
                if target_sheet_name not in self._workbook.sheetnames:
                    return None
                source_sheet = self._workbook[target_sheet_name]
            else:
                source_sheet = self._worksheet
                range_part = formula_string

            range_part = range_part.replace("$", "")
            values: list[str] = []
            for row_or_cell in source_sheet[range_part]:
                cells = row_or_cell if isinstance(row_or_cell, tuple) else (row_or_cell,)
                for cell in cells:
                    if cell.value is not None and str(cell.value).strip():
                        values.append(str(cell.value).strip())
            return values or None
        except Exception:
            return None

    def _resolve_named_range(self, definition) -> list[str] | None:
        """Resolve a named-range definition to its cell values.

        Args:
            definition: An openpyxl ``DefinedName`` object.

        Returns:
            A list of non-empty string values, or ``None`` on failure.
        """
        values: list[str] = []
        try:
            for sheet_title, cell_range in definition.destinations:
                if sheet_title in self._workbook.sheetnames:
                    source_sheet = self._workbook[sheet_title]
                    for row_or_cell in source_sheet[cell_range]:
                        cells = (
                            row_or_cell
                            if isinstance(row_or_cell, tuple)
                            else (row_or_cell,)
                        )
                        for cell in cells:
                            if cell.value is not None and str(cell.value).strip():
                                values.append(str(cell.value).strip())
        except Exception:
            return None
        return values or None
