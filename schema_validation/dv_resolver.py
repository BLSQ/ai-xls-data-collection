"""Extract and resolve Excel data-validation rules (dropdown lists).

This module provides :class:`DataValidationResolver`, which walks through the
``data_validations`` attached to an openpyxl worksheet, resolves list formulas
(including INDIRECT references and named ranges), and exposes the results as
two maps consumed by the section processors and the schema builder:

* ``data_validation_map`` -- cell reference  ->  validation info
* ``conditional_map``     -- child column    ->  parent-dependent choices
"""

from __future__ import annotations

import re

from text_helpers import (
    cell_reference_to_column,
    column_index,
    column_letter,
)


class DataValidationResolver:
    """Extracts and resolves Excel data-validation rules from a worksheet.

    Provides the data_validation_map and conditional_map needed by the
    section processors and schema builder.
    """

    def __init__(self, worksheet_full, workbook_full) -> None:
        """Initialise the resolver with an openpyxl worksheet and workbook.

        Args:
            worksheet_full: The openpyxl ``Worksheet`` that owns the data
                validations.
            workbook_full: The parent ``Workbook`` (used to follow cross-sheet
                references and named ranges).
        """
        self._worksheet = worksheet_full
        self._workbook = workbook_full
        self._data_validation_map: dict = {}
        self._conditional_map: dict = {}

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def extract(self) -> dict:
        """Build ``data_validation_map`` from the worksheet's data validations.

        Each *list*-type validation is resolved into either a concrete list of
        values or an ``indirect`` marker that records which parent column the
        dropdown depends on.

        Returns:
            The completed ``data_validation_map`` (also stored internally).
        """
        self._data_validation_map = {}

        if (
            not hasattr(self._worksheet, "data_validations")
            or self._worksheet.data_validations is None
        ):
            return self._data_validation_map

        for data_validation in self._worksheet.data_validations.dataValidation:
            if data_validation.type != "list":
                continue

            formula = data_validation.formula1
            if not formula:
                continue

            formula_str = str(formula).strip('"')
            is_indirect = formula_str.upper().startswith("INDIRECT(")
            choices = self._resolve_dv_formula(formula_str)

            if is_indirect and not choices:
                parent_col = self._parse_indirect_parent_column(formula_str)
                if parent_col:
                    dv_info = {
                        "type": "indirect",
                        "depends_on_column": parent_col,
                        "formula": formula_str,
                        "allow_blank": bool(data_validation.allow_blank),
                    }
                else:
                    continue
            elif choices:
                dv_info = {
                    "type": "list",
                    "values": choices,
                    "allow_blank": bool(data_validation.allow_blank),
                    "formula": formula_str,
                }
            else:
                continue

            for cell_range in data_validation.sqref.ranges:
                for row in range(cell_range.min_row, cell_range.max_row + 1):
                    for col in range(cell_range.min_col, cell_range.max_col + 1):
                        ref = f"{column_letter(col)}{row}"
                        self._data_validation_map[ref] = dv_info

        return self._data_validation_map

    def resolve_indirect_dependencies(self) -> dict:
        """Build ``conditional_map`` for INDIRECT-based cascading dropdowns.

        For every column that uses an ``INDIRECT`` formula, this method finds
        the possible parent values and resolves named-range variants for each
        one, producing a mapping of child-column -> { parent column, values
        keyed by parent selection }.

        Returns:
            The completed ``conditional_map`` (also stored internally).
        """
        self._conditional_map = {}
        indirect_columns: dict = {}

        for ref, info in self._data_validation_map.items():
            if info.get("type") != "indirect":
                continue
            child_col = cell_reference_to_column(ref)
            if child_col not in indirect_columns:
                indirect_columns[child_col] = info

        for child_col, info in indirect_columns.items():
            parent_col = info["depends_on_column"]
            parent_values = self._find_column_values(parent_col)
            if not parent_values:
                continue

            values_by_parent: dict = {}
            for parent_value in parent_values:
                resolved = self._resolve_named_range_variants(parent_value)
                if resolved:
                    values_by_parent[parent_value] = resolved

            if values_by_parent:
                self._conditional_map[child_col] = {
                    "depends_on_column": parent_col,
                    "values_by_parent": values_by_parent,
                }

        return self._conditional_map

    def merge_proposal_validations(self, proposal: dict) -> None:
        """Merge AI-proposed data-validation entries into the map.

        Entries from ``proposal["data_validations_resolved"]`` that reference
        cells not yet present in the map are added.  Existing entries are
        never overwritten.

        Args:
            proposal: The structure-proposal dictionary, expected to contain
                an optional ``data_validations_resolved`` list.
        """
        resolved_list = proposal.get("data_validations_resolved", [])

        for entry in resolved_list:
            values = entry.get("resolved_values")
            if not values:
                continue

            sqref = entry.get("cells", "")
            dv_info = {
                "type": entry.get("type", "list"),
                "values": values,
                "allow_blank": entry.get("allow_blank", True),
                "formula": entry.get("formula", ""),
            }

            for cell_ref in self._expand_sqref(sqref):
                if cell_ref not in self._data_validation_map:
                    self._data_validation_map[cell_ref] = dv_info

    # ------------------------------------------------------------------
    # Properties
    # ------------------------------------------------------------------

    @property
    def data_validation_map(self) -> dict:
        """Return the cell-reference -> validation-info mapping."""
        return self._data_validation_map

    @property
    def conditional_map(self) -> dict:
        """Return the child-column -> parent-dependent-choices mapping."""
        return self._conditional_map

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _parse_indirect_parent_column(formula_str: str) -> str | None:
        """Extract the parent column letter from an INDIRECT formula.

        Args:
            formula_str: The raw formula string, e.g. ``INDIRECT($B7)``.

        Returns:
            The uppercase column letter, or ``None`` if parsing fails.
        """
        match = re.search(
            r"INDIRECT\(\$?([A-Z]+)\d+\)", formula_str, re.IGNORECASE
        )
        if match:
            return match.group(1).upper()
        return None

    def _find_column_values(self, col_letter_str: str) -> list:
        """Find the concrete list values for a given column in the map.

        Args:
            col_letter_str: The column letter(s) to look up (e.g. ``'B'``).

        Returns:
            The list of values if found, otherwise an empty list.
        """
        for ref, info in self._data_validation_map.items():
            ref_col = cell_reference_to_column(ref)
            if ref_col == col_letter_str and info.get("type") == "list":
                return info.get("values", [])
        return []

    def _resolve_named_range_variants(self, name: str) -> list:
        """Try multiple name variants when resolving a named range.

        Excel users sometimes define names with trailing underscores, spaces
        replaced by underscores, or stripped accents.  This method tries the
        literal name first and then a set of common transformations.

        Args:
            name: The base name to resolve.

        Returns:
            The list of cell values from the first matching named range,
            or an empty list if none matched.
        """
        if not hasattr(self._workbook, "defined_names"):
            return []

        result = self._resolve_named_range(name)
        if result:
            return result

        variants = [
            name + "_",
            name.replace(" ", "_"),
            name.replace(" ", "_") + "_",
            name.replace("'", ""),
            name.replace("'", "") + "_",
            re.sub(r"[^a-zA-Z0-9_àâäéèêëïîôùûüÿçœæ]", "_", name),
            re.sub(r"[^a-zA-Z0-9_àâäéèêëïîôùûüÿçœæ]", "_", name) + "_",
        ]

        seen: set = {name}
        for variant in variants:
            if variant not in seen:
                seen.add(variant)
                result = self._resolve_named_range(variant)
                if result:
                    return result

        return []

    def _resolve_dv_formula(self, formula_str: str) -> list:
        """Dispatch formula resolution to the appropriate strategy.

        Handles INDIRECT formulas, cell-range references, named ranges,
        and inline semicolon/comma-separated lists.

        Args:
            formula_str: The raw formula string from the data validation.

        Returns:
            A list of resolved string values (may be empty).
        """
        if formula_str.upper().startswith("INDIRECT("):
            return self._resolve_indirect(formula_str)

        if "$" in formula_str or ":" in formula_str or "!" in formula_str:
            return self._resolve_range_values(formula_str)

        is_identifier = formula_str.isidentifier() or (
            formula_str.replace("é", "e")
            .replace("è", "e")
            .replace("ê", "e")
            .replace("à", "a")
            .isidentifier()
        )
        if is_identifier:
            resolved = self._resolve_named_range(formula_str)
            if resolved:
                return resolved

        if ";" in formula_str:
            return [v.strip() for v in formula_str.split(";") if v.strip()]

        if "," in formula_str:
            return [v.strip() for v in formula_str.split(",") if v.strip()]

        return [formula_str.strip()] if formula_str.strip() else []

    def _resolve_named_range(self, name: str) -> list:
        """Resolve a workbook-level named range to its cell values.

        Args:
            name: The defined name to look up in the workbook.

        Returns:
            A list of non-empty string values from the named range,
            or an empty list on failure.
        """
        if not hasattr(self._workbook, "defined_names"):
            return []

        definition = self._workbook.defined_names.get(name)
        if definition is None:
            return []

        values: list = []
        try:
            for sheet_title, cell_range in definition.destinations:
                if sheet_title in self._workbook.sheetnames:
                    source_worksheet = self._workbook[sheet_title]
                    for row_or_cell in source_worksheet[cell_range]:
                        row_cells = (
                            row_or_cell
                            if isinstance(row_or_cell, tuple)
                            else (row_or_cell,)
                        )
                        for cell in row_cells:
                            if (
                                cell.value is not None
                                and str(cell.value).strip()
                            ):
                                values.append(str(cell.value).strip())
        except Exception:
            return []

        return values

    def _resolve_indirect(self, formula_str: str) -> list:
        """Resolve an INDIRECT formula by reading the referenced cell value.

        Args:
            formula_str: The full ``INDIRECT(...)`` formula string.

        Returns:
            The named-range values pointed to by the cell, or an empty list.
        """
        match = re.search(r"INDIRECT\((.+)\)", formula_str, re.IGNORECASE)
        if not match:
            return []

        ref = match.group(1).replace("$", "").strip("'\"")
        try:
            cell_value = self._worksheet[ref].value
            if cell_value and isinstance(cell_value, str):
                return self._resolve_named_range(cell_value.strip())
        except Exception:
            pass

        return []

    def _resolve_range_values(self, formula_str: str) -> list:
        """Resolve a cell-range formula (possibly cross-sheet) to values.

        Args:
            formula_str: A range reference such as ``Sheet1!$A$1:$A$10`` or
                ``A1:A10``.

        Returns:
            A list of non-empty string cell values from the range.
        """
        values: list = []
        try:
            if "!" in formula_str:
                sheet_part, range_part = formula_str.rsplit("!", 1)
                sheet_name = sheet_part.strip("'\"")
                if sheet_name not in self._workbook.sheetnames:
                    return values
                source_worksheet = self._workbook[sheet_name]
            else:
                source_worksheet = self._worksheet
                range_part = formula_str

            range_part = range_part.replace("$", "")

            for row_or_cell in source_worksheet[range_part]:
                row_cells = (
                    row_or_cell
                    if isinstance(row_or_cell, tuple)
                    else (row_or_cell,)
                )
                for cell in row_cells:
                    if cell.value is not None and str(cell.value).strip():
                        values.append(str(cell.value).strip())
        except Exception:
            pass

        return values

    @staticmethod
    def _expand_sqref(sqref: str) -> list:
        """Expand a square-reference string into individual cell references.

        Handles both single references (``"B3"``) and ranges (``"B3:D10"``).
        Ranges are capped at 500 rows to avoid runaway expansion.

        Args:
            sqref: The sqref string from the data validation.

        Returns:
            A list of individual A1-notation cell references.
        """
        if ":" not in sqref:
            return [sqref.strip()]

        parts = sqref.strip().split(":")
        if len(parts) != 2:
            return [sqref.strip()]

        start, end = parts
        start_col = cell_reference_to_column(start)
        start_row = int("".join(c for c in start if c.isdigit()) or "0")
        end_col = cell_reference_to_column(end)
        end_row = int("".join(c for c in end if c.isdigit()) or "0")

        if not start_row or not end_row:
            return [sqref.strip()]

        refs: list = []
        start_col_index = column_index(start_col)
        end_col_index = column_index(end_col)

        for col_idx in range(start_col_index, end_col_index + 1):
            col_letter_str = column_letter(col_idx)
            for row_idx in range(start_row, min(end_row + 1, start_row + 500)):
                refs.append(f"{col_letter_str}{row_idx}")

        return refs
