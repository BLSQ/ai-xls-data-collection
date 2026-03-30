"""Section processors that convert detected sections into schema entries.

Each processor handles one section layout (key-value, grouped key-value,
or records) and produces the corresponding labels, values, and metadata
dictionaries consumed by the schema builder.

This module is a faithful refactoring of ``process_key_value``,
``process_grouped_key_value``, ``process_records``, and their helpers
from the original monolithic pipeline.
"""

from __future__ import annotations

import re
from abc import ABC, abstractmethod

from text_helpers import (
    cell_reference_to_column,
    cell_reference_to_row,
    column_index,
    format_example,
    to_canonical,
    to_column_canonical,
)


class SectionProcessor(ABC):
    """Base class for processing a detected section into schema entries.

    Subclasses must implement the three-stage pipeline (register, qualify,
    add_validation) and may override ``process`` to combine them when the
    original logic is naturally single-pass.
    """

    def __init__(self, worksheet, data_validation_map: dict) -> None:
        """Initialise the processor with the worksheet and validation map.

        Args:
            worksheet: An openpyxl worksheet object.
            data_validation_map: A mapping from cell references to their
                resolved data-validation information.
        """
        self.worksheet = worksheet
        self.data_validation_map = data_validation_map

    @abstractmethod
    def register(self, section: dict) -> dict:
        """Extract raw structural information from the section definition.

        Args:
            section: The section dictionary from the structure proposal.

        Returns:
            A dict of registered structural data.
        """

    @abstractmethod
    def qualify(self, registered: dict) -> dict:
        """Enrich registered data with examples and type information.

        Args:
            registered: Output of ``register``.

        Returns:
            A dict of qualified data.
        """

    @abstractmethod
    def add_validation(self, qualified: dict) -> dict:
        """Attach validation rules (enums, conditional enums) to qualified data.

        Args:
            qualified: Output of ``qualify``.

        Returns:
            A dict of fully validated data.
        """

    def process(self, section: dict) -> dict:
        """Convenience method that chains register -> qualify -> add_validation.

        Args:
            section: The section dictionary from the structure proposal.

        Returns:
            The fully processed result.
        """
        registered = self.register(section)
        qualified = self.qualify(registered)
        return self.add_validation(qualified)

    @staticmethod
    def derive_module_key(section: dict) -> str:
        """Derive a module key from the section title or fall back to the section key.

        Looks for a pattern like ``module 1`` in the title and returns
        ``module_1``.  Falls back to ``section["key"]`` or ``"unknown"``.

        Args:
            section: The section dictionary from the structure proposal.

        Returns:
            A string suitable for use as a module identifier.
        """
        title = (section.get("title") or {}).get("expected", "")
        match = re.search(r"module\s*(\d+)", title, re.IGNORECASE)
        if match:
            return f"module_{match.group(1)}"
        return section.get("key", "unknown")

    def _read_example(self, row: int, column_letter_str: str) -> str:
        """Safely read a cell value and format it as an example string.

        Args:
            row: The 1-based row number.
            column_letter_str: The column letter (e.g. ``'C'``).

        Returns:
            A formatted example string, or empty string on failure.
        """
        try:
            raw_value = self.worksheet.cell(
                row=row, column=column_index(column_letter_str)
            ).value
            return format_example(raw_value)
        except Exception:
            return ""


class KeyValueProcessor(SectionProcessor):
    """Processor for simple key-value sections.

    Each field has a label cell and a value cell, producing flat label and
    value dictionaries keyed by field key.
    """

    def register(self, section: dict) -> dict:
        """Extract field definitions and cell references from the section.

        Args:
            section: The section dictionary.

        Returns:
            A dict containing module_key, title_info, and raw field data.
        """
        return {
            "module_key": self.derive_module_key(section),
            "title_info": section.get("title") or {},
            "fields": section.get("fields", []),
        }

    def qualify(self, registered: dict) -> dict:
        """Pass-through; qualification happens in ``process`` for this layout.

        Args:
            registered: Output of ``register``.

        Returns:
            The same dict, unchanged.
        """
        return registered

    def add_validation(self, qualified: dict) -> dict:
        """Pass-through; validation is applied in ``process`` for this layout.

        Args:
            qualified: Output of ``qualify``.

        Returns:
            The same dict, unchanged.
        """
        return qualified

    def process(self, section: dict) -> tuple[str, dict, dict]:
        """Process a key-value section into labels and values.

        Args:
            section: The section dictionary from the structure proposal.

        Returns:
            A tuple of (module_key, labels, values).
        """
        module_key = self.derive_module_key(section)
        labels: dict = {}
        values: dict = {}

        title_info = section.get("title") or {}
        if title_info.get("cell"):
            labels["title"] = {
                "cell": title_info["cell"],
                "expected": title_info.get("expected", ""),
            }

        for field in section.get("fields", []):
            field_key = field.get("key", "unnamed")

            # --- label ---
            label_info = field.get("label", {})
            label_cell = label_info.get("cell", "")
            labels[field_key] = {
                "cell": label_cell,
                "expected": label_info.get("expected", ""),
            }

            # --- value ---
            value_info = field.get("value", {})
            value_cell = value_info.get("cell", "")
            value_cell_single = (
                value_cell.split(":")[0] if ":" in value_cell else value_cell
            )
            value_type = value_info.get("type", "string")

            example = ""
            if value_cell_single:
                try:
                    col_letter = cell_reference_to_column(value_cell_single)
                    row_number = cell_reference_to_row(value_cell_single)
                    raw_value = self.worksheet.cell(
                        row=row_number, column=column_index(col_letter)
                    ).value
                    example = format_example(raw_value)
                except Exception:
                    pass

            label_display = label_info.get("expected", "").rstrip(":").strip()
            value_key = f"{field_key}_value"
            value_entry = {
                "cell": value_cell_single,
                "label_ref": f"{module_key}.{field_key}",
                "label_display": label_display,
                "type": value_type,
                "example": example,
                "required": False,
                "message": (
                    f"Le champ '{label_display}' "
                    f"(cellule {value_cell_single}) est optionnel."
                ),
            }

            if value_cell_single and self.data_validation_map.get(value_cell_single):
                dv_info = self.data_validation_map[value_cell_single]
                if dv_info.get("values"):
                    value_entry["validation"] = {"enum": dv_info["values"]}
                    value_entry["type"] = "string"

            values[value_key] = value_entry

        return module_key, labels, values


class GroupedKeyValueProcessor(SectionProcessor):
    """Processor for grouped key-value sections.

    Multiple elements share the same set of property rows, arranged under
    a set of columns (e.g. a table where each column is an element and
    each row is a property).
    """

    def register(self, section: dict) -> dict:
        """Extract grouped structure metadata from the section.

        Args:
            section: The section dictionary.

        Returns:
            A dict containing section_key, module_key, and raw field data.
        """
        return {
            "section_key": section.get("key", "group"),
            "module_key": self.derive_module_key(section),
            "section": section,
        }

    def qualify(self, registered: dict) -> dict:
        """Pass-through; qualification happens in ``process`` for this layout.

        Args:
            registered: Output of ``register``.

        Returns:
            The same dict, unchanged.
        """
        return registered

    def add_validation(self, qualified: dict) -> dict:
        """Pass-through; validation is applied in ``process`` for this layout.

        Args:
            qualified: Output of ``qualify``.

        Returns:
            The same dict, unchanged.
        """
        return qualified

    def process(self, section: dict) -> tuple[str, str, dict, dict, dict | None]:
        """Process a grouped key-value section.

        Args:
            section: The section dictionary from the structure proposal.

        Returns:
            A tuple of (section_key, module_key, labels, values,
            column_labels_entry).
        """
        section_key = section.get("key", "group")
        module_key = self.derive_module_key(section)
        labels: dict = {}
        values: dict = {}

        title_info = section.get("title") or {}
        section_title = title_info.get("expected", section_key)
        if title_info.get("cell"):
            field_key = to_canonical(section_title)
            labels[field_key] = {
                "cell": title_info["cell"],
                "expected": section_title,
            }

        # --- column labels ---
        elements = section.get("elements", {})
        element_columns = elements.get("columns", [])
        element_labels = elements.get("labels", [])
        header_row = elements.get("header_row")

        column_labels_items = []
        for index, column_letter_str in enumerate(element_columns):
            label_text = (
                element_labels[index]
                if index < len(element_labels)
                else f"Element {index + 1}"
            )
            column_labels_items.append({
                "cell": f"{column_letter_str}{header_row}",
                "expected": label_text,
            })

        column_labels_entry = None
        if column_labels_items:
            display_name = (
                section_title
                if section_title and section_title != section_key
                else section_key.replace("_", " ").title()
            )
            column_labels_entry = {
                "_description": (
                    f"Column labels for grouped structure "
                    f"'{section_key}' (row {header_row})."
                ),
                "layout": "grouped_key_value",
                "display_name": display_name,
                "count_is_dynamic": elements.get("count_is_dynamic", False),
                "header_row": header_row,
                "items": column_labels_items,
            }

        # --- property rows ---
        property_rows: dict = {}
        for field_index, field in enumerate(section.get("fields", [])):
            field_key = field.get("key", "unnamed")
            label_info = field.get("label", {})
            row_number = label_info.get("row")
            col_letter = label_info.get("col", "B")
            label_cell = f"{col_letter}{row_number}" if row_number else ""
            labels[field_key] = {
                "cell": label_cell,
                "expected": label_info.get("expected", ""),
            }

            example = ""
            first_column = element_columns[0] if element_columns else None
            if first_column and row_number:
                example = self._read_example(row_number, first_column)

            is_first = field_index == 0
            field_type = field.get("type", "string")
            label_display = label_info.get("expected", "")

            if is_first:
                message = (
                    f"Au moins un '{label_display}' "
                    f"(ligne {row_number}) doit \u00eatre renseign\u00e9."
                )
            else:
                message = (
                    f"Le '{label_display}' (ligne {row_number}) "
                    f"est requis pour chaque \u00e9l\u00e9ment actif."
                )

            row_entry = {
                "row": row_number,
                "label_ref": f"{module_key}.{field_key}",
                "label_display": label_display,
                "type": field_type,
                "example": example,
                "required_rule": field.get(
                    "required_rule", "anchor" if is_first else "if_active"
                ),
                "message": message,
            }
            property_rows[field_key] = row_entry

        grouped_key = f"{section_key}_values"
        values[grouped_key] = {
            "_description": (
                f"Grouped rows for '{section_key}' "
                f"under columns {', '.join(element_columns)}."
            ),
            "columns": element_columns,
            "rows": property_rows,
        }

        return section_key, module_key, labels, values, column_labels_entry


class RecordsProcessor(SectionProcessor):
    """Processor for tabular record sections.

    Each section describes a table with fixed or grouped columns, a title
    row, and data rows starting at ``data_start_row``.
    """

    def __init__(
        self,
        worksheet,
        data_validation_map: dict,
        data_start_row: int,
        conditional_map: dict | None = None,
    ) -> None:
        """Initialise the records processor.

        Args:
            worksheet: An openpyxl worksheet object.
            data_validation_map: A mapping from cell references to their
                resolved data-validation information.
            data_start_row: The first row of actual data in the table.
            conditional_map: Optional mapping from column letters to
                conditional validation rules.
        """
        super().__init__(worksheet, data_validation_map)
        self.data_start_row = data_start_row
        self.conditional_map = conditional_map or {}

    def register(self, section: dict) -> dict:
        """Extract the title, area, and raw column definitions.

        Args:
            section: The section dictionary from the structure proposal.

        Returns:
            A dict with module_key, title_info, start_column, and raw
            column definitions.
        """
        title_info = section.get("title") or {}
        headers_info = section.get("headers") or {}
        return {
            "module_key": self.derive_module_key(section),
            "title_info": title_info,
            "start_column": section.get("area", {}).get("start_col", ""),
            "column_definitions": headers_info.get("columns", []),
        }

    def qualify(self, registered: dict) -> dict:
        """Build the column entries with examples and type information.

        Args:
            registered: Output of ``register``.

        Returns:
            A dict with module_key, header_entry (containing built columns),
            and title_label.
        """
        columns = []
        for column_definition in registered["column_definitions"]:
            column_type = column_definition.get(
                "type", column_definition.get("position", "single")
            )
            if column_type in ("single", "fixed"):
                columns.append(self._build_fixed_column(column_definition))
            elif column_type == "group":
                columns.append(self._build_group_column(column_definition))

        title_info = registered["title_info"]
        header_entry = {
            "title": {
                "cell": title_info.get("cell", ""),
                "expected": title_info.get("expected", ""),
            },
            "start_column": registered["start_column"],
            "columns": columns,
        }
        title_label = {
            "cell": title_info.get("cell", ""),
            "expected": title_info.get("expected", ""),
        }
        return {
            "module_key": registered["module_key"],
            "header_entry": header_entry,
            "title_label": title_label,
        }

    def add_validation(self, qualified: dict) -> dict:
        """Pass-through; validation is applied inside column builders.

        Args:
            qualified: Output of ``qualify``.

        Returns:
            The same dict, unchanged.
        """
        return qualified

    def process(self, section: dict) -> tuple[str, dict, dict]:
        """Process a records section into a header entry and title label.

        Args:
            section: The section dictionary from the structure proposal.

        Returns:
            A tuple of (module_key, header_entry, title_label).
        """
        registered = self.register(section)
        qualified = self.qualify(registered)
        result = self.add_validation(qualified)
        return (
            result["module_key"],
            result["header_entry"],
            result["title_label"],
        )

    def _build_fixed_column(self, column_definition: dict) -> dict:
        """Build a schema entry for a fixed (single) column.

        Reads an example value from the first data row and attaches any
        data-validation or conditional-validation rules found for that column.

        Args:
            column_definition: The column definition dict from the structure
                proposal.

        Returns:
            A fully populated column entry dict.
        """
        column_letter_str = column_definition.get("column", "")
        header_row = column_definition.get("header_row", 15)

        example = ""
        if column_letter_str and self.data_start_row:
            example = self._read_example(self.data_start_row, column_letter_str)

        entry = {
            "position": "fixed",
            "column": column_letter_str,
            "row": header_row,
            "expected": column_definition.get("expected", ""),
            "canonical_name": to_column_canonical(
                column_definition.get("expected", "")
            ),
            "value_type": column_definition.get("value_type", "string"),
            "example": example,
            "required": column_definition.get("required", False),
            "known_variants": [],
        }

        data_cell_ref = f"{column_letter_str}{self.data_start_row}"

        if column_letter_str in self.conditional_map:
            condition = self.conditional_map[column_letter_str]
            entry["validation"] = {
                "conditional_enum": {
                    "depends_on_column": condition["depends_on_column"],
                    "values_by_parent": condition.get("values_by_parent", {}),
                },
            }
            entry["value_type"] = "string"
        elif self.data_validation_map.get(data_cell_ref):
            dv_info = self.data_validation_map[data_cell_ref]
            if dv_info.get("values"):
                entry["validation"] = {"enum": dv_info["values"]}
                entry["value_type"] = "string"

        return entry

    def _build_group_column(self, column_definition: dict) -> dict:
        """Build a schema entry for a grouped column with sub-columns.

        A group column has a parent header spanning multiple sub-columns,
        each of which may be fixed or dynamic.

        Args:
            column_definition: The column definition dict from the structure
                proposal.

        Returns:
            A group entry dict containing parent info and sub-column entries.
        """
        parent = column_definition.get("parent", {})
        group = {
            "position": "group",
            "parent": {
                "row": parent.get("header_row", 15),
                "expected": parent.get("expected", ""),
                "known_variants": [],
            },
            "sub_columns": [],
        }

        parent_canonical = to_column_canonical(parent.get("expected", ""))

        for sub in column_definition.get("sub_columns", []):
            is_fixed = sub.get("fixed", sub.get("position") == "fixed")
            is_dynamic = (
                not is_fixed
                if "fixed" in sub
                else sub.get("position") == "dynamic"
            )

            if is_dynamic:
                group["sub_columns"].append({
                    "position": "dynamic",
                    "row": sub.get("header_row", 16),
                    "pattern": sub.get("pattern", ".*"),
                    "canonical_prefix": parent_canonical,
                    "value_type": sub.get("value_type", "string"),
                    "description": sub.get("description", ""),
                })
            else:
                sub_column_letter = sub.get("column", "")
                sub_canonical = (
                    f"{parent_canonical}"
                    f"_{to_column_canonical(sub.get('expected', ''))}"
                )

                example = ""
                if sub_column_letter and self.data_start_row:
                    example = self._read_example(
                        self.data_start_row, sub_column_letter
                    )

                sub_entry = {
                    "position": "fixed",
                    "column": sub_column_letter,
                    "row": sub.get("header_row", 16),
                    "expected": sub.get("expected", ""),
                    "canonical_name": sub_canonical,
                    "value_type": sub.get("value_type", "string"),
                    "example": example,
                    "required": sub.get("required", False),
                    "known_variants": [],
                }

                data_cell_ref = f"{sub_column_letter}{self.data_start_row}"
                if sub_column_letter and self.data_validation_map.get(
                    data_cell_ref
                ):
                    dv_info = self.data_validation_map[data_cell_ref]
                    if dv_info.get("values"):
                        sub_entry["validation"] = {"enum": dv_info["values"]}
                        sub_entry["value_type"] = "string"

                group["sub_columns"].append(sub_entry)

        return group
