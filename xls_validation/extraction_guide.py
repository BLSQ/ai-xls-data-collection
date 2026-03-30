"""Build the extraction guide JSON from a validated schema.

The extraction guide is consumed by the ingest pipeline to know exactly
which cells to read and how to map them to database columns.

Supports both flat-key schemas (``header_modules`` / ``table_definitions``)
and sections-based schemas (``sections[layout=records]``).

No OpenHEXA SDK dependency — pure data transformation.
"""

from __future__ import annotations

from openpyxl.utils import column_index_from_string, get_column_letter


class ExtractionGuideBuilder:
    """Builds the extraction guide for a single sheet.

    After construction, call :meth:`build` to produce the guide dict.
    """

    def __init__(
        self,
        schema: dict,
        column_map: dict[str, str],
        module_columns: dict[str, list[str]],
        data_rows: list[int],
        grouped_info: dict,
    ) -> None:
        """Initialise the builder.

        Args:
            schema: The full validation schema dict.
            column_map: The canonical_name → column_letter mapping
                (from :class:`RecordsValidator`).
            module_columns: The module_key → column letters mapping.
            data_rows: The list of validated data row numbers.
            grouped_info: Detected grouped column structures from the
                pipeline orchestrator.
        """
        self.schema = schema
        self.column_map = column_map
        self.module_columns = module_columns
        self.data_rows = data_rows
        self.grouped_info = grouped_info

    def build(self) -> dict:
        """Build the extraction guide dict.

        Returns:
            A dict with ``metadata`` and ``data`` sections.
        """
        guide = {
            "metadata": self._build_metadata(),
            "data": {
                "start_row": min(self.data_rows) if self.data_rows else 2,
                "end_row": max(self.data_rows) if self.data_rows else 2,
                "row_count": len(self.data_rows),
                "columns": {},
            },
        }

        if self._has_sections():
            self._populate_sections_columns(guide["data"]["columns"])
        else:
            self._populate_module_columns(guide["data"]["columns"])

        return guide

    # -- Metadata -----------------------------------------------------------

    def _build_metadata(self) -> dict:
        """Build the metadata section of the extraction guide.

        Reads from ``table_definitions.metadata`` generically.

        Returns:
            A dict with ``fields`` and optionally ``grouped_sections``.
        """
        table_definition = self.schema.get("table_definitions", {}).get("metadata", {})
        fields: list[dict] = []

        for field in table_definition.get("fields", []):
            fields.append(
                {
                    "canonical_name": field["canonical_name"],
                    "value_cell": field["value_cell"],
                    "type": field["type"],
                }
            )

        # Build grouped sections generically
        groups: dict = {}
        for section_key, info in self.grouped_info.items():
            fields_key = f"{section_key}_fields"
            table_grouped = table_definition.get(fields_key, {})
            groups[section_key] = {
                "columns": info["columns"],
                "active_columns": info["active_columns"],
                "fields_per_column": table_grouped.get("fields_per_column", []),
            }

        result: dict = {"fields": fields}
        if groups:
            result["grouped_sections"] = groups

        return result

    # -- Data columns -------------------------------------------------------

    def _has_sections(self) -> bool:
        """Check whether the schema uses sections-based format.

        Returns:
            True if the schema has ``sections`` but no ``label_fields``.
        """
        return "sections" in self.schema and "label_fields" not in self.schema

    def _populate_sections_columns(self, columns_dict: dict) -> None:
        """Populate columns from sections-based schema.

        Args:
            columns_dict: The dict to populate with column entries.
        """
        for section in self.schema.get("sections", []):
            if section.get("layout") != "records":
                continue
            for column_definition in section.get("headers", {}).get("columns", []):
                self._walk_sections_column(columns_dict, column_definition)

    def _walk_sections_column(
        self,
        columns_dict: dict,
        column_definition: dict,
    ) -> None:
        """Walk a sections-based column entry and populate columns_dict.

        Args:
            columns_dict: The dict to populate.
            column_definition: A column definition dict from the schema.
        """
        column_type = column_definition.get("type", "single")

        if column_type == "single":
            canonical = column_definition.get("canonical_name", "")
            column_letter = column_definition.get("column", "")
            if canonical and column_letter:
                columns_dict[canonical] = {
                    "col_letter": column_letter,
                    "header": column_definition.get("expected", ""),
                    "value_type": column_definition.get("value_type", "string"),
                    "required": column_definition.get("required", False),
                }

        elif column_type == "group":
            parent = column_definition.get("parent", {})
            group_name = parent.get("expected", "")
            parent_canonical = parent.get("canonical_name", "")

            for sub_column in column_definition.get("sub_columns", []):
                canonical = sub_column.get("canonical_name", "")
                sub_column_letter = sub_column.get("column", "")
                full_canonical = (
                    f"{parent_canonical}__{canonical}"
                    if parent_canonical
                    else canonical
                )
                if full_canonical and sub_column_letter:
                    columns_dict[full_canonical] = {
                        "col_letter": sub_column_letter,
                        "header": sub_column.get("expected", ""),
                        "value_type": sub_column.get("value_type", "string"),
                        "group": group_name,
                    }

    def _populate_module_columns(self, columns_dict: dict) -> None:
        """Populate columns from flat-key header_modules.

        Args:
            columns_dict: The dict to populate with column entries.
        """
        header_modules = self.schema.get("header_modules", {})

        for module_key, module_definition in header_modules.items():
            if module_key.startswith("_") or not isinstance(module_definition, dict):
                continue
            if "columns" not in module_definition:
                continue

            column_cursor = column_index_from_string(
                module_definition.get("start_column", "B")
            )
            for column_entry in module_definition["columns"]:
                column_cursor = self._walk_module_column(
                    columns_dict,
                    column_entry,
                    module_key,
                    "default",
                    column_cursor,
                )

    def _walk_module_column(
        self,
        columns_dict: dict,
        column_entry: dict,
        module: str,
        category: str,
        column_cursor: int,
    ) -> int:
        """Walk a flat-key column entry and populate columns_dict.

        Args:
            columns_dict: The dict to populate.
            column_entry: A column entry dict from the schema.
            module: The parent module key.
            category: The column category (e.g. group name).
            column_cursor: The current column cursor position.

        Returns:
            The next column cursor position.
        """
        position = column_entry.get("position", "fixed")

        if position == "fixed":
            canonical = column_entry.get("canonical_name", "")
            column_letter = (
                column_entry.get("column") or get_column_letter(column_cursor)
            )
            if canonical:
                columns_dict[canonical] = {
                    "col_letter": column_letter,
                    "module": module,
                    "category": category,
                    "header": column_entry.get("expected", ""),
                    "value_type": column_entry.get("value_type", "string"),
                }
            if column_entry.get("column"):
                return column_index_from_string(column_entry["column"]) + 1
            return column_cursor + 1

        elif position == "group":
            parent = column_entry.get("parent", {})
            group_category = parent.get("expected", category)
            cursor = column_cursor
            for sub_column in column_entry.get("sub_columns", []):
                cursor = self._walk_module_column(
                    columns_dict,
                    sub_column,
                    module,
                    group_category,
                    cursor,
                )
            return cursor

        elif position == "dynamic":
            prefix = column_entry.get("canonical_prefix", "")
            value_type = column_entry.get("value_type", "string")
            count = 0
            for canonical, column_letter in self.column_map.items():
                if canonical.startswith(prefix + "_"):
                    header = canonical[len(prefix) + 1:].replace("_", " ").title()
                    columns_dict[canonical] = {
                        "col_letter": column_letter,
                        "module": module,
                        "category": category,
                        "header": header,
                        "value_type": value_type,
                    }
                    count += 1
            return column_cursor + count

        return column_cursor
