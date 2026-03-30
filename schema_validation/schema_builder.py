"""Orchestrates section processors and assembles the final schema_validation.json.

This module contains the :class:`SchemaBuilder` class, which is responsible for
iterating over the sections declared in a structure proposal, delegating each
section to the appropriate processor, and combining the results into a single
schema dictionary suitable for downstream validation.

All logging is deferred: instead of calling ``current_run.log_info()`` directly,
messages are accumulated in :pyattr:`SchemaBuilder.log_messages` so that the
orchestrating pipeline can emit them at the appropriate time.
"""

from __future__ import annotations

from section_processors import (
    GroupedKeyValueProcessor,
    KeyValueProcessor,
    RecordsProcessor,
)
from text_helpers import to_canonical, to_column_canonical


class SchemaBuilder:
    """Orchestrates section processors and assembles the final schema."""

    def __init__(
        self,
        proposal: dict,
        worksheet,
        dv_resolver,
        conditional_map: dict | None = None,
    ) -> None:
        """Initialise the builder with the inputs required for schema generation.

        Args:
            proposal:        The parsed structure_proposal.json dictionary.
            worksheet:       An openpyxl worksheet (data_only=True) for the target sheet.
            dv_resolver:     The data_validation_map (cell ref → validation info).
            conditional_map: Column letter → conditional enum info (from
                             DataValidationResolver.resolve_indirect_dependencies).
        """
        self.proposal = proposal
        self.worksheet = worksheet
        self.dv_resolver = dv_resolver
        self.conditional_map = conditional_map or {}
        self.log_messages: list[str] = []

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def build(self) -> dict:
        """Build and return the complete schema dictionary.

        This is the main entry point.  It walks every section declared in the
        proposal, delegates processing to the relevant section processor, and
        assembles the label fields, value fields, header modules, and table
        definitions into a single schema dict.

        Returns:
            A dictionary ready to be serialised as schema_validation.json.
        """
        proposal = self.proposal
        conditional_map = self.conditional_map

        landmarks = proposal.get("structure_landmarks", {})
        title_row = landmarks.get("title_row", 14)
        header_rows = landmarks.get("header_rows", [15, 16])
        data_start_row = landmarks.get("data_start_row", 17)
        header_row = header_rows[0] if header_rows else 15

        label_fields: dict = {}
        value_fields: dict = {}
        header_modules: dict = {
            "_description": (
                "Each module defines its own header columns independently. "
                "Columns can be 'fixed', 'dynamic', or 'group'."
            ),
            "mismatch_message_template": (
                "L'en-tete en colonne {column} ('{actual_value}') ne correspond "
                "pas a la valeur attendue ('{expected}'). Confirmez-vous que "
                "cette colonne correspond bien a '{canonical_name}' ?"
            ),
        }
        module_headers_row: dict = {}
        metadata_module_keys: list[str] = []
        primary_key: str | None = None

        for section in proposal.get("sections", []):
            layout = section.get("layout")

            if layout == "key_value":
                self._process_key_value_section(
                    section,
                    label_fields,
                    value_fields,
                    metadata_module_keys,
                )

            elif layout == "grouped_key_value":
                self._process_grouped_key_value_section(
                    section,
                    label_fields,
                    value_fields,
                    metadata_module_keys,
                )

            elif layout == "records":
                records_processor = RecordsProcessor(
                    self.worksheet,
                    self.dv_resolver,
                    data_start_row,
                    conditional_map,
                )
                module_key, header_entry, title_label = records_processor.process(section)

                header_modules[module_key] = header_entry
                module_headers_row[module_key] = title_label

                if primary_key is None:
                    primary_key = self._find_primary_key(
                        header_entry,
                        section.get("headers", {}),
                    )

                column_count = len(header_entry.get("columns", []))
                self.log_messages.append(
                    f"  records section '{module_key}': {column_count} column entries"
                )

        if module_headers_row:
            row_key = f"module_headers_row{title_row}"
            label_fields[row_key] = module_headers_row

        self._deduplicate_canonicals(header_modules)

        table_definitions = self._build_table_definitions(
            label_fields,
            value_fields,
            metadata_module_keys,
            header_row,
            data_start_row,
            primary_key,
        )

        schema: dict = {
            "$schema": "https://json-schema.org/draft/2020-12/schema",
            "version": "1.0.0",
            "description": (
                f"Validation schema for '{proposal.get('source_sheet', '')}' "
                "sheet -- generated from structure_proposal.json."
            ),
            "sheet_name": proposal.get("source_sheet", ""),
            "label_fields": label_fields,
            "value_fields": value_fields,
            "header_modules": header_modules,
            "table_definitions": table_definitions,
            "validation_config": self._build_validation_config(data_start_row),
        }

        generation = proposal.get("_generation", {})
        if generation.get("source_file"):
            schema["template_file"] = generation["source_file"]

        return schema

    # ------------------------------------------------------------------
    # Section dispatch helpers
    # ------------------------------------------------------------------

    def _process_key_value_section(
        self,
        section: dict,
        label_fields: dict,
        value_fields: dict,
        metadata_module_keys: list[str],
    ) -> None:
        """Process a single key_value section and merge results into the accumulators.

        Args:
            section: The section definition from the proposal.
            label_fields: Accumulator for label field mappings.
            value_fields: Accumulator for value field mappings.
            metadata_module_keys: Ordered list of metadata module keys seen so far.
        """
        module_key, labels, values = KeyValueProcessor(
            self.worksheet, self.dv_resolver
        ).process(section)

        label_fields[module_key] = labels
        value_fields[f"{module_key}_values"] = values
        metadata_module_keys.append(module_key)

        self.log_messages.append(
            f"  key_value section '{module_key}': {len(labels)} labels, {len(values)} values"
        )

    def _process_grouped_key_value_section(
        self,
        section: dict,
        label_fields: dict,
        value_fields: dict,
        metadata_module_keys: list[str],
    ) -> None:
        """Process a single grouped_key_value section and merge results into the accumulators.

        Args:
            section: The section definition from the proposal.
            label_fields: Accumulator for label field mappings.
            value_fields: Accumulator for value field mappings.
            metadata_module_keys: Ordered list of metadata module keys seen so far.
        """
        section_key, module_key, labels, values, column_labels = (
            GroupedKeyValueProcessor(
                self.worksheet, self.dv_resolver
            ).process(section)
        )

        parent_key = self._find_parent_module(metadata_module_keys) or module_key

        if parent_key in label_fields:
            label_fields[parent_key].update(labels)
        else:
            label_fields[module_key] = labels
            metadata_module_keys.append(module_key)

        grouped_key = f"{section_key}_values"

        if parent_key != module_key:
            grouped_values = values.get(grouped_key, {})
            for _row_name, row_data in grouped_values.get("rows", {}).items():
                if "label_ref" in row_data:
                    row_data["label_ref"] = row_data["label_ref"].replace(
                        f"{module_key}.", f"{parent_key}.", 1
                    )

        parent_values_key = f"{parent_key}_values"
        if parent_values_key in value_fields:
            value_fields[parent_values_key].update(values)
        else:
            value_fields[f"{module_key}_values"] = values
            if module_key not in metadata_module_keys:
                metadata_module_keys.append(module_key)

        if column_labels:
            actual_parent = parent_key if parent_key in label_fields else module_key
            column_labels["values_section"] = f"{actual_parent}_values"
            column_labels["values_key"] = grouped_key
            column_labels["label_module"] = actual_parent
            label_fields[f"{section_key}_columns"] = column_labels

        self.log_messages.append(
            f"  grouped_key_value section '{section_key}': "
            f"{len(labels)} labels, merged into '{parent_key}'"
        )

    # ------------------------------------------------------------------
    # Table definitions
    # ------------------------------------------------------------------

    def _build_table_definitions(
        self,
        label_fields: dict,
        value_fields: dict,
        metadata_module_keys: list[str],
        header_row: int,
        data_start_row: int,
        primary_key: str | None = None,
    ) -> dict:
        """Assemble the ``table_definitions`` section of the schema.

        Combines metadata fields (from key-value and grouped key-value sections)
        and records-data configuration into a single dictionary.

        Args:
            label_fields: Accumulated label field mappings.
            value_fields: Accumulated value field mappings.
            metadata_module_keys: Ordered list of metadata module keys.
            header_row: The first header row number.
            data_start_row: The first data row number.
            primary_key: The canonical name of the primary key column, if found.

        Returns:
            A dictionary describing the metadata and records_data tables.
        """
        table_definitions: dict = {}
        fields_list: list[dict] = []

        for module_key in metadata_module_keys:
            labels = label_fields.get(module_key, {})
            values = value_fields.get(f"{module_key}_values", {})

            for value_key, value_field in values.items():
                if value_key.startswith("_"):
                    continue
                if isinstance(value_field, dict) and "columns" in value_field and "rows" in value_field:
                    continue
                if not isinstance(value_field, dict) or "cell" not in value_field:
                    continue

                label_ref = value_field.get("label_ref", "")
                label_key = label_ref.split(".")[-1] if "." in label_ref else ""
                label_cell = labels.get(label_key, {}).get("cell", "")

                fields_list.append({
                    "canonical_name": label_key or to_canonical(
                        value_field.get("label_display", value_key)
                    ),
                    "label_cell": label_cell,
                    "value_cell": value_field.get("cell", ""),
                    "type": value_field.get("type", "string"),
                })

        grouped_fields: dict = {}
        for module_key in metadata_module_keys:
            values = value_fields.get(f"{module_key}_values", {})

            for value_key, value_field in values.items():
                if not isinstance(value_field, dict):
                    continue
                if "columns" not in value_field or "rows" not in value_field:
                    continue

                section_key = (
                    value_key.removesuffix("_values")
                    if value_key.endswith("_values")
                    else value_key
                )

                grouped_fields[f"{section_key}_fields"] = {
                    "_description": (
                        f"Repeating group '{section_key}' "
                        f"(columns {', '.join(value_field['columns'])})."
                    ),
                    "columns": value_field["columns"],
                    "fields_per_column": [
                        {
                            "canonical_name": f"{section_key}_{row_key}",
                            "row": row_data["row"],
                            "type": row_data["type"],
                        }
                        for row_key, row_data in value_field.get("rows", {}).items()
                    ],
                }

        table_definitions["metadata"] = {
            "_description": "Metadata table built from label-value pairs.",
            "source": "metadata sections",
            "fields": fields_list,
        }
        table_definitions["metadata"].update(grouped_fields)

        primary_column_name = primary_key or "primary_column"
        table_definitions["records_data"] = {
            "_description": "Data table built from data rows. Each row becomes a record.",
            "header_row_start": header_row,
            "header_row_end": header_row + 1,
            "data_row_start": data_start_row,
            "data_row_end": "dynamic",
            "header_source": "header_modules",
            "empty_row_terminates": True,
            "row_validation": {
                primary_column_name: {
                    "required": True,
                    "message": (
                        f"Chaque ligne de donnees doit avoir un '{primary_column_name}'."
                    ),
                },
            },
        }

        return table_definitions

    # ------------------------------------------------------------------
    # Validation config
    # ------------------------------------------------------------------

    def _build_validation_config(self, data_start_row: int) -> dict:
        """Build the ``validation_config`` section of the schema.

        Args:
            data_start_row: The first data row number.

        Returns:
            A dictionary describing label comparison rules, header matching
            strategy, dynamic column handling, and data row termination.
        """
        return {
            "_description": "Configuration for validation and header matching behavior.",
            "label_comparison": {
                "case_sensitive": False,
                "strip_whitespace": True,
                "strip_trailing_colon": True,
                "normalize_accents": True,
            },
            "header_matching": {
                "strategy": "exact_then_variants_then_ask",
                "on_mismatch": "warn_and_ask_confirmation",
                "on_missing": "error",
            },
            "dynamic_columns": {
                "_description": (
                    "Columns with position='dynamic' are matched by regex pattern. "
                    "Canonical names are generated as '{canonical_prefix}_{normalized_value}'."
                ),
            },
            "data_rows": {
                "start_row": data_start_row,
                "termination": "first_fully_empty_row",
            },
        }

    # ------------------------------------------------------------------
    # Primary key detection
    # ------------------------------------------------------------------

    def _find_primary_key(
        self,
        header_entry: dict,
        headers_proposal: dict,
    ) -> str | None:
        """Identify the primary key column from a records header entry.

        The method first checks whether the AI proposal nominated a primary key.
        If so, it attempts to match it against the fixed columns by canonical
        name.  Failing that, the first required column is used, and as a last
        resort the very first fixed column.

        Args:
            header_entry: The header module entry (with a ``columns`` list).
            headers_proposal: The raw ``headers`` dict from the section proposal,
                which may contain a ``primary_key`` hint.

        Returns:
            The canonical name of the chosen primary key column, or ``None`` if
            no columns are available.
        """
        gemini_primary_key = headers_proposal.get("primary_key", "")
        columns = header_entry.get("columns", [])

        flat_columns: list[dict] = []
        for column in columns:
            if column.get("position") == "fixed":
                flat_columns.append(column)
            elif column.get("position") == "group":
                for sub_column in column.get("sub_columns", []):
                    if sub_column.get("position") == "fixed":
                        flat_columns.append(sub_column)

        if gemini_primary_key:
            for column in flat_columns:
                expected = column.get("expected", "")
                if to_column_canonical(expected) == to_column_canonical(gemini_primary_key):
                    return column.get("canonical_name")
                if column.get("canonical_name") == gemini_primary_key:
                    return column["canonical_name"]

        for column in flat_columns:
            if column.get("required"):
                return column.get("canonical_name")

        if flat_columns:
            return flat_columns[0].get("canonical_name")

        return None

    # ------------------------------------------------------------------
    # Canonical name deduplication
    # ------------------------------------------------------------------

    def _deduplicate_canonicals(self, header_modules: dict) -> None:
        """Ensure canonical names are unique across all header modules.

        When two or more columns in different modules share the same canonical
        name, each is prefixed with its module key to make it unique.

        Args:
            header_modules: The assembled header_modules dictionary (mutated
                in place).
        """
        name_map: dict[str, list[tuple[str, dict]]] = {}

        for module_key, module_definition in header_modules.items():
            if module_key.startswith("_") or module_key == "mismatch_message_template":
                continue
            if not isinstance(module_definition, dict):
                continue
            for column_entry in module_definition.get("columns", []):
                self._collect_canonicals(name_map, module_key, column_entry)

        for name, entries in name_map.items():
            if len(entries) > 1:
                for module_key, entry_dict in entries:
                    entry_dict["canonical_name"] = f"{module_key}_{name}"

    def _collect_canonicals(
        self,
        name_map: dict[str, list[tuple[str, dict]]],
        module_key: str,
        column_entry: dict,
    ) -> None:
        """Recursively collect canonical names from a column entry into *name_map*.

        Handles both ``fixed`` columns (collected directly) and ``group``
        columns (whose ``sub_columns`` are inspected).

        Args:
            name_map: Accumulator mapping canonical names to a list of
                ``(module_key, entry_dict)`` tuples.
            module_key: The key of the header module this column belongs to.
            column_entry: A single column (or group) entry dictionary.
        """
        position = column_entry.get("position", "")

        if position == "fixed":
            canonical = column_entry.get("canonical_name", "")
            if canonical:
                name_map.setdefault(canonical, []).append((module_key, column_entry))
        elif position == "group":
            for sub_column in column_entry.get("sub_columns", []):
                if sub_column.get("position") == "fixed":
                    canonical = sub_column.get("canonical_name", "")
                    if canonical:
                        name_map.setdefault(canonical, []).append(
                            (module_key, sub_column)
                        )

    # ------------------------------------------------------------------
    # Utilities
    # ------------------------------------------------------------------

    def _find_parent_module(self, metadata_module_keys: list[str]) -> str | None:
        """Return the first metadata module key, or ``None`` if none exist.

        Grouped key-value sections are merged into the first (parent) metadata
        module when one is available.

        Args:
            metadata_module_keys: Ordered list of metadata module keys seen so far.

        Returns:
            The first key, or ``None``.
        """
        return metadata_module_keys[0] if metadata_module_keys else None
