"""OpenHEXA Pipeline: Excel Validation (AI-schema based).

Validates Excel files against schema_validation.json produced by the
schema_validation pipeline. Fully generalized — no hardcoded domain terms.

Produces a JSON report with errors, warnings, and confirmations per sheet,
plus an extraction_guide.json for the ingest pipeline.

This file is the thin orchestrator — all domain logic lives in:

- :mod:`text_helpers`          — text normalization and matching
- :mod:`cell_helpers`          — cell reading and issue building
- :mod:`json_logic`            — JsonLogic rule evaluation
- :mod:`section_validators`    — KeyValue / GroupedKeyValue / Records validators
- :mod:`extraction_guide`      — ExtractionGuideBuilder
"""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

import openpyxl
from openhexa.sdk import File, current_run, parameter, pipeline, workspace

from cell_helpers import find_data_rows, read_cell, read_cell_by_row_column
from extraction_guide import ExtractionGuideBuilder
from section_validators import (
    GroupedKeyValueValidator,
    KeyValueValidator,
    RecordsValidator,
)
from text_helpers import text_matches

SKIP_SHEETS = {"guide", "observations", "liste des catégories"}
SCHEMA_FILENAME = "schema_validation.json"


# ---------------------------------------------------------------------------
# Main pipeline
# ---------------------------------------------------------------------------


@pipeline("xls-validation", timeout=3600)
@parameter(
    "excel_file",
    name="Excel File Path",
    type=File,
    required=True,
    help="Path to the Excel file to validate (relative to workspace files).",
)
def xls_validation(excel_file: File):
    """Validate an Excel file against the AI-generated schema."""
    current_run.log_info(f"Starting validation of: {excel_file}")

    excel_path = excel_file.path
    schema_file = Path(workspace.files_path) / SCHEMA_FILENAME

    schema = _load_schema(schema_file)
    config = schema.get("validation_config", {}).get("label_comparison", {})

    current_run.log_info(f"Loading workbook: {excel_path}")
    workbook = openpyxl.load_workbook(str(excel_path), data_only=True)

    report, extraction_guide = _build_report(workbook, schema, config)

    report_path = Path(workspace.files_path) / "validation_report.json"
    _save_json(report, report_path, "validation report")

    guide_path = Path(workspace.files_path) / "extraction_guide.json"
    _save_json(extraction_guide, guide_path, "extraction guide")

    _log_summary(report)


def _load_schema(schema_path: Path) -> dict:
    """Load and return the validation schema from disk.

    Args:
        schema_path: The path to schema_validation.json.

    Returns:
        The parsed schema dict.
    """
    current_run.log_info(f"Loading schema from: {schema_path}")
    with schema_path.open(encoding="utf-8") as schema_file:
        return json.load(schema_file)


# ---------------------------------------------------------------------------
# Report + extraction guide
# ---------------------------------------------------------------------------


def _build_report(
    workbook,
    schema: dict,
    config: dict,
) -> tuple[dict, dict]:
    """Validate all eligible sheets and return (report, extraction_guide).

    Args:
        workbook: An openpyxl ``Workbook`` (data_only mode).
        schema: The validation schema dict.
        config: The label_comparison config dict.

    Returns:
        A tuple of (report_dict, extraction_guide_dict).
    """
    report = {
        "generated_at": datetime.now().isoformat(),
        "schema_version": schema.get("version", "unknown"),
        "sheets": {},
        "summary": {
            "total_sheets_validated": 0,
            "total_sheets_skipped": 0,
            "skipped_sheet_names": [],
            "total_errors": 0,
            "total_warnings": 0,
            "total_info": 0,
        },
    }

    extraction_guide = {
        "generated_at": datetime.now().isoformat(),
        "schema_version": schema.get("version", "unknown"),
        "sheets": {},
    }

    # Merge hardcoded skip list with schema-defined skip_sheets
    skip_sheets = set(SKIP_SHEETS)
    for sheet_name in schema.get("schema_meta", {}).get("skip_sheets", []):
        skip_sheets.add(sheet_name.strip().lower())

    for sheet_name in workbook.sheetnames:
        if sheet_name.strip().lower() in skip_sheets:
            report["summary"]["total_sheets_skipped"] += 1
            report["summary"]["skipped_sheet_names"].append(sheet_name)
            current_run.log_info(f"Skipping sheet: '{sheet_name}'")
            continue

        current_run.log_info(f"Validating sheet: '{sheet_name}'")
        worksheet = workbook[sheet_name]

        # Fingerprint check
        fingerprint = _get_fingerprint(schema)
        if fingerprint:
            fingerprint_value = read_cell(worksheet, fingerprint["cell"])
            if fingerprint_value is None or not text_matches(
                str(fingerprint_value),
                fingerprint["expected"],
                fingerprint.get("known_variants", []),
                config,
            ):
                report["summary"]["total_sheets_skipped"] += 1
                report["summary"]["skipped_sheet_names"].append(sheet_name)
                current_run.log_info(
                    f"Sheet '{sheet_name}' does not match expected format. Skipping.",
                )
                continue

        # Validate sheet
        sheet_issues, column_map, module_columns = _validate_sheet(
            worksheet, schema, config
        )

        # Find data rows
        data_start_row = _get_data_start_row(schema)
        data_rows = find_data_rows(worksheet, data_start_row)

        # Detect grouped column structures for extraction guide
        grouped_info = _detect_grouped_columns(worksheet, schema)

        # Build extraction guide for this sheet
        guide_builder = ExtractionGuideBuilder(
            schema, column_map, module_columns, data_rows, grouped_info,
        )
        extraction_guide["sheets"][sheet_name] = guide_builder.build()

        # Aggregate issue counts
        error_count = sum(1 for i in sheet_issues if i["severity"] == "error")
        warning_count = sum(1 for i in sheet_issues if i["severity"] == "warning")
        info_count = sum(1 for i in sheet_issues if i["severity"] == "info")

        report["sheets"][sheet_name] = {
            "issues": sheet_issues,
            "counts": {
                "errors": error_count,
                "warnings": warning_count,
                "info": info_count,
                "total": len(sheet_issues),
            },
            "valid": error_count == 0,
        }

        report["summary"]["total_sheets_validated"] += 1
        report["summary"]["total_errors"] += error_count
        report["summary"]["total_warnings"] += warning_count
        report["summary"]["total_info"] += info_count

    return report, extraction_guide


# ---------------------------------------------------------------------------
# Schema introspection helpers
# ---------------------------------------------------------------------------


def _get_fingerprint(schema: dict) -> dict | None:
    """Get the fingerprint definition for sheet matching.

    Checks ``schema_meta.fingerprint`` first, then falls back to the
    first module title in ``label_fields``.

    Args:
        schema: The validation schema dict.

    Returns:
        A dict with ``cell`` and ``expected`` keys, or None.
    """
    fingerprint = schema.get("schema_meta", {}).get("fingerprint")
    if fingerprint and isinstance(fingerprint, dict) and "cell" in fingerprint:
        return fingerprint

    label_fields = schema.get("label_fields", {})
    for module_key, module_definition in label_fields.items():
        if module_key.startswith("_"):
            continue
        if isinstance(module_definition, dict) and "title" in module_definition:
            title = module_definition["title"]
            if isinstance(title, dict) and "cell" in title:
                return title
    return None


def _get_data_start_row(schema: dict) -> int:
    """Extract the data start row from the schema.

    Handles both flat-key and sections-based schemas.

    Args:
        schema: The validation schema dict.

    Returns:
        The 1-based data start row number.
    """
    # Sections-based schema
    if "sections" in schema and "table_definitions" not in schema:
        for section in schema.get("sections", []):
            if section.get("layout") == "records":
                return section.get("data", {}).get("start_row", 2)
        return 2

    # Flat-key schema
    records_definition = schema.get("table_definitions", {}).get("records_data", {})
    return records_definition.get("data_row_start", 17)


def _has_sections(schema: dict) -> bool:
    """Check whether the schema uses sections-based format.

    Args:
        schema: The validation schema dict.

    Returns:
        True if sections-based (has ``sections``, no ``label_fields``).
    """
    return "sections" in schema and "label_fields" not in schema


# ---------------------------------------------------------------------------
# Validation dispatch
# ---------------------------------------------------------------------------


def _validate_sheet(
    worksheet,
    schema: dict,
    config: dict,
) -> tuple[list[dict], dict, dict]:
    """Run all validation passes on a single sheet.

    Instantiates the appropriate validators based on the schema format
    and returns the combined issues plus column mappings.

    Args:
        worksheet: An openpyxl ``Worksheet``.
        schema: The validation schema dict.
        config: The label_comparison config dict.

    Returns:
        A tuple of (issues, column_map, module_columns).
    """
    issues: list[dict] = []

    if _has_sections(schema):
        # Sections-based schema — use RecordsValidator in sections mode
        records_sections = [
            section
            for section in schema.get("sections", [])
            if section.get("layout") == "records"
        ]
        records_validator = RecordsValidator(
            worksheet, config, sections=records_sections,
        )
        issues.extend(records_validator.validate_structure())
        issues.extend(records_validator.validate_data())
        return issues, records_validator.column_map, {}

    # Flat-key schema — use all three validator types
    # 1. KeyValueValidator for labels and single-cell values
    key_value_validator = KeyValueValidator(
        worksheet,
        config,
        label_fields=schema.get("label_fields", {}),
        value_fields=schema.get("value_fields", {}),
    )
    issues.extend(key_value_validator.validate_structure())
    issues.extend(key_value_validator.validate_data())

    # 2. GroupedKeyValueValidator for columns × rows entries
    grouped_definitions = _extract_grouped_definitions(schema)
    if grouped_definitions:
        grouped_validator = GroupedKeyValueValidator(
            worksheet, config, grouped_definitions,
        )
        issues.extend(grouped_validator.validate_data())

    # 3. RecordsValidator for headers and data rows
    records_validator = RecordsValidator(
        worksheet,
        config,
        header_modules=schema.get("header_modules", {}),
        table_definitions=schema.get("table_definitions", {}),
    )
    issues.extend(records_validator.validate_structure())
    issues.extend(records_validator.validate_data())

    return issues, records_validator.column_map, records_validator.module_columns


def _extract_grouped_definitions(schema: dict) -> list[dict]:
    """Extract grouped field definitions from value_fields.

    Grouped entries are identified structurally by the presence of both
    ``columns`` and ``rows`` keys.

    Args:
        schema: The validation schema dict.

    Returns:
        A list of grouped field definition dicts.
    """
    grouped: list[dict] = []
    value_fields = schema.get("value_fields", {})

    for section_key, section in value_fields.items():
        if section_key.startswith("_") or not isinstance(section, dict):
            continue
        for field_key, field_definition in section.items():
            if not isinstance(field_definition, dict):
                continue
            if "columns" in field_definition and "rows" in field_definition:
                grouped.append(field_definition)

    return grouped


# ---------------------------------------------------------------------------
# Grouped column detection (for extraction guide)
# ---------------------------------------------------------------------------


def _detect_grouped_columns(worksheet, schema: dict) -> dict:
    """Detect grouped column structures from value_fields.

    Returns a dict mapping section_key to column activity information.
    Used by the extraction guide builder.

    Args:
        worksheet: An openpyxl ``Worksheet``.
        schema: The validation schema dict.

    Returns:
        A dict mapping section_key to ``{columns, active_columns, anchor_row}``.
    """
    grouped_info: dict = {}
    value_fields = schema.get("value_fields")
    if not value_fields:
        return grouped_info

    for section_key, section in value_fields.items():
        if not isinstance(section, dict):
            continue
        for field_key, field_definition in section.items():
            if not isinstance(field_definition, dict):
                continue
            if "columns" not in field_definition or "rows" not in field_definition:
                continue

            section_key_clean = (
                field_key.removesuffix("_values")
                if field_key.endswith("_values")
                else field_key
            )
            columns = field_definition["columns"]
            rows_definition = field_definition.get("rows", {})

            # Find anchor row
            anchor_row = None
            for row_key, row_def in rows_definition.items():
                if row_def.get("required_rule") == "anchor":
                    anchor_row = row_def.get("row")
                    break

            active_columns: list[str] = []
            if anchor_row:
                for column in columns:
                    value = read_cell_by_row_column(worksheet, anchor_row, column)
                    if value is not None and str(value).strip():
                        active_columns.append(column)

            grouped_info[section_key_clean] = {
                "columns": columns,
                "active_columns": active_columns,
                "anchor_row": anchor_row,
            }

    return grouped_info


# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------


def _save_json(data: dict, output_path: Path, label: str) -> None:
    """Write a dict as JSON and register it as a pipeline output.

    Args:
        data: The dict to serialize.
        output_path: The destination file path.
        label: A human-readable label for logging.
    """
    current_run.log_info(f"Saving {label} to: {output_path}")
    with output_path.open("w", encoding="utf-8") as output_file:
        json.dump(data, output_file, ensure_ascii=False, indent=2)
    current_run.add_file_output(str(output_path))


def _log_summary(report: dict) -> None:
    """Log a human-readable summary of the validation report.

    Args:
        report: The complete validation report dict.
    """
    summary = report["summary"]
    current_run.log_info(
        f"Validation complete — "
        f"{summary['total_sheets_validated']} sheets validated, "
        f"{summary['total_sheets_skipped']} skipped"
    )
    if summary["total_errors"] > 0:
        current_run.log_warning(
            f"Found {summary['total_errors']} errors across all sheets"
        )
    if summary["total_warnings"] > 0:
        current_run.log_warning(
            f"Found {summary['total_warnings']} warnings requiring confirmation"
        )
    if summary["total_errors"] == 0 and summary["total_warnings"] == 0:
        current_run.log_info(
            "All sheets passed validation with no errors or warnings."
        )


if __name__ == "__main__":
    xls_validation()
