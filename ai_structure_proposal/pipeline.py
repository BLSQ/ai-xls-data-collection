"""OpenHEXA Pipeline: Excel Structure Proposal via Gemini AI.

Analyzes an Excel template's visual and structural layout using Gemini's
multimodal capabilities.  Produces a ``structure_proposal.json`` describing
every section found: metadata key-value pairs, grouped key-value tables,
and record-based data tables.

The structure proposal is designed to be reviewed/edited by a human
(or a webapp) before being consumed by a deterministic schema generator.

This file is the thin orchestrator — all domain logic lives in:

- :mod:`excel_reader`          — workbook loading and text representation
- :mod:`sheet_renderer`        — PIL-based image rendering
- :mod:`gemini_client`         — Gemini API calls (secure header-based auth)
- :mod:`prompt_builder`        — prompt assembly from separated concerns
- :mod:`response_parser`       — JSON extraction from LLM output
- :mod:`validation_extractor`  — deterministic data-validation resolution
"""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

from openhexa.sdk import (
    CustomConnection,
    File,
    current_run,
    parameter,
    pipeline,
    workspace,
)

from excel_reader import ExcelReader
from gemini_client import DEFAULT_MODEL, GeminiClient
from prompt_builder import build_full_prompt
from response_parser import parse_structure_json
from sheet_renderer import SheetRenderer
from validation_extractor import ValidationExtractor

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

OUTPUT_FILENAME = "structure_proposal.json"


# ---------------------------------------------------------------------------
# Pipeline definition
# ---------------------------------------------------------------------------


@pipeline("ai-structure-proposal", timeout=3600)
@parameter(
    "excel_file",
    name="Excel File",
    type=File,
    required=True,
    help="The Excel template to analyse.",
)
@parameter(
    "gemini_connection",
    name="Gemini API Connection",
    type=CustomConnection,
    required=True,
    help="Custom connection containing the Gemini API key (field: api_key).",
)
@parameter(
    "sheet_name",
    name="Sheet Name",
    type=str,
    required=False,
    help="Specific sheet to analyse.  If empty, auto-selects the first data sheet.",
)
@parameter(
    "user_guidelines",
    name="User Guidelines",
    type=str,
    required=False,
    help=(
        "Free-text instructions to guide the analysis. "
        "E.g.: 'Rows 8-12 are a donor/funder group, not simple key-value. "
        "Columns after column Y contain province-level percentages that must "
        "sum to 100%. The primary key column is the budget line label in column B.'"
    ),
)
def ai_structure_proposal(
    excel_file: File,
    gemini_connection: CustomConnection,
    sheet_name: str = None,
    user_guidelines: str = None,
):
    """Analyse an Excel template and produce a structure_proposal.json."""
    current_run.log_info(f"Starting structure analysis of: {excel_file}")

    # -- Load workbook --
    reader = ExcelReader(excel_file.path)
    current_run.log_info("Workbook loaded (values + full modes)")

    # -- Select sheet --
    worksheet_values, worksheet_full = reader.select_sheets(sheet_name)
    current_run.log_info(f"Analysing sheet: '{worksheet_values.title}'")

    # -- Build text representation --
    current_run.log_info("Building text representation...")
    text_representation = reader.build_text_representation(
        worksheet_values, worksheet_full
    )
    current_run.log_info(f"Text representation: {len(text_representation)} chars")

    # -- Render image --
    current_run.log_info("Rendering sheet image...")
    renderer = SheetRenderer()
    image_bytes = renderer.render(worksheet_values)
    if image_bytes:
        current_run.log_info(f"Image rendered: {len(image_bytes)} bytes")
    else:
        current_run.log_warning(
            "PIL not available — proceeding with text-only analysis"
        )

    # -- Extract data validations deterministically --
    current_run.log_info("Extracting data validations (deterministic)...")
    extractor = ValidationExtractor(worksheet_full, reader.workbook_full)
    resolved_validations = extractor.extract()
    current_run.log_info(
        f"Found {len(resolved_validations)} data validation rules"
    )
    reader.workbook_full.close()

    # -- Build and send prompt --
    if user_guidelines:
        current_run.log_info(f"User guidelines: {user_guidelines[:200]}")
    prompt = build_full_prompt(
        text_representation, worksheet_values.title, user_guidelines
    )

    current_run.log_info("Sending to Gemini for analysis...")
    client = GeminiClient(api_key=gemini_connection.api_key)
    raw_response = client.generate(prompt, image_bytes)

    # -- Parse response --
    current_run.log_info("Parsing Gemini response...")
    try:
        proposal = parse_structure_json(raw_response)
    except (json.JSONDecodeError, ValueError) as exc:
        current_run.log_warning(f"JSON parse failed: {exc}")
        current_run.log_info(
            f"Raw response (first 2000 chars):\n{raw_response[:2000]}"
        )
        raw_path = Path(workspace.files_path) / "structure_proposal_raw.txt"
        raw_path.write_text(raw_response, encoding="utf-8")
        current_run.add_file_output(str(raw_path))
        raise RuntimeError(
            "Gemini returned a response that could not be parsed as JSON. "
            "Raw response saved to structure_proposal_raw.txt"
        ) from exc

    # -- Merge deterministic validations --
    if resolved_validations:
        proposal["data_validations_resolved"] = resolved_validations
        current_run.log_info(
            f"Merged {len(resolved_validations)} deterministic validation "
            "rules into proposal"
        )

    # -- Add generation metadata --
    proposal["_generation"] = {
        "generated_at": datetime.now().isoformat(),
        "model": DEFAULT_MODEL,
        "source_file": str(excel_file.path),
        "prompt_chars": len(prompt),
        "image_bytes": len(image_bytes) if image_bytes else 0,
    }

    # -- Save output --
    output_path = Path(workspace.files_path) / OUTPUT_FILENAME
    current_run.log_info(f"Saving structure proposal to: {output_path}")
    with output_path.open("w", encoding="utf-8") as output_file:
        json.dump(proposal, output_file, ensure_ascii=False, indent=2)
    current_run.add_file_output(str(output_path))

    # -- Summary --
    section_count = len(proposal.get("sections", []))
    layout_counts: dict[str, int] = {}
    for section in proposal.get("sections", []):
        layout = section.get("layout", "unknown")
        layout_counts[layout] = layout_counts.get(layout, 0) + 1
    current_run.log_info(
        f"Structure proposal complete: {section_count} sections "
        f"({', '.join(f'{count} {layout}' for layout, count in layout_counts.items())})"
    )

    reader.close()


if __name__ == "__main__":
    ai_structure_proposal()
