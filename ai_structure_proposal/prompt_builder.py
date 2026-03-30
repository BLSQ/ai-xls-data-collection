"""Prompt construction for the Gemini structure-analysis request.

The prompt is assembled from four distinct concerns:

1. **System context** — what the LLM is and how to reason about spreadsheets.
2. **Output format** — the JSON schema the LLM must produce.
3. **User guidelines** — optional domain hints from the admin.
4. **Sheet data** — the actual text representation and metadata.

Each concern is a separate constant or function, making it easy to update
one part without touching the others.
"""


# ---------------------------------------------------------------------------
# 1. SYSTEM CONTEXT — role and layout-pattern definitions
# ---------------------------------------------------------------------------

SYSTEM_CONTEXT = """\
You are a spreadsheet structure analyst.  Given the cell data, merged cell
ranges, styles, and data validations from an Excel worksheet, identify
every distinct SECTION and classify its layout.

## The Three Layout Patterns

### 1. `key_value`  — Simple label -> value metadata
A cell contains a label (usually bold, often ending with ":"), and the
adjacent cell (to the right, or across a merged range) contains a value.
These are typically found in the upper part of a sheet, above data tables.

Example:
```
Row 3: B="Nom du projet:" [bold] | C="PROSANTE" [merged C3:I3]
Row 4: B="Date début:" [bold]    | C="2023-01-01"
```

### 2. `grouped_key_value`  — Repeating column group (transposed table)
A label cell is followed by N element columns (e.g., "Bailleur 1",
"Bailleur 2", ...), and below them are property rows with a value per
element.  This is a TABLE embedded in metadata, but oriented
column-per-entity rather than row-per-record.

How to recognise:
- A header row with a group label + N numbered/named element columns
- Below it: property rows where each element column has a value
- The number of elements may vary (some may be empty)

Example:
```
Row 8:  B="Bailleur(s)" [bold] | C="Bailleur 1" | D="Bailleur 2" | ...
Row 9:  B="Nom"                | C="USAID"      | D="Gavi"       | ...
Row 10: B="Montant"            | C=500000       | D=300000       | ...
```

### 3. `records`  — Data table with column headers + data rows
One or two header rows define columns, followed by many data rows.
Headers may have GROUPS: a parent header merged across several columns,
with sub-column headers in the row below.
Some columns may be DYNAMIC (variable count, e.g., year columns, province names).

How to recognise:
- Bold / coloured title row(s) spanning the full width -> module titles
- Below: column header row(s) with individual column names
- Below: data rows (numbers, text, dates)
- Data validations (dropdowns) often apply to entire data columns

Example:
```
Row 14: B:I merged "Module 3: Budget data" [bold, colored bg]
Row 15: B="Budget line" | C="Total" [merged C15:D15] | E="Province 1" | ...
Row 16:                 | C="Planned" | D="Actual"   |               | ...
Row 17: B="Vaccines"    | C=50000     | D=45000       | E=30%         | ...
```
"""

# ---------------------------------------------------------------------------
# 2. DETECTION GUIDELINES — how to identify sections and columns
# ---------------------------------------------------------------------------

DETECTION_GUIDELINES = """\
## How to identify sections

1. **Title rows**: Bold cells with coloured background spanning multiple
   columns are section/module titles.  They mark section boundaries.
2. **Metadata vs. data**: Above the first record-table title row is
   metadata (key_value or grouped_key_value).  Below it is data (records).
3. **Grouped key_value**: Look for a label + numbered/named element
   columns in the same row, with property rows below.
4. **Records columns** come in two types:
   - `"type": "single"` — a standalone column with one header in one row.
     Produces a flat indicator name: `canonical_name`.
   - `"type": "group"` — a parent header (row 15, merged across columns)
     with sub-column headers (row 16).  Produces concatenated indicator
     names: `parent_canonical_name + "_" + sub_canonical_name`.
   A single module CAN MIX both types.  For example, Module 8 has three
   single columns (NA, National, Central) followed by a group column
   (Provincial → province sub-columns).
   Within a group's sub_columns, each sub is either `"fixed": true`
   (known, stable name) or `"fixed": false` (dynamic — the count and
   names may vary across template versions, e.g., province names).
5. **Data validations**: Dropdowns on data columns indicate enum constraints.

## Important guidelines

- Detect ALL sections, including ones that span only a few rows.
- For `records` sections, carefully identify: which row(s) are headers,
  which row is the first data row, and the column structure.
- `headers.rows` is PER-SECTION.  If a module has ONLY `"type": "single"`
  columns, set `"rows": [15]` (one row).  If the module has ANY
  `"type": "group"` column (parent header + sub-columns), include both
  rows: `"rows": [15, 16]`.  A module CAN mix single and group columns.
- For group headers in records, identify the parent and its sub-columns.
- For dynamic columns (years, provinces, categories), note the pattern.
- Assign a short snake_case `key` to each section and field.
- `canonical_name` for record columns should be snake_case of the header.
- Include the `area` (row and column range) for each section.
- For value types, infer from example values: "string", "number", "date",
  "percentage".
- Note data validations that apply to columns (enum constraints).

## Using the image

An image of the spreadsheet is attached. Use it to:
- Confirm section boundaries where the visual layout shifts (color bands,
  spacing, or density changes between regions).
- Identify grouped columns that share a visual pattern (same background
  color, similar formatting) even if their names don't follow a numbered
  convention.
- Distinguish metadata regions (sparse, label-value pairs) from data
  tables (dense grids) when the text representation is ambiguous.
Do not rely solely on the image — the text representation is authoritative
for exact cell values and references.
"""

# ---------------------------------------------------------------------------
# 3. OUTPUT FORMAT — the JSON schema the LLM must produce
# ---------------------------------------------------------------------------

OUTPUT_FORMAT = """\
## Output format

Produce a single JSON object with this structure:

```json
{
  "version": "2.0.0",
  "source_sheet": "<sheet name>",
  "schema_meta": {
    "fingerprint": {
      "cell": "<cell ref of the first section title>",
      "expected": "<title text>"
    },
    "locale": "<detected language: fr, en, etc.>",
    "skip_sheets": ["<sheet names that look like reference/guide sheets>"]
  },
  "structure_landmarks": {
    "metadata_area": {"start_row": 1, "end_row": 13},
    "title_row": 14,
    "header_rows": [15, 16],
    "data_start_row": 17,
    "last_data_col": "BF"
  },
  "sections": [
    {
      "key": "<snake_case section name>",
      "title": {
        "cell": "<cell ref>",
        "expected": "<title text>"
      },
      "layout": "key_value",
      "area": {"start_row": N, "end_row": N, "start_col": "B", "end_col": "I"},
      "fields": [
        {
          "key": "<snake_case>",
          "label": {"cell": "<ref>", "expected": "<label text>"},
          "value": {
            "cell": "<ref>",
            "type": "string|number|date",
            "required": true|false,
            "validation": {"enum": ["...", "..."]}
          }
        }
      ]
    },
    {
      "key": "<snake_case>",
      "title": {"cell": "<ref>", "expected": "<text>"},
      "layout": "grouped_key_value",
      "area": {"start_row": N, "end_row": N, "start_col": "B", "end_col": "I"},
      "elements": {
        "header_row": N,
        "columns": ["C", "D", "E"],
        "labels": ["Element 1", "Element 2", "..."],
        "count_is_dynamic": true
      },
      "fields": [
        {
          "key": "<snake_case>",
          "label": {"row": N, "col": "B", "expected": "<text>"},
          "type": "string|number|date",
          "required_rule": "always|first_required|if_element_active"
        }
      ]
    },
    {
      "key": "<snake_case>",
      "title": {"cell": "<ref>", "expected": "<module title text>"},
      "layout": "records",
      "area": {"start_row": N, "end_row": "dynamic", "start_col": "B", "end_col": "I"},
      "headers": {
        "rows": [15, 16],
        "primary_key": "<canonical_name of the required column>",
        "columns": [
          {
            "type": "single",
            "column": "B",
            "header_row": 15,
            "expected": "<header text>",
            "canonical_name": "<snake_case>",
            "value_type": "string|number|date|percentage",
            "required": true|false
          },
          {
            "type": "group",
            "parent": {
              "header_row": 15,
              "expected": "<parent header text>",
              "canonical_name": "<snake_case of group>",
              "start_col": "C",
              "end_col": "F"
            },
            "sub_columns": [
              {
                "fixed": true,
                "column": "C",
                "header_row": 16,
                "expected": "<sub-header text>",
                "canonical_name": "<snake_case>",
                "value_type": "number"
              },
              {
                "fixed": false,
                "header_row": 16,
                "pattern": "regex to match dynamic headers",
                "canonical_prefix": "<prefix>",
                "value_type": "number",
                "description": "what these dynamic columns represent"
              }
            ]
          }
        ]
      },
      "data": {
        "start_row": 17,
        "end_rule": "first_empty_row"
      }
    }
  ],
  "data_validations_summary": [
    {
      "cells": "<sqref>",
      "type": "list",
      "formula": "<formula1 value>",
      "resolved_values": ["val1", "val2"]
    }
  ],
  "validation_config": {
    "label_comparison": {
      "case_sensitive": false,
      "normalize_accents": true,
      "strip_whitespace": true
    }
  }
}
```
"""

# ---------------------------------------------------------------------------
# 4. FIELD EXPLANATIONS — key fields and their semantics
# ---------------------------------------------------------------------------

FIELD_EXPLANATIONS = """\
## Key fields explained

- **`structure_landmarks`**: Global row/column boundaries. `metadata_area`
  covers key_value and grouped_key_value sections. `title_row` is the row
  where record module titles are (bold colored cells spanning columns).
  `header_rows` are the column header row(s) for data tables.
  `data_start_row` is where records begin. `last_data_col` is the
  rightmost column with actual data.

- **`schema_meta.fingerprint`**: The cell + text used to verify whether a
  submitted workbook matches this template (usually the first module title).

- **`data_validations_summary`**: List the validations from the DATA
  VALIDATIONS section. Set `resolved_values` to null if you cannot
  determine the actual values (named ranges, INDIRECT formulas) — the
  deterministic pipeline will resolve them later.

- **`headers.primary_key`**: The canonical_name of the column that must
  not be empty for a data row to be valid (usually the first/leftmost
  column — the "line item" identifier).

- **Column `type` (single vs group)**: Each column in `headers.columns`
  has `"type": "single"` or `"type": "group"`.  A module CAN mix both.
  - `single`: one header cell, one column.  Indicator = `canonical_name`.
  - `group`: merged parent header spanning N sub-columns.  Each sub has
    `"fixed": true` (stable name) or `"fixed": false` (dynamic — name/count
    varies across template versions).  Indicator = `parent_canonical + "_" + sub_canonical`.
  Example: Module 8 has single columns (NA, National, Central) then a
  group (Provincial → 26 province sub-columns marked `"fixed": false` \
if the provinces can change from one submission to another.).

Now analyse the following spreadsheet and produce the structure_proposal JSON.
"""


# ---------------------------------------------------------------------------
# Public API — prompt assembly
# ---------------------------------------------------------------------------


def build_full_prompt(
    text_representation: str,
    sheet_name: str,
    user_guidelines: str | None = None,
) -> str:
    """Assemble the complete LLM prompt from all concerns.

    Combines:
    - System context (role + layout patterns)
    - Detection guidelines (how to identify sections)
    - Output format (JSON schema)
    - Field explanations
    - Optional user guidelines
    - The actual sheet data

    Args:
        text_representation: Rich text dump of the worksheet (from
            :meth:`ExcelReader.build_text_representation`).
        sheet_name:          Name of the worksheet being analysed.
        user_guidelines:     Optional free-text hints from the admin.

    Returns:
        The complete prompt string ready to send to the Gemini API.
    """
    sections: list[str] = [
        SYSTEM_CONTEXT,
        DETECTION_GUIDELINES,
        OUTPUT_FORMAT,
        FIELD_EXPLANATIONS,
    ]

    if user_guidelines and user_guidelines.strip():
        sections.append(_build_user_guidelines_section(user_guidelines))

    sections.append(_build_sheet_data_section(text_representation, sheet_name))

    return "\n".join(sections)


def _build_user_guidelines_section(user_guidelines: str) -> str:
    """Format the optional user guidelines as a prompt section.

    Args:
        user_guidelines: Free-text instructions from the admin.

    Returns:
        Formatted guidelines block ready for inclusion in the prompt.
    """
    return (
        "=== USER GUIDELINES ===\n"
        "The user provided the following domain-specific instructions. "
        "Use them to improve your analysis — they take priority over "
        "your own heuristics when there is a conflict.\n"
        "\n"
        f"{user_guidelines.strip()}\n"
    )


def _build_sheet_data_section(text_representation: str, sheet_name: str) -> str:
    """Format the sheet data as the final prompt section.

    Args:
        text_representation: The text dump of the worksheet.
        sheet_name:          The name of the sheet.

    Returns:
        Formatted data block including the action instruction.
    """
    return (
        f"=== SPREADSHEET DATA FOR SHEET '{sheet_name}' ===\n"
        "\n"
        f"{text_representation}\n"
        "\n"
        "=== END OF DATA ===\n"
        "\n"
        "Produce the structure_proposal JSON for this sheet.  "
        "Be thorough — identify every section, every field, every column."
    )
