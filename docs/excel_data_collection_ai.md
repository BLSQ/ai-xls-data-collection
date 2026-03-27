# AEDES — Excel Data Collection & Validation Workflow

A step-by-step guide for administrators and organisations using the AEDES resource-mapping platform. This document covers the entire lifecycle: from teaching the system a new Excel template, to collecting validated data, to publishing indicators on a public dashboard.

> For the full schema field reference (`required`, `required_rule`, JsonLogic rules, etc.), see [`data_structures_reference.md`](data_structures_reference.md).

---

## Overview

The workflow is a **five-stage pipeline chain** orchestrated through OpenHEXA. Each stage produces an artifact consumed by the next:

```
  ┌──────────────────────┐
  │  1. Structure        │   Gemini AI analyses the Excel template layout
  │     Proposal (AI)    │──▶ structure_proposal.json
  └──────────────────────┘
             │
  ┌──────────────────────┐
  │  2. Schema           │   Deterministic pipeline converts the proposal
  │     Generation       │──▶ schema_validation.json
  └──────────────────────┘
             │
  ┌──────────────────────┐
  │  3. Admin: Review    │   Admin tweaks rules, required fields, types
  │     (Private Webapp) │──▶ schema_validation.json (edited)
  └──────────────────────┘
             │
  ┌──────────────────────┐
  │ 4. Viewer: Validation│   Organisation uploads their filled Excel
  │     & Ingestion      │──▶ validation_report.json → program_data (DB)
  └──────────────────────┘
             │
  ┌──────────────────────┐
  │  5. Indicators       │   Scheduled pipeline computes indicators
  │     & Dashboard      │──▶ parquet files → Public dashboard
  └──────────────────────┘
```

**Stages 1–4** are template-agnostic — they work with any Excel layout. **Stage 5** (indicators & dashboard) contains domain-specific logic tailored to the current AEDES use case.

| Stage | Pipeline code | Input | Output |
|-------|--------------|-------|--------|
| 1. Structure proposal | `ai-structure-proposal` | Excel template + Gemini API key | `structure_proposal.json` |
| 2. Schema generation | `schema-validation` | `structure_proposal.json` + Excel template | `schema_validation.json` |
| 3. Admin review | Private webapp (`index_ai.html`) | `schema_validation.json` | Edited schema (same file) |
| 4a. Validation | `xls-validation` | Filled Excel + `schema_validation.json` | `validation_report.json` + `extraction_guide.json` |
| 4b. Ingestion | `xls-ingest` | Filled Excel + `extraction_guide.json` | `program_metadata` and `program_data` DB tables |
| 5. Indicators | `aedes-indicators` | `program_data` DB table | `indicators` DB table + parquet files → public dashboard |

---

## Key Concepts

These terms appear throughout the document. Read this section first for easier understanding.

### Canonical names

Every Excel header and field is converted to a deterministic **canonical name** — a lowercase, ASCII-only, underscore-separated identifier. This name is used consistently across the schema, extraction guide, database columns, and indicator pipeline.

The conversion (`to_column_canonical()`) applies these steps:

1. Strip parenthetical descriptions: `"Niveau central (Directions, Programmes, …)"` → `"Niveau central"`
2. NFKD Unicode normalisation: strip accents
3. Expand ligatures: `œ` → `oe`, `æ` → `ae`
4. Replace non-alphanumeric characters with `_`
5. Collapse consecutive underscores, strip leading/trailing, lowercase

**Group sub-columns** are always prefixed with their parent group's canonical name:
- `"Réalisé" > "2024"` → `realise_2024`
- `"Prévisionnel" > "Total Ligne"` → `previsionnel_total_ligne`

**Collision handling**: if two modules produce the same canonical name (e.g., both have an "NA" column), the pipeline automatically prefixes them with the module key: `module_6_na`, `module_8_na`.

Examples:
- `"Intitulé budgétaire / Libellé d'activité"` → `intitule_budgetaire_libelle_d_activite`
- `"Mode de mise en œuvre"` → `mode_de_mise_en_oeuvre` (ligature expanded)
- `"Précision sur le piler"` → `precision_sur_le_piler` (typo preserved — the system does not correct spelling)

### Section types

The AI classifies each region of the Excel template as one of three types: **`key_value`** (single metadata fields), **`grouped_key_value`** (repeating column structures), or **`records`** (data tables with headers and rows). See Stage 1 below for how each type is detected.

### Fingerprint matching

The validation pipeline uses a **fingerprint** to determine whether a sheet matches the expected template. The fingerprint is a specific cell value (e.g., a module title like "Module 1 : Informations générales") defined in the schema. Sheets that don't match are skipped, allowing workbooks with mixed content.

### JsonLogic rules

[JsonLogic](https://jsonlogic.com/) is a portable rule format used for validation rules. It allows the admin to define conditions declaratively — for example, `{"==": [{"var": "value"}, {"+":[{"var": "realise_2024"}, {"var": "realise_2025"}]}]}` means "this cell must equal the sum of two other columns." Rules are attached to fields in the schema and evaluated at validation time.

### Private vs. Public

- The **private webapp** (`index_ai.html`) is deployed as an OpenHEXA webapp with restricted access. Only authorised users (admin and partner organisations) can upload files, run validations, and edit the schema.
- The **public dashboard** (`aedes_indicators/index.html`) is a separate OpenHEXA webapp. It reads only from pre-computed parquet files and exposes no write operations or sensitive data.

---

## Stage 1 — Structure Proposal (AI)

**Who**: Administrator (one-time setup per template design)
**Pipeline**: `ai-structure-proposal`
**Code**: `ai_structure_proposal/pipeline.py`

### What happens

The pipeline converts the Excel template into a **text representation** and sends it to Gemini along with a **rendered screenshot** of the sheet (generated via Pillow, falls back to text-only if unavailable). The text representation includes, for every cell: its value, style annotations (`[bold]`, `[bg:#RRGGBB]`, `[fg:#RRGGBB]`), merged cell ranges (`[merged:A1:C3]`), plus a summary of data validation rules (dropdown lists) and non-default column widths.

Both inputs are sent together as a single **multimodal request** (text + inline base64 PNG). The text is the authoritative source for exact cell values and references, while the image provides the **visual gestalt** — the spatial layout, colour bands between sections, density shifts from sparse metadata to dense data grids, and how merged parent headers visually group sub-columns. This redundancy helps Gemini disambiguate edge cases where the text alone is ambiguous (e.g., distinguishing a few key-value rows from the start of a data table, or detecting grouped columns that share formatting but don't follow a numbered naming convention).

Gemini returns a JSON structure classifying each detected region as one of three section types:

#### `key_value` — single metadata fields

A **bold label** (often ending with `:`) with an adjacent value cell (to the right, or in a merged range). Typically found in the upper part of the sheet above data tables. Examples: project name, start date, currency.

#### `grouped_key_value` — repeating column structures

A label column with **N numbered or named element columns** (e.g., "Bailleur 1", "Bailleur 2", "Bailleur 3"). Property rows below hold values for each element — essentially a transposed table where each column is an entity rather than each row. Visual signals: repeating numbered columns, bold group label spanning the header.

#### `records` — data tables

**Bold/coloured title rows** spanning the full width mark module boundaries. Below them, one or two **header rows** define columns (single columns or grouped columns with a merged parent header and sub-headers). Data rows follow with numbers, text, dates, or dropdown values. **Data validations** on entire columns indicate enum constraints. Columns can be:
- `"type": "single"` — standalone column with one header cell.
- `"type": "group"` — parent header (merged across columns) with individual sub-headers beneath.

### Parameters

| Parameter | Required | Description |
|-----------|----------|-------------|
| `excel_file` | Yes | The blank Excel template (the reference design, not a filled-in file) |
| `gemini_connection` | Yes | OpenHEXA custom connection containing a Gemini API key (field: `api_key`) |
| `sheet_name` | No | Specific sheet to analyse (auto-selects the first data sheet if empty) |
| `user_guidelines` | No | Free-text hints for Gemini, e.g., *"Rows 8–12 are a donor group, not simple key-value"* |

### Output

`structure_proposal.json` — saved to the workspace files. This file describes the template's layout in a structured format. It is designed to be **human-readable and editable** before proceeding to stage 2.

### Tips

- Provide `user_guidelines` if the template has unusual structures that Gemini might misinterpret (e.g., a grouped donor table that looks like regular key-value pairs).
- Review the proposal JSON before proceeding. Check that the section boundaries, column assignments, and `primary_key` values make sense.

---

## Stage 2 — Schema Generation (Deterministic)

**Who**: Administrator (one-time, or after editing the structure proposal)
**Pipeline**: `schema-validation`
**Code**: `schema_validation/pipeline.py`

### What happens

This pipeline is **fully deterministic** — no AI calls. It reads the structure proposal and the original Excel template, then produces a `schema_validation.json` containing:

1. **Sections** — each detected region (key_value, grouped_key_value, records) with its cell references, expected labels, and validation rules.
2. **Column definitions** — for each records section: header text, canonical name, value type, position (fixed, group, or dynamic), and data validations (dropdown lists resolved from the Excel file).
3. **Required fields** — each value field gets a `required` flag (default: `false`). For grouped_key_value sections, each row gets a `required_rule` (`"anchor"` for the row that determines active columns, `"if_active"` for rows that only validate active columns, or `"optional"`).
4. **Table definitions** — metadata table structure and records data table structure with primary key enforcement (`row_validation`).
5. **Validation config** — comparison settings (case sensitivity, accent normalisation, etc.).

All column names are computed as canonical names (see Key Concepts above).

### Parameters

| Parameter | Required | Description |
|-----------|----------|-------------|
| `structure_proposal_file` | Yes | The `structure_proposal.json` from stage 1 |
| `excel_file` | Yes | The same Excel template used in stage 1 |
| `sheet_name` | No | Override the sheet specified in the proposal |

### Output

`schema_validation.json` — the validation schema used by all downstream pipelines.

---

## Stage 3 — Admin Review (Private Webapp)

**Who**: Administrator
**Webapp**: `index_ai.html` (private, accessible only to authorised users)

### What the webapp does

The private webapp provides a visual interface for the full pipeline chain. It has three main functions:

#### a) Run pipelines

The webapp can trigger all four pipelines (structure proposal → schema generation → validation → ingestion) in sequence directly from the browser using the OpenHEXA GraphQL API. The pipeline orchestration panel shows real-time status for each step.

#### b) Review and edit the schema

After schema generation, the admin can:

- **Edit field properties**: mark fields as required or optional, change value types (string, number, date, percentage).
- **Add known variants**: alternative spellings for header labels that the validation should accept.
- **Add validation rules**: attach JsonLogic rules to cells (see Key Concepts), defining conditions like "this cell must equal the sum of these other cells" or "if column X = 'yes', then column Y is required".
- **Reorder columns**: drag-and-drop column entries within the schema to match the expected layout.
- **Preview the Excel mapping**: a colour-coded grid view shows which cells map to which schema fields (labels in yellow, values in white, headers in purple, data rows in green).

#### c) Validate uploaded Excel files

Organisations (or the admin on their behalf) upload a filled-in Excel file. The webapp runs the validation pipeline and displays results:

- **Errors** (red) — the file cannot be ingested until these are fixed.
- **Warnings** (amber) — headers that don't exactly match but are close enough (the admin can confirm or reject).
- **Info** (blue) — confirmations that fields matched as expected.

If validation passes with zero errors, a "Save to Database" button appears to trigger the ingestion pipeline.

### How it connects to OpenHEXA

The webapp communicates with the OpenHEXA backend through the GraphQL API (same-origin, cookie-authenticated). It:

- Uploads Excel files to the workspace bucket via presigned URLs.
- Triggers pipelines with specific parameters.
- Polls pipeline runs for completion.
- Downloads output files (validation report, schema, extraction guide) from the workspace.

### Configuration

Click the gear icon in the webapp header to configure:

- **Workspace slug** — the OpenHEXA workspace identifier (e.g., `mc-aedes-ressources-mapping`).
- **Pipeline codes** — the registered codes for each pipeline step (normally unchanged).
- **Webhook URLs** — optional, for notification integrations.

---

## Stage 4a — Validation

**Who**: Any organisation submitting data (via the webapp, or directly via OpenHEXA)
**Pipeline**: `xls-validation`
**Code**: `xls_validation/pipeline.py` + `xls_validation/validators.py`

### What happens

The validation pipeline reads the submitted Excel file and checks it against `schema_validation.json`:

1. **Sheet filtering** — skips sheets listed in `skip_sheets` (e.g., "Guide", "Observations") and sheets that don't match the expected fingerprint (see Key Concepts).
2. **Label validation** — checks that expected labels are present at their designated cells.
3. **Value validation** — checks metadata values (types, required fields, enum constraints). Fields marked `required: true` produce an error if empty; optional fields produce an info-level note.
4. **Header validation** — verifies column headers match expected text (with accent/case tolerance), builds a column map linking each canonical name to its Excel column letter.
5. **Data row validation** — enforces the primary key (`required: true` in `row_validation` produces an error if the PK column is empty), evaluates any JsonLogic rules attached to columns.

### Dynamic column discovery

For columns marked as `position: "dynamic"` in the schema, the validator scans consecutive cells matching a regex pattern. Each discovered column is assigned a canonical name using the same `to_canonical()` function as the schema generator, ensuring consistency.

### Output

- `validation_report.json` — per-sheet list of issues (error/warning/info), counts, and a valid/invalid flag.
- `extraction_guide.json` — maps each canonical column name to its Excel column letter and row range, used by the ingestion pipeline.

---

## Stage 4b — Ingestion

**Who**: Triggered after successful validation (via the webapp or directly)
**Pipeline**: `xls-ingest`
**Code**: `xls_ingest/pipeline.py`

### What happens

The ingestion pipeline reads the validated Excel file together with the `extraction_guide.json` and writes two database tables:

| Table | Content |
|-------|---------|
| `program_metadata` | One row per sheet — key-value pairs extracted from metadata sections (e.g., project name, organisation, currency, dates) |
| `program_data` | One row per Excel data row — all columns from records sections as database columns, using their canonical names |

The pipeline is fully generic: it reads section boundaries and column mappings from the extraction guide, with no assumptions about the template's domain or structure.

### Upsert behaviour

Each programme (Excel file + sheet) gets a deterministic UUID (`program_id`) computed from the file path and sheet name. When re-ingesting:

1. All existing rows with matching `program_id` values are **deleted**.
2. All new rows are **inserted**.

This means re-uploading the same file replaces the previous data entirely — there is no row-level merge.

### Schema evolution

If the new Excel has columns that don't exist in the database table yet, the pipeline automatically runs `ALTER TABLE` to add them (as TEXT or the appropriate PostgreSQL type, nullable). Existing columns are never removed or renamed.

---

## Stage 5 — Indicators & Dashboard (AEDES-specific)

**Who**: Scheduled pipeline (admin sets up the schedule in OpenHEXA)
**Pipeline**: `aedes-indicators`
**Code**: `aedes_indicators/pipeline.py`

> **Note**: Unlike Stages 1–4 which are template-agnostic, this stage contains domain-specific logic tailored to the current AEDES resource-mapping use case. Adapting it to a different domain requires modifying the indicator definitions and dashboard layout.

### What happens

The indicator pipeline reads the `program_data` table and computes strategic health indicators. Each indicator is defined by a **filter function** that selects matching rows based on column values (e.g., pilier, type de dépense, thématique).

For each indicator × programme, it produces:

- **Budget totals**: `budget_realise` (sum of realised amounts) and `budget_previsionnel` (sum of planned amounts).
- **Year breakdown**: sums per year column (e.g., `realise_2024`, `previsionnel_2026`).
- **Population distribution**: budget-weighted average of percentage columns (e.g., children <5, adolescents, pregnant women).
- **Care levels**: budget-weighted average across facility levels (e.g., central, provincial, operational, community).
- **Geographic distribution**: budget-weighted average across geographic columns (e.g., 26 DRC provinces discovered dynamically).

#### Budget-weighted distributions

Percentage columns (e.g., 40% Central, 60% Provincial) are aggregated as **budget-weighted averages**: if a budget line of 100,000€ is 60% Provincial and another of 50,000€ is 100% Central, the weighted distribution reflects where the money actually went, not just a simple average of the percentages.

### Indicator catalogue

The indicators are defined in the `INDICATORS` list in the pipeline code. Each entry specifies:

- `source_sheet` — which reference framework the indicator belongs to (e.g., PNDS, CSU, SRMNEA-NUT, LF Macro).
- `indicator_name` — the full strategic indicator label.
- `filter_fn` — a Python function that returns a boolean mask selecting matching rows.
- `notes` — a human-readable description of the filter logic.

A special `_TOTAL_` sentinel indicator matches all rows and is used for KPI totals and distribution charts (avoiding double-counting across indicators).

### Output

1. **`indicators` database table** — replaced entirely on each run.
2. **Parquet files** pushed to the public dashboard webapp via the OpenHEXA GraphQL API:
   - `indicators.parquet` — all indicator rows.
   - `program_metadata.parquet` — programme metadata for the dashboard filters.

### Public dashboard

**Code**: `aedes_indicators/index.html`
**Access**: Public — anyone with the URL can view the dashboard.

The dashboard reads the parquet files (using hyparquet in the browser) and renders KPI cards, budget breakdowns by indicator and framework, year trends, care level and population group distributions, geographic charts, and a programmes table. Filters allow drilling down by programme and by reference framework.

---

## File Map

```
precipitation_worldwide/
├── ai_structure_proposal/
│   └── pipeline.py            ← Stage 1: Gemini AI structure analysis
│
├── schema_validation/
│   └── pipeline.py            ← Stage 2: Deterministic schema generation
│
├── index_ai.html              ← Stage 3: Private admin webapp
│
├── xls_validation/
│   ├── pipeline.py            ← Stage 4a: Validation orchestrator
│   └── validators.py          ← Validation logic (label, header, data row checks)
│
├── xls_ingest/
│   └── pipeline.py            ← Stage 4b: Excel → Database ingestion
│
├── aedes_indicators/
│   ├── pipeline.py            ← Stage 5: Indicator computation (AEDES-specific)
│   └── index.html             ← Public dashboard
│
├── data_structures_reference.md  ← Full schema field reference
├── schema_validation.json     ← Current active schema (generated + admin-edited)
├── structure_proposal.json    ← Current structure proposal (from Gemini)
└── extraction_guide.json      ← Column mapping (generated by validation)
```

---

## Typical Workflows

### First-time setup (new template)

1. Admin uploads the blank Excel template in the private webapp.
2. The webapp runs **Stage 1** (structure proposal) — Gemini analyses the layout.
3. The webapp runs **Stage 2** (schema generation) — deterministic pipeline produces the schema.
4. Admin reviews the schema in the webapp sidebar: adjusts required fields, adds validation rules, confirms column mappings.
5. Admin saves the edited schema.
6. Setup is complete — organisations can now submit filled-in files.

### Organisation data submission

1. An organisation opens the private webapp.
2. They upload their filled-in Excel file.
3. The webapp runs **Stage 4a** (validation).
4. If errors are found, the organisation corrects their Excel file and re-uploads.
5. Once validation passes (zero errors), they click "Save to Database".
6. The webapp runs **Stage 4b** (ingestion) — data lands in the database.

### Indicator refresh

1. Admin schedules the `aedes-indicators` pipeline in OpenHEXA (e.g., daily, or after each new ingestion).
2. The pipeline reads all ingested data, computes indicators, and pushes parquet files to the public dashboard.
3. The public dashboard automatically reflects the latest data on next page load.

### Template modification

If the Excel template design changes (new columns, moved cells, renamed headers):

1. Re-run **Stage 1** with the updated template.
2. Re-run **Stage 2** to regenerate the schema.
3. Review and adjust the schema in the webapp.
4. Previously ingested data remains in the database. New submissions will use the updated schema. The ingestion pipeline handles schema evolution (adds new columns automatically).
