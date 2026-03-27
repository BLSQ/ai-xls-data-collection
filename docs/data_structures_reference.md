# Data Structures Reference

> How `structure_proposal.json` and `schema_validation.json` are built,
> and where validation rules live.

---

## Overview

The AEDES pipeline transforms Excel workbooks into validated, structured data through two core JSON files:

| File | Producer | Nature |
|------|----------|--------|
| `structure_proposal.json` | Gemini AI | *Descriptive* — "here is what the AI sees in the Excel" |
| `schema_validation.json` | Deterministic pipeline | *Prescriptive* — "here is what the validator must enforce" |

```
┌────────────┐    Gemini AI    ┌──────────────────────────┐    Deterministic    ┌─────────────────────────┐
│ Excel file │───────────────▶│ structure_proposal.json   │──────────────────▶│ schema_validation.json   │
└────────────┘                └──────────────────────────┘     pipeline        └────────────┬────────────┘
                                                                                            │
                                                                                        Validator
                                                                                            │
                                                                                            ▼
                                                                              ┌─────────────────────────┐
                                                                              │ extraction_guide.json    │
                                                                              └────────────┬────────────┘
                                                                                           │
                                                                                        Ingest
                                                                                           │
                                                                                           ▼
                                                                              ┌─────────────────────────┐
                                                                              │       Database           │
                                                                              └─────────────────────────┘
```

The proposal **describes** what exists. The schema **prescribes** what is valid.

---

## The Three Section Types

Before diving into the schema fields, it's essential to understand that the pipeline must handle three fundamentally different ways data is laid out in Excel sheets. Every section the AI identifies falls into one of these three types, and each type has its own extraction logic, its own constraints, and produces different keys in `schema_validation.json`.

### Key-Value (`key_value`)

A flat list of label → value pairs. One label in a cell, its corresponding value in an adjacent cell.

**Example** — a project header section:

```
         │       A                    │       B              │
─────────┼───────────────────────────┼──────────────────────┤
  Row 3  │ Nom du projet             │ PROGRAMME SANTE      │
  Row 4  │ Date de soumission        │ 2024-01-15           │
  Row 5  │ Pays                      │ RDC                  │
```

**Main constraint**: each field is a single cell at a known, fixed position. The validator checks that the label cell still contains the expected text (layout hasn't shifted), then validates the value cell against its type and rules. If the field is marked `required: true`, an empty cell in the submitted file is an error.

**Produces in schema**: `label_fields` (one entry per label cell) and `value_fields` (one entry per value cell).

### Grouped Key-Value (`grouped_key_value`)

A repeating pattern of columns sharing the same row labels. Think of a "Donors" section where each donor occupies a column, and the same fields (name, amount, currency) repeat on fixed rows.

**Example** — a donors section with 4 column slots:

```
         │       A              │   C       │   D       │   E       │   F       │
─────────┼──────────────────────┼───────────┼───────────┼───────────┼───────────┤
  Row 10 │ Nom du bailleur      │ USAID     │ WHO       │           │           │
  Row 11 │ Montant              │ 500000    │ 300000    │           │           │
  Row 12 │ Devise               │ USD       │ EUR       │           │           │
```

**Main constraint**: the number of filled elements is unknown at design time — the schema defines the column slots (C, D, E, F) but only some may be filled. The first row of the group acts as the **anchor row**: it determines which columns are "active" (have data). Only active columns are then checked for the remaining rows. At least one column must be active, otherwise it's an error.

**Produces in schema**: `label_fields` (shared labels in column A) and `value_fields` with a grouped structure containing `columns` (the slot letters) and `rows` (one entry per field, each with a `required_rule`).

### Records / Table (`records`)

A data table with column headers and an arbitrary number of data rows — the most complex layout.

**Example** — a budget table:

```
         │       A                    │    B        │    C       │    D       │    E       │
─────────┼───────────────────────────┼─────────────┼────────────┼────────────┼────────────┤
  Row 15 │ Intitulé budgétaire       │ Montant     │ Unité      │ Zone Nord  │ Zone Sud   │
─────────┼───────────────────────────┼─────────────┼────────────┼────────────┼────────────┤
  Row 16 │ Activité 1                │ 50000       │ Forfait    │ 60         │ 40         │
  Row 17 │ Activité 2                │ 30000       │ Jour       │ 100        │ 0          │
  Row 18 │                           │             │            │            │            │  ← empty = end
```

**Main constraints**: headers can be single columns (fixed position) or groups with sub-columns — and sub-columns can be **fixed** (known at design time, like "Zone Nord") or **dynamic** (discovered at validation time by regex pattern, like any province name matching `^[A-Z]`). Data rows start after the header and end at the first fully empty row. A primary key column is identified to enforce that every row has an identifier.

**Produces in schema**: `header_modules` (column definitions with position, type, validation rules) and `table_definitions` (row range, primary key, row-level validation).

### Summary

```
┌───────────────────────────────────────────────────────────────────────────────────────┐
│                                 Section Types                                          │
│                                                                                       │
│  ┌─────────────────────┐  ┌─────────────────────────┐  ┌───────────────────────────┐  │
│  │  key_value           │  │  grouped_key_value       │  │  records                   │  │
│  │                      │  │                          │  │                            │  │
│  │  Fixed label → value │  │  Repeating columns,      │  │  Header row + data rows,   │  │
│  │  pairs at known      │  │  unknown fill count,     │  │  fixed & dynamic columns,  │  │
│  │  positions            │  │  anchor row determines   │  │  ends at first empty row   │  │
│  │                      │  │  active elements         │  │                            │  │
│  │  → label_fields      │  │  → label_fields          │  │  → header_modules          │  │
│  │  → value_fields      │  │  → value_fields (grouped)│  │  → table_definitions       │  │
│  └─────────────────────┘  └─────────────────────────┘  └───────────────────────────┘  │
└───────────────────────────────────────────────────────────────────────────────────────┘
```

With this foundation, the next sections define every field that appears in `schema_validation.json` — knowing *which section type* each field belongs to is key to understanding its purpose.

---

## Canonical Naming

Both JSON files use **canonical names** — a deterministic transformation of human-readable headers into database-safe identifiers, performed by `to_column_canonical()`.

```
"Intitulé budgétaire / Libellé d'activité"  →  "intitule_budgetaire_libelle_d_activite"
"Population (%)"                             →  "population"
"Répartition géographique"                   →  "repartition_geographique"
```

In short: strip parenthetical text, remove accents and ligatures, replace non-alphanumeric characters with `_`, collapse, lowercase. The full algorithm is in `aedes_schema_from_ai_template/utils.py`.

**Why it matters**: canonical names are the thread that connects every stage of the pipeline. The same function is used in schema generation, validation, and indicator computation to ensure a header always maps to the same identifier. At ingestion, canonical names become the **PostgreSQL column headers** in `program_metadata` and `program_data` tables.

---

## Field Glossary

All JSON fields used in `schema_validation.json` are defined here, with how they are produced and consumed across the pipelines. They are described in context later in the document.

### Schema-level keys

| Field | Defined as | Used in |
|-------|-----------|---------|
| `label_fields` | Dictionary of section → field → `{cell, expected}`. Produced from `key_value` and `grouped_key_value` layouts by `aedes_schema_from_ai_template`. | `aedes_xls_ai_validation`: `validate_labels()` reads each entry, checks that the Excel cell at `cell` contains the `expected` text. Detects shifted layouts. Not used in ingestion. |
| `value_fields` | Dictionary of section → field → `{cell, type, validation, ...}`. Produced from `key_value` and `grouped_key_value` layouts. | `aedes_xls_ai_validation`: `validate_values()` reads each entry, enforces type, required, and validation rules on the cell value. `xls_ingest`: metadata extraction reads `canonical_name` and `cell` to build the `program_metadata` table. |
| `header_modules` | Dictionary of `module_N` → `{title, columns[], validation}`. Produced from `records` layouts. | `aedes_xls_ai_validation`: `validate_headers()` checks header cells match expectations; `validate_data_rows()` applies column-level and module-level rules. `xls_ingest`: `extract_data()` uses column canonical names as DataFrame/database column headers in the `program_data` table. |
| `table_definitions` | Dictionary of table_name → `{metadata, records_data}`. Produced from all layouts (metadata from key_value/grouped_key_value, records_data from records). | `aedes_xls_ai_validation`: `build_extraction_guide()` reads metadata fields and records_data config. `xls_ingest`: `extract_metadata()` and `extract_data()` read the extraction guide derived from this. |
| `validation_config` | Global settings object. Added manually or via webapp. | `aedes_xls_ai_validation`: controls `text_matches()` behavior (accent normalization, case sensitivity, whitespace stripping, etc.) across all label comparisons. |

### Field-level keys (inside `value_fields` and `header_modules` columns)

| Field | Defined as | Used in |
|-------|-----------|---------|
| `canonical_name` | Database-safe snake_case identifier, generated by `to_column_canonical()` (see Canonical Naming above). | **Validation**: stored in `column_map` to link canonical names to column letters. **Ingestion**: becomes the dict key for each extracted value → DataFrame column name → PostgreSQL column name in `program_metadata` or `program_data` tables (truncated to 63 chars if needed). |
| `expected` | The human-readable text expected in the header/label cell. Copied from the AI proposal's label or column name. | **Validation**: compared against the actual cell value via `text_matches()`. A mismatch produces a warning (headers) or error (labels). **Ingestion**: not used directly; stored in the extraction guide as `header` for documentation. |
| `cell` | A1-notation cell reference (e.g., `"B3"`). Points to the exact cell in the Excel sheet. | **Validation**: `ws[cell].value` reads the actual cell content. **Ingestion**: `ws[cell_ref].value` extracts the data value for metadata fields. |
| `type` / `value_type` | Data type: `"string"`, `"number"`, `"date"`, `"percentage"`. Inferred from the template cell value by `infer_type()`. | **Validation**: enforces type-appropriate parsing (`float()` for number, `datetime` check for date, etc.). **Ingestion**: `_coerce(raw, value_type)` converts cell values during extraction. |
| `required` | Boolean. Indicates whether a field must have data. Used across **all layout types**: key_value fields, records columns, and row_validation PK enforcement. Defaults to `false` for key_value fields (user sets it via the webapp). For records columns, set by the AI proposal. | **Validation**: empty cell + `required: true` → severity `"error"`; `required: false` + empty → severity `"info"`. Also used by `_find_primary_key()` as a fallback to select the primary key. |
| `example` | The actual value found in the template Excel cell at schema generation time. | **Validation**: included in error hint messages (`"Exemple attendu : '{example}'"`). Never validated itself — purely informational. **Ingestion**: not used. |
| `required_rule` | String. Used only for grouped key-value rows. Values: `"anchor"` (auto-assigned to the first field — the anchor row) or `"if_active"` (auto-assigned to subsequent fields). | **Validation**: the anchor row (`"anchor"`) determines which columns are "active" — at least one column must have a value there. Then, for each other row with `"if_active"`, only the active columns are checked (empty → error). See the grouped structure JSON in Part 1.2 for a full example. |
| `message` | Pre-composed error message string for null/type violations. | **Validation**: used as the error text when validation fails. |
| `label_ref` | Cross-reference string `"module_N.field_key"` linking a value field to its label field. | **Validation**: used internally to correlate label and value entries. |

### Position types (in `header_modules` columns)

| Field | Defined as | Used in |
|-------|-----------|---------|
| `position: "fixed"` | Header is at a known absolute cell position. | **Validation**: `_validate_fixed_header()` reads that specific cell, compares to `expected`, adds to `column_map`. |
| `position: "dynamic"` | Header is discovered at runtime by scanning consecutive cells matching `pattern`. | **Validation**: `_validate_dynamic_headers()` scans cells, builds canonical names as `{canonical_prefix}_{to_canonical(cell_value)}`, adds each to `column_map`. |
| `position: "group"` | Parent header spanning multiple sub-columns (which are themselves fixed or dynamic). | **Validation**: `_validate_group_header()` checks the parent cell, then recurses into `sub_columns`. |
| `pattern` | Regex string for dynamic columns. Applied to each scanned cell via `re.match()`. | **Validation**: only cells matching this pattern are included; first non-match stops the scan. |
| `canonical_prefix` | Prefix for building dynamic canonical names: `{prefix}_{to_canonical(cell_value)}`. | **Validation**: used to generate canonicals. **Extraction guide**: used to match resolved dynamic canonicals by `startswith(prefix + "_")`. |
| `min_count` | Minimum number of dynamic columns expected. Default: 0. | **Validation**: if fewer matching columns found → error. |
| `known_variants` | Alternative acceptable texts for a header. | **Validation**: if actual value matches a variant → severity `"info"` (recognized but non-standard). |

> **What is `column_map`?** — During header validation, the validator builds a dictionary called `column_map` that maps each canonical name to its resolved column letter (e.g., `{"zone_nord": "D", "provincial_bas_uele": "E", "provincial_haut_katanga": "F"}`). This is essential because **dynamic columns don't have a known column letter at design time** — the schema only contains a `pattern` and `canonical_prefix`. The `column_map` is populated as headers are validated and then used throughout data row validation and extraction guide generation to translate canonical names into actual Excel column positions. Without it, there would be no way to read dynamic column data.

### Validation rule types

Rules can be attached at four scopes (described in detail in Part 4):

| Scope | Location in schema | Applied to |
|-------|-------------------|------------|
| **Cell-level** | `value_fields[section][field].validation` | One specific cell value |
| **Column-level** | `header_modules[module].columns[i].validation` | Every cell in that column across all data rows |
| **Module-level** | `header_modules[module].validation` | Cross-column rules for an entire group of columns |
| **Row-level** | `table_definitions[table].records_data.row_validation` | Relationships between fields in the same row |

| Rule format | Defined as | Used in |
|-------------|-----------|---------|
| `min` | Minimum allowed value (numeric or date). | **Validation**: after successful type cast, checks `value < min` → error. |
| `max` | Maximum allowed value (numeric or date). | Defined in schema but **not currently enforced** by the validator. |
| `enum` | List of allowed string values (e.g., `["Oui", "Non", "N/A"]`). | **Validation**: checks membership → error if not in list. |
| `conditional_enum` | Dependent dropdown: `{depends_on_column, values_by_parent}`. | **Validation**: reads the parent column first, then checks allowed values for that parent. |
| `json_logic` | List of `{rule, description}` objects using JsonLogic expressions. | **Validation**: evaluated by `evaluate_json_logic()`. Supports `==`, `!=`, `>`, `>=`, `<`, `<=`, `+`, `-`, `*`, `/`, `!`, `if`, `and`, `or`, `var`, `reduce`. Data context: `{"value": cell_value}` at cell-level, full row dict at column/module-level. |
| `json_logic_op` | How to combine multiple `json_logic` rules: `"and"` (default) or `"or"`. | **Validation**: `"and"` requires all rules to pass; `"or"` requires at least one. |

### Table definitions fields

| Field | Defined as | Used in |
|-------|-----------|---------|
| `header_row_start` / `header_row_end` | Row numbers of the header band. | Informational only — not read by the validator or ingestion. |
| `data_row_start` | First row number containing data. | **Validation**: `find_data_rows()` and `validate_data_rows()` start scanning from this row. **Ingestion** (via extraction guide): `extract_data()` iterates from `start_row`. |
| `data_row_end` | `"dynamic"` — end row determined at runtime. | **Validation**: `find_data_rows()` scans until first fully empty row (`empty_row_terminates: true`). |
| `empty_row_terminates` | Boolean. If `true`, stop at first fully empty row. | **Validation**: `_is_row_empty()` checks columns B through BF; first empty row ends the data range. |
| `primary_key` | Canonical name of the primary key column. Determined by `_find_primary_key()`. | Stored as the key in `row_validation`. Every data row must have a non-empty value in this column. |
| `row_validation` | Dict of `{canonical_name: {required: true, message: "..."}}`. | **Validation**: enforces non-null on the primary key column for each data row. |

### Validation config sub-fields

| Field | Defined as | Used in |
|-------|-----------|---------|
| `label_comparison.strip_whitespace` | Default `true`. | Calls `.strip()` and collapses internal whitespace before comparison. |
| `label_comparison.strip_trailing_colon` | Default `true`. | Calls `.rstrip(":")` to tolerate labels with/without trailing colons. |
| `label_comparison.case_sensitive` | Default `false`. | If falsy, converts to `.lower()` before comparison. |
| `label_comparison.normalize_accents` | Default `true`. | Strips diacritical marks and expands ligatures (e.g., `"œ"` → `"oe"`). |

---

## Part 1 — JSON Representations by Layout Type

The section types introduced above are now shown with their full JSON structures in both `structure_proposal.json` and `schema_validation.json`.

---

### 1.1 — Key-Value (`key_value`)

#### In `structure_proposal.json`

```json
{
  "section_name": "Informations Générales",
  "layout": "key_value",
  "fields": [
    {
      "label": "Nom du projet / Programme",
      "value": "EXAMPLE PROJECT",
      "cell_label": "A3",
      "cell_value": "B3"
    },
    {
      "label": "Date de soumission",
      "value": "2024-01-15",
      "cell_label": "A5",
      "cell_value": "B5"
    }
  ]
}
```

Each field has: a **label** (what the cell says), a **value** (what was found), and **cell coordinates**.

#### In `schema_validation.json`

The proposal fields become two parallel dictionaries:

```json
{
  "label_fields": {
    "informations_generales": {
      "nom_du_projet_programme": {
        "cell": "A3",
        "expected": "Nom du projet / Programme"
      }
    }
  },
  "value_fields": {
    "informations_generales": {
      "nom_du_projet_programme": {
        "cell": "B3",
        "type": "string",
        "required": true,
        "example": "EXAMPLE PROJECT",
        "message": "Le champ 'Nom du projet / Programme' (cellule B3) est requis.",
        "canonical_name": "nom_du_projet_programme"
      }
    }
  }
}
```

```
┌─────────────────────────────────────┐          ┌──────────────────────────────────────┐
│       Proposal (key_value)          │          │              Schema                   │
│                                     │          │                                      │
│  field.label + field.cell_label ────┼────────▶│  label_fields → expected text at cell │
│                                     │          │                                      │
│  field.value + field.cell_value ────┼────────▶│  value_fields → type, rules,          │
│                                     │          │                 canonical name        │
└─────────────────────────────────────┘          └──────────────────────────────────────┘
```

**Key insight**: `label_fields` lets the validator confirm the Excel layout hasn't shifted. `value_fields` carries the actual validation rules.

---

### 1.2 — Grouped Key-Value (`grouped_key_value`)

A repeating pattern of columns. Think of a "Donors" section where each donor has the same fields (Name, Amount, Currency) repeated across columns.

#### In `structure_proposal.json`

```json
{
  "section_name": "Bailleurs",
  "layout": "grouped_key_value",
  "elements": [
    {
      "element_name": "Bailleur 1",
      "columns": ["C", "D"],
      "labels": {
        "A10": "Nom du bailleur",
        "A11": "Montant"
      },
      "fields": {
        "C10": "USAID",
        "C11": "500000"
      }
    },
    {
      "element_name": "Bailleur 2",
      "columns": ["E", "F"],
      "labels": {
        "A10": "Nom du bailleur",
        "A11": "Montant"
      },
      "fields": {
        "E10": "WHO",
        "E11": "300000"
      }
    }
  ]
}
```

Each **element** is one instance of the repeating group. They share the same label rows but occupy different columns.

#### In `schema_validation.json`

Grouped key-values are also stored in `label_fields` / `value_fields`, but with **element-indexed keys**:

```json
{
  "value_fields": {
    "bailleurs": {
      "bailleur_1__nom_du_bailleur": {
        "cell": "C10",
        "type": "string",
        "canonical_name": "bailleur_1__nom_du_bailleur"
      },
      "bailleur_2__nom_du_bailleur": {
        "cell": "E10",
        "type": "string",
        "canonical_name": "bailleur_2__nom_du_bailleur"
      }
    }
  }
}
```

The double-underscore `__` separates element name from field name in the canonical key.

The schema also stores the **grouped validation structure** under a separate `value_fields` key, which contains the column slots and row-level `required_rule` entries:

```json
{
  "value_fields": {
    "bailleurs_values": {
      "columns": ["C", "D", "E", "F"],
      "rows": {
        "nom_du_bailleur": {
          "row": 10,
          "type": "string",
          "required_rule": "anchor",
          "message": "Au moins un 'Nom du bailleur' (ligne 10) doit être renseigné."
        },
        "montant": {
          "row": 11,
          "type": "number",
          "required_rule": "if_active",
          "message": "Le 'Montant' (ligne 11) est requis pour chaque élément actif."
        },
        "devise": {
          "row": 12,
          "type": "string",
          "required_rule": "if_active",
          "message": "La 'Devise' (ligne 12) est requis pour chaque élément actif."
        }
      }
    }
  }
}
```

This is where the **anchor row mechanism** lives. The validator finds the row with `required_rule: "anchor"` (here: `nom_du_bailleur`, row 10), reads it across all `columns`, and determines which columns are **active** (have a non-empty value). Then for every other row with `required_rule: "if_active"`, only the active columns are checked — empty cells in inactive columns are silently skipped.

---

### 1.3 — Records (`records`)

The most complex layout. A data table with column headers and data rows — like a spreadsheet grid.

#### In `structure_proposal.json`

```json
{
  "section_name": "Budget Détaillé",
  "layout": "records",
  "headers": {
    "header_row": 15,
    "data_start_row": 16,
    "primary_key": "Intitulé budgétaire / Libellé d'activité",
    "columns": [
      {
        "name": "Intitulé budgétaire / Libellé d'activité",
        "type": "single",
        "column_letter": "A",
        "expected_values": ["Activité 1", "Activité 2"]
      },
      {
        "name": "Répartition géographique (%)",
        "type": "group",
        "column_letter": "D",
        "sub_columns": [
          {
            "name": "Zone Nord",
            "column_letter": "D",
            "fixed": true
          },
          {
            "name": "Autre zone",
            "column_letter": "E",
            "fixed": false,
            "pattern": "^[A-Z]",
            "canonical_prefix": "repartition_geographique"
          }
        ]
      }
    ]
  },
  "records_data": [
    {"Intitulé budgétaire / Libellé d'activité": "Activité 1", "Zone Nord": 60}
  ]
}
```

This introduces the **column type hierarchy**:

```
                        ┌────────────┐
                        │   Column   │
                        └──────┬─────┘
                               │
                ┌──────────────┴──────────────┐
                │                             │
                ▼                             ▼
    ┌───────────────────┐         ┌───────────────────┐
    │   Single Column   │         │   Group Column    │
    │  (always fixed)   │         │  (parent header)  │
    └───────────────────┘         └─────────┬─────────┘
                                            │
                                 ┌──────────┴──────────┐
                                 │                     │
                                 ▼                     ▼
                     ┌───────────────────┐ ┌───────────────────────┐
                     │   Sub-column      │ │   Sub-column          │
                     │   fixed: true     │ │   fixed: false        │
                     │  (known at design │ │  (discovered at       │
                     │   time)           │ │   validation time)    │
                     └───────────────────┘ └───────────────────────┘
```

| Column kind | Known at design time? | Example |
|-------------|----------------------|---------|
| `single` (fixed) | Yes | "Intitulé budgétaire" — always present |
| `group` > `fixed: true` | Yes | "Zone Nord" — a specific known zone |
| `group` > `fixed: false` | **No** | "Autre zone" — discovered by regex pattern |

**Dynamic columns** carry a `pattern` (regex) and `canonical_prefix`. The validator scans consecutive cells matching the pattern and assigns canonical names like `repartition_geographique_zone_sud`, `repartition_geographique_zone_est`, etc.

#### In `schema_validation.json`

Records produce a `header_modules` dictionary:

```json
{
  "header_modules": {
    "module_1": {
      "label": "Intitulé budgétaire / Libellé d'activité",
      "position": "fixed",
      "columns": [
        {
          "expected": "Intitulé budgétaire / Libellé d'activité",
          "canonical_name": "intitule_budgetaire_libelle_d_activite",
          "position": "fixed",
          "required": true,
          "type": "string"
        }
      ]
    },
    "module_8": {
      "label": "Répartition géographique (%)",
      "position": "group",
      "columns": [
        {
          "expected": "Zone Nord",
          "canonical_name": "zone_nord",
          "position": "fixed",
          "type": "number"
        },
        {
          "expected": null,
          "canonical_name": null,
          "position": "dynamic",
          "pattern": "^[A-Z]",
          "canonical_prefix": "repartition_geographique",
          "type": "number"
        }
      ]
    }
  }
}
```

And a `table_definitions` entry:

```json
{
  "table_definitions": {
    "budget_detaille": {
      "header_row": 15,
      "data_start_row": 16,
      "primary_key": "intitule_budgetaire_libelle_d_activite",
      "records_data": [...]
    }
  }
}
```

```
┌──────────────────────────────────────┐          ┌──────────────────────────────────────────────┐
│       Proposal (records)             │          │              Schema                           │
│                                      │          │                                              │
│  headers.columns[] ──────────────────┼────────▶│  header_modules{}                             │
│                                      │          │    module_N → columns[]                       │
│                                      │          │                                              │
│  records_data[] ─────────────────────┼────────▶│  table_definitions{}                          │
│                                      │          │    header_row, data_start_row,                │
│                                      │          │    primary_key (canonical), records_data      │
│                                      │          │                                              │
│  headers.primary_key ────────────────┼────────▶│  (resolved via _find_primary_key())           │
└──────────────────────────────────────┘          └──────────────────────────────────────────────┘
```

---

## Part 2 — Complete Schema Structure

Here is the full top-level structure of `schema_validation.json`:

```
                              ┌──────────────────────────────┐
                              │    schema_validation.json     │
                              └──────────────┬───────────────┘
                                             │
              ┌──────────────┬───────────────┼───────────────┬──────────────┐
              │              │               │               │              │
              ▼              ▼               ▼               ▼              ▼
   ┌──────────────┐ ┌──────────────┐ ┌──────────────┐ ┌────────────┐ ┌────────────────┐
   │ label_fields │ │ value_fields │ │header_modules│ │table_defs  │ │validation_config│
   │              │ │              │ │              │ │            │ │                │
   │ section →    │ │ section →    │ │ module_N →   │ │ table →    │ │ global         │
   │ field →      │ │ field →      │ │ label +      │ │ metadata + │ │ settings       │
   │ cell +       │ │ cell + type  │ │ position +   │ │ records_   │ │                │
   │ expected     │ │ + rules      │ │ columns[]    │ │ data       │ │                │
   └──────┬───────┘ └──────┬───────┘ └──────┬───────┘ └─────┬──────┘ └────────────────┘
          │                │                │               │
          │  from:         │  from:         │  from:        │  from:
          ▼                ▼                ▼               ▼
  ┌──────────────┐ ┌──────────────┐ ┌──────────────┐ ┌──────────────┐
  │ key_value    │ │ key_value    │ │ records      │ │ records data │
  │ fields       │ │ values       │ │ headers      │ │ + metadata   │
  │              │ │              │ │              │ │              │
  │ grouped_kv   │ │ grouped_kv   │ │              │ │              │
  │ labels       │ │ values       │ │              │ │              │
  └──────────────┘ └──────────────┘ └──────────────┘ └──────────────┘
```

See the **Field Glossary** above for detailed definitions of each key.

---

## Part 3 — Validation Rules & JsonLogic

Rules can be attached at four scopes. Each scope stores rules differently.

```
┌───────────────────────────────────────────────────────────────────────────────────────────┐
│                                     Rule Scopes                                           │
│                                                                                           │
│  ┌──────────────────┐  ┌──────────────────┐  ┌──────────────────┐  ┌──────────────────┐  │
│  │  Cell-level       │  │  Column-level     │  │  Module-level    │  │  Row-level        │  │
│  │  One specific     │  │  One column in    │  │  An entire group │  │  Cross-column     │  │
│  │  value field      │  │  a data table     │  │  of columns      │  │  in same row      │  │
│  └────────┬─────────┘  └────────┬─────────┘  └────────┬─────────┘  └────────┬─────────┘  │
│           │                     │                      │                     │            │
│       stored in             stored in              stored in             stored in        │
│           │                     │                      │                     │            │
│           ▼                     ▼                      ▼                     ▼            │
│  value_fields               header_modules          header_modules       table_definitions│
│  [section][field]           [module]                [module]             [table]           │
│  .validation                .columns[i]             .validation          .records_data    │
│                             .validation             .json_logic          .row_validation  │
└───────────────────────────────────────────────────────────────────────────────────────────┘
```

### 4.1 — Cell-level rules

Located in `value_fields` → specific field → `validation`.

**Simple (compact) format** — for common constraints:

```json
{
  "value_fields": {
    "informations_generales": {
      "date_de_soumission": {
        "cell": "B5",
        "type": "date",
        "validation": {
          "min": "2020-01-01",
          "max": "2030-12-31"
        }
      }
    }
  }
}
```

**Enum format** — for allowed values:

```json
{
  "validation": {
    "enum": ["Oui", "Non", "N/A"]
  }
}
```

**Conditional enum format** — for dependent dropdowns where the allowed values depend on another column's value (typically generated from Excel's INDIRECT data validation):

```json
{
  "validation": {
    "conditional_enum": {
      "depends_on_column": "E",
      "values_by_parent": {
        "Santé": ["Vaccination", "Nutrition", "WASH"],
        "Éducation": ["Primaire", "Secondaire", "Formation professionnelle"]
      }
    }
  }
}
```

The validator reads the parent column (`E`) for the current row first, then looks up the allowed values. If the parent value is `"Santé"`, only `["Vaccination", "Nutrition", "WASH"]` are accepted. An unknown parent value produces a warning; a value not in the allowed list produces an error.

**JsonLogic format** — for complex conditions:

```json
{
  "validation": {
    "json_logic": [
      {
        "rule": { "<=": [{ "var": "value" }, 100] },
        "description": "Must be ≤ 100"
      }
    ],
    "json_logic_op": "and"
  }
}
```

At cell-level, `{ "var": "value" }` resolves to the single cell's value. This is the simplest context — you only have access to the one value being validated.

### 4.2 — Column-level rules

Located in `header_modules` → module → `columns[i]` → `validation`.

Same format options (compact or JsonLogic). Applied to every cell in that column across all data rows.

```json
{
  "header_modules": {
    "module_6": {
      "columns": [
        {
          "canonical_name": "population_cible_h",
          "position": "fixed",
          "type": "number",
          "validation": {
            "min": 0,
            "json_logic": [
              {
                "rule": { "<=": [{ "var": "value" }, 1000000] },
                "description": "Population must be ≤ 1,000,000"
              }
            ]
          }
        }
      ]
    }
  }
}
```

At column-level, the `var` context changes: the data passed to JsonLogic is a **full row dictionary** `{canonical_name: cell_value, ...}` containing all columns in the current data row. This means:
- `{ "var": "value" }` still resolves to the current column's cell value (it's injected as a convenience)
- `{ "var": "population_cible_f" }` resolves to another column's value in the same row — enabling cross-column comparisons like `population_h <= population_total`

### 4.3 — Module-level (group) rules

Located in `header_modules` → module → `validation`.

These rules apply **across columns** — typically sum constraints. All cross-column constraints use JsonLogic exclusively.

**Example** — a "Répartition géographique" group with fixed sub-columns Zone Nord (col F), Zone Est (col G), Zone Sud (col H). The rule checks that the three zones sum to 100%:

```json
{
  "header_modules": {
    "module_8": {
      "columns": [
        {
          "position": "group",
          "parent": { "row": 15, "expected": "Répartition géographique (%)" },
          "sub_columns": [
            { "position": "fixed", "column": "F", "canonical_name": "zone_nord", "expected": "Zone Nord" },
            { "position": "fixed", "column": "G", "canonical_name": "zone_est", "expected": "Zone Est" },
            { "position": "fixed", "column": "H", "canonical_name": "zone_sud", "expected": "Zone Sud" }
          ]
        }
      ],
      "validation": {
        "json_logic": [
          {
            "rule": {
              "==": [
                { "+": [{ "var": "zone_nord" }, { "var": "zone_est" }, { "var": "zone_sud" }] },
                100
              ]
            },
            "description": "Zone Nord + Zone Est + Zone Sud must equal 100%"
          }
        ],
        "json_logic_op": "and"
      }
    }
  }
}
```

At module-level, `{ "var": "zone_nord" }` resolves to column F's value in the current row. The `+` operator accepts any number of arguments, so the sum is straightforward.

For **dynamic columns** (where the number of sub-columns is discovered at runtime), the `row_values` variable provides a list of all numeric values in the module's columns. In that case, use `reduce` to sum the array:

```json
{ "==": [
    { "reduce": [
        { "var": "row_values" },
        { "+": [{ "var": "accumulator" }, { "var": "current" }] },
        0
    ]},
    100
]}
```

Here `row_values` might be `[60, 25, 15]` — the `reduce` walks the list, accumulating the sum starting from 0. This is only needed when you can't name the columns explicitly.

### JsonLogic `var` context summary

```
┌─────────────────────┬────────────────────────────────────────────────────────┐
│  Scope              │  Data available to { "var": "..." }                    │
├─────────────────────┼────────────────────────────────────────────────────────┤
│  Cell-level         │  { "value": <cell_value> }                            │
│  (value_fields)     │  Only the single cell being validated.                │
├─────────────────────┼────────────────────────────────────────────────────────┤
│  Column-level       │  { "value": <cell_value>,                             │
│  (header_modules    │    "canonical_a": <val>, "canonical_b": <val>, ... }  │
│   .columns[i])      │  Full row dict. Access any column by canonical name.  │
├─────────────────────┼────────────────────────────────────────────────────────┤
│  Module-level       │  { "value": <cell_value>,                             │
│  (header_modules    │    "canonical_a": <val>, "canonical_b": <val>, ...,   │
│   .validation)      │    "row_values": [<num>, <num>, ...] }                │
│                     │  Full row dict + row_values list for reduce/sum.      │
└─────────────────────┴────────────────────────────────────────────────────────┘
```

### 4.4 — Primary key enforcement (`row_validation`)

The `row_validation` entry in `table_definitions.records_data` enforces that the primary key column is never empty. It is **not** a general-purpose cross-column rule engine — for cross-column logic (e.g., "total must equal sum of parts"), use column-level or module-level JsonLogic (see section 4.3).

```json
{
  "table_definitions": {
    "records_data": {
      "row_validation": {
        "intitule_budgetaire_libelle_d_activite": {
          "required": true,
          "message": "Chaque ligne de données doit avoir un 'intitule_budgetaire_libelle_d_activite'."
        }
      }
    }
  }
}
```

For cross-column validation, use a **column-level JsonLogic** rule on the total column. For example, to check that `realise_total_ligne` equals the sum of `realise_2024` and `realise_2025`:

```json
{
  "header_modules": {
    "module_5": {
      "columns": [
        {
          "position": "fixed",
          "column": "H",
          "canonical_name": "realise_total_ligne",
          "expected": "Total ligne",
          "validation": {
            "json_logic": [
              {
                "rule": {
                  "==": [
                    { "var": "realise_total_ligne" },
                    { "+": [{ "var": "realise_2024" }, { "var": "realise_2025" }] }
                  ]
                },
                "description": "Total ligne must equal réalisé 2024 + réalisé 2025"
              }
            ]
          }
        }
      ]
    }
  }
}
```

At column-level scope, `{ "var": "realise_2024" }` resolves to the value of that column in the same row (the full row dict is available — see the `var` context table above).

### Rule format summary

```
                        ┌─────────────────┐
                        │ Validation Rule  │
                        └────────┬────────┘
                                 │
              ┌──────────────────┼──────────────────┐
              │                  │                   │
              ▼                  ▼                   ▼
   ┌──────────────────┐ ┌──────────────────┐ ┌──────────────────┐
   │  Compact format  │ │ JsonLogic format │ │  Enum format     │
   │  min, max        │ │ json_logic[] +   │ │  allowed values  │
   │                  │ │ json_logic_op    │ │  list            │
   │  simple          │ │ complex logic    │ │  standalone      │
   │  constraints     │ │ any scope        │ │  any scope       │
   │  any scope       │ │                  │ │                  │
   └──────────────────┘ └──────────────────┘ └──────────────────┘
```

| Format | Use case | Combining |
|--------|----------|-----------|
| `min` / `max` | Numeric or date range | Implicit AND with other rules |
| `enum` | Allowed values list | Standalone |
| `json_logic` | Arbitrary conditions (including sum constraints) | `json_logic_op`: `"and"` (default) or `"or"` |

---

## Part 4 — Transformation Summary

This diagram shows the full transformation from proposal to schema for all three layout types:

```
┌──────────────────────────────────────────────────────────────────────────────────────────┐
│                              structure_proposal.json                                      │
│                                                                                          │
│  ┌─────────────────────────────┐  ┌──────────────────────────────┐  ┌──────────────────┐ │
│  │ key_value section           │  │ grouped_key_value section    │  │ records section   │ │
│  │ fields[]: label, value,     │  │ elements[]: columns,         │  │ headers:          │ │
│  │   cell_label, cell_value    │  │   labels, fields             │  │   columns[],      │ │
│  │                             │  │                              │  │   primary_key     │ │
│  │                             │  │                              │  │ records_data[]    │ │
│  └──────────┬──────────────────┘  └──────────────┬───────────────┘  └────────┬─────────┘ │
└─────────────┼────────────────────────────────────┼──────────────────────────┼────────────┘
              │                                    │                          │
              │  label → label_fields              │  labels → label_fields   │  columns → header_modules
              │  value → value_fields              │  fields → value_fields   │  metadata + data → table_defs
              │                                    │  (element__field keys)   │
              ▼                                    ▼                          ▼
┌──────────────────────────────────────────────────────────────────────────────────────────┐
│                              schema_validation.json                                       │
│                                                                                          │
│  ┌─────────────┐  ┌─────────────┐  ┌───────────────┐  ┌───────────────┐  ┌────────────┐ │
│  │ label_fields │  │ value_fields │  │header_modules │  │table_          │  │validation_ │ │
│  │              │  │              │  │               │  │definitions    │  │config      │ │
│  │ section →    │  │ section →    │  │ module →      │  │               │  │            │ │
│  │ field →      │  │ field →      │  │ columns[]     │  │ table →       │  │ (added by  │ │
│  │ cell,        │  │ cell, type,  │  │ (fixed /      │  │ metadata,     │  │  webapp or │ │
│  │ expected     │  │ rules        │  │  group /      │  │ primary_key,  │  │  manually) │ │
│  │              │  │              │  │  dynamic)     │  │ row_valid.,   │  │            │ │
│  │              │  │              │  │               │  │ records_data  │  │            │ │
│  └─────────────┘  └─────────────┘  └───────────────┘  └───────────────┘  └────────────┘ │
└──────────────────────────────────────────────────────────────────────────────────────────┘
```

---

## Part 5 — The Extraction Guide

The `extraction_guide.json` is the **output** of the validation pipeline (`aedes_xls_ai_validation`). It is a pure data-extraction map — it contains no validation rules, no label checks, and no `required` fields. Its sole purpose is to tell the ingestion pipeline **exactly where to find each piece of data** in the Excel file.

The validation pipeline produces it by resolving dynamic columns (scanning the actual Excel headers), computing active grouped columns, and determining the data row range. The ingestion pipeline reads **only** the extraction guide — it never reads `schema_validation.json` directly.

### Structure

```json
{
  "generated_at": "2026-03-25T15:38:09",
  "schema_version": "1.0.0",
  "sheets": {
    "<sheet_name>": {
      "metadata": {
        "fields": [
          { "canonical_name": "project_name", "value_cell": "C3", "type": "string" },
          { "canonical_name": "total_budget",  "value_cell": "C6", "type": "number" }
        ],
        "grouped_sections": {
          "donors": {
            "columns": ["C", "D", "E", "F", "G"],
            "active_columns": ["C", "D"],
            "fields_per_column": [
              { "canonical_name": "donors_donor_name", "row": 9, "type": "string" },
              { "canonical_name": "donors_amount",     "row": 10, "type": "number" }
            ]
          }
        }
      },
      "data": {
        "start_row": 17,
        "end_row": 42,
        "row_count": 26,
        "columns": {
          "budget_line_item": {
            "col_letter": "B",
            "module": "module_3",
            "category": "default",
            "header": "Intitulé budgétaire",
            "value_type": "string"
          },
          "2024": {
            "col_letter": "D",
            "module": "module_3",
            "category": "Réalisé",
            "header": "2024",
            "value_type": "number"
          }
        }
      }
    }
  }
}
```

### How ingestion reads it

| Guide path | Ingestion usage |
|-----------|----------------|
| `metadata.fields[]` | `extract_metadata()` reads `canonical_name` + `value_cell` → reads `ws[value_cell].value` → coerces via `type` → one row in `program_metadata` table |
| `metadata.grouped_sections` | `extract_metadata()` reads `active_columns` + `fields_per_column` → reads each `(column, row)` intersection → keys like `donors_1_donor_name` in `program_metadata` |
| `data.start_row` / `data.end_row` | `extract_data()` iterates rows in this range |
| `data.columns` | For each `canonical_name → {col_letter, value_type}`: reads `ws.cell(row, col_idx(col_letter)).value` → coerces via `value_type` → one column in `program_data` table |

### Key differences from `schema_validation.json`

| | `schema_validation.json` | `extraction_guide.json` |
|--|--------------------------|------------------------|
| **Purpose** | What to validate | Where to extract data |
| **Dynamic columns** | Pattern + prefix (unresolved) | Concrete canonical names (resolved) |
| **Row range** | `data_row_start` + `"dynamic"` | `start_row` + `end_row` (computed) |
| **Grouped sections** | `required_rule` for validation logic | `active_columns` (which slots have data) |
| **Validation rules** | `enum`, `json_logic`, `min`, etc. | None |
| **Label checks** | `label_fields` with `expected` text | None |

### Modifying the extraction guide

The extraction guide is **regenerated** each time the validation pipeline runs. To change what gets extracted:

- **To change column mappings or types**: modify `schema_validation.json` (via the webapp rule editor or directly), then re-run validation.
- **To force-include/exclude columns**: the guide reflects whatever `header_modules` and `column_map` produce — adjust the schema accordingly.
- **Manual edits**: you can edit `extraction_guide.json` directly for one-off adjustments (e.g., changing `end_row` to include more rows), but these changes will be overwritten on the next validation run.

---

## Appendix — File Locations

| Artifact | Default path |
|----------|-------------|
| Structure proposal | `structure_proposal.json` (output of `aedes-structure-proposal`) |
| Validation schema | `schema_validation.json` (output of `aedes-schema-from-ai-template`) |
| Extraction guide | `extraction_guide.json` (output of `aedes-xls-ai-validation`) |
| Webapp (rule editor) | `index.html` |

Rules are edited through the webapp's rule builder interface, which writes directly into `schema_validation.json` at the appropriate scope (cell, column, or module).
