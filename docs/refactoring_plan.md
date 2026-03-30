# OOP Refactoring Plan — Excel Data Collection AI

## Context

The codebase has ~3,700 lines of procedural Python across 4 OpenHEXA pipelines that process
Excel templates through: AI structure proposal → schema generation → validation → ingestion.

Three **section types** are the central domain concept:

| Type | What it represents | Complexity |
|------|--------------------|------------|
| `key_value` | Label → value pairs at fixed cells (e.g., project name, date) | Low |
| `grouped_key_value` | Repeating column groups sharing row labels (e.g., donors) | Medium |
| `records` | Data tables with headers and rows (budget lines, activities) | High |

Currently all logic lives in flat functions within monolithic files (989, 1247, 1459, 290 lines).
This makes it hard for colleagues to navigate and maintain.

**Goal**: Refactor into classes organized around the three section types, with a shared abstract
interface (`register`, `qualify`, `add_validation` for schema; `validate_structure`, `validate_data`
for validation). Keep `pipeline.py` as thin orchestrator in each directory. Preserve all behavior.
Fix C1 (API key in URL).

**Constraint**: Each OpenHEXA pipeline directory is deployed independently — **no shared modules
across directories**. Utility code (e.g., `to_canonical`) is duplicated per pipeline that needs it.

---

## Design Principles

### 1. Polymorphic section types

Each section type is a class that implements the same abstract interface. This means a colleague
reading the code can understand "key_value sections are handled here" without reading 1,200 lines.

```
Pipeline 2 (schema generation)          Pipeline 3 (validation)
┌────────────────────────┐              ┌────────────────────────┐
│  SectionProcessor(ABC) │              │  SectionValidator(ABC) │
│  ─ register()          │              │  ─ validate_structure() │
│  ─ qualify()           │              │  ─ validate_data()      │
│  ─ add_validation()    │              └──────────┬─────────────┘
└──────────┬─────────────┘                         │
           │                                       │
    ┌──────┼──────────────┐               ┌────────┼────────────┐
    │      │              │               │        │            │
  KV    Grouped       Records           KV     Grouped     Records
```

### 2. Column definitions with 3-level depth

Records tables support up to 3 levels of headers: **module title → group parent → leaf column**.
Canonical names are built by concatenating the parent chain (as the current code already does).

```python
@dataclass
class ColumnDef:
    """Column definition — leaf or group with up to 2 levels of sub-columns."""
    name: str
    canonical_name: str
    column_letter: str | None = None    # known for fixed, None for dynamic
    position: str = "fixed"             # "fixed" | "dynamic"
    value_type: str = "string"
    required: bool = False
    sub_columns: list["ColumnDef"] = field(default_factory=list)  # max 1 nesting level
    validation: dict | None = None
    # Dynamic-only:
    pattern: str | None = None
    canonical_prefix: str | None = None
```

### 3. Multiple records tables (designed, one for now)

Each `records` section can produce its own database table. The AI proposes a `section_key`
(e.g., `"budget_lines"`); the schema builder uses that as the table name. Fallback if no key
is provided: `"program_data"`. The ingest pipeline already receives the table name from the
extraction guide, so this is forward-compatible.

### 4. Duplicate header detection (pipeline 2)

When building the schema, the `SchemaBuilder` checks that no two leaf columns produce the same
`canonical_name` without a distinguishing parent (module key or group parent). If a collision is
detected, it emits a warning and auto-prefixes with the module key (existing behavior in
`_deduplicate_canonicals`, now made explicit as a validation step).

### 5. No v1/v2 distinction

There is one schema format. The current `is_v2_schema()` branching and all `_v2_`-prefixed
functions are merged into the main section-type classes. The section type (key_value, grouped,
records) is the dispatch axis — not a version number.

---

## Pipeline 1: `ai_structure_proposal/`

**Current**: 989 lines in `pipeline.py`

### New file structure
```
ai_structure_proposal/
  pipeline.py              ← orchestrator (~80 lines)
  excel_reader.py          ← ExcelReader class
  sheet_renderer.py        ← SheetRenderer class
  gemini_client.py         ← GeminiClient class  ★ fixes C1
  prompt_builder.py        ← SCHEMA_FORMAT_SPEC constant + build_prompt()
  response_parser.py       ← parse_structure_json()
  validation_extractor.py  ← ValidationExtractor class
  workspace.yaml           ← unchanged
```

### ExcelReader (`excel_reader.py`)

Bundles the two workbook handles (`data_only=True` and `data_only=False`) that are threaded
through every function today.

| Method | From function | Purpose |
|--------|--------------|---------|
| `__init__(path)` | — | loads `wb_values` + `wb_full` |
| `select_sheets(name)` | `_select_sheets()` | returns `(ws_values, ws_full)` |
| `build_text_representation(ws, ws_full, max_rows)` | `build_text_representation()` | text dump for LLM |
| `close()` | — | close both workbooks |
| `_extract_hex_color(color)` | `_extract_hex_color()` | static |
| `_cell_bg_hex(cell)` | `_cell_bg_hex()` | static |
| `_cell_fg_hex(cell)` | `_cell_fg_hex()` | static |

### SheetRenderer (`sheet_renderer.py`)

Isolated image rendering concern — PIL-based, no dependency on other classes.

| Method | From function | Purpose |
|--------|--------------|---------|
| `__init__(max_rows, max_cols)` | — | stores config |
| `render(ws) -> bytes \| None` | `render_sheet_image()` | renders top-left as PNG |
| `_load_fonts()` | `_load_fonts()` | static, font loading |

### GeminiClient (`gemini_client.py`) — fixes C1

| Method | From function | Purpose |
|--------|--------------|---------|
| `__init__(api_key, model)` | — | stores `_api_key` privately |
| `generate(prompt, image_bytes) -> str` | `call_gemini()` | API call |

**C1 fix**: uses `headers={"x-goog-api-key": self._api_key}` instead of `params={"key": api_key}`.
The API key never appears in the URL.

### ValidationExtractor (`validation_extractor.py`)

Groups the formula resolution functions that all share `ws_full` + `wb_full`.

| Method | From function | Purpose |
|--------|--------------|---------|
| `__init__(ws_full, wb_full)` | — | stores worksheet + workbook |
| `extract() -> list[dict]` | `extract_resolved_validations()` | main entry |
| `_try_resolve_formula(formula)` | `_try_resolve_formula()` | private |
| `_resolve_range(formula)` | `_resolve_range()` | private |
| `_resolve_named_range(defn)` | `_resolve_named_range()` | private |

### Free functions (no state to bundle)

- **`prompt_builder.py`**: `SCHEMA_FORMAT_SPEC` constant (~280 lines) + `build_prompt(text_repr, sheet_name, user_guidelines)` — pure text assembly.
- **`response_parser.py`**: `parse_structure_json(raw_text)` — extracts JSON from LLM response.

### Orchestrator (`pipeline.py`)

```python
@pipeline("ai-structure-proposal")
@parameter("excel_file", ...)
@parameter("gemini_connection", ...)
@parameter("sheet_name", ...)
@parameter("user_guidelines", ...)
def ai_structure_proposal(excel_file, gemini_connection, sheet_name=None, user_guidelines=None):
    reader = ExcelReader(excel_file.path)
    ws_val, ws_full = reader.select_sheets(sheet_name)

    text_repr = reader.build_text_representation(ws_val, ws_full)

    renderer = SheetRenderer()
    image = renderer.render(ws_val)

    extractor = ValidationExtractor(ws_full, reader.wb_full)
    resolved_dvs = extractor.extract()
    reader.close()   # close wb_full early (no longer needed)

    prompt = build_prompt(text_repr, ws_val.title, user_guidelines)

    client = GeminiClient(api_key=gemini_connection.api_key)
    raw = client.generate(prompt, image)

    proposal = parse_structure_json(raw)
    proposal["resolved_validations"] = resolved_dvs
    # ... add metadata, save JSON ...
```

---

## Pipeline 2: `schema_validation/`

**Current**: 1,247 lines in `pipeline.py`. This is the most complex pipeline — it
processes each section type differently and assembles the final schema.

### New file structure
```
schema_validation/
  pipeline.py              ← orchestrator (~100 lines)
  text_helpers.py          ← to_canonical, to_column_canonical, col_idx, etc.
  dv_resolver.py           ← DataValidationResolver class
  section_processors.py    ← SectionProcessor ABC + 3 concrete processors
  schema_builder.py        ← SchemaBuilder class
  workspace.yaml           ← unchanged
```

### text_helpers.py (free functions)

Stateless utilities used by multiple classes within this pipeline:

- `to_canonical(text) -> str`
- `to_column_canonical(text) -> str` — strips parentheticals first
- `col_idx(letter) -> int` / `col_letter(idx) -> str`
- `cell_ref_to_col(ref) -> str` / `cell_ref_to_row(ref) -> int`
- `format_example(value) -> str`
- `infer_type(value) -> str`

### DataValidationResolver (`dv_resolver.py`)

Groups the 12+ DV-related functions that all share `ws`, `wb`, and the accumulated `dv_map`.

| Method | From function | Purpose |
|--------|--------------|---------|
| `__init__(ws_full, wb_full)` | — | stores ws/wb, initializes empty maps |
| `extract()` | `extract_data_validations()` | builds `self.dv_map` |
| `resolve_indirect_dependencies()` | `resolve_indirect_dependencies()` | builds `self.conditional_map` |
| `merge_proposal_validations(proposal)` | `_merge_proposal_validations()` | enriches dv_map with pre-resolved DVs |
| **Properties** | | |
| `dv_map -> dict` | — | col_letter → validation info |
| `conditional_map -> dict` | — | parent_col → child values mapping |
| **Private** | | |
| `_resolve_dv_formula(formula)` | `_resolve_dv_formula()` | |
| `_resolve_named_range(name)` | `_resolve_named_range()` | |
| `_resolve_named_range_variants(name)` | `_resolve_named_range_variants()` | |
| `_resolve_indirect(formula)` | `_resolve_indirect()` | |
| `_resolve_range_values(formula)` | `_resolve_range_values()` | |
| `_parse_indirect_parent_col(formula)` | `_parse_indirect_parent_col()` | static |
| `_find_column_values(col_letter)` | `_find_column_values()` | reads self.dv_map |
| `_expand_sqref(sqref)` | `_expand_sqref()` | static |

### SectionProcessor hierarchy (`section_processors.py`)

This is the core of the refactoring. The three section types share the same
abstract interface, making it clear what each section type does at each stage.

```python
class SectionProcessor(ABC):
    """Base class for processing a detected section into schema entries."""

    def __init__(self, ws, dv_map: dict):
        self.ws = ws
        self.dv_map = dv_map

    @abstractmethod
    def register(self, section: dict) -> dict:
        """Declare fields for this section.

        For key_value:         returns label_fields + value_fields entries
        For grouped_key_value: returns label_fields + grouped value_fields
        For records:           returns header_module entry with ColumnDefs
        """

    @abstractmethod
    def qualify(self, registered: dict) -> dict:
        """Enrich registered fields with types, examples, resolved enums.

        Reads actual cell values from self.ws and validation info from self.dv_map.
        """

    @abstractmethod
    def add_validation(self, qualified: dict) -> dict:
        """Attach validation rules to each field.

        For key_value:         required flag, type checks
        For grouped_key_value: required_rule (anchor/if_active)
        For records:           column-level enum/json_logic, primary key
        """
```

**KeyValueProcessor** — simplest section type.

| Method | From function | Notes |
|--------|--------------|-------|
| `register(section)` | first half of `process_key_value()` | declares label_fields + value_fields |
| `qualify(registered)` | second half of `process_key_value()` | reads cell values, infers types |
| `add_validation(qualified)` | embedded in `process_key_value()` | attaches DV enum if present |

**GroupedKeyValueProcessor** — medium complexity (anchor row logic).

| Method | From function | Notes |
|--------|--------------|-------|
| `register(section)` | first half of `process_grouped_key_value()` | declares columns × rows structure |
| `qualify(registered)` | reads cells for each column slot | infers types from first non-empty |
| `add_validation(qualified)` | `required_rule` assignment | anchor for first row, if_active for rest |

**RecordsProcessor** — most complex (columns, groups, dynamic).

| Method | From function | Notes |
|--------|--------------|-------|
| `__init__(ws, dv_map, data_start_row, conditional_map)` | — | extra state for records |
| `register(section)` | first part of `process_records()` | builds column tree from proposal |
| `qualify(registered)` | `_build_fixed_column()` + `_build_group_column()` | resolves DVs, types, examples |
| `add_validation(qualified)` | enum attachment, PK detection | attaches column-level + conditional rules |
| `_build_fixed_column(col_def)` | `_build_fixed_column()` | private |
| `_build_group_column(col_def)` | `_build_group_column()` | private, handles sub-columns (3 levels max) |

### SchemaBuilder (`schema_builder.py`)

Orchestrates all section processors and assembles the final `schema_validation.json`.

| Method | From function | Purpose |
|--------|--------------|---------|
| `__init__(proposal, ws, dv_resolver)` | — | stores inputs |
| `build() -> dict` | `generate_schema()` | main entry — iterates sections, dispatches to processors |
| `_build_table_definitions(...)` | `build_table_definitions()` | metadata + records table defs |
| `_build_validation_config(start_row)` | `build_validation_config()` | comparison settings |
| `_find_primary_key(hdr, proposal)` | `_find_primary_key()` | PK detection with fallback |
| `_deduplicate_canonicals(header_modules)` | `_deduplicate_canonicals()` | **+ duplicate header warning** |
| `_collect_canonicals(...)` | `_collect_canonicals()` | helper |
| `_find_parent_module(meta_keys)` | `_find_parent_module()` | finds metadata parent |

**Duplicate header detection**: `_deduplicate_canonicals()` is enhanced to log a warning
when two columns share the same canonical name without a distinguishing parent (module key
or group parent). It still auto-prefixes to resolve the collision, but now the user is alerted.

**Multiple records tables**: `build()` assigns each `records` section its own table name
using `section.get("key", "program_data")`. The table definitions dict can hold multiple
records table entries. Current behavior (one table called `"program_data"`) is the default fallback.

### Orchestrator (`pipeline.py`)

```python
@pipeline("schema-validation")
@parameter("structure_proposal_file", ...)
@parameter("excel_file", ...)
@parameter("sheet_name", ...)
def schema_validation(structure_proposal_file, excel_file, sheet_name=None):
    proposal = load_proposal(structure_proposal_file)
    wb, wb_full = load_workbooks(excel_file.path)
    ws, ws_full = select_sheet(wb, wb_full, proposal, sheet_name)

    resolver = DataValidationResolver(ws_full, wb_full)
    resolver.extract()
    resolver.resolve_indirect_dependencies()
    resolver.merge_proposal_validations(proposal)
    wb_full.close()

    builder = SchemaBuilder(proposal, ws, resolver)
    schema = builder.build()

    save_schema(schema)
```

---

## Pipeline 3: `xls_validation/`

**Current**: 1,459 lines in `validators.py` + 497 lines in `pipeline.py`

### New file structure
```
xls_validation/
  pipeline.py              ← orchestrator (report building, fingerprint, sheet loop)
  text_helpers.py          ← normalize_text, to_canonical, text_matches
  cell_helpers.py          ← read_cell, read_cell_rc, is_row_empty, make_issue, find_data_rows
  json_logic.py            ← evaluate_json_logic + helpers
  section_validators.py    ← SectionValidator ABC + KeyValueValidator, GroupedValidator, RecordsValidator
  extraction_guide.py      ← ExtractionGuideBuilder class
  workspace.yaml           ← unchanged
```

### text_helpers.py (free functions)

- `normalize_text(text, config) -> str`
- `to_canonical(text) -> str`
- `text_matches(actual, expected, variants, config) -> str`
- `_remove_accents(text) -> str` (private)

### cell_helpers.py (free functions)

- `read_cell(ws, cell_ref) -> value`
- `read_cell_rc(ws, row, col_letter) -> value`
- `cell_ref(col_letter, row) -> str`
- `make_issue(severity, cell, message, field_ref, group) -> dict`
- `find_data_rows(ws, schema) -> list[int]`
- `_is_row_empty(ws, row) -> bool` (private)

### json_logic.py (free functions)

- `evaluate_json_logic(rule, data) -> value`
- `_jl_var(data, path)` (private)
- `_jl_short_circuit(op, raw_args, data)` (private)

### SectionValidator hierarchy (`section_validators.py`)

Mirrors the section processor hierarchy from pipeline 2 — same section types, same
abstract interface, but for validation instead of schema generation.

```python
class SectionValidator(ABC):
    """Base class for validating a section type against its schema."""

    def __init__(self, ws, config: dict):
        self.ws = ws
        self.config = config

    @abstractmethod
    def validate_structure(self, section_schema: dict) -> list[dict]:
        """Check that labels/headers match expectations.

        For key_value:         check label cells exist at expected positions
        For grouped_key_value: check label cells + detect active columns
        For records:           check header row(s), build column_map
        """

    @abstractmethod
    def validate_data(self, section_schema: dict, **context) -> list[dict]:
        """Check that values/rows satisfy rules.

        For key_value:         check value cells (type, required, enum)
        For grouped_key_value: check grouped values (anchor + if_active)
        For records:           check data rows (PK, enums, json_logic)
        """
```

**KeyValueValidator**

| Method | From function | Purpose |
|--------|--------------|---------|
| `validate_structure(schema)` | `validate_labels()` | checks label_fields |
| `validate_data(schema)` | `validate_values()` (single values) | checks value_fields |
| `_check_label(field_def, field_ref)` | `_check_label()` | private |
| `_validate_single_value(key, field)` | `_validate_single_value()` | private |

**GroupedKeyValueValidator**

| Method | From function | Purpose |
|--------|--------------|---------|
| `validate_structure(schema)` | `validate_labels()` (grouped part) | checks grouped labels |
| `validate_data(schema)` | `_validate_grouped_values()` | anchor detection, active columns, if_active rules |
| `_validate_grouped_values(grouped_def, section_key)` | `_validate_grouped_values()` | private |

**RecordsValidator** — the most complex validator.

| Method | From function | Purpose |
|--------|--------------|---------|
| `validate_structure(schema)` | `validate_headers()` + `validate_sections_headers()` | checks header rows, builds column_map |
| `validate_data(schema, column_map, module_columns)` | `validate_data_rows()` + `_validate_data_rows_v2()` | checks data rows |
| `column_map -> dict` | — | built during `validate_structure()`, maps canonical → col_letter |
| `module_columns -> dict` | — | built during `validate_structure()`, maps mod_key → col_letters |
| `_validate_fixed_header(col_def, template)` | `_validate_fixed_header()` | private |
| `_validate_group_header(group_def, cursor, ...)` | `_validate_group_header()` + `_validate_v2_group_header()` | merged, handles sub-columns |
| `_validate_dynamic_headers(dyn_def, cursor)` | `_validate_dynamic_headers()` + `_match_dynamic_columns()` | private |
| `_check_column_enum(col_entry, data_rows, mod_key)` | `_check_column_enum()` | private |
| `_validate_data_column(col_def, data_rows, pk)` | `_validate_v2_data_column()` | merged into main path |

**Note**: The current `is_v2_schema()` branching in `validate_sheet()` is absorbed. The
`RecordsValidator` handles both the `header_modules` format and the `sections` format
internally — detected by the schema structure, not by a version label. The `_v2_` prefixed
functions are merged into the corresponding methods above.

### ExtractionGuideBuilder (`extraction_guide.py`)

Moved from `pipeline.py` — builds the extraction guide that tells the ingest pipeline where
to find each canonical field in the Excel file.

| Method | From function | Purpose |
|--------|--------------|---------|
| `__init__(schema, column_map, module_columns, data_rows, grouped_info)` | — | stores context |
| `build() -> dict` | `build_extraction_guide()` | main entry |
| `_build_metadata_guide()` | `_build_metadata_guide()` | KV + grouped metadata |
| `_walk_column_entry(columns, col_def)` | `_walk_column_entry()` + `_walk_v2_column_entry()` | merged |

### Orchestrator (`pipeline.py`)

```python
def validate_sheet(ws, schema, config):
    issues = []

    # Key-value sections (labels + single values)
    kv = KeyValueValidator(ws, config)
    issues.extend(kv.validate_structure(schema))
    issues.extend(kv.validate_data(schema))

    # Grouped key-value sections
    gkv = GroupedKeyValueValidator(ws, config)
    issues.extend(gkv.validate_structure(schema))
    issues.extend(gkv.validate_data(schema))

    # Records sections (headers + data rows)
    rec = RecordsValidator(ws, config)
    issues.extend(rec.validate_structure(schema))
    data_rows = find_data_rows(ws, schema)
    issues.extend(rec.validate_data(schema,
        column_map=rec.column_map,
        module_columns=rec.module_columns))

    return issues, rec.column_map, rec.module_columns
```

---

## Pipeline 4: `xls_ingest/` — not refactored

At 290 lines, already manageable. Skip unless requested.

---

## Summary: function → class mapping

### Pipeline 1 — ai_structure_proposal (989 lines → 7 files)

| Current function | New location |
|-----------------|-------------|
| `build_text_representation()` | `ExcelReader.build_text_representation()` |
| `_extract_hex_color/bg/fg()` | `ExcelReader._extract_hex_color/bg/fg()` (static) |
| `render_sheet_image()` | `SheetRenderer.render()` |
| `_load_fonts()` | `SheetRenderer._load_fonts()` |
| `call_gemini()` | `GeminiClient.generate()` ★ C1 fix |
| `parse_structure_json()` | `response_parser.parse_structure_json()` (free function) |
| `SCHEMA_FORMAT_SPEC` | `prompt_builder.SCHEMA_FORMAT_SPEC` (constant) |
| `build_prompt()` | `prompt_builder.build_prompt()` (free function) |
| `extract_resolved_validations()` | `ValidationExtractor.extract()` |
| `_try_resolve_formula/range/named()` | `ValidationExtractor._*()` (private) |
| `_select_sheets()` | `ExcelReader.select_sheets()` |
| `ai_structure_proposal()` | `pipeline.ai_structure_proposal()` (thin orchestrator) |

### Pipeline 2 — schema_validation (1,247 lines → 5 files)

| Current function | New location |
|-----------------|-------------|
| `to_canonical()`, `to_column_canonical()`, etc. | `text_helpers.*` (free functions) |
| `extract_data_validations()` | `DataValidationResolver.extract()` |
| `resolve_indirect_dependencies()` | `DataValidationResolver.resolve_indirect_dependencies()` |
| `_resolve_dv_formula/named/indirect/range()` | `DataValidationResolver._*()` (private) |
| `_merge_proposal_validations()` | `DataValidationResolver.merge_proposal_validations()` |
| `_expand_sqref()` | `DataValidationResolver._expand_sqref()` (static) |
| `process_key_value()` | `KeyValueProcessor.register/qualify/add_validation()` |
| `process_grouped_key_value()` | `GroupedKeyValueProcessor.register/qualify/add_validation()` |
| `process_records()` | `RecordsProcessor.register/qualify/add_validation()` |
| `_build_fixed_column()` | `RecordsProcessor._build_fixed_column()` |
| `_build_group_column()` | `RecordsProcessor._build_group_column()` |
| `generate_schema()` | `SchemaBuilder.build()` |
| `build_table_definitions()` | `SchemaBuilder._build_table_definitions()` |
| `build_validation_config()` | `SchemaBuilder._build_validation_config()` |
| `_find_primary_key()` | `SchemaBuilder._find_primary_key()` |
| `_deduplicate_canonicals()` | `SchemaBuilder._deduplicate_canonicals()` ★ + warning |
| `schema_validation()` | `pipeline.schema_validation()` (thin orchestrator) |

### Pipeline 3 — xls_validation (1,956 lines → 7 files)

| Current function | New location |
|-----------------|-------------|
| `normalize_text()`, `to_canonical()`, `text_matches()` | `text_helpers.*` |
| `read_cell()`, `read_cell_rc()`, `make_issue()` | `cell_helpers.*` |
| `evaluate_json_logic()` | `json_logic.evaluate_json_logic()` |
| `validate_labels()` + `_check_label()` | `KeyValueValidator.validate_structure()` |
| `validate_values()` + `_validate_single_value()` | `KeyValueValidator.validate_data()` |
| `_validate_grouped_values()` | `GroupedKeyValueValidator.validate_data()` |
| `validate_headers()` + `_validate_fixed/group/dynamic_header()` | `RecordsValidator.validate_structure()` |
| `validate_sections_headers()` + `_validate_v2_single/group_header()` | merged into `RecordsValidator.validate_structure()` |
| `validate_data_rows()` + `_check_column_enum()` | `RecordsValidator.validate_data()` |
| `_validate_data_rows_v2()` + `_validate_v2_data_column()` | merged into `RecordsValidator.validate_data()` |
| `find_data_rows()`, `_is_row_empty()` | `cell_helpers.*` |
| `is_v2_schema()` | **removed** — dispatch by section type, not schema version |
| `build_extraction_guide()` + helpers | `ExtractionGuideBuilder.build()` |
| `_walk_column_entry()` + `_walk_v2_column_entry()` | merged into `ExtractionGuideBuilder._walk_column_entry()` |

---

## Implementation Order

1. **Pipeline 1: `ai_structure_proposal/`** — medium complexity, includes C1 fix
2. **Pipeline 2: `schema_validation/`** — highest complexity (section processors + schema builder)
3. **Pipeline 3: `xls_validation/`** — largest file (1,459 → 7 files, merges v2 code)

Within each pipeline:
1. Create helper files (`text_helpers.py`, `cell_helpers.py`)
2. Create domain classes (extractors, processors, validators)
3. Refactor `pipeline.py` to import and orchestrate the new classes
4. Verify `python -c "from pipeline import *"` works

## Verification checklist

- [ ] `@pipeline` and `@parameter` decorators remain on the main function in each `pipeline.py`
- [ ] No cross-directory imports exist
- [ ] Output filenames and JSON structures are identical
- [ ] All current functions have a home in the new structure (see mapping tables above)
- [ ] Duplicate header detection emits warning in pipeline 2
- [ ] GeminiClient uses `x-goog-api-key` header (C1 fix)
