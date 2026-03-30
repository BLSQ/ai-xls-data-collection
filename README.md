# Setting Up a New Workspace — Step-by-Step Guide

This guide walks through deploying the Excel Data Collection & Validation system in a fresh OpenHEXA workspace. By the end, the workspace will be ready for organisations to upload Excel files, validate them against an AI-generated schema, and ingest the data into the workspace database.

> **Prerequisite**: You must have **admin** permissions on the target OpenHEXA workspace.

---

## 1. Add All Pipelines to the Workspace

Push or import the four core pipelines into your OpenHEXA workspace. Each pipeline lives in its own directory in this repository:

| # | Pipeline code | Directory | Purpose |
|---|--------------|-----------|---------|
| 1 | `ai-structure-proposal` | `ai_structure_proposal/` | Gemini AI analyses the Excel template layout and produces `structure_proposal.json` |
| 2 | `schema-validation` | `schema_validation/` | Deterministic conversion of the AI proposal into `schema_validation.json` |
| 3 | `xls-validation` | `xls-validation/` | Validates uploaded Excel files against the schema, produces `validation_report.json` + `extraction_guide.json` |
| 4 | `xls-ingest` | `xls_ingest/` | Ingests validated Excel data into `program_metadata` and `program_data` database tables |

**How to add each pipeline:**
1. In the OpenHEXA workspace, go to **Pipelines** > **Create pipeline**
2. Connect the Git repository (or push directly via `openhexa CLI`)
3. Ensure the pipeline code matches exactly (e.g. `ai-structure-proposal`, not `ai_structure_proposal`)
4. Verify each pipeline appears in the workspace pipeline list and can be triggered manually

---

## 2. Create the Gemini API Connection

The `ai-structure-proposal` pipeline requires a **Custom Connection** to access the Gemini API (used for multimodal Excel analysis).

1. Go to **Workspace Settings** > **Connections** > **New connection**
2. Select type: **Custom**
3. Fill in:
   - **Connection identifier**: `gemini_api_key` (this name matters — it matches the `workspace.yaml`)
   - **Field `api_key`**: paste your Google Gemini API key (obtain from [Google AI Studio](https://aistudio.google.com/apikey))
4. Save the connection
5. Go to **Pipelines** > `ai-structure-proposal` > **Settings**
6. Set the **Gemini API Connection** parameter default to the `gemini_api_key` connection you just created

> **Note**: Only the `ai-structure-proposal` pipeline needs this connection. The other three pipelines have no external API dependencies.

---

## 3. Deploy the Admin Webapp

1. In the OpenHEXA workspace, go to **Webapps** > **Create webapp**
2. Upload the file `webapp/index.html`
3. Open the webapp — it will:
   - Automatically detect the workspace slug via the GraphQL API
   - Use cookie-based authentication (no token needed)
   - Show you as **Admin** (if you have `manageMembers` permission)

> **First-time behaviour**: On a fresh workspace there is no `schema_validation.json` yet. The left panel will show *"No schema found yet — use Update Template above to generate one"*. This is normal.

---

## 4. Import a Template and Generate the Schema

This is the core setup step: teaching the system your Excel template structure.

### 4.1 Prepare Your Excel Template

Your template file should be a clean, unfilled (or minimally filled) Excel workbook with:
- **Headers** — column names the system should expect
- **Example data** — at least one row of sample values so the AI can infer types
- **Data validation dropdowns** — these are automatically extracted as `enum` constraints

### 4.2 Write User Guidelines

Before running the AI analysis, prepare a short text note describing the template structure. This dramatically improves the AI's accuracy. Include:

- **What is metadata vs data**: *"Rows 3–12 are metadata key-value pairs (project name, budget, currency…). Data starts at row 17."*
- **Grouped key-value sections**: *"Rows 9–12, columns C–I, are a repeating group (one column per funder/donor). The number of donors varies between files."*
- **Fixed vs variable columns**: *"Columns B–Y are fixed (same in every file). Columns Z onward are province-level percentages — the provinces change depending on the country, so these columns are dynamic."*
- **Groups with sub-headers**: *"Row 14 has group headers ('Réalisé', 'Prévisionnel'). Row 15 has sub-columns under each group (years: 2024, 2025…)."*
- **Percentage rules**: *"Province percentages (columns Z onward) must sum to 100% per row."*
- **Primary key column**: *"Column B ('Intitulé budgétaire') is the primary key — each row must have a unique label."*

### 4.3 Generate the Schema (Admin Only)

In the admin webapp:

1. Expand the **Update Template** section in the left sidebar
2. Click **Select template file** and choose your Excel template
3. Click **Add guidelines for AI** and paste your user guidelines text
4. Click **Generate Schema (AI)**

This triggers a two-step process:
- **Step 1** — `ai-structure-proposal`: Gemini analyses the template (takes ~30–60 seconds). Produces `structure_proposal.json`.
- **Step 2** — `schema-validation`: Deterministic pipeline converts the proposal into `schema_validation.json`, enriched with example values and dropdown validations extracted from the Excel file.

Once complete, the schema appears in the left sidebar with all detected modules, fields, and validation rules.

### 4.4 Review and Refine the Schema

After generation, review the schema in the webapp:
- **Duplication Rules** - select if the submissions ingestion should be enforced at metadata or data level (go to duplicates settings in the left panel)
- **Check field types** — click any column card to edit: change `string` ↔ `number` ↔ `date`
- **Mark required fields** — toggle the "Required" checkbox for mandatory columns
- **Edit expected headers** — fix any AI misdetections (typos in labels are preserved; fix them here if needed)
- **Add/remove columns** — use the + buttons to add columns or the trash icon to remove
- **Drag to reorder** — columns can be dragged to reorder or moved into/out of groups
- **Add validation rules** — click cells in the Excel preview to open the Rule Builder and create JsonLogic formulas (e.g. `SUM(provinces) == 100`)
- **Save** — click the **Save** button to upload the edited `schema_validation.json` back to the workspace

> For detailed schema field reference (required_rule, JsonLogic syntax, enum constraints, etc.), see [`data_structures_reference.md`](data_structures_reference.md).
>
> For the full pipeline workflow and architecture, see [`excel_data_collection_ai.md`](excel_data_collection_ai.md).

---

## 5. Configure Optional Settings (Admin Only)

Open **Pipeline Settings** (gear icon) in the webapp sidebar to configure:

### 5.1 Dashboard URL

If you have a dashboard (e.g. built with ECharts/gridstack on OpenHEXA), paste its URL here. This adds a **Dashboard** button in the webapp header for quick access.

- Leave empty to hide the button entirely
- Example: `https://app.demo.openhexa.org/workspaces/my-workspace/webapps/my-dashboard/play/`

### 5.2 Post-Ingestion Pipeline

If you need to compute indicators, create derived tables, or run any custom logic after data is ingested:

1. Check **Enable post-ingestion pipeline**
2. Enter the **Pipeline code** (e.g. `my-indicators`)
3. Optionally set a **Webhook URL** for faster triggering

This pipeline will run automatically after each successful "Save to Database" action.

**Typical use cases:**
- Compute indicators from `program_data` → write to an `indicators` table
- Generate aggregated tables optimised for the dashboard
- Export data to external systems

### 5.3 Webhook URLs (Core Pipelines)

For faster pipeline triggering, you can set webhook URLs for each core pipeline. Webhooks bypass the GraphQL pipeline resolution step and are self-authenticated.

To get a webhook URL: go to the pipeline in OpenHEXA > **Settings** > **Webhook** > copy the URL.

### 5.4 Save Configuration

- **Save to workspace** — stores `app_config.json` in the workspace bucket (shared across all users)
- **Local** — stores in browser localStorage only (for development/testing)

---

## 6. Test the Full Workflow

Once the schema is generated and reviewed:

1. **As a viewer** — open the webapp in a different browser/incognito (or as a non-admin user)
2. **Upload a filled Excel file** — drag & drop onto the upload zone
3. **Click Validate** — this triggers the `xls-validation` pipeline
4. **Review the report** — errors, warnings, and info messages grouped by category
5. **If no errors** — the **Save to Database** button appears
6. **Click Save to Database** — triggers the `xls-ingest` pipeline, writing to `program_metadata` and `program_data` tables
7. **Post-ingestion** — if configured, the indicators pipeline runs automatically

---

## Summary Checklist

| Step | Action | Who |
|------|--------|-----|
| 1 | Push 4 pipelines to workspace | Admin |
| 2 | Create `gemini_api_key` Custom Connection with `api_key` field | Admin |
| 3 | Deploy `webapp/index.html` as a webapp | Admin |
| 4 | Upload template + guidelines → Generate Schema → Review & Save | Admin |
| 5 | (Optional) Set dashboard URL and post-ingestion pipeline | Admin |
| 6 | Test: upload filled Excel → validate → ingest | Admin/Viewer |

---

## Troubleshooting

| Issue | Cause | Fix |
|-------|-------|-----|
| Role shows "Viewer" instead of "Admin" | Workspace permissions not detected | Check you have `manageMembers` permission on the workspace |
| "No schema found" in left panel | `schema_validation.json` not yet generated | Use **Update Template** to generate it (Step 4) |
| AI structure proposal fails | Missing or invalid Gemini API key | Check the `gemini_api_key` connection has a valid `api_key` field |
| Validation pipeline finds no schema | `schema_validation.json` not in workspace root | Ensure the schema was saved via the webapp (not just downloaded locally) |
| "Save to Database" button hidden | Validation report has errors | Fix errors in the Excel file and re-upload |
| Dashboard button not visible | No dashboard URL configured | Set it in Pipeline Settings (Step 5.1) |
