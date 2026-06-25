from pathlib import Path

import polars as pl
from config import DONOR_IDS, FRENCH_MONTHS, META_JOIN_COLS, PROVINCE_MAPPING
from dateutil.relativedelta import relativedelta
from db_ops import read_table, write_table  # type: ignore
from openhexa.sdk import File, PostgreSQLConnection, current_run, parameter, pipeline, workspace


@pipeline("aedes-compute-indicators")
@parameter(
    "target_connection",
    name="Target PostgreSQL connection",
    type=PostgreSQLConnection,
    default="oh-db-aedes-prod",
    required=False,
    help=(
        "Identifier of a PostgreSQL workspace connection to export the "
        "indicators to. Leave empty to write to the local workspace database."
    ),
)
@parameter(
    "attach_geometry",
    name="Attach organisation-unit geometry",
    type=bool,
    required=False,
    default=True,
    help=(
        "Join the geographic ventilation table with the pre-downloaded "
        "organisation-unit geometry. Disable to skip the geometry join."
    ),
)
@parameter(
    "org_units_file",
    name="Organisation units file",
    type=File,
    required=False,
    default="geometries/org_units.parquet",
    help=(
        "Path (relative to the workspace files) of the pre-downloaded "
        "organisation-unit file with the 'coordinates' column."
    ),
)
def aedes_compute_indicators(
    target_connection: PostgreSQLConnection | None = None,
    attach_geometry: bool = True,
    org_units_file: File | None = None,
):
    """Compute the AEDES indicator tables and export them to the database."""
    current_run.log_info("Starting AEDES compute-indicators pipeline...")

    source_uri = workspace.database_url
    target_uri = target_connection.url if target_connection else workspace.database_url
    org_units_path = org_units_file.path if org_units_file else ""

    meta_done = build_program_metadata_compile(source_uri, target_uri)
    build_contribution_bailleurs(target_uri, meta_done)
    build_dim_date(target_uri, meta_done)

    data_done = build_program_data_compile(source_uri, target_uri, meta_done)
    pivot_done = build_pivot_realise_previsionnel(target_uri, data_done)
    build_ventilation_geo(target_uri, attach_geometry, org_units_path, pivot_done)

    current_run.log_info("All indicator tables computed.")


@aedes_compute_indicators.task
def build_program_metadata_compile(source_uri: str, target_uri: str) -> bool:
    """Enrich program_metadata with the computed project end date."""
    df = read_table(source_uri, "program_metadata")

    df = df.with_columns(
        pl.col("date_debut_du_projet").cast(pl.Datetime),
    ).with_columns(
        pl.struct(["date_debut_du_projet", "duree_du_projet_en_mois"])
        .map_elements(_compute_end_date, return_dtype=pl.Datetime)
        .alias("date_fin_du_projet"),
    )

    write_table(df, "program_metadata_compile", target_uri)
    current_run.log_info(f"program_metadata_compile: {df.height} rows")
    return True


@aedes_compute_indicators.task
def build_contribution_bailleurs(target_uri: str, _dep: bool) -> bool:
    """Build the per-donor contribution table from the project metadata."""
    df = read_table(target_uri, "program_metadata_compile")

    id_cols = [
        "entity_id",
        "nom_du_projet_programme_intitule_du_budget",
        "nom_de_l_organisation",
        "date_debut_du_projet",
        "date_fin_du_projet",
        "sigle",
        "monnaie",
        "budget_total_du_projet",
    ]
    cols_nom = [f"donors_{i}_nom_du_bailleur" for i in DONOR_IDS]
    cols_montant = [f"donors_{i}_montant_mis_a_disposition" for i in DONOR_IDS]

    df_nom = df.unpivot(
        index=id_cols, on=cols_nom, variable_name="donor_id", value_name="bailleur"
    ).with_columns(pl.col("donor_id").str.extract(r"donors_(\d+)"))

    df_montant = df.unpivot(
        index=id_cols,
        on=cols_montant,
        variable_name="donor_id",
        value_name="contribution",
    ).with_columns(pl.col("donor_id").str.extract(r"donors_(\d+)"))

    df_part = (
        df_nom.join(df_montant, on=id_cols + ["donor_id"], how="inner")
        .drop_nulls(subset=["bailleur", "contribution"])
        .drop("donor_id")
        .with_columns(pl.col("contribution").cast(pl.Float64).alias("contribution"))
    )

    write_table(df_part, "program_metadata_contribution_bailleurs", target_uri)
    current_run.log_info(f"program_metadata_contribution_bailleurs: {df_part.height} rows")
    return True


@aedes_compute_indicators.task
def build_dim_date(target_uri: str, _dep: bool) -> bool:
    """Build a date dimension from the project start/end dates."""
    df = read_table(target_uri, "program_metadata_compile")

    df_dim = (
        df.select(["date_debut_du_projet", "date_fin_du_projet"])
        .with_columns(pl.col("date_debut_du_projet").cast(pl.Datetime))
        .select(
            pl.col("date_debut_du_projet"),
            pl.col("date_debut_du_projet")
            .map_elements(_french_period, return_dtype=pl.String)
            .alias("period"),
            (
                pl.col("date_debut_du_projet").dt.year() * 100
                + pl.col("date_debut_du_projet").dt.month()
            )
            .cast(pl.Int32)
            .alias("date_order"),
            pl.col("date_debut_du_projet").dt.year().cast(pl.Int32).alias("annee"),
            pl.col("date_fin_du_projet"),
        )
        .unique()
        .sort("date_order")
    )

    write_table(df_dim, "dim_date", target_uri)
    current_run.log_info(f"dim_date: {df_dim.height} rows")
    return True


@aedes_compute_indicators.task
def build_program_data_compile(source_uri: str, target_uri: str, _dep: bool) -> bool:
    """Enrich program_data with the project metadata columns."""
    df_data = read_table(source_uri, "program_data")
    df_meta = read_table(target_uri, "program_metadata_compile").select(META_JOIN_COLS)

    df = df_data.join(df_meta, on="entity_id", how="inner")

    write_table(df, "program_data_compile", target_uri)
    current_run.log_info(f"program_data_compile: {df.height} rows")
    return True


@aedes_compute_indicators.task
def build_pivot_realise_previsionnel(target_uri: str, _dep: bool) -> bool:
    """Unpivot the realised/forecast amount columns into a long format."""
    df = read_table(target_uri, "program_data_compile")

    index_cols = [
        "entity_id",
        "sigle",
        "date_debut_du_projet",
        "date_fin_du_projet",
        "sheet_name",
        "intitule_budgetaire_libelle_d_activite",
        "mode_de_gestion",
        "mode_de_mise_en_oeuvre",
        "type_de_depense",
        "detail_type_de_depense",
        "pilier",
        "precision_sur_le_piler",
        "thematique_principale",
        "thematique_secondaire",
    ]
    value_cols = [
        c
        for c in df.columns
        if (c.startswith("realise_") or c.startswith("previsionnel_")) and not c.endswith("_ligne")
    ]

    df_pivot = (
        df.unpivot(
            index=index_cols,
            on=value_cols,
            variable_name="realisation_ou_prevision",
            value_name="montant_depense",
        )
        .with_columns(
            [
                pl.col("montant_depense").replace("-", "0").cast(pl.Float64),
                pl.col("realisation_ou_prevision")
                .str.split("_")
                .list.get(0)
                .alias("type_realisation"),
                pl.col("realisation_ou_prevision")
                .str.split("_")
                .list.get(-1)
                .cast(pl.Int32)
                .cast(pl.Utf8)
                .str.strptime(pl.Date, "%Y")
                .alias("date_mise_en_oeuvre"),
            ]
        )
        .drop("realisation_ou_prevision")
    )

    write_table(df_pivot, "program_data_pivot_realise_previsionnel", target_uri)
    current_run.log_info(f"program_data_pivot_realise_previsionnel: {df_pivot.height} rows")
    return True


@aedes_compute_indicators.task
def build_ventilation_geo(
    target_uri: str,
    attach_geometry: bool,
    org_units_path: str,
    _dep: bool,
) -> bool:
    """Break down amounts by geographic level, optionally with org-unit geometry."""
    df_data = read_table(target_uri, "program_data_compile")
    df_pivot = read_table(target_uri, "program_data_pivot_realise_previsionnel")

    geo_index = [
        "entity_id",
        "sigle",
        "date_debut_du_projet",
        "date_fin_du_projet",
        "sheet_name",
        "intitule_budgetaire_libelle_d_activite",
    ]
    geo_cols = [
        c
        for c in df_data.columns
        if c.startswith("provincial_") or c in ("module_8_na", "national", "central")
    ]

    df_vent = (
        df_data.unpivot(
            index=geo_index,
            on=geo_cols,
            variable_name="level_name",
            value_name="pct_affectation",
        )
        .with_columns(
            [
                pl.col("level_name")
                .str.replace("module_8_", "")
                .str.replace("provincial_", "")
                .alias("level_name"),
                (pl.col("pct_affectation").cast(pl.Float64) / 100).alias("pct_affectation"),
            ]
        )
        .filter(pl.col("pct_affectation").is_not_null())
        .join(
            df_pivot,
            on=[
                "entity_id",
                "sigle",
                "date_debut_du_projet",
                "date_fin_du_projet",
                "sheet_name",
                "intitule_budgetaire_libelle_d_activite",
            ],
        )
        .with_columns(
            (pl.col("pct_affectation") * pl.col("montant_depense")).alias("montant_alloue")
        )
        .select(
            [
                "entity_id",
                "sigle",
                "date_debut_du_projet",
                "date_fin_du_projet",
                "date_mise_en_oeuvre",
                "sheet_name",
                "intitule_budgetaire_libelle_d_activite",
                "mode_de_gestion",
                "mode_de_mise_en_oeuvre",
                "type_de_depense",
                "detail_type_de_depense",
                "pilier",
                "precision_sur_le_piler",
                "thematique_principale",
                "thematique_secondaire",
                "type_realisation",
                "pct_affectation",
                "montant_depense",
                "montant_alloue",
                "level_name",
            ]
        )
    )

    df_geom = (
        _load_org_unit_geometry(org_units_path) if attach_geometry and org_units_path else None
    )

    if df_geom is not None:
        df_vent = df_vent.join(df_geom, on="level_name", how="left")
    else:
        current_run.log_info(
            "Geometry not attached — ventilation table written without coordinates."
        )

    write_table(df_vent, "program_data_ventilation_geo", target_uri)
    current_run.log_info(f"program_data_ventilation_geo: {df_vent.height} rows")
    return True


def _compute_end_date(row: dict):
    """Add the project duration (in months) to the start date."""
    start = row.get("date_debut_du_projet")
    duration = row.get("duree_du_projet_en_mois")
    if start is None or duration is None:
        return None
    try:
        months = int(float(duration))
    except (ValueError, TypeError):
        return None
    return start + relativedelta(months=months)


def _french_period(value) -> str | None:
    """Return a capitalised French 'Month Year' label, e.g. 'Janvier 2025'."""
    if value is None:
        return None
    return f"{FRENCH_MONTHS[value.month]} {value.year}"


def _load_org_unit_geometry(org_units_path: str) -> pl.DataFrame | None:
    """Load pre-downloaded organisation units with their coordinates."""
    path = Path(workspace.files_path) / org_units_path
    if not path.exists():
        current_run.log_warning(
            f"Organisation units file not found at '{path}' — skipping geometry attachment."
        )
        return None

    org_units = pl.read_parquet(path) if path.suffix.lower() == ".parquet" else pl.read_csv(path)
    current_run.log_info(f"Loaded organisation units from '{path}'.")

    df_geom = (
        org_units.filter(pl.col("level").is_in([1, 2]))
        .with_columns(
            pl.col("name").replace_strict(PROVINCE_MAPPING, default=None).alias("level_name")
        )
        .filter(pl.col("level_name").is_not_null())
        .select(["id", "name", "level_name", "coordinates"])
    )
    return df_geom


if __name__ == "__main__":
    aedes_compute_indicators()
