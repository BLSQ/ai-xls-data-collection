import polars as pl
from sqlalchemy import create_engine


def read_table(uri: str, table_name: str) -> pl.DataFrame:
    """Read a full table from a PostgreSQL database into a Polars DataFrame."""
    return pl.read_database_uri(query=f'SELECT * FROM "{table_name}"', uri=uri)


def write_table(df: pl.DataFrame, table_name: str, uri: str) -> None:
    """Replace a table in the target database with the given DataFrame."""
    engine = create_engine(uri)
    try:
        df.write_database(
            table_name=table_name,
            connection=engine,
            if_table_exists="replace",
        )
    finally:
        engine.dispose()
