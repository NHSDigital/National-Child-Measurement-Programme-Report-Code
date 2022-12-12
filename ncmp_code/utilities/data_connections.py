"""
Purpose of script: handles reading data in from sql.
"""
import sqlalchemy as sa
import pandas as pd
import logging

logger = logging.getLogger(__name__)


def df_from_sql(query, server, database) -> pd.DataFrame:
    """
    Use sqlalchemy to connect to the server and database with the help
    of mssql and pyodbc packages

    Inputs:
        server: server name
        database: database name
        query: string containing a sql query

    Output:
        pandas Dataframe
    """
    conn = sa.create_engine(
        f"mssql+pyodbc://{server}/{database}?driver=SQL+Server",
        fast_executemany=True)
    conn.execution_options(autocommit=True)
    logger.info(f"Getting dataframe from SQL database {database}")
    logger.info(f"Running query:\n\n {query}")
    df = pd.read_sql_query(query, conn)
    return df

    print("This message shows that you have successfully imported \
the get_df_from_sql() function from the data connections module")
