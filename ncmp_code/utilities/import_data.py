import pandas as pd
import datetime
import logging

import ncmp_code.utilities.data_connections as dbc
import ncmp_code.parameters as param

logger = logging.getLogger(__name__)


def import_asset_data(server, database, table, year):
    """
    This function will import data from the NCMP asset SQL database
    Uses the df_from_sql function

    Parameters
    ----------

    Returns
    -------
    pandas.DataFrame

    """

    sql_folder = r'ncmp_code\sql_code'

    with open(sql_folder + '\query_asset.sql', 'r') as sql_file:
        data = sql_file.read()

    # The parameters in the sql file - query_asset.sql
    # are replaced with our user defined parameters
    data = data.replace("<Database>", database)
    data = data.replace("<Table>", table)
    data = data.replace("<Year>", year)

    # Get SQL data
    df = dbc.df_from_sql(data, server, database)

    # If the extract date field is not null, filter for latest data
    if df["ExtractDate"].notnull().any():

        # replace any null extract dates with string before getting max
        df["ExtractDate"].fillna("", inplace=True)
        df = df[df["ExtractDate"] == max(df["ExtractDate"])]

    return df


def import_la_ref(server, database, year):
    """
    This function will import LA reference data

    Parameters
    ----------

    Returns
    -------
    pandas.DataFrame

    """

    sql_folder = r'ncmp_code\sql_code'

    with open(sql_folder + '\query_laref.sql', 'r') as sql_file:
        data = sql_file.read()

    # define the start and end period the la data should refer to
    # based on the year we run it for
    fy_start = str(datetime.date(int(year[:4]), 4, 1))
    fy_end = str(datetime.date(int(year[:4])+1, 3, 31))

    # plug in start and end date in sql query defined in the file
    data = data.replace("<FYStart>", fy_start)
    data = data.replace("<FYEnd>", fy_end)

    # Get SQL data
    df = dbc.df_from_sql(data, server, database)

    return df


def import_la_dq_file(ladq_file, ladq_cols_dtypes=param.LA_DQ_COLS_DTYPES,
                      sheet_name="Extract"):
    """
    This function will import LA DQ file
    containing participation rate data for each LA.

    Parameters
    ----------
    ladq_file: str
        LA data quality data containing participation rates for each LA
    ladq_cols_dtypes: dict
        dictionary of columns to be imported from LA DQ file and their data types
        default = parameter LA_DQ_COLS_DTYPES set in parameters.py
    sheet_name: str, default = "Extract"
        sheet name in Excel where the data is

    Returns
    -------
    pandas.DataFrame

    """
    logging.info(f"Importing LA DQ file")

    # read the ladq file containing the participation rate in 'Extract' sheet
    df = pd.read_excel(ladq_file, sheet_name)

    # make first row into column names
    df.columns = df.iloc[0, :]
    df.drop(df.index[0], inplace=True)

    # rename the org code column to match org code column in raw data
    df = df.rename(columns={"LocalAuthorityCode": "OrgCode"})

    # extract columns required
    df = df[list(ladq_cols_dtypes)].copy()

    # set data types of columns
    df = df.astype(ladq_cols_dtypes)

    return df
