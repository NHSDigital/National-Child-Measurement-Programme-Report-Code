"""
Purpose of the script: contains the Excel automation script.
"""
import math
import xlwings as xw
import ncmp_code.parameters as param
import logging


def col_to_number(s, write_cell):
    '''
    converts Excel column letters to numbers

    Parameters
    ----------
    s: str
        Excel column letter to convert to number
    write_cell: str
        cell to write data to, used to calculate order of letter

    Returns
    -------
    order of letter value: int
        number indicating which position to insert new column into dataframe
    '''

    if len(s) == 1:
        return (ord(s[0])) - (ord(write_cell[0]))
    if len(s) == 2:
        return (int(math.pow(26, len(s)-1)*(ord(s[0]) - ord('A') + 1)
                    + col_to_number(s[1:], write_cell)))


def insert_empty_columns(df, empty_cols, write_cell):
    '''
    Inserts empty columns into the dataframe at the specified locations

    Parameters
    ----------
    df : pandas.DataFrame
    empty_cols: list[str]
        A list of letters representing any empty (section seperator) excel
        columns in the worksheet. Empty columns will be inserted into the
        dataframe in these positions.
        Default is None
    write_cell: str
        cell to write data to, used for reference in col_to_number function

    Returns
    -------
    df : pandas.DataFrame
    '''

    # Add empty list to add the column location numbers that need an empty column
    col_numbers = []
    # Set a base adjustment value - used in for loop when adding multiple
    # empty columns to account for change in number of columns each time
    adjustment = 0
    # For each column letter, convert to a number and apply the adjustment
    for col in empty_cols:
        col_num = col_to_number(col, write_cell) - adjustment
        # Then add to the list and update the adjustment for the presence of
        # the extra column
        col_numbers.append(col_num)
        adjustment = adjustment + 1

    # Sort the column list descending so that the last columns are added first
    col_numbers.sort(reverse=True)

    # Insert the empty columns as per the col_numbers list
    for col_num in col_numbers:
        df.insert(col_num, col_num, "")

    return df


def write_to_excel_static(df, sheetname, empty_cols, write_cell, output_file):
    """
    Write data to an excel template. Assumes the table length remains constant.

    Parameters
    ----------
    df : pandas.DataFrame
    sheetname : str
        Name of the destination Excel worksheet.
    empty_cols: list[str]
        A list of letters representing any empty (section seperator) excel
        columns in the worksheet. Empty columns will be inserted into the
        dataframe in these positions.
        Default is None
    write_cell: str
        identifies the cell location in the Excel worksheet where the data
        will be pasted (top left of data)
    Returns
    -------
    df : pandas.DataFrame
    """

    logging.info("Writing data to specified output file")
    # Load the template and select the required table sheet
    wb = xw.Book(output_file)
    sht = wb.sheets[sheetname]
    sht.select()

    # Add empty columns where present in target Excel worksheet
    if empty_cols is not None:
        df = insert_empty_columns(df, empty_cols, write_cell)

    # Remove index and column headers and write to the write_cell
    df_values = df.values
    sht.range(write_cell).value = df_values


def write_to_excel_variable(df, sheetname, empty_cols, write_cell, output_file):
    """
    Write data to an excel template. Can accommodate dataframes where the
    number of rows may change e.g. LA data where the number of LAs may change
    each year

    Parameters
    ----------
    df : pandas.DataFrame
    sheetname : str
        Name of the destination Excel worksheet.
    empty_cols: list[str]
        A list of letters representing any empty (section seperator) excel
        columns in the worksheet. Empty columns will be inserted into the
        dataframe in these positions.
        Default is None
    write_cell: str
        identifies the cell location in the Excel worksheet where the data
        will be pasted (top left of data)
        This should be the first cell in the master file where the variable
        data currently exists as it also determines which row to delete first
        e.g. for LAs would be the first cell of the first row of LA data
    Returns
    -------
    df : pandas.DataFrame
    """

    logging.info("Writing data to specified output file")
    # Load the template and select the required table sheet
    wb = xw.Book(output_file)
    sht = wb.sheets[sheetname]
    sht.select()

    # Add empty columns where present in target Excel worksheet
    if empty_cols is not None:
        df = insert_empty_columns(df, empty_cols, write_cell)

    # Get row number of write cell
    firstrownum = int(''.join(filter(str.isdigit, write_cell)))

    # Get row number of last row of existing data
    lastrownum_current = sht.range(write_cell).end('down').row

    # Clear all existing data rows from write_cell to end of data
    delete_rows = str(firstrownum) + ":" + str(lastrownum_current)
    sht.range(delete_rows).delete()

    # Count number of rows in dataframe
    df_rowcount = len(df)

    # Create range for new set of rows and insert into sheet
    lastrownnum_new = firstrownum + df_rowcount - 1
    df_rowsrange = str(firstrownum) + ":" + str(lastrownnum_new)

    sht.range(df_rowsrange).insert(shift='down')

    # Write dataframe to the Excel sheet starting at the write_cell reference
    df_values = df.values
    sht.range(write_cell).value = df_values


def save_close_output(output_file):
    """
    Save updated output file and close Excel
    """
    logging.info(
        f"Saving and closing {output_file}"
        )

    # Open output file
    wb = xw.Book(output_file)

    # Select first sheet of workbook
    sht = wb.sheets[0]
    sht.select()

    # Save workbook
    wb.save()
    xw.apps.active.api.Quit()
