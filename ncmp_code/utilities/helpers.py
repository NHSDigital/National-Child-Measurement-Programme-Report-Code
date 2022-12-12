def check_update_column(df, df_compare, check_column, check_compare_column,
                        check_type, flag_value):
    """
    Appends value to contents of an existing column in a dataframe,
    based on the result of a check of values in the specified check_column
    against a set of values in the specified check_content_column,
    and the check type specified

    e.g. if value in check_column = "Breach", value_if_false = "_CHECK!",
    check_type = False and value isn't present in check_compare_column
    then the new value in check_column would be "Breach_CHECK!"

    Currently only works for string values

    Parameters
    ----------
    df : pandas.DataFrame
        Dataframe containing column to be checked
    df_compare: pandas.Dataframe
        Dataframe containing column to be compared against
    check_column: str
        Name of the existing dataframe column containing the values to be
        checked.
    check_compare_column: str
        Name of the existing dataframe column containing the values to be
        checked against
    check_type: bool
        Set to True to append flag_value when value in check_column is in
        check_compare_column
        Set to False to append when value is not found in check_compare_column
    flag_value: str
        Value that will be appended if the check result matches the
        check_type selected. Will not append if value in check_column is blank

    Returns
    -------
    df : pandas.DataFrame
        with amended column
    """
    # Get the data types of the columns and flag_value
    format_flag_value = type(flag_value)
    format_check_column = type(check_column)
    format_check_compare_column = type(check_compare_column)

    # Raise an error if any of the inputted values/columns are not string
    if (format_flag_value != str) | (format_check_column != str) | (format_check_compare_column != str):
        raise ValueError("One or more of the inputs (flag_value, check_column, check_compare_column) are not in a string format")

    if check_type is True:
        # appends flag_value where value found in check_compare_column
        df.loc[(df[check_column].isin(df_compare[check_compare_column])) &
               (df[check_column] != ""), check_column] = (df[check_column] +
                                                          flag_value)

    if check_type is False:
        df.loc[(~df[check_column].isin(df_compare[check_compare_column])) &
               (df[check_column] != ""), check_column] = (df[check_column] +
                                                          flag_value)

    return df
