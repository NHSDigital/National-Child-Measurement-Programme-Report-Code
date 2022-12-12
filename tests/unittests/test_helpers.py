import pandas as pd
import ncmp_code.utilities.helpers as helpers
import ncmp_code.parameters as params
import pytest
import numpy as np


def test_check_update_column_true():
    """
    Tests check_update_column function where check_type = True
    (flag if value is found in check_compare_column)

    """

    input_df_check = pd.DataFrame(
        {"value_to_check": ["standard response1", "standard response2",
                            "non standard response?"]}
        )

    input_df_compare = pd.DataFrame(
        {"value_to_compare": ["standard response1", "standard response2"]}
        )

    df = input_df_check
    df_compare = input_df_compare
    check_column = "value_to_check"
    check_compare_column = "value_to_compare"
    flag_value = "_CHECK!"
    check_type = True

    return_df = helpers.check_update_column(df, df_compare, check_column,
                                            check_compare_column, check_type,
                                            flag_value)

    expected = ["standard response1_CHECK!", "standard response2_CHECK!",
                "non standard response?"]

    actual = list(return_df[check_column])

    assert actual == expected, f"When checking for check_update_column, expected to find {expected} but found {actual}"


def test_check_update_column_false():
    """
    Tests check_update_column function where check_type = False
    (flag if value is not found in check_compare_column)

    """

    input_df_check = pd.DataFrame(
        {"value_to_check": ["standard response1", "standard response2",
                            "non standard response?"]}
        )

    input_df_compare = pd.DataFrame(
        {"value_to_compare": ["standard response1", "standard response2"]}
        )

    df = input_df_check
    df_compare = input_df_compare
    check_column = "value_to_check"
    check_compare_column = "value_to_compare"
    flag_value = "_CHECK!"
    check_type = False

    return_df = helpers.check_update_column(df, df_compare, check_column,
                                            check_compare_column, check_type,
                                            flag_value)

    expected = ["standard response1", "standard response2",
                "non standard response?_CHECK!"]

    actual = list(return_df[check_column])

    assert actual == expected, f"When checking for check_update_column, expected to find {expected} but found {actual}"
