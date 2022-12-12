import pandas as pd
import ncmp_code.utilities.processing_steps as processing
import ncmp_code.parameters as params
import pytest
import numpy as np


def test_suppress_count_column():
    input_df = pd.DataFrame(
        {
            "org": ['org', 'org', 'org', 'org', 'org', 'org', 'org', 'org', 'org', 'org',
                    'org', 'org', 'org', 'org', 'grand_total'],
            "to_suppress": [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 13]
        }
    )

    actual = processing.suppress_count_column(input_df, input_df["to_suppress"])

    expected = pd.Series([0, "*", "*", "*", "*", "*", "*", "*",
                          10, 10, 10, 10, 10, 15, 13], name="to_suppress")

    pd.testing.assert_series_equal(actual, expected)


def test_suppress_percentage_column():
    input_df = pd.DataFrame(
        {
            "org": ['grand_total', 'grand_total', 'grand_total',
                    'grand_total', 'grand_total', 'grand_total',
                    'org', 'org', 'org', 'org', 'org'],
            "to_suppress": [0, 0, 20, 25, 20, 62.5, 0, 0, 20, 25, 20],
            "numerator": [0, 0, 1, 3, 8, 10, 0, 0, 1, 3, 8],
            "denominator": [1, 8, 5, 12, 40, 16, 1, 8, 5, 12, 40]
        }
    )

    actual = processing.suppress_percentage_column(input_df,
                                                   input_df["to_suppress"],
                                                   input_df["numerator"],
                                                   input_df["denominator"])

    expected = pd.Series([0, 0, 20, 25, 20, 62.5, '*',
                          0, '*', '*', 25], name="to_suppress", dtype='object')

    pd.testing.assert_series_equal(actual, expected)


def test_suppress_percentage_column_round_to_dp():
    input_df = pd.DataFrame(
        {
            "org": ['grand_total', 'grand_total', 'grand_total', 'grand_total',
                    'org', 'org', 'org', 'org'],
            "to_suppress": [0, 22.5, 62.5, 56.25, 0, 22.5, 62.5, 56.25],
            "numerator": [0, 9, 10, 9, 0, 9, 10, 9],
            "denominator": [39, 40, 16, 16, 39, 40, 16, 16]
        }
    )

    actual = processing.suppress_percentage_column(input_df, input_df["to_suppress"],
                                                   input_df["numerator"],
                                                   input_df["denominator"],
                                                   round_to_dp=0)

    expected = pd.Series([0, 22.5, 62.5, 56.25, 0, 25,
                          67, 67], name="to_suppress", dtype='object')

    pd.testing.assert_series_equal(actual, expected)


def test_add_bmi_categories():
    input_df = pd.DataFrame(
        {"BmiPScore": [0, 0.02, 0.021, 0.11, 0.85, 0.90, 0.95, 0.967, 0.996, 1.01]
         }
        )
    actual = processing.add_bmi_categories(input_df, params.BMI_CATEGORIES)

    expected = pd.DataFrame(
        {"BmiPScore": [0, 0.02, 0.021, 0.11, 0.85, 0.90, 0.95, 0.967, 0.996, 1.01],
         "Underweight": [1, 1, np.nan, np.nan, np.nan, np.nan,
                         np.nan, np.nan, np.nan, np.nan],
         "Healthy weight": [np.nan, np.nan, 1, 1, np.nan, np.nan, np.nan, np.nan,
                            np.nan, np.nan],
         "Overweight": [np.nan, np.nan, np.nan, np.nan, 1, 1, np.nan, np.nan,
                        np.nan, np.nan],
         "Obese": [np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, 1, 1, 1, 1],
         "Severely obese": [np.nan, np.nan, np.nan, np.nan, np.nan, np.nan,
                            np.nan, np.nan, 1, 1],
         "Overweight and obese": [np.nan, np.nan, np.nan, np.nan, 1, 1, 1, 1, 1, 1]
         }
        )

    pd.testing.assert_frame_equal(actual, expected)


def test_replace_null_in_col():
    input_df = pd.DataFrame(
        {
            "Column1": ['org', 'org', 'org', 'org', 'org', 'org', 'org',
                        'org', 'org', 'org'],
            "Column2": [1, 2, np.nan, 4, 5, 6, np.nan, 8, 9, np.nan],
            "Column3": ['a', np.nan, 'c', 'd', np.nan, 'f', 'g', np.nan, 'i', np.nan]
            }
        )
    replace_null_logic = {'Column2': 'Not known',
                          'Column3': 'Not Stated'}
    actual = processing.replace_null_in_col(input_df, replace_null_logic)

    expected = pd.DataFrame(
        {"Column1": ['org', 'org', 'org', 'org', 'org', 'org', 'org',
                     'org', 'org', 'org'],
         "Column2": [1, 2, 'Not known', 4, 5, 6, 'Not known', 8, 9, 'Not known'],
         "Column3": ['a', 'Not Stated', 'c', 'd', 'Not Stated', 'f', 'g',
                     'Not Stated', 'i', 'Not Stated']
         }

        )

    pd.testing.assert_frame_equal(actual, expected)


def test_replace_null_in_col_none():
    input_df = pd.DataFrame(
        {
            "Column1": ['org', 'org', 'org', 'org', 'org', 'org', 'org',
                        'org', 'org', 'org'],
            "Column2": [1, 2, np.nan, 4, 5, 6, np.nan, 8, 9, np.nan],
            "Column3": ['a', np.nan, 'c', 'd', np.nan, 'f', 'g', np.nan, 'i', np.nan]
            }
        )
    replace_null_logic = None
    actual = processing.replace_null_in_col(input_df, replace_null_logic)

    expected = pd.DataFrame(
        {
            "Column1": ['org', 'org', 'org', 'org', 'org', 'org', 'org',
                        'org', 'org', 'org'],
            "Column2": [1, 2, np.nan, 4, 5, 6, np.nan, 8, 9, np.nan],
            "Column3": ['a', np.nan, 'c', 'd', np.nan, 'f', 'g', np.nan, 'i', np.nan]
            }
        )

    pd.testing.assert_frame_equal(actual, expected)


def test_row_total_both_true():
    input_df = pd.DataFrame(
            {"Column1": ["Org1", "Org1", "Org2", "Org2", "Org2"],
             "Column2": ["Org1", "Org2", "Org3", "Org4", "Org5"],
             "Underweight": [2, 3, 5, 6, 1],
             "Healthy weight": [1, 2, 4, 6, 7],
             "Overweight": [3, 4, 3, 4, 5],
             "Obese": [4, 5, 8, 6, 5],
             "Severely obese": [0, 1, 2, 5, 6],
             "Overweight and obese": [1, 3, 5, 6, 7]
             }
            )
    input_df = input_df.set_index(["Column1", "Column2"])
    total = True
    subtotal = True
    actual = processing.apply_row_totals(input_df, total, subtotal)

    expected = pd.DataFrame(
        {"Column1": ["Org1", "Org1", "Org1", "Org2",
                     "Org2", "Org2", "Org2", "grand_total"],
         "Column2": ["Org1", "Org2", "Org1 - subtotal",
                     "Org3", "Org4", "Org5", "Org2 - subtotal", " "],
         "Underweight": [2, 3, 5, 5, 6, 1, 12, 17],
         "Healthy weight": [1, 2, 3, 4, 6, 7, 17, 20],
         "Overweight": [3, 4, 7, 3, 4, 5, 12, 19],
         "Obese": [4, 5, 9, 8, 6, 5, 19, 28],
         "Severely obese": [0, 1, 1, 2, 5, 6, 13, 14],
         "Overweight and obese": [1, 3, 4, 5, 6, 7, 18, 22]
         }
        )
    expected = expected.set_index(["Column1", "Column2"])
    pd.testing.assert_frame_equal(actual, expected)


def test_row_total_both_false():
    input_df = pd.DataFrame(
            {"Column1": ["Org1", "Org1", "Org2", "Org2", "Org2"],
             "Column2": ["Org1", "Org2", "Org3", "Org4", "Org5"],
             "Underweight": [2, 3, 5, 6, 1],
             "Healthy weight": [1, 2, 4, 6, 7],
             "Overweight": [3, 4, 3, 4, 5],
             "Obese": [4, 5, 8, 6, 5],
             "Severely obese": [0, 1, 2, 5, 6],
             "Overweight and obese": [1, 3, 5, 6, 7]
             }
            )
    input_df = input_df.set_index(["Column1", "Column2"])
    total = False
    subtotal = False
    actual = processing.apply_row_totals(input_df, total, subtotal)

    expected = pd.DataFrame(
        {"Column1": ["Org1", "Org1", "Org2", "Org2", "Org2"],
         "Column2": ["Org1", "Org2", "Org3", "Org4", "Org5"],
         "Underweight": [2, 3, 5, 6, 1],
         "Healthy weight": [1, 2, 4, 6, 7],
         "Overweight": [3, 4, 3, 4, 5],
         "Obese": [4, 5, 8, 6, 5],
         "Severely obese": [0, 1, 2, 5, 6],
         "Overweight and obese": [1, 3, 5, 6, 7]
         }
        )
    expected = expected.set_index(["Column1", "Column2"])
    pd.testing.assert_frame_equal(actual, expected)


def test_add_prevalence():
    """
    Test prevalence has been calculated correctly
    df[observedcol + "_prevalence"] = 100 * df[observedcol] / df[samplecol]

    return df
    """
    input_df = pd.DataFrame({"observedcol": [0, 10111, 2629, 8291],
                             "samplecol": [8500, 10111, 15150, 9678]})
    return_df = processing.add_prevalence(input_df, "observedcol", "samplecol")

    expected = [0.00000, 100.00000, 17.35314, 85.66853]
    actual = list(return_df["observedcol_prevalence"].astype(float).round(5))
    assert actual == expected, f"when checking prevelance, expected to find {expected} but found {actual}"


def test_count_bmi_categories():
    input_df = pd.DataFrame(
        {"Org": ["Org1", "Org2", "Org1", "Org1", "Org2", "Org1", "Org2",
         "Org2", "Org1", "Org2"],
         "BmiPScore": [0, 0.02, 0.021, 0.11, 0.85, 0.90, 0.95, 0.967, 0.996, 1.01],
         "Underweight": [1, 1, np.nan, np.nan, np.nan, np.nan,
                         np.nan, np.nan, np.nan, np.nan],
         "Healthy weight": [np.nan, np.nan, 1, 1, np.nan, np.nan, np.nan, np.nan,
                            np.nan, np.nan],
         "Overweight": [np.nan, np.nan, np.nan, np.nan, 1, 1, np.nan, np.nan,
                        np.nan, np.nan],
         "Obese": [np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, 1, 1, 1, 1],
         "Severely obese": [np.nan, np.nan, np.nan, np.nan, np.nan, np.nan,
                            np.nan, np.nan, 1, 1],
         "Overweight and obese": [np.nan, np.nan, np.nan, np.nan, 1, 1, 1, 1, 1, 1]
         }
    )
    groups = ["Org"]
    actual = processing.count_bmi_category(input_df, groups)

    expected = pd.DataFrame({
        "Org": ["Org1", "Org2"],
        "underweight": [1, 1],
        "healthy_weight": [2, 0],
        "overweight": [1, 1],
        "obese": [1, 3],
        "severely_obese": [1, 1],
        "overweight_obese": [2, 4],
        "total": [5, 5]
        }
        )
    expected = expected.set_index(["Org"])
    expected = expected.astype('float64')
    expected = expected.astype({"total": "int64"})
    pd.testing.assert_frame_equal(actual, expected)


def test_subgroup_rows():
    input_df = pd.DataFrame({
        "Row_Def": ["1", "2", "3", "4", "5", "6", "7", "8"],
        "Total": [45, 70, 82, 15, 34, 64, 83, 103],
        }
        )
    groups = ["Row_Def"]
    col1_row_subgroups = {"Urban": ["1", "5"], "Town_Fringe": ["2", "6"],
                          "Village_Hamlet_Isolated": ["3", "7", "4", "8"]}

    actual = processing.add_subgroup_rows(input_df, groups, col1_row_subgroups)

    expected = pd.DataFrame(
        {"Row_Def": ["1", "2", "3", "4", "5", "6", "7", "8",
                     "Urban", "Town_Fringe", "Village_Hamlet_Isolated"],
         "Total": [45, 70, 82, 15, 34, 64, 83, 103, 79, 134, 283]
         }
        )

    actual = actual.reset_index(drop=True)
    pd.testing.assert_frame_equal(actual, expected)


def test_calc_conf_intervals():

    input_df = pd.DataFrame(
        {
            "count": [3733, 5724, 23],
            "total": [304820, 597812, 5460]
        }
    )

    actual = processing.calc_conf_intervals(input_df,
                                            "count",
                                            "total",
                                            "percentage")

    expected = pd.DataFrame(
        {
            "count": [3733, 5724, 23],
            "total": [304820, 597812, 5460],
            "count_ci_lower": [0.0118622293714433, 0.00933119236318462,
                               0.00280869323449586],
            "count_ci_upper": [0.0126432076805657, 0.00982494346432261,
                               0.00631336112896509]
        }
    )

    pd.testing.assert_frame_equal(actual, expected)


def test_map_la_code_to_name():
    input_df = pd.DataFrame({
        "OrgCode_ONS": ["E00000001", "E00000002", "E00000003"],
        "PupilTier1LocalAuthority": ["E00000004", "E00000005", "E00000006"],
        "PupilTier2LocalAuthority": ["E00000007", "E00000008", "E00000009"],
        "SchoolTier1LocalAuthority": ["E00000010", "E00000011", "E00000012"],
        "SchoolTier2LocalAuthority": ["E00000013", "E00000014", "E00000015"],
        }
        )

    input_df_reference = pd.DataFrame({
        "OrgCode": ["E00000001", "E00000002", "E00000003",
                    "E00000004", "E00000005", "E00000006",
                    "E00000007", "E00000008", "E00000009",
                    "E00000010", "E00000011", "E00000012",
                    "E00000013", "E00000014", "E00000015"],
        "OrgName": ["London", "Coventry", "York",
                    "Oxford", "Bath", "Bristol",
                    "Warrington", "Luton", "Blackpool",
                    "Derby", "Manchester", "Darlington",
                    "Cambridge", "Hull", "Middlesbrough"],
        "Other": [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15],
        }
        )

    col_ref = ['OrgCode_ONS',
               'PupilTier1LocalAuthority',
               'PupilTier2LocalAuthority',
               'SchoolTier1LocalAuthority',
               'SchoolTier2LocalAuthority']

    actual = processing.map_la_code_to_name(input_df, input_df_reference, col_ref)

    expected = pd.DataFrame({
        "OrgCode_ONS": ["E00000001", "E00000002", "E00000003"],
        "PupilTier1LocalAuthority": ["E00000004", "E00000005", "E00000006"],
        "PupilTier2LocalAuthority": ["E00000007", "E00000008", "E00000009"],
        "SchoolTier1LocalAuthority": ["E00000010", "E00000011", "E00000012"],
        "SchoolTier2LocalAuthority": ["E00000013", "E00000014", "E00000015"],
        "OrgCode_ONS_Name": ["London", "Coventry", "York"],
        "PupilTier1LocalAuthority_Name": ["Oxford", "Bath", "Bristol"],
        "PupilTier2LocalAuthority_Name": ["Warrington", "Luton", "Blackpool"],
        "SchoolTier1LocalAuthority_Name": ["Derby", "Manchester", "Darlington"],
        "SchoolTier2LocalAuthority_Name": ["Cambridge", "Hull", "Middlesbrough"],
        }
        )

    pd.testing.assert_frame_equal(actual, expected)


def test_add_participation():

    participation = {'org_code_col': 'RegionCode',
                     'participation_col': 'derived_participation_R'}

    df = pd.DataFrame({
        "RegionCode": ["E00000001", "E00000002", "E00000003"],
        "values": [1, 2, 3],
        }
        )

    df_part_ref = pd.DataFrame({
        "derived_code": ["E00000001", "E00000002", "E00000003"],
        "derived_participation_R": [90.5, 91.2, 97],
        }
        )

    actual = processing.add_participation(df, df_part_ref, participation)
    actual = actual.drop(columns=['index'])

    expected = pd.DataFrame({
        "RegionCode": ["E00000001", "E00000002", "E00000003"],
        "values": [1, 2, 3],
        "derived_participation_R": [90.5, 91.2, 97],
        }
        )

    pd.testing.assert_frame_equal(actual, expected)
