import pandas as pd
import numpy as np
import sidetable as stb
import scipy.stats
import logging

import ncmp_code.utilities.write_excel as write
import ncmp_code.utilities.helpers as helpers


logger = logging.getLogger(__name__)


def calc_conf_intervals(df, observedcol, samplecol, measure=None):
    """ Calculates lower and upper confidence intervals based on methodology
    used in NCMP annual report
    https://digital.nhs.uk/data-and-information/publications/statistical/national-child-measurement-programme/2020-21-school-year/appendices#appendix-d-confidence-intervals)

    Parameters
    ----------
    df: pandas.DataFrame
    observedcol: str
        columns for observed number of feature of interest e.g. numerator
    samplecol: str
        column for sample size e.g. denominator
    measure: str
        will output CIs as percentages if set to 'percent'

    Returns
    -------
    df : pandas.DataFrame
    """

    df["r"] = df[observedcol]
    df["n"] = df[samplecol]

    df["p"] = df["r"]/df["n"]  # proportion with feature of interest
    df["q"] = 1 - df["p"]  # proportion without feature of interest
    df["z"] = scipy.stats.norm.ppf(0.975)  # 𝑧(1−∝/2) from the standard Normal distribution

    df["A"] = (2*df["r"]) + (df["z"]**2)
    df["B"] = df["z"] * np.sqrt((df["z"]**2) + (4*df["r"]*df["q"]))
    df["C"] = 2*(df["n"] + df["z"]**2)

    if measure == "percent":
        df[observedcol + "_ci_lower"] = (df["A"]-df["B"])/df["C"]*100
        df[observedcol + "_ci_upper"] = (df["A"]+df["B"])/df["C"]*100

    else:
        df[observedcol + "_ci_lower"] = (df["A"]-df["B"])/df["C"]
        df[observedcol + "_ci_upper"] = (df["A"]+df["B"])/df["C"]

    df.drop(columns=["r", "n", "p", "q", "z", "A", "B", "C"], inplace=True)

    return df


def update_org_columns(df, small_la_combine,
                       region_code_cols, region_code_name_map):
    """ Makes any required updates to org columns in the dataset

    Parameters
    ----------
    df: pandas.DataFrame
        dataframe containing NCMP data for this year
    small_la_combine: dict
        dictionary showing small LAs (key of dictionary)
        to be combined with larger LAs (value of dictionary)
    region_code_cols: list
        list of regional columns we need to map to names
    region_code_name_map: dict
        dictionary showing region codes (key of dictionary)
        to region names (value of dictionary)

    Returns
    -------
    df : pandas.DataFrame
        dataframe containing NCMP data for this year with updated org columns
    """
    logging.info("Updating org columns")

    # combine small LAs if needed
    df = df.replace(small_la_combine)

    # Update incorrect codes for submitting LAs from FACT
    # TEMP until DME process can be updated to accommodate org codes length >3

    # North Northamptonshire
    df.loc[df["OrgCode"] == "Z9D", "OrgCode"] = "Z9D4Z"
    df.loc[df["OrgCode"] == "Z9D4Z", "OrgCode_ONS"] = "E06000061"

    # West Northamptonshire
    df.loc[df["OrgCode"] == "U6Q", "OrgCode"] = "U6Q5Z"
    df.loc[df["OrgCode"] == "U6Q5Z", "OrgCode_ONS"] = "E06000062"

    # create region names column based on region code
    # for each region column map the code to name
    for region_code_col in region_code_cols:
        for region_code, region_name in region_code_name_map.items():
            df.loc[df[region_code_col] == region_code,
                   region_code_col + "_Name"] = region_name

    return df


def map_la_code_to_name(df, df_ref, col_ref):
    """ Function to map user specified LA code to its name
    in the referance dataframe and append to the raw data

    Parameters
    ----------
    df: pandas.DataFrame
    df_ref: pandas.DataFrame
        reference dataframe with the la names
    col_ref: list
        list of org columns we want to find the la name for
        specified in params

    Returns
    -------
    df : pandas.DataFrame
    """
    logging.info("Mapping LA code to name")

    # filtering ref data for code and name cols
    df_ref = df_ref[["OrgCode", "OrgName"]]

    # for each column in the list specified in param
    for i in col_ref:
        # rename reference columns
        df_i = df_ref.rename(columns={"OrgCode": i, "OrgName": i+"_Name"})
        # merge dataframe with ref data on the ref column name
        df = pd.merge(df, df_i, how="left", on=i)

    return df


def suppress_count_column(df, col_to_suppress,
                          national_level="grand_total",
                          lower=1, upper=7, base=5):
    """Follows HES disclosure control guidance.
    https://digital.nhs.uk/data-and-information/data-tools-and-services/data-services/hospital-episode-statistics/change-to-disclosure-control-methodology-for-hes-and-ecds-from-september-2018
    For sub-national counts, suppress the values of a count column
    based on upper and lower bounds, and round the values above the upper bound
    to the nearest base.

    If not national level, then apply suppression and rounding as per below logic.
    If more than or equal to lower bound and less than or equal to upper bound,
    then replace the values with "*".
    If more than upper value, then round to the nearest 5.

    Parameters
    ----------
    col_to_suppress: pd.Series
        A numeric column that should be suppressed
    national_level: str
        the name used for national level data - default is grand_total
    lower: int
        Lower bound - default is 1
        Used to filter for values more than or equal to 1 (>=1).
    upper: int
        Upper bound - default is 7
        Used to filter for values less than or equal to 7 (<=7).
    base: int - default is 5
        Round to the nearest base.
        E.g. a value of 21 or 22 would round to 20,
        while value of 23 or 24 would round to 25.

    Returns
    -------
    pd.Series

    """

    # copy of the column to suppress
    suppression = col_to_suppress.copy(deep=True)

    # define filters
    # filter national level data by searching for grand_total anywhere in the row
    national_level = (df.isin([national_level]).any(axis=1))
    # filter data between lower and upper bound that should be suppressed
    # for sub-national level
    should_suppress = ((~national_level) &
                       (col_to_suppress.between(lower, upper, inclusive="both")))
    # filter data above upper limit to be rounded for sub-national level
    should_round = ((~national_level) & (col_to_suppress > upper))

    # suppression and rounding logic for relevant data defined by above filters
    # if data should be suppressed, replace with *
    suppression.loc[should_suppress] = "*"

    # if data should be rounded, round to the nearest base
    suppression.loc[should_round] = (
        suppression[should_round]
        .apply(
            lambda p: base * round(p/base)
            )
        )

    return suppression


def suppress_percentage_column(df, col_to_suppress,
                               col_numerator, col_denominator,
                               national_level="grand_total",
                               round_to_dp=1):
    """Follows HES disclosure control guidance.
    https://digital.nhs.uk/data-and-information/data-tools-and-services/data-services/hospital-episode-statistics/change-to-disclosure-control-methodology-for-hes-and-ecds-from-september-2018
    For sub-national level, suppress the values of a percentage column
    based on the numerator and denominator columns suppression and rounding.

    If not national level, then apply suppression and rounding as per below
    suppress_count_logic to the numerator and denominator.
    If numerator and denominator are suppressed, then suppress the percentage.
    Calculate percentage based on rounded numerator and denominator otherwise.
    Also round the percentage to specified decimal places in that case.

    Parameters
    ----------
    col_to_suppress: pd.Series
        A numeric column with percentage values that should be suppressed
    col_numerator: pd.Series
        A numeric column with the numerator for the percentage
    col_denominator: pd.Series - default is df["total"]
        A numeric column with the denominator for the percentage
    national_level: str
        the name used for national level data - default is grand_total
    round_to_dp: int - default is None
        Round to the percentages to a number of decimal places.
        None means no rounding is applied to the value.

    Returns
    -------
    pd.Series

    """

    # copy of the column to suppress
    suppression = col_to_suppress.copy(deep=True)

    # suppressed the numerator and denominator
    numerator_suppress = suppress_count_column(df, col_numerator)
    denominator_suppress = suppress_count_column(df, col_denominator)

    # define filters
    # filter national level data by searching for grand_total anywhere in the row
    national_level = (df.isin([national_level]).any(axis=1))
    # filter percentages that need to be suppressed by ffiltering for
    # numerator or denominator being suppressed
    should_suppress = (numerator_suppress == "*") | (denominator_suppress == "*")
    # filter percentages that need to be calculated on rounded numbers
    # based on if numerator and denominator are not suppressed
    should_round = (~national_level) & (~should_suppress)

    # supress percentages that need to be suppressed
    suppression.loc[should_suppress] = "*"

    # calculate percentages that need to be based on rounded numerator and denominatoe
    suppression.loc[should_round] = (100 * numerator_suppress.loc[should_round]
                                     / denominator_suppress.loc[should_round])

    # round the final percentage if round_to_dp not None
    if round_to_dp is not None:
        suppression.loc[should_round] = (
            suppression[should_round]
            .apply(
                lambda p: round(p, round_to_dp)
            )
        )

    return suppression


def create_participation_reference(df, df_ref):
    """ Function to extract the participation data from the LADQ extract, map
    to LA ONS codes using the raw data mapping between OrgCode and OrgCode_ONS,
    add subtotoals for each region and grand_total for England level,and
    calculate the participation rates for LAs, regional subtotals and England
    grand_total. Remap all org codes to a common derived code column that can
    be used for reference for all levels.

    Parameters
    ----------
    df: pandas.DataFrame
        raw data
    df_ref: pandas.DataFrame
        reference dataframe with participation data for each LA

    Returns
    -------
    df_ref : pandas.DataFrame
    """
    logging.info("Adding participation rates data")

    # extract org code (short code), ONS org code and region code from raw data
    df = df[["OrgCode", "OrgCode_ONS", "RegionCode"]]

    # keep unique combinations of the three columns above
    df = df.drop_duplicates()

    # keep only relevant columns from participation reference file
    df_ref = df_ref[["OrgCode",
                     "TotalEligibleYearR",
                     "TotalEligibleYear6",
                     "TotalEligibleMeasuredYearR",
                     "TotalEligibleMeasuredYear6"]]

    # pick the columns in reference file that contain the partcipation data
    # we need to calculate the rates
    participation_cols = ["TotalEligibleYearR",
                          "TotalEligibleYear6",
                          "TotalEligibleMeasuredYearR",
                          "TotalEligibleMeasuredYear6"]

    # convert above columns' values to integer so we can calculate the rates
    df_ref[participation_cols] = df_ref[participation_cols].astype(int)

    # merge the unique combinations of org mapping from raw data
    # to the mapping to participation columns in reference data
    # on short org code
    df_ref = pd.merge(df, df_ref, how="left", on="OrgCode")

    # set the index to region code and ONS LA code
    # so can create the total and subtotal on region level
    df_ref = df_ref.set_index(["RegionCode", "OrgCode_ONS"])

    # create total and subtotals
    df_ref = df_ref.stb.subtotal()

    # derive the participation rate based total measured and total eligible
    # for both reception and year 6
    df_ref["derived_participation_R"] = (100 *
                                         df_ref["TotalEligibleMeasuredYearR"] /
                                         df_ref["TotalEligibleYearR"])

    df_ref["derived_participation_6"] = (100 *
                                         df_ref["TotalEligibleMeasuredYear6"] /
                                         df_ref["TotalEligibleYear6"])

    df_ref["TotalEligibleMeasuredOverall"] = (df_ref["TotalEligibleMeasuredYearR"] +
                                              df_ref["TotalEligibleMeasuredYear6"])

    df_ref["TotalEligibleOverall"] = (df_ref["TotalEligibleYearR"] +
                                      df_ref["TotalEligibleYear6"])

    df_ref["derived_participation_overall"] = (100 *
                                               df_ref["TotalEligibleMeasuredOverall"] /
                                               df_ref["TotalEligibleOverall"])

    # apply suppression to participation rate (percentage) columns
    df_ref["derived_participation_R"] = suppress_percentage_column(df_ref,
                                                                   df_ref["derived_participation_R"],
                                                                   df_ref["TotalEligibleMeasuredYearR"],
                                                                   df_ref["TotalEligibleYearR"])

    df_ref["derived_participation_6"] = suppress_percentage_column(df_ref,
                                                                   df_ref["derived_participation_6"],
                                                                   df_ref["TotalEligibleMeasuredYear6"],
                                                                   df_ref["TotalEligibleYear6"])

    df_ref["derived_participation_overall"] = suppress_percentage_column(df_ref,
                                                                         df_ref["derived_participation_overall"],
                                                                         df_ref["TotalEligibleMeasuredOverall"],
                                                                         df_ref["TotalEligibleOverall"])
    # leave only particpitation rates in the reference table
    df_ref = df_ref[["derived_participation_R", "derived_participation_6",
                     "derived_participation_overall"]]

    # reset index so region code and ONS LA code become columns instead of index
    df_ref = df_ref.reset_index()

    # make a new code column that will contain ONS LA codes, region codes
    # and england totals
    df_ref["derived_code"] = df_ref["OrgCode_ONS"]

    # rename the region subtotal rows to the region code
    df_ref["derived_code"][df_ref["derived_code"]
                           .str.endswith("subtotal")] = df_ref["RegionCode"]

    # rename the region total to grand_total specified in region code column
    df_ref["derived_code"][df_ref["derived_code"] == " "] = df_ref["RegionCode"]

    return df_ref


def add_bmi_categories(df, bmi_categories):
    """ Loop through the dictionary of categories defined in parameters
    that defines the filter based on BmiPScore for each category.
    Filter the raw data for the category, ceate a column with 1 assigned
    if the filter is True and the merge to the raw data.
    The final result is the raw data with additional column for each category
    showing 1 if the filter based on BmiPScore (from dict) is True.

    Parameters
    ----------
    df: pandas.DataFrame
    bmi_categories: dict
        categories and the filter based on BmiPScore that defines them

    Returns
    -------
    df_final : pandas.DataFrame
    """
    logging.info("Adding BMI category columns")

    # columns of the raw data that would be used for merging
    cols = df.columns.tolist()
    df_final = df.copy()

    # loop through the dictionary for categories and filters
    for category, filter in bmi_categories.items():
        # for each category query the raw data based on filter and assign to new df
        df_category = df.query(filter, engine="python")
        # assign 1 to all rows in new df to show the filter condition is True
        df_category = df_category.copy()
        df_category[category] = 1
        # drop duplicates
        df_category = df_category.drop_duplicates(cols)
        # merge the new category df to the raw data
        df_final = pd.merge(df_final, df_category, how="left", on=cols)

    return df_final


def count_bmi_category(df, groups):
    """ Group by the groups to create the required breakdown.
    Sum all 1s for each weight category column to find the total number of
    children in each category and counts the total number of children overall.

    Parameters
    ----------
    df: pandas.DataFrame
    groups: list
        list of columns to group by to create the breakdown
        e.g SchoolYear and GenderCode

    Returns
    -------
    df : pandas.DataFrame
    """
    logging.info("Adding counts of BMI categories")

    df = (df
          .groupby(groups, as_index=True)
          .agg(underweight=("Underweight", "sum"),
               healthy_weight=("Healthy weight", "sum"),
               overweight=("Overweight", "sum"),
               obese=("Obese", "sum"),
               severely_obese=("Severely obese", "sum"),
               overweight_obese=("Overweight and obese", "sum"),
               total=("BmiPScore", "count"))
          )

    return df


def replace_null_in_col(df, replace_null_logic=None):
    """ Replace NAs using the dictionary replace_null_logic
    showing the column in which we want to replace the NAs
    and the value the NAs will be replaced with

    Parameters
    ----------
    df: pandas.DataFrame
    replace_null_logic: dict, default = None
        dictionary showing the column in which we want to replace the NAs
        and the value the NAs will be replaced with

    Returns
    -------
    df : pandas.DataFrame
    """

    # if we want to replace NAs we set replace_null_logic to be a dict
    # showing the column in which we want to replace the NAs
    # and the value the NAs will be replaced with
    if replace_null_logic is not None:
        # going through the columns names and the values to replaces NAs with
        for col_name, replace_value in replace_null_logic.items():
            # and replacing the NAs in the column k with the value v
            df = df.copy()
            df[col_name] = df[col_name].fillna(replace_value)

    return df


def apply_row_totals(df, total=False, subtotal=False):
    """ Calculate the total and subtotals for all levels of the index
    which are the groups we break down by like SchoolYear and GenderCode.
    Total and subtotal are flags that if False will excude the totals and
    subtotals, if True, they will be included in the table.

    Parameters
    ----------
    df: pandas.DataFrame
    total: bool, default = False
        flag that will be set to False if we don't want row with total included
        defaults to False
    subtotal: bool, default = False
        flag that will be set to False if we don't want row with subtotals included

    Returns
    -------
    df : pandas.DataFrame
    """

    # setting type of the index values on each level to string
    # so that subtotal function works
    if df.index.nlevels == 1:
        df.index = df.index.get_level_values(0).astype(str)
    if df.index.nlevels == 2:
        df.index = [df.index.get_level_values(0).astype(str),
                    df.index.get_level_values(1).astype(str)]

    # using subtotal function
    df = df.stb.subtotal()

    # exclude grand_total row if total flag is False
    if total is False:
        df = df[df.index.get_level_values(0) != "grand_total"]

    # if we have more than one level of the index (if we group by more than one thing)
    # and if subtotal flag is False for each subindex level
    if (df.index.nlevels > 1) & (subtotal is False):
        for i in range(df.index.nlevels):
            df = df[~df.index.get_level_values(i-1).str.endswith("subtotal")]

    return df


def add_subgroup_rows(df, columns, subgroup):
    """
    Combines groups of values into a single grouped value. The subgrouping will
    be based on the first column in the list.
    This is added to the dataframe as additional rows.

    Parameters
    ----------
    df : pandas.DataFrame
        Data with a breakdown
    column: list[str]
        Column to apply the subgroups to.
    subgroup: dictionary
        Contains the variable content that will be assigned to the
        new grouping(s), and the value that will be assigned to the group.

    Returns
    -------
    pandas.DataFrame with breakdown subgroup

    """
    # Extracts the first column only to perform the subgrouping on.
    column = columns[0]

    # Loops through the dictionary and creates each of the subgroups, appending them back into the data
    for subgroup_code, subgroup_values in subgroup.items():
        subgroup = df[df[column].isin(subgroup_values)].copy(deep=True)
        subgroup[column] = subgroup_code
        subgroup = (
            subgroup.groupby(columns)
            .agg("sum")
            .reset_index()
            )
        df = df.append(subgroup)

    return df


def add_prevalence(df, observedcol, samplecol):
    """ Calculate prevalence by dividing the number of children
    in each weight category by the total number of children.

    Parameters
    ----------
    df: pandas.DataFrame
    observedcol: str
        columns for observed number of feature of interest e.g. numerator
    samplecol: str
        column for sample size e.g. denominator

    Returns
    -------
    df : pandas.DataFrame
    """

    # calculate prevalence and assign to a new column
    df[observedcol + "_prevalence"] = 100 * df[observedcol] / df[samplecol]

    return df


def add_participation(df, df_part_ref, participation=None):
    """ Get relevant participation rate and merge to final data based on
    specified org code column.

    Parameters
    ----------
    df: pandas.DataFrame
    df_part_ref: pandas.DataFrame
        reference dataframe with participation rate data for each org_code
        including LAs, regions, and England total
    participation: dict, default = None
        dictionary showing the name of the org code column to refer to
        and the name of the participation column to include from reference data

    Returns
    -------
    df : pandas.DataFrame
    """

    # assign the org code column defined in the participation dictionary
    # to use to map to the participation reference dataframe
    org_code_col = participation["org_code_col"]
    # assign the participation column defined in the participation
    # dictionary to use to add to the final output
    participation_col = participation["participation_col"]

    # assign derived_code column in reference dataframe the a new column
    # which name matches the org code column name in the final data
    df_part_ref[org_code_col] = df_part_ref["derived_code"]

    # keep only org column and relevant participation rate in reference
    df_part_ref = df_part_ref[[org_code_col, participation_col]]

    # reset index of final data
    df = df.reset_index()

    # merge final data with relevant participation rate in reference
    df = pd.merge(df, df_part_ref, how="left", on=org_code_col)

    return df


def format_table(df, groups, col_order, sort_logic=None,
                 col1_row_order=None, col2_row_order=None,
                 table_bmi_category=None, participation=None,
                 write_variable=False):
    """ Format table by sorting on the groups we breakdown by,
    then rename any of the groups if needed and set the order of the columns.

    Parameters
    ----------
    df: pandas.DataFrame
    groups: list
        list of columns to group by to create the breakdown
        e.g SchoolYear and GenderCode
    sort_logic: dict, default = None
        sorting logic for each column in groups used for breakdown
        defining if ascending or descenting
        True order it ascending
        False orders it descenting
    col1_row_order: list, , default = None
        list with the order of values in the first level of the index
    col2_row_order: list, , default = None
        list with the order of values in the second level of the index
    table_bmi_category: str, default = None
        parameter used to specify if the table should output only specific
        category; the type of table that shows the only one category have a
        different structure than the rest, where the break down by groups
        happens on columns rathen than rows
        e.g. 'obese' or 'severely_obese
    participation: dict, default = None
        dictionary showing the name of the org code column to refer to
        and the name of the participation column to include from reference data
    write_variable: bool, default = False
        flag to show if the table should use dynamic write to excel function.
        Set to True if the number of rows might change like lower and upper
        tier org tables.

    Returns
    -------
    df : pandas.DataFrame
    """

    # set index to groups
    df = df.set_index(groups)

    if table_bmi_category is None:

        # setting the order of the index defined as input to the function
        # row order for first level of the index
        if (df.index.nlevels == 1) & (col1_row_order is not None):
            df = df.reindex(col1_row_order)

        # row order for first and second level of the index if more than one level
        if (df.index.nlevels > 1) & (col1_row_order is not None) & (col2_row_order is not None):
            df = df.reindex(col1_row_order, level=0)
            df = df.reindex(col2_row_order, level=1)

        # reset index so groups can become columns to sort by
        df = df.reset_index()

        # renaming the grand_total to 0 in order to make sure it is sorted on top
        df = df.replace({"grand_total": "0"})

        # sorting logic for each group - ascending vs descenting
        # defined as input in the function
        if sort_logic is not None:
            df = df.sort_values(by=list(sort_logic.keys()),
                                ascending=list(sort_logic.values()))

        # if participation is not None and needs to be added to the table,
        # then we add the participation rate column name to the end of the
        # column order
        if participation is not None:
            participation_col = participation["participation_col"]
            col_order_participation = col_order.copy()
            col_order_participation.append(participation_col)
            col_order = col_order_participation.copy()

        # if need to dynamically write to excel include groups but exclude
        # region code columns
        if write_variable is True:
            col_order_la_tables = col_order.copy()
            col_order_la_tables = groups + col_order_la_tables
            region_codes = ["RegionCode", "SchoolRegionCode", "PupilRegionCode"]

            for code in region_codes:
                if code in col_order_la_tables:
                    col_order_la_tables.remove(code)

            col_order = col_order_la_tables.copy()

        # include col_order columns in final dataframe
        df = df[col_order]

    # if we only want to display specific category in the table
    else:

        # selecting the category columns we want in the final table
        col_order = [table_bmi_category,
                     table_bmi_category + "_prevalence",
                     table_bmi_category + "_ci_lower",
                     table_bmi_category + "_ci_upper"]

        df = df[col_order]

        # renaming columns so they can be sorted alphabetically later
        df = df.rename(columns={table_bmi_category + "_ci_lower":
                                table_bmi_category + "_prevalence_ci_lower",
                                table_bmi_category + "_ci_upper":
                                table_bmi_category + "_prevalence_ci_upper"})

        # renaming index for both genders together to 0 so it is sorted first
        df = df.rename(index={"1.0 - subtotal": "0",
                              "10.0 - subtotal": "0"})

        # create the final format of the table
        # set the selected groups as the columns in the table
        df = df.unstack().unstack().reset_index()
        df = df.set_index([groups[1], groups[0], "level_0"])
        df = df.T.sort_index(axis=1, ascending=True)

    return df


# create table broken down by groups
def create_output(df, groups, col_order,
                  sheetname, empty_cols, write_cell,
                  output_file, df_part_ref,
                  filter_condition=None, replace_null_logic=None,
                  total=False, subtotal=False, col1_row_subgroups=None,
                  sort_logic=None, col1_row_order=None,
                  col2_row_order=None, disclosure_control=True,
                  la_lower_tier=False,
                  table_bmi_category=None, participation=None,
                  write_variable=False):
    """
    Function to create a table output (for tables and charts)
    and write to the master excel sheet.
    Includes filter, counting of children in each category,
    including totals and subtotals, adding prevalence
    and confidence interval for the prevalence, and formatting.

    Parameters
    ----------
    df: pandas.DataFrame
    groups: list
        list of columns to group by to create the breakdown
        e.g SchoolYear and GenderCode
    col_order: list
        list of columns to include and their order
        defined in parameters under COL_ORDER
    df_part_ref: pandas.DataFrame
        reference dataframe with participation for each LA
    sheetname: str
        shows which sheet you write data to in the Excel master file
    empty_cols: list
        list of column that are left blank in the Excel master file
    write_cell: str
        cell from which to start the pasting the output in the Excel master file
    output_file: str
            Excel file to output data to
    filter_condition: str, default = None
        query logic to use to filter the data - e.g. (SchoolYear not in ("R"))
    replace_null_logic: dict, default = None
        dictionary showing the column in which we want to replace the NAs
        and the value the NAs will be replaced with
    total: bool, default = False
        flag that will be set to False if we don't want row with total included
        defaults to False
    subtotal: bool, default = False
        flag that will be set to False if we don't want row with subtotals included
    col1_row_subgroups: dict, default = None
        dictionary containing the subgroups and the values to populate thoes subgroups
    sort_logic: dict, default = None
        sorting logic for each column in groups used for breakdown
        defining if ascending or descenting
        True order it ascending
        False orders it descenting
    col1_row_order: list, , default = None
        list with the order of values in the first level of the index
    col2_row_order: list, , default = None
        list with the order of values in the second level of the index
    col_ref: str, default = None
        org column we want to find the la name for
    participation: dict, default = None
        dictionary showing the name of the org code column to refer to
        and the name of the participation column to include from reference data
    disclosure_control: bool, default = True
        flag set to True if HES disclosure control for suppressing and rounding
        should be applied
    la_lower_tier: bool, default = False
        flag to show if the table is includes lower tier LAs.
        Set to True if lower tier LAs need to be used for the grouping of data.
    table_bmi_category: str, default = None
        parameter used to specify if the table should output only specific
        category; the type of table that shows the only one category have a
        different structure than the rest, where the break down by groups
        happens on columns rathen than rows
        e.g. "obese" or "severely_obese
    write_variable: bool, default = False
        flag to show if the table should use dynamic write to excel function.
        Set to True if the number of rows might change like lower and upper
        tier org tables.

    Returns
    -------
    df : pandas.DataFrame
    """
    # Apply the optional table filter
    if filter_condition is not None:
        df = df.query(filter_condition, engine="python")

    # replace NAs with a value
    df = replace_null_in_col(df, replace_null_logic)

    # count the children in each weight category and break down by groups
    # e.g school year and gender
    df = count_bmi_category(df, groups)

    # weight category column
    cols = df.columns

    # add totals and subtotals
    df = apply_row_totals(df, total, subtotal)

    # reset index
    df = df.reset_index()

    # add subgroups to column 1 if required
    if col1_row_subgroups is not None:
        df = add_subgroup_rows(df, groups, col1_row_subgroups)

    # for each weight category column
    for col in cols:

        # calculate prevalence and assign to a new column
        add_prevalence(df, col, "total")
        # calculate lower and upper confidence interval for each prevalence
        # and assign to new columns
        calc_conf_intervals(df, col, "total", measure="percent")

        # if disclosure control is applied
        if disclosure_control is True:
            # apply suppression to prevalence (percentage) columns
            df[col + "_prevalence"] = suppress_percentage_column(df,
                                                                 df[col + "_prevalence"],
                                                                 df[col],
                                                                 df["total"])
            # suppress confidence intervals if the prevalence is suppressed
            df.loc[df[col + "_prevalence"] == "*", col + "_ci_lower"] = "*"
            df.loc[df[col + "_prevalence"] == "*", col + "_ci_upper"] = "*"
            # apply suppression to count columns
            df[col] = suppress_count_column(df, df[col])

    # add participation rates if needed
    if participation is not None:
        df = add_participation(df, df_part_ref, participation)

    # filter rows where lower tier is equivalent to upper tier LA
    if la_lower_tier is True:
        if "SchoolTier2LocalAuthority_Name" in groups:
            df = df[df["SchoolTier2LocalAuthority_Name"] !=
                    df["SchoolTier1LocalAuthority_Name"]]
        if "PupilTier2LocalAuthority_Name" in groups:
            df = df[df["PupilTier2LocalAuthority_Name"] !=
                    df["PupilTier1LocalAuthority_Name"]]

    # format final table
    df = format_table(df, groups, col_order, sort_logic,
                      col1_row_order, col2_row_order,
                      table_bmi_category, participation,
                      write_variable)

    # Write to the master table template
    if write_variable is True:
        write.write_to_excel_variable(df, sheetname, empty_cols, write_cell,
                                      output_file)
    else:
        write.write_to_excel_static(df, sheetname, empty_cols, write_cell,
                                    output_file)

    return df


def create_output_dq(df_la, df_la_dq, df_la_dq_sysbreachref,
                     la_dq_col_order, sheetname, empty_cols,
                     write_cell, output_file, write_type,
                     total_only=True, col_order=None):
    """
    Function to create the output for DQ tables and charts
    and write to the master excel sheet.
    It uses columns from the LA DQ file which is on LA level only to create
    the derive the England total figures and then formats it to fit in the
    masterfile templates.

    Parameters
    ----------
    df_la: pandas.DataFrame
        Derived mapping between OrgCode and LA ONS code and name
    df_la_dq: pandas.DataFrame
        Imported LA DQ file
    df_la_dq_sysbreachref: pandas.Dataframe
        Contains details of standard breach responses available in NCMP system
    la_dq_col_order: list
        List of default columns to use from the LA DQ file for all outputs
        defined in parameters under LA_DQ_COL_ORDER
        If col_order is set to None, these columns, in the defined order,
        will be outputted
    sheetname: str
        shows which sheet you write data to in the Excel master file
    empty_cols: list
        list of column that are left blank in the Excel master file
    write_cell: str
        cell from which to start the pasting the output in the Excel master file
    output_file: str
        Excel file to output DQ data to
    write_type: str
        Two options -
        'excel_variable' to use the Excel variable write function for dataframes
        that vary frequently e.g. org level outputs
        'excel_static' to use Excel static write function for dataframes that
        always have the same number of rows
    total_only: str, default = True
        flag that will be set to True if we want only England level figure;
        if set to False, it would include the LA level in addition to England level
    col_order: list, default = None
        list with the columns needed to create the final output;
        if set to None, includes all columns in la_dq_col_order

    Returns
    -------
    df_la_dq : pandas.DataFrame
    """

    # take a copy of df_la_dq and set the index to the org code column
    df = df_la_dq.copy()
    df = df.set_index(["OrgCode"])

    # replace nulls with 0 for Total columns
    df["TotalEligibleMeasuredYearR"].fillna(0, inplace=True)
    df["TotalEligibleMeasuredYear6"].fillna(0, inplace=True)
    df["TotalEligibleYearR"].fillna(0, inplace=True)
    df["TotalEligibleYear6"].fillna(0, inplace=True)

    # derive the total eligible and measured number of children
    df["TotalEligibleMeasured_derived"] = (df["TotalEligibleMeasuredYearR"] +
                                           df["TotalEligibleMeasuredYear6"])

    df["TotalEligible_derived"] = (df["TotalEligibleYearR"] +
                                   df["TotalEligibleYear6"])

    # delete all LAs with 0 children measured
    df = df[df["TotalEligibleMeasured_derived"] != 0].copy()

    # get list of percentage columns (those beginning with Percentage)
    la_dq_perc_cols = [col for col in df if col.startswith("Percentage")]

    # calculate the counts for each percentage column
    for percentage_col in la_dq_perc_cols:
        col = percentage_col.replace("Percentage", "")
        df["Count" + col] = (df[percentage_col] *
                             df["TotalEligibleMeasured_derived"])

    # calculate the England total counts
    df = df.stb.subtotal()

    # calculate the percentage columns including the derived England total level
    for percentage_col in la_dq_perc_cols:
        col = percentage_col.replace("Percentage", "")
        df[percentage_col] = (df["Count" + col] /
                              df["TotalEligibleMeasured_derived"])

    # derive participation rates
    df["derived_participation_R"] = (100 *
                                     df["TotalEligibleMeasuredYearR"] /
                                     df["TotalEligibleYearR"])

    df["derived_participation_6"] = (100 *
                                     df["TotalEligibleMeasuredYear6"] /
                                     df["TotalEligibleYear6"])

    df["derived_participation_overall"] = (100 *
                                           df["TotalEligibleMeasured_derived"] /
                                           df["TotalEligible_derived"])

    # reset index
    df = df.reset_index()

    # filter for England level only if total_only flag is set to True
    if total_only is True:
        df = df[df["OrgCode"] == "grand_total"].copy()

    # charts outputs are presented in vertical format
    # we use melt to select the measures and present the values in a column
    # if col_order is not None, melt on the list of specified columns
    # if col_order is None, all columns will be included in the output
    # and additional manipulations are done to get the required table output
    if col_order is not None:
        df = pd.melt(df, id_vars=["OrgCode"],
                     value_vars=col_order,
                     var_name="Measure", value_name="Value")
        df = df[["Value"]].copy()
    else:
        # include the mapped org code to ONS org code and name
        df_la_orgs = df_la[["OrgCode", "OrgCode_ONS", "OrgCode_ONS_Name"]].copy()

        df_la_orgs = df_la_orgs.groupby(["OrgCode", "OrgCode_ONS",
                                         "OrgCode_ONS_Name"]).count().reset_index()

        df = pd.merge(df, df_la_orgs, on="OrgCode", how="left")

        # add ONS name and code for England
        df.loc[df["OrgCode"] == "grand_total",
               "OrgCode_ONS_Name"] = "A_England"

        df.loc[df["OrgCode"] == "grand_total",
               "OrgCode_ONS"] = "E92000001"

        # make the percentage of records sharing same ethnicity check
        df.loc[((df["PercentageEthnicGroupOther"] == 100) |
                (df["PercentageEthnicGroupAsian"] == 100) |
                (df["PercentageEthnicGroupBlack"] == 100) |
                (df["PercentageEthnicGroupChinese"] == 100) |
                (df["PercentageEthnicGroupMixed"] == 100) |
                (df["PercentageEthnicGroupWhite"] == 100)
                ), "Percentage of records sharing the same ethnicity"] = 1

        df["Percentage of records sharing the same ethnicity"].replace({np.nan: 0},
                                                                       inplace=True)

        # set the column order of the table
        df = df[la_dq_col_order].copy()

        # sort the table
        df = df.sort_values(by="OrgCode_ONS_Name")

        # replace values in the table
        df["OrgCode_ONS_Name"].replace({"A_England": "ENGLAND"}, inplace=True)

    # get list of breach reasons columns (beginning with DQ)
    la_dq_breach_cols = [col for col in df if col.startswith("DQ")]

    # replace any na values in breach columns
    df[la_dq_breach_cols] = df[la_dq_breach_cols].replace({"nan": ""})
    df[la_dq_breach_cols] = df[la_dq_breach_cols].fillna("")

    # check if supplied breach reasons match standard ones from system
    # add "_CHECK!" to end of breach reason if not
    df_compare = df_la_dq_sysbreachref
    check_type = False
    flag_value = "_CHECK!"

    for breach_col in la_dq_breach_cols:

        check_column = breach_col
        check_compare_column = breach_col + "_System"

        helpers.check_update_column(df, df_compare, check_column,
                                    check_compare_column, check_type, flag_value)

    # write to the master table template using specified write option
    if write_type == "excel_variable":
        write.write_to_excel_variable(df, sheetname, empty_cols, write_cell,
                                      output_file)

    elif write_type == "excel_static":
        write.write_to_excel_static(df, sheetname, empty_cols, write_cell,
                                    output_file)

    return df
