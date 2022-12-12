import pandas as pd
import numpy as np
from datetime import datetime
import logging
import ncmp_code.utilities.write_excel as write
import ncmp_code.utilities.processing_steps as processing
import ncmp_code.parameters as param

logger = logging.getLogger(__name__)


def get_validations_group1():
    """
    Creates a list of functions relating to the first group of validations
    that use the same parameters: df_compyear, df_weight, output_file
    """

    validation_list = [pd_validation_1_num_measured,
                       pd_validation_3_bmi_prevalence
                       ]

    return validation_list


def get_validations_group2():
    """
    Creates a list of functions relating to the second group of validations
    that use the same parameters: df_weight, output_file
    """

    validation_list = [pd_validation_5_extreme_measurements,
                       pd_validation_6_pupil_postcode,
                       pd_validation_9_num_measured_agerange
                       ]

    return validation_list


def get_validations_group3():
    """
    Creates a list of functions relating to the first group of validations
    that use the same parameters: df_school_level, df_la_dq, output_file
    """

    validation_list = [pd_validation_2_eligible_pupil,
                       pd_validation_7_schools_removed]

    return validation_list


def pd_validation_1_num_measured(df_compyear, df_weight, output_file):
    """
    Creates validation 1 of the post deadline validations.
    This checks that the number of children measured this year is not more
    than 10% different to the previous year.

    Results are outputted to an Excel file

    Parameters:
    df_compyear: pandas.Dataframe
        Contains measurement data for NCMP schools, for comparison year

    df_weight: pandas.Dataframe
        Contains measurement data for NCMP schools, for this year, with BMI
        categories assigned

    output_file: str
        Excel file to output final data to - based on which part of the
        process is set to run using parameters in create_publication.py

    Returns:
    None
    """
    logging.info("Running validation 1 (num_measured)")

    # Create the fact comparison df
    df_compyear = df_compyear.groupby(["OrgCode",
                                       "SchoolYear"]
                                      )["NcmpSystemId"].count().reset_index()

    df_compyear = df_compyear.rename(columns={"NcmpSystemId": "CompYearValue"})
    df_compyear = df_compyear.astype({"OrgCode": str, "SchoolYear": str})

    # Create this years df to check
    df_thisyear = df_weight.groupby(["OrgCode", "OrgCode_ONS",
                                     "OrgCode_ONS_Name",
                                     "SchoolYear"])["NcmpSystemId"].count().reset_index()

    df_thisyear = df_thisyear.rename(columns={"NcmpSystemId": "ThisYearValue"})

    # Merge data for both years
    df = pd.merge(df_thisyear, df_compyear,  how="left",
                  left_on=["OrgCode", "SchoolYear"],
                  right_on=["OrgCode", "SchoolYear"])

    # Create the % change column
    df["PercentChange"] = ((df["ThisYearValue"] - df["CompYearValue"]) /
                           df["CompYearValue"] * 100)

    # Create absolute % change column for sort (with negative values converted)
    df["AbsPercentChange"] = df["PercentChange"]
    df.loc[df["PercentChange"] < 0,
           "AbsPercentChange"] = df["PercentChange"] * -1

    # Create the breach flag
    df["Breach"] = 0
    df.loc[(df["PercentChange"] > 10) | (df["PercentChange"] < -10),
           "Breach"] = 1

    # Add threshold and validation description
    df["BreachThreshold"] = "+/- 10%"
    df["Validation"] = "Changes in number of children measured"

    # Add run date/time and year values to file before export
    df["CompYear"] = param.COMP_YEAR
    df["ThisYear"] = param.YEAR
    df["RunDate"] = datetime.now().strftime("%Y-%m-%d, %H:%M:%S")

    # Sort outputs by absolute % change column descending
    df.sort_values(by=["AbsPercentChange"], ascending=False, inplace=True)

    # Define parameter values for write to Excel
    sheetname = "Val1_NumMeas"
    write_cell = "A2"
    empty_cols = None

    # Write the dataframe to Excel
    write.write_to_excel_variable(df, sheetname, empty_cols, write_cell,
                                  output_file)


def pd_validation_3_bmi_prevalence(df_compyear, df_weight, output_file):
    """
    This function will create validation 3 of the post deadline validations.
    This checks the prevalence rate for all BMI categories does not change by
    more than 5 percentage points from the previous collection year.

    Results are outputted to an Excel file

    Parameters:
    df_compyear: pandas.Dataframe
        Contains measurement data for NCMP schools, for comparison year

    df_weight: pandas.Dataframe
        Contains measurement data for NCMP schools, for this year, with BMI
        categories assigned

    output_file: str
        Excel file to output final data to - based on which part of the
        process is set to run using parameters in create_publication.py

    Returns:
    None
    """
    logging.info("Running validation 3 (bmi_prevalence)")

    # define variables used in function
    bmi_categories = param.BMI_CATEGORIES
    CompYearGroups = ["OrgCode", "SchoolYear"]
    ThisYearGroups = ["OrgCode", "OrgCode_ONS", "OrgCode_ONS_Name", "SchoolYear"]

    # Adds the BMI categories to the comparison year data
    df_compyear = processing.add_bmi_categories(df_compyear, bmi_categories)

    # Sum up the totals for each of the BMI categories for comparison year
    df_compyear = processing.count_bmi_category(df_compyear, CompYearGroups)

    # Using reset index to solve the nested columns
    df_compyear = df_compyear.reset_index()

    # Unpivot comparison year columns to match SAS outputs
    df_compyear = pd.melt(df_compyear, id_vars=["OrgCode", "SchoolYear", "total"],
                          value_vars=["underweight",
                                      "healthy_weight",
                                      "overweight",
                                      "obese",
                                      "severely_obese"])

    # Calculate the prevelance for the comparison year
    df_compyear = processing.add_prevalence(df_compyear, "value", "total")

    # Rename the columns using a dictionary
    df_compyear = df_compyear.rename(columns={"value": "CompYearValue",
                                              "total": "CompYearTotal",
                                              "variable": "BMICategory",
                                              "value_prevalence": "CompYearPrevalence"})

    # Sum up the totals for each of the BMI categories for the current year
    df_thisyear = processing.count_bmi_category(df_weight, ThisYearGroups)

    # Using reset index to solve the nested columns
    df_thisyear = df_thisyear.reset_index()

    # Unpivot current year's columns  to macth SAS outputs
    df_thisyear = pd.melt(df_thisyear,
                          id_vars=["OrgCode", "OrgCode_ONS", "OrgCode_ONS_Name",
                                   "SchoolYear", "total"],
                          value_vars=["underweight",
                                      "healthy_weight",
                                      "overweight",
                                      "obese",
                                      "severely_obese"])

    # Caculate the prevelance for the current year
    df_thisyear = processing.add_prevalence(df_thisyear, "value", "total")

    # Rename the current years columns and include year
    df_thisyear = df_thisyear.rename(columns={"value": "ThisYearValue",
                                              "total": "ThisYearTotal",
                                              "variable": "BMICategory",
                                              "value_prevalence": "ThisYearPrevalence"})

    # Merge the current and comparison year tables
    df = pd.merge(df_thisyear, df_compyear,  how="left",
                  left_on=["OrgCode", "SchoolYear", "BMICategory"],
                  right_on=["OrgCode", "SchoolYear", "BMICategory"])

    # Create the % change column
    df["PercentChange"] = (df["ThisYearPrevalence"] - df["CompYearPrevalence"])

    # Create absolute % change column for sort (with negative values converted)
    df["AbsPercentChange"] = df["PercentChange"]
    df.loc[df["PercentChange"] < 0,
           "AbsPercentChange"] = df["PercentChange"] * -1

    # Create the breach flag
    df["Breach"] = 0
    df.loc[(df["PercentChange"] > 5) | (df["PercentChange"] < -5),
           "Breach"] = 1

    # Add threshold and validation description
    df["BreachThreshold"] = "+/- 5%"
    df["Validation"] = "Changes in BMI prevalence"

    # Add run date/time and year values to file before export
    df["CompYear"] = param.COMP_YEAR
    df["ThisYear"] = param.YEAR
    df["RunDate"] = datetime.now().strftime("%Y-%m-%d, %H:%M:%S")

    df_final = df[["OrgCode", "OrgCode_ONS", "OrgCode_ONS_Name", "SchoolYear",
                   "BMICategory", "CompYearPrevalence",
                   "ThisYearPrevalence", "AbsPercentChange",
                   "Breach", "PercentChange", "BreachThreshold", "Validation",
                   "CompYear", "ThisYear", "RunDate"]].copy()

    # Sort the final dataframe by the absolute % change
    df_final.sort_values(by=["AbsPercentChange"], ascending=False, inplace=True)

    # Define parameter values for write to Excel
    sheetname = "Val3_BMIPrev"
    write_cell = "A2"
    empty_cols = None

    # Write the dataframe to Excel
    write.write_to_excel_variable(df_final, sheetname, empty_cols, write_cell,
                                  output_file)


def pd_validation_8_num_measured_ncmpschoolstatus(df_la, output_file):
    """
    Creates an output containing the total number and percentage of measurements
    submitted this year, by NCMP school status (NCMP or non-NCMP) and LA

    Outputted to Excel file

    Parameters
    ----------
    df_la: pandas.Dataframe
        All submitted data for this year, with LA details

    output_file: str
        Filepath of Excel file to output final data to - based on which part
        of the process is set to run using parameters in create_publication.py

    Returns
    -------
    None.

    """
    logging.info("Running validation 8 (num_measured_ncmpschoolstatus")

    # select rows with measurements
    df = df_la[df_la["Bmi"].notnull()].copy()

    # add school status flags
    df.loc[df["NcmpSchoolStatus"] == "NCMP", "NCMPMeasured"] = 1
    df.loc[df["NcmpSchoolStatus"] == "Non-NCMP", "nonNCMPMeasured"] = 1

    # sum count columns and groupby LA
    df = (df
          .groupby(["OrgCode", "OrgCode_ONS", "OrgCode_ONS_Name"])
          [["NCMPMeasured", "nonNCMPMeasured"]]
          .apply(sum).reset_index())

    # calculate England level
    df_eng = df.copy()
    df_eng["OrgCode"] = "ENG"
    df_eng["OrgCode_ONS"] = "E92000001"
    df_eng["OrgCode_ONS_Name"] = "ENGLAND"

    df_eng = (df_eng
              .groupby(["OrgCode", "OrgCode_ONS", "OrgCode_ONS_Name"])
              [["NCMPMeasured", "nonNCMPMeasured"]]
              .apply(sum).reset_index())

    # append England level to rest of data
    df = df.append(df_eng)

    # calculate total and percentages
    df["TotalMeasured"] = df["NCMPMeasured"] + df["nonNCMPMeasured"]
    df["NCMPMeasuredPerc"] = df["NCMPMeasured"] / df["TotalMeasured"] * 100
    df["nonNCMPMeasuredPerc"] = df["nonNCMPMeasured"] / df["TotalMeasured"] * 100

    # add breach threshold and validation descriptions
    df["BreachThreshold"] = "No threshold currently set"
    df["Validation"] = "Percentage of measurements from non-NCMP schools"

    # add reporting year and run date/time
    df["ReportingYear"] = param.YEAR
    df["RunDate"] = datetime.now().strftime("%Y-%m-%d, %H:%M:%S")

    # assign sort order so England values outputted at top
    df["OrgOrder"] = 0
    df.loc[df["OrgCode_ONS_Name"] == "ENGLAND", "OrgOrder"] = 1

    # re-order columns and sort
    df = df[["ReportingYear", "OrgCode", "OrgCode_ONS",
             "OrgCode_ONS_Name", "NCMPMeasured", "nonNCMPMeasured",
             "TotalMeasured", "NCMPMeasuredPerc", "nonNCMPMeasuredPerc",
             "BreachThreshold", "Validation", "RunDate",
             "OrgOrder"]].sort_values(by=["OrgOrder", "nonNCMPMeasuredPerc"],
                                      ascending=False)

    # drop the order column
    df.drop(["OrgOrder"], axis=1, inplace=True)

    # define parameter values for write to Excel
    sheetname = "Val8_NumMeas_SchoolStatus"
    empty_cols = None
    write_cell = "A2"

    # write output to validations Excel file
    write.write_to_excel_variable(df, sheetname, empty_cols, write_cell,
                                  output_file)


def pd_validation_9_num_measured_agerange(df_weight, output_file):
    """
    Creates an output containing the total number and percentage of measurements
    submitted this year by NCMP schools, by school year, LA and age range category:
    'InAgeRange', 'OutsideAgeRange'

    Outputted to Excel file

    Parameters
    ----------
    df_weight: pandas.Dataframe
        Contains measurement data for NCMP schools, for this year, with BMI
        categories assigned

    output_file: str
        Filepath of Excel file to output final data to - based on which part
        of the process is set to run using parameters in create_publication.py

    Returns
    -------
    None.

    """
    logging.info("Running validation 9 (num_measured_agerange")

    # add age in years column for grouping
    df_weight["AgeInYears"] = df_weight["AgeInMonths"]/12

    # define lower and upper limits for flag, by school year
    lowerage = [4.5, 10.5]
    upperage = [5.5, 11.5]
    schoolyear = ["R", "6"]

    # create empty dataframe for append
    total_dfs = []

    # add flags for in or outside one year age range, by school year
    for lower, upper, schyear in zip(lowerage, upperage, schoolyear):

        # filter for selected school year's data
        df_schyear = df_weight[df_weight["SchoolYear"] == schyear].copy()

        # check for ages outside one year age range
        df_schyear.loc[(df_schyear["AgeInYears"] < lower) |
                       (df_schyear["AgeInYears"] > upper),
                       "OutsideAgeRange"] = 1

        # check for ages inside one year age range
        df_schyear.loc[(df_schyear["AgeInYears"] >= lower) &
                       (df_schyear["AgeInYears"] <= upper),
                       "InAgeRange"] = 1

        # add age range category
        df_schyear["AgeRange"] = str(lower) + " to " + str(upper)

        # sum count columns and groupby LA
        df_schyear = (df_schyear
                      .groupby(["OrgCode", "OrgCode_ONS", "OrgCode_ONS_Name",
                                "AgeRange", "SchoolYear"])
                      [["InAgeRange", "OutsideAgeRange"]]
                      .apply(sum).reset_index())

        # append to total_dfs
        total_dfs.append(df_schyear)

    # combine school data
    df_comb = pd.concat(total_dfs)

    # calculate England level
    df_eng = df_comb.copy()
    df_eng["OrgCode"] = "ENG"
    df_eng["OrgCode_ONS"] = "E92000001"
    df_eng["OrgCode_ONS_Name"] = "ENGLAND"

    df_eng = (df_eng
              .groupby(["OrgCode", "OrgCode_ONS", "OrgCode_ONS_Name",
                        "AgeRange", "SchoolYear"])
              [["InAgeRange", "OutsideAgeRange"]]
              .apply(sum).reset_index())

    # append England level to rest of data
    df = df_comb.append(df_eng)

    # calculate school year totals
    df["Total"] = df["InAgeRange"] + df["OutsideAgeRange"]

    # calculate percentages
    df["InAgeRange_Perc"] = df["InAgeRange"] / df["Total"] * 100
    df["OutsideAgeRange_Perc"] = df["OutsideAgeRange"] / df["Total"] * 100

    # add breach threshold and validation descriptions
    df["BreachThreshold"] = "No threshold currently set"
    df["Validation"] = "Percentage of records with age outside one year range"

    # add reporting year and run date/time
    df["ReportingYear"] = param.YEAR
    df["RunDate"] = datetime.now().strftime("%Y-%m-%d, %H:%M:%S")

    # assign sort order so England values outputted at top
    df["OrgOrder"] = 0
    df.loc[df["OrgCode_ONS_Name"] == "ENGLAND", "OrgOrder"] = 1

    # sort output by OrgOrder and % outside age range descending
    df = df[["ReportingYear", "OrgCode", "OrgCode_ONS", "OrgCode_ONS_Name",
             "AgeRange", "SchoolYear", "InAgeRange", "OutsideAgeRange",
             "Total", "InAgeRange_Perc", "OutsideAgeRange_Perc", "BreachThreshold",
             "Validation", "RunDate",
             "OrgOrder"]].sort_values(by=["OrgOrder", "OutsideAgeRange_Perc"],
                                      ascending=False)

    # drop the order column
    df.drop(["OrgOrder"], axis=1, inplace=True)

    # write output to validations Excel file
    sheetname = "Val9_NumMeas_AgeRange"
    empty_cols = None
    write_cell = "A2"

    write.write_to_excel_variable(df, sheetname, empty_cols, write_cell,
                                  output_file)


def pd_validation_4_change_ethnicity(df_compyear_raw, df_la, output_file):
    """
    This function will create validation 4 of the post deadline validations.
    The proportion of children in each ethnic category is more than 20%
    and the ethnic group proportion was in the top five.
    LAs will not be queried if the ethnic categories of “not stated” or “unknown”
    have decreased by more than 20%.

    Results are outputted to an Excel file

    Parameters:
    df_compyear_raw: pandas.Dataframe
        All submitted data for comparison year

    df_la: pandas.Dataframe
        All submitted data for this year, with LA details

    output_file: str
        Excel file to output final data to - based on which part of the
        process is set to run using parameters in create_publication.py

    Returns:
    None
    """
    logging.info("Running validation 4 (ethnicity_change)")

    df_compyear = df_compyear_raw[(df_compyear_raw["NcmpSchoolStatus"] == "NCMP")].copy()

    df_compyear["NcmpEthnicityCode"].replace(to_replace={None: "Unknown",
                                                         "Not stated": "Unknown"},
                                             inplace=True)

    df_compyear = df_compyear[(df_compyear["NcmpSchoolStatus"] == "NCMP")]

    df_compyear1 = df_compyear.groupby(["OrgCode",
                                        "NcmpEthnicityCode"]
                                       )["NcmpSystemId"].count().reset_index()

    df_compyear2 = df_compyear.groupby(["OrgCode"])["NcmpSystemId"].count().reset_index()

    df_compyear = pd.merge(df_compyear1, df_compyear2,  how="left",
                           left_on=["OrgCode"],
                           right_on=["OrgCode"])

    df_compyear = df_compyear.rename(columns={"NcmpSystemId_y": "ComparisonYear_Total",
                                              "NcmpSystemId_x": "ComparisonYear"})
    df_compyear = df_compyear.astype({"OrgCode": str})
    df_compyear["ComparisonYear_Prev"] = ((df_compyear["ComparisonYear"] /
                                           df_compyear["ComparisonYear_Total"])*100)

    # Create this years df to check
    df_la = df_la[(df_la["NcmpSchoolStatus"] == "NCMP")]
    df_la["NcmpEthnicityCode"].replace(to_replace={None: "Unknown",
                                       "Not stated": "Unknown"}, inplace=True)

    df_thisyear1 = df_la.groupby(["OrgCode",
                                  "OrgCode_ONS_Name",
                                  "NcmpEthnicityCode"]
                                 )["NcmpSystemId"].count().reset_index()

    df_thisyear2 = df_la.groupby(["OrgCode",
                                  "OrgCode_ONS_Name"]
                                 )["NcmpSystemId"].count().reset_index()

    df_thisyear = pd.merge(df_thisyear1, df_thisyear2, how="left",
                           left_on=["OrgCode", "OrgCode_ONS_Name"],
                           right_on=["OrgCode", "OrgCode_ONS_Name"])

    df_thisyear = df_thisyear.rename(columns={"NcmpSystemId_y": "ThisYear_Total",
                                              "NcmpSystemId_x": "ThisYear"})

    df_thisyear["ThisYear_Prev"] = ((df_thisyear["ThisYear"] /
                                     df_thisyear["ThisYear_Total"])*100)

    df_thisyear["Rank"] = df_thisyear.groupby(["NcmpEthnicityCode"]
                                              )["ThisYear_Prev"].rank(ascending=False)

    # Merge data for both years
    df = pd.merge(df_thisyear, df_compyear[["ComparisonYear_Prev",
                                            "OrgCode", "NcmpEthnicityCode"]],
                  how="left",
                  left_on=["OrgCode", "NcmpEthnicityCode"],
                  right_on=["OrgCode", "NcmpEthnicityCode"])

    # Create the % change column
    df["PercentChange"] = (df["ThisYear_Prev"] - df["ComparisonYear_Prev"])

    # Create absolute % change column for sort (with negative values converted)
    df["AbsPercentChange"] = df["PercentChange"]
    df.loc[df["PercentChange"] < 0,
           "AbsPercentChange"] = df["PercentChange"] * -1

    # Create the unknown breach flag
    dfUkn = df.copy()
    dfUkn.loc[(dfUkn["AbsPercentChange"] > 20) &
              (dfUkn["NcmpEthnicityCode"] == "Unknown"), "UnknownBreach"] = 1

    dfUkn = dfUkn[dfUkn.UnknownBreach == 1]

    df = pd.merge(df, dfUkn[["OrgCode", "UnknownBreach", "NcmpEthnicityCode"]],
                  how="left",
                  on="OrgCode").drop(columns=["NcmpEthnicityCode_y"])

    # Create the breach flag
    df["Breach"] = 0
    df.loc[(df["AbsPercentChange"] > 20) & (df["Rank"] <= 5), "Breach"] = 1

    # Add threshold and validation description
    df["BreachThreshold"] = "+/- 20%"
    df["Validation"] = "Changes in ethnicity groups"
    df["RunDate"] = datetime.now().strftime("%Y-%m-%d, %H:%M:%S")
    df.sort_values(by=["Breach", "AbsPercentChange"], ascending=[False, False],
                   inplace=True)

    # Define parameter values for write to Excel
    sheetname = "Val4_NumEthn"
    write_cell = "A2"
    empty_cols = None

    # Write the dataframe to Excel
    write.write_to_excel_variable(df, sheetname, empty_cols, write_cell,
                                  output_file)


def pd_validation_5_extreme_measurements(df_weight, output_file):
    """
    This function will create validation 5 of the post deadline validations.
    This checks schools with a high number of extreme pupil measurements for
    Bmi, Weight and height.

    Results are outputted to an Excel file

    Parameters:

    df_weight: pandas.Dataframe
        Contains measurement data for NCMP schools, for this year, with BMI
        categories assigned

    output_file: str
        Excel file to output final data to - based on which part of the
        process is set to run using parameters in create_publication.py

    Returns:
    None
    """
    logging.info("Running validation 5 (extreme_measurements)")

    # columns to use in this validation synthax
    columns_to_use = ["OrgCode", "OrgCode_ONS_Name",
                      "SchoolUrn", "SchoolName", "SchoolPostcode", "SchoolYear",
                      "NcmpSystemId", "WeightZScore", "HeightZScore", "BmiZScore"]

    # reduce dataframe to only columns of interest
    df = df_weight[columns_to_use].copy()

    # add extreme flags
    # Extreme flag for Weight
    df["Weight_Extreme"] = 0

    df.loc[(df["WeightZScore"] < -3) | (df["WeightZScore"] > 4),
           "Weight_Extreme"] = 1

    # Extreme flag for Height
    df["Height_Extreme"] = 0

    df.loc[(df["HeightZScore"] < -3) | (df["HeightZScore"] > 4),
           "Height_Extreme"] = 1

    # Extreme flag for Bmi
    df["Bmi_Extreme"] = 0

    df.loc[(df["BmiZScore"] < -3) | (df["BmiZScore"] > 4),
           "Bmi_Extreme"] = 1

    # Create the final dataframe by school, school year and by count of pupils
    df = (df.groupby(columns_to_use[:6])
          .agg({"NcmpSystemId": "count", "Weight_Extreme": "sum",
                "Height_Extreme": "sum", "Bmi_Extreme": "sum"})
          .reset_index()
          .rename(columns={"NcmpSystemId": "Measured_Pupils"}))

    # Calculate extreme percentages for each measurement
    df["Weight_Extreme_Percent"] = (df["Weight_Extreme"]/df["Measured_Pupils"])*100
    df["Height_Extreme_Percent"] = (df["Height_Extreme"]/df["Measured_Pupils"])*100
    df["Bmi_Extreme_Percent"] = (df["Bmi_Extreme"]/df["Measured_Pupils"])*100

    # remove total measured pupil less than 20
    df = df.loc[df["Measured_Pupils"] > 20]

    # add a breach flag where any of the extreme percentages is greater than 10.
    df["Breach"] = 0
    df.loc[(df["Weight_Extreme_Percent"] > 10) |
           (df["Height_Extreme_Percent"] > 10) |
           (df["Bmi_Extreme_Percent"] > 10),
           "Breach"] = 1

    # remove non-breaching values
    df = df.loc[df["Breach"] == 1]

    # sort final dataframe
    df.sort_values(by=["Measured_Pupils"], ascending=False, inplace=True)

    # add breach threshold and validation descriptions
    df["BreachThreshold"] = "Any Extreme_Percent value +10% (schools with +20 pupils in school year)"
    df["Validation"] = "Schools with a high proportion of extreme measurements"

    # Add run date/time to file before export
    df["RunDate"] = datetime.now().strftime("%Y-%m-%d, %H:%M:%S")

    # Define parameter values for write to Excel
    sheetname = "Val5_Extreme_Measurements"
    write_cell = "A2"
    empty_cols = None

    # Write the dataframe to Excel
    write.write_to_excel_variable(df, sheetname, empty_cols, write_cell,
                                  output_file)


def pd_validation_6_pupil_postcode(df_weight, output_file):
    """
    This function will create validation 6 of the post deadline validations.
    This checks schools with a high number of extreme pupil postcode to school
    postcode distance – the number of pupils in a school where the distance
    from their home postcode to school is greater than 60km should not be 3
    or more.

    Results are outputted to an Excel file

    Parameters:

    df_weight: pandas.Dataframe
        Contains measurement data for NCMP schools, for this year, with BMI
        categories assigned

    output_file: str
        Excel file to output final data to - based on which part of the
        process is set to run using parameters in create_publication.py

    Returns:
    None
    """
    logging.info("Running validation 6 (pupil_postcode)")

    # columns to use in this validation synthax
    columns_to_use = ["OrgCode", "OrgCode_ONS_Name", "SchoolUrn", "SchoolName",
                      "SchoolPostcode", "PupilSchoolPostcodeDistance",
                      "NcmpSystemId"]

    # reduce dataframe to only columns of interest
    df = df_weight[columns_to_use].copy()

    # add extreme flag for distance > 60
    df["Extreme_PupilPostcode"] = "False"

    df.loc[(df["PupilSchoolPostcodeDistance"] > 60),
           "Extreme_PupilPostcode"] = "True"

    # exclude pupils with no extreme distance value
    df = df.loc[df["Extreme_PupilPostcode"] == "True"]

    # Create the final dataframe by school details and by count of pupils
    df = df.groupby(columns_to_use[:5]
                    )["NcmpSystemId"].count().reset_index().rename(
                        columns={"NcmpSystemId": "ExtremeCount"})

    # add a breach flag where pupil count is 3 or more.
    df["Breach_PupilPostcode"] = 0

    df.loc[(df["ExtremeCount"] >= 3),
           "Breach_PupilPostcode"] = 1

    # sort final dataframe
    df.sort_values(by=["ExtremeCount"], ascending=False, inplace=True)

    # add breach threshold and validation descriptions
    df["BreachThreshold"] = "3 or more pupils with distance from home postcode to school >60km"
    df["Validation"] = "Schools with a high number of extreme pupil postcode to school postcode distance"

    # Add run date/time to file before export
    df["RunDate"] = datetime.now().strftime("%Y-%m-%d, %H:%M:%S")

    # Define parameter values for write to Excel
    sheetname = "Val6_Pupil_Postcode"
    write_cell = "A2"
    empty_cols = None

    # Write the dataframe to Excel
    write.write_to_excel_variable(df, sheetname, empty_cols, write_cell,
                                  output_file)


def pd_validation_2_eligible_pupil(df_school_level, df_la_dq, output_file):
    """
    This function will create validation 2 of the post deadline validations.
    This checks the eligible pupil numbers for each LA.
    Flag any LA not providing their own headcounts for more than 90%
    of schools on their list and with a participation rate below 90%.
    This is carried out separately for children in reception and year 6.
    Results are outputted to an Excel file.

    Parameters:
    df_school_level: pandas.Dataframe
        Contains school level data for current year

    df_la_dq: pandas.Dataframe
        Contains DQ LA data for current year

    output_file: str
        Excel file to output final data to - based on which part of the
        process is set to run using parameters in create_publication.py

    Returns:
    None
    """

    logging.info("Running validation 2 (eligible_pupil")

    # select needed columns from the DQ LA file
    df_la_dq = df_la_dq[["OrgCode",
                         "LocalAuthorityName",
                         "PercentageParticipationYearR",
                         "PercentageParticipationYear6"]]

    # rename the LA code and name columns
    df_la_dq = df_la_dq.rename(columns={"OrgCode":
                                        "LA_submitted_code",
                                        "LocalAuthorityName":
                                        "LA_submitted_name"})

    # set the LA code and name as index
    df_la_dq = df_la_dq.set_index(["LA_submitted_code", "LA_submitted_name"])

    # select needed columns from the school file
    df_school_level = df_school_level[["LA_submitted_code",
                                       "LA_submitted_name",
                                       "NCMP_category",
                                       "Measured_Count_R",
                                       "Measured_Count_6",
                                       "LA_Pupil_Count_R",
                                       "LA_Pupil_Count_6"]]

    # calculate number of meaured schools for reception per LA
    df_school_measured_R = (df_school_level
                            .query("NCMP_category == 'NCMP' and Measured_Count_R > 0")
                            .groupby(["LA_submitted_code",
                                      "LA_submitted_name"], as_index=True)
                            .agg(measured_count_R=("Measured_Count_R",
                                                   "count"))
                            )

    # calculate number of meaured schools for year 6 per LA
    df_school_measured_6 = (df_school_level
                            .query("NCMP_category == 'NCMP' and Measured_Count_6 > 0")
                            .groupby(["LA_submitted_code",
                                      "LA_submitted_name"], as_index=True)
                            .agg(measured_count_6=("Measured_Count_6",
                                                   "count"))
                            )

    # calculate number of meaured schools with pupulated LA headcounts
    # for reception per LA
    df_school_pupil_R = (df_school_level
                         .query("NCMP_category == 'NCMP' and Measured_Count_R > 0 and LA_Pupil_Count_R > 0")
                         .groupby(["LA_submitted_code",
                                   "LA_submitted_name"], as_index=True)
                         .agg(la_pupil_count_R=("LA_Pupil_Count_R",
                                                "count"))
                         )

    # calculate number of meaured schools with pupulated LA headcounts
    # for year 6 per LA
    df_school_pupil_6 = (df_school_level
                         .query("NCMP_category == 'NCMP' and Measured_Count_6 > 0 and LA_Pupil_Count_6 > 0")
                         .groupby(["LA_submitted_code",
                                   "LA_submitted_name"], as_index=True)
                         .agg(la_pupil_count_6=("LA_Pupil_Count_6",
                                                "count"))
                         )

    # concatinate all dataframes with the new measures and participation rates
    result = pd.concat([df_school_measured_R,
                        df_school_measured_6,
                        df_school_pupil_R,
                        df_school_pupil_6,
                        df_la_dq], axis=1)

    # calculate percentage measured schools with populated LA headcounts
    result["la_pupil_count_R_perc"] = 100 * (result["la_pupil_count_R"] /
                                             result["measured_count_R"])
    result["la_pupil_count_6_perc"] = 100 * (result["la_pupil_count_6"] /
                                             result["measured_count_6"])

    # replave empty cells with 0
    result = result.fillna(0)

    # create flag showing if percentage measured schools with populated
    # LA headcounts or participation rate lower than 90
    result.loc[((result["PercentageParticipationYearR"] < 90) &
                (result["la_pupil_count_R_perc"] < 90)), "validation_R"] = 1
    result.loc[((result["PercentageParticipationYear6"] < 90) &
                (result["la_pupil_count_6_perc"] < 90)), "validation_6"] = 1

    # if denominaor is 0, make the percentage display "No meas"
    result.loc[result["measured_count_R"] == 0,
               "la_pupil_count_R_perc"] = "No meas"
    result.loc[result["measured_count_6"] == 0,
               "la_pupil_count_6_perc"] = "No meas"

    # if percentage is not valid measure, set the flag to 0
    result.loc[result["la_pupil_count_R_perc"] == "No meas",
               "validation_R"] = 0
    result.loc[result["la_pupil_count_6_perc"] == "No meas",
               "validation_6"] = 0

    # order the columns in the final table
    result = result.reset_index()

    col_order = ["LA_submitted_name", "LA_submitted_code",
                 "measured_count_R", "measured_count_6",
                 "la_pupil_count_R", "la_pupil_count_6",
                 "la_pupil_count_R_perc", "la_pupil_count_6_perc",
                 "PercentageParticipationYearR", "PercentageParticipationYear6",
                 "validation_R", "validation_6"]

    df_result = result[col_order].copy()

    # sort by LA name
    df_result.sort_values(by=["LA_submitted_name"], inplace=True)

    # add breach descriptions and runtime
    df_result["BreachThreshold"] = ">90% of LA's schools headcounts not provided by LA and participation rate <90%"
    df_result["Validation"] = "Eligible pupil numbers"
    df_result["RunDate"] = datetime.now().strftime("%Y-%m-%d, %H:%M:%S")

    # write to master file
    sheetname = "Val2_EligiblePupil"
    write_cell = "A2"
    empty_cols = None
    write.write_to_excel_variable(df_result, sheetname, empty_cols,
                                  write_cell, output_file)


def pd_validation_7_schools_removed(df_school_level, df_la_dq, output_file):
    """
    This function will create validation 7 of the post deadline validations.
    This checks the schools removed and added by the local authority –
    the number of schools removed from their school list by a local authority
    should not exceed 3 or more than the number added.
    Results are outputted to an Excel file.

    Parameters:
    df_school_level: pandas.Dataframe
        Contains school level data for current year

    df_la_dq: pandas.Dataframe
        Contains DQ LA data for current year

    output_file: str
        Excel file to output final data to - based on which part of the
        process is set to run using parameters in create_publication.py

    Returns:
    None
    """
    logging.info("Running validation 7 (schools removed)")

    # select needed columns from the DQ LA file
    df_la_dq = df_la_dq[["OrgCode", "LocalAuthorityName"]]

    # rename LA code and name to be consistent with the other dataframes
    df_la_dq = df_la_dq.rename(columns={"OrgCode": "code",
                                        "LocalAuthorityName": "name"})

    # select needed columns from the school file
    df_school_level = df_school_level[["Original_LA_assigned_code",
                                       "Original_LA_assigned_name",
                                       "LA_assigned_code",
                                       "LA_assigned_name",
                                       "LA_added_code",
                                       "LA_added_name",
                                       "NCMP_category",
                                       "Removed flag",
                                       "Schl_name"]]

    # rename Removed flag column as I can"t query for it if it has a blank space
    df_school_level = df_school_level.rename(columns={"Removed flag":
                                                      "removed_flag"})

    # calculate number of schools removed by LA
    df_school_removed = (df_school_level
                         .query("NCMP_category == 'NCMP' and removed_flag == 1")
                         .groupby(["Original_LA_assigned_code",
                                   "Original_LA_assigned_name"], as_index=True)
                         .agg(removed=("Schl_name", "count"))
                         .reset_index()
                         )

    # rename LA code and name to be consistent with the other dataframes
    df_school_removed = df_school_removed.rename(columns={"Original_LA_assigned_code":
                                                          "code",
                                                          "Original_LA_assigned_name":
                                                          "name"})

    # calculate number of schools assigned by LA
    df_school_assigned = (df_school_level
                          .query("NCMP_category == 'NCMP'")
                          .groupby(["LA_assigned_code",
                                    "LA_assigned_name"], as_index=True)
                          .agg(assigned=("Schl_name", "count"))
                          .reset_index()
                          )

    # rename LA code and name to be consistent with the other dataframes
    df_school_assigned = df_school_assigned.rename(columns={"LA_assigned_code":
                                                            "code",
                                                            "LA_assigned_name":
                                                            "name"})

    # calculate number of schools added by LA
    df_school_added = (df_school_level
                       .query("NCMP_category == 'NCMP'")
                       .groupby(["LA_added_code",
                                 "LA_added_name"], as_index=True)
                       .agg(added=("Schl_name", "count"))
                       .reset_index()
                       )

    # rename LA code and name to be consistent with the other dataframes
    df_school_added = df_school_added.rename(columns={"LA_added_code": "code",
                                                      "LA_added_name": "name"})

    # combine all dataframes with the new measures
    df_school_combined = (df_la_dq
                          .merge(df_school_removed, how="left", on=["code", "name"])
                          .merge(df_school_assigned, how="left", on=["code", "name"])
                          .merge(df_school_added, how="left", on=["code", "name"]))

    # replace all blank values with 0
    df_school_combined.replace(to_replace={np.nan: 0}, inplace=True)

    # calculate number of schools assigned or added
    df_school_combined["added_assigned"] = (df_school_combined["assigned"] +
                                            df_school_combined["added"])

    # calculate number of schools removed - assigned or added
    df_school_combined["removed_ex_added_assigned"] = (df_school_combined["removed"] -
                                                       df_school_combined["added_assigned"])

    # flag if number of schools removed - assigned or added is above 3
    df_school_combined.loc[df_school_combined["removed_ex_added_assigned"] >= 3,
                           "breach_flag"] = 1

    # replace blank values in breach flag with 0
    df_school_combined['breach_flag'].replace(to_replace={np.nan: 0}, inplace=True)

    # set the column order
    col_order = ["name", "code", "removed", "assigned", "added",
                 "added_assigned", "removed_ex_added_assigned", "breach_flag"]

    df_school_combined = df_school_combined[col_order]

    # sort values
    df_school_combined.sort_values(by=["breach_flag", "name"],
                                   ascending=[False, True], inplace=True)

    # add breach descriptions and runtime
    df_school_combined["BreachThreshold"] = "Number of schools removed from school list 3+ than number added"
    df_school_combined["Validation"] = "Schools removed and added by LA"
    df_school_combined["RunDate"] = datetime.now().strftime("%Y-%m-%d, %H:%M:%S")

    # write to master file
    sheetname = "Val7_School_Remove"
    write_cell = "A2"
    empty_cols = None
    write.write_to_excel_variable(df_school_combined, sheetname, empty_cols,
                                  write_cell, output_file)
