from ncmp_code.utilities.processing_steps import create_output
from ncmp_code.utilities.processing_steps import create_output_dq


def get_charts():

    chart_list = [chart_R,
                  chart_6,
                  chart_sex_R,
                  chart_sex_6,
                  chart_obese_R,
                  chart_obese_6,
                  chart_obese_pupil_imd_R,
                  chart_severely_obese_pupil_imd_R,
                  chart_obese_pupil_imd_6,
                  chart_severely_obese_pupil_imd_6,
                  chart_obese_school_imd_R,
                  chart_obese_school_imd_gender_R,
                  chart_obese_school_imd_6,
                  chart_obese_school_imd_gender_6,
                  chart_severely_obese_school_imd_R,
                  chart_severely_obese_school_imd_gender_R,
                  chart_severely_obese_school_imd_6,
                  chart_severely_obese_school_imd_gender_6,
                  chart_region_R_multi,
                  chart_region_6_multi,
                  chart_rurality_R_multi,
                  chart_rurality_6_multi,
                  chart_ons_area_R_multi,
                  chart_ons_area_6_multi,
                  chart_ethnicity_R_multi,
                  chart_ethnicity_6_multi,
                  chart_obese_pupil_imd_R_ts,
                  chart_obese_pupil_imd_gender_R_ts,
                  chart_obese_pupil_imd_6_ts,
                  chart_obese_pupil_imd_gender_6_ts,
                  chart_severely_obese_pupil_imd_R_ts,
                  chart_severely_obese_pupil_imd_gender_R_ts,
                  chart_severely_obese_pupil_imd_6_ts,
                  chart_severely_obese_pupil_imd_gender_6_ts]

    return chart_list


def get_charts_dq():

    chart_list = [chart_dq_participation,
                  chart_dq_HW_whole_num,
                  chart_dq_missing]

    return chart_list


def chart_R(df, output_file, df_part_ref):
    groups = ["AcademicYear"]
    filter_condition = 'SchoolYear == "R"'
    replace_null_logic = None
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = None
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    write_variable = False
    sheetname = "BMICat_R"
    empty_cols = []
    write_cell = "B3"
    col_order = ["underweight_prevalence",
                 "healthy_weight_prevalence",
                 "overweight_prevalence",
                 "obese_prevalence",
                 "severely_obese_prevalence"]

    return create_output(df, groups, col_order,
                         sheetname, empty_cols, write_cell,
                         output_file, df_part_ref,
                         filter_condition, replace_null_logic,
                         total, subtotal,
                         col1_row_subgroups,
                         sort_logic,
                         col1_row_order, col2_row_order,
                         disclosure_control,
                         la_lower_tier,
                         table_bmi_category,
                         participation,
                         write_variable)


def chart_6(df, output_file, df_part_ref):
    groups = ["AcademicYear"]
    filter_condition = 'SchoolYear == "6"'
    replace_null_logic = None
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = None
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    write_variable = False
    sheetname = "BMICat_6"
    empty_cols = []
    write_cell = "B3"
    col_order = ["underweight_prevalence",
                 "healthy_weight_prevalence",
                 "overweight_prevalence",
                 "obese_prevalence",
                 "severely_obese_prevalence"]

    return create_output(df, groups, col_order,
                         sheetname, empty_cols, write_cell,
                         output_file, df_part_ref,
                         filter_condition, replace_null_logic,
                         total, subtotal,
                         col1_row_subgroups,
                         sort_logic,
                         col1_row_order, col2_row_order,
                         disclosure_control,
                         la_lower_tier,
                         table_bmi_category,
                         participation,
                         write_variable)


def chart_sex_R(df, output_file, df_part_ref):
    groups = ["GenderCode"]
    filter_condition = 'SchoolYear == "R"'
    replace_null_logic = None
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = None
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    write_variable = False
    sheetname = "BMICat_Sex_R"
    empty_cols = []
    write_cell = "C3"
    col_order = ["underweight_prevalence",
                 "underweight_ci_lower",
                 "underweight_ci_upper",
                 "healthy_weight_prevalence",
                 "healthy_weight_ci_lower",
                 "healthy_weight_ci_upper",
                 "overweight_prevalence",
                 "overweight_ci_lower",
                 "overweight_ci_upper",
                 "obese_prevalence",
                 "obese_ci_lower",
                 "obese_ci_upper",
                 "severely_obese_prevalence",
                 "severely_obese_ci_lower",
                 "severely_obese_ci_upper"]

    return create_output(df, groups, col_order,
                         sheetname, empty_cols, write_cell,
                         output_file, df_part_ref,
                         filter_condition, replace_null_logic,
                         total, subtotal,
                         col1_row_subgroups,
                         sort_logic,
                         col1_row_order, col2_row_order,
                         disclosure_control,
                         la_lower_tier,
                         table_bmi_category,
                         participation,
                         write_variable)


def chart_sex_6(df, output_file, df_part_ref):
    groups = ["GenderCode"]
    filter_condition = 'SchoolYear == "6"'
    replace_null_logic = None
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = None
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    write_variable = False
    sheetname = "BMICat_Sex_6"
    empty_cols = []
    write_cell = "C3"
    col_order = ["underweight_prevalence",
                 "underweight_ci_lower",
                 "underweight_ci_upper",
                 "healthy_weight_prevalence",
                 "healthy_weight_ci_lower",
                 "healthy_weight_ci_upper",
                 "overweight_prevalence",
                 "overweight_ci_lower",
                 "overweight_ci_upper",
                 "obese_prevalence",
                 "obese_ci_lower",
                 "obese_ci_upper",
                 "severely_obese_prevalence",
                 "severely_obese_ci_lower",
                 "severely_obese_ci_upper"]

    return create_output(df, groups, col_order,
                         sheetname, empty_cols, write_cell,
                         output_file, df_part_ref,
                         filter_condition, replace_null_logic,
                         total, subtotal,
                         col1_row_subgroups,
                         sort_logic,
                         col1_row_order, col2_row_order,
                         disclosure_control,
                         la_lower_tier,
                         table_bmi_category,
                         participation,
                         write_variable)


def chart_obese_R(df, output_file, df_part_ref):
    groups = ["AcademicYear"]
    filter_condition = 'SchoolYear == "R"'
    replace_null_logic = None
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = None
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    write_variable = False
    sheetname = "BMICat_Obese_TSeries_R"
    empty_cols = []
    write_cell = "B3"
    col_order = ["obese_prevalence",
                 "severely_obese_prevalence",
                 "overweight_obese_prevalence"]

    return create_output(df, groups, col_order,
                         sheetname, empty_cols, write_cell,
                         output_file, df_part_ref,
                         filter_condition, replace_null_logic,
                         total, subtotal,
                         col1_row_subgroups,
                         sort_logic,
                         col1_row_order, col2_row_order,
                         disclosure_control,
                         la_lower_tier,
                         table_bmi_category,
                         participation,
                         write_variable)


def chart_obese_6(df, output_file, df_part_ref):
    groups = ["AcademicYear"]
    filter_condition = 'SchoolYear == "6"'
    replace_null_logic = None
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = None
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    write_variable = False
    sheetname = "BMICat_Obese_TSeries_6"
    empty_cols = []
    write_cell = "B3"
    col_order = ["obese_prevalence",
                 "severely_obese_prevalence",
                 "overweight_obese_prevalence"]

    return create_output(df, groups, col_order,
                         sheetname, empty_cols, write_cell,
                         output_file, df_part_ref,
                         filter_condition, replace_null_logic,
                         total, subtotal,
                         col1_row_subgroups,
                         sort_logic,
                         col1_row_order, col2_row_order,
                         disclosure_control,
                         la_lower_tier,
                         table_bmi_category,
                         participation,
                         write_variable)


def chart_obese_pupil_imd_R(df, output_file, df_part_ref):
    groups = ["PupilIndexOfMultipleDeprivationD"]
    filter_condition = 'SchoolYear == "R"'
    replace_null_logic = {"PupilIndexOfMultipleDeprivationD": "Unknown"}
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = ["1.0", "2.0", "3.0", "4.0", "5.0", "6.0", "7.0", "8.0",
                      "9.0", "10.0"]
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    write_variable = False
    sheetname = "Obese_PupIMD_R"
    empty_cols = []
    write_cell = "C3"
    col_order = ["obese_prevalence", "obese_ci_lower", "obese_ci_upper"]

    return create_output(df, groups, col_order,
                         sheetname, empty_cols, write_cell,
                         output_file, df_part_ref,
                         filter_condition, replace_null_logic,
                         total, subtotal,
                         col1_row_subgroups,
                         sort_logic,
                         col1_row_order, col2_row_order,
                         disclosure_control,
                         la_lower_tier,
                         table_bmi_category,
                         participation,
                         write_variable)


def chart_severely_obese_pupil_imd_R(df, output_file, df_part_ref):
    groups = ["PupilIndexOfMultipleDeprivationD"]
    filter_condition = 'SchoolYear == "R"'
    replace_null_logic = {"PupilIndexOfMultipleDeprivationD": "Unknown"}
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = ["1.0", "2.0", "3.0", "4.0", "5.0", "6.0", "7.0", "8.0",
                      "9.0", "10.0"]
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    write_variable = False
    sheetname = "SevObese_PupIMD_R"
    empty_cols = []
    write_cell = "C3"
    col_order = ["severely_obese_prevalence", "severely_obese_ci_lower",
                 "severely_obese_ci_upper"]

    return create_output(df, groups, col_order,
                         sheetname, empty_cols, write_cell,
                         output_file, df_part_ref,
                         filter_condition, replace_null_logic,
                         total, subtotal,
                         col1_row_subgroups,
                         sort_logic,
                         col1_row_order, col2_row_order,
                         disclosure_control,
                         la_lower_tier,
                         table_bmi_category,
                         participation,
                         write_variable)


def chart_obese_pupil_imd_6(df, output_file, df_part_ref):
    groups = ["PupilIndexOfMultipleDeprivationD"]
    filter_condition = 'SchoolYear == "6"'
    replace_null_logic = {"PupilIndexOfMultipleDeprivationD": "Unknown"}
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = ["1.0", "2.0", "3.0", "4.0", "5.0", "6.0", "7.0", "8.0",
                      "9.0", "10.0"]
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    write_variable = False
    sheetname = "Obese_PupIMD_6"
    empty_cols = []
    write_cell = "C3"
    col_order = ["obese_prevalence", "obese_ci_lower", "obese_ci_upper"]

    return create_output(df, groups, col_order,
                         sheetname, empty_cols, write_cell,
                         output_file, df_part_ref,
                         filter_condition, replace_null_logic,
                         total, subtotal,
                         col1_row_subgroups,
                         sort_logic,
                         col1_row_order, col2_row_order,
                         disclosure_control,
                         la_lower_tier,
                         table_bmi_category,
                         participation,
                         write_variable)


def chart_severely_obese_pupil_imd_6(df, output_file, df_part_ref):
    groups = ["PupilIndexOfMultipleDeprivationD"]
    filter_condition = 'SchoolYear == "6"'
    replace_null_logic = {"PupilIndexOfMultipleDeprivationD": "Unknown"}
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = ["1.0", "2.0", "3.0", "4.0", "5.0", "6.0", "7.0", "8.0",
                      "9.0", "10.0"]
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    write_variable = False
    sheetname = "SevObese_PupIMD_6"
    empty_cols = []
    write_cell = "C3"
    col_order = ["severely_obese_prevalence", "severely_obese_ci_lower",
                 "severely_obese_ci_upper"]

    return create_output(df, groups, col_order,
                         sheetname, empty_cols, write_cell,
                         output_file, df_part_ref,
                         filter_condition, replace_null_logic,
                         total, subtotal,
                         col1_row_subgroups,
                         sort_logic,
                         col1_row_order, col2_row_order,
                         disclosure_control,
                         la_lower_tier,
                         table_bmi_category,
                         participation,
                         write_variable)


def chart_obese_school_imd_R(df, output_file, df_part_ref):
    groups = ["SchoolIndexOfMultiDeprivationD"]
    filter_condition = 'SchoolYear == "R"'
    replace_null_logic = {"SchoolIndexOfMultiDeprivationD": "Unknown"}
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = ["1.0", "10.0"]
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    write_variable = False
    sheetname = "Obese_SchIMDGap_TSeries_R"
    empty_cols = []
    write_cell = "C3"
    col_order = ["obese_prevalence"]

    return create_output(df, groups, col_order,
                         sheetname, empty_cols, write_cell,
                         output_file, df_part_ref,
                         filter_condition, replace_null_logic,
                         total, subtotal,
                         col1_row_subgroups,
                         sort_logic,
                         col1_row_order, col2_row_order,
                         disclosure_control,
                         la_lower_tier,
                         table_bmi_category,
                         participation,
                         write_variable)


def chart_obese_school_imd_gender_R(df, output_file, df_part_ref):
    groups = ["GenderCode", "SchoolIndexOfMultiDeprivationD"]
    filter_condition = 'SchoolYear == "R"'
    replace_null_logic = {"SchoolIndexOfMultiDeprivationD": "Unknown"}
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = ["1", "2"]
    col2_row_order = ["1.0", "10.0"]
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    write_variable = False
    sheetname = "Obese_SchIMDGap_Sex_TSeries_R"
    empty_cols = []
    write_cell = "D3"
    col_order = ["obese_prevalence"]

    return create_output(df, groups, col_order,
                         sheetname, empty_cols, write_cell,
                         output_file, df_part_ref,
                         filter_condition, replace_null_logic,
                         total, subtotal,
                         col1_row_subgroups,
                         sort_logic,
                         col1_row_order, col2_row_order,
                         disclosure_control,
                         la_lower_tier,
                         table_bmi_category,
                         participation,
                         write_variable)


def chart_obese_school_imd_6(df, output_file, df_part_ref):
    groups = ["SchoolIndexOfMultiDeprivationD"]
    filter_condition = 'SchoolYear == "6"'
    replace_null_logic = {"SchoolIndexOfMultiDeprivationD": "Unknown"}
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = ["1.0", "10.0"]
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    write_variable = False
    sheetname = "Obese_SchIMDGap_TSeries_6"
    empty_cols = []
    write_cell = "C3"
    col_order = ["obese_prevalence"]

    return create_output(df, groups, col_order,
                         sheetname, empty_cols, write_cell,
                         output_file, df_part_ref,
                         filter_condition, replace_null_logic,
                         total, subtotal,
                         col1_row_subgroups,
                         sort_logic,
                         col1_row_order, col2_row_order,
                         disclosure_control,
                         la_lower_tier,
                         table_bmi_category,
                         participation,
                         write_variable)


def chart_obese_school_imd_gender_6(df, output_file, df_part_ref):
    groups = ["GenderCode", "SchoolIndexOfMultiDeprivationD"]
    filter_condition = 'SchoolYear == "6"'
    replace_null_logic = {"SchoolIndexOfMultiDeprivationD": "Unknown"}
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = ["1", "2"]
    col2_row_order = ["1.0", "10.0"]
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    write_variable = False
    sheetname = "Obese_SchIMDGap_Sex_TSeries_6"
    empty_cols = []
    write_cell = "D3"
    col_order = ["obese_prevalence"]

    return create_output(df, groups, col_order,
                         sheetname, empty_cols, write_cell,
                         output_file, df_part_ref,
                         filter_condition, replace_null_logic,
                         total, subtotal,
                         col1_row_subgroups,
                         sort_logic,
                         col1_row_order, col2_row_order,
                         disclosure_control,
                         la_lower_tier,
                         table_bmi_category,
                         participation,
                         write_variable)


def chart_severely_obese_school_imd_R(df, output_file, df_part_ref):
    groups = ["SchoolIndexOfMultiDeprivationD"]
    filter_condition = 'SchoolYear == "R"'
    replace_null_logic = {"SchoolIndexOfMultiDeprivationD": "Unknown"}
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = ["1.0", "10.0"]
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    write_variable = False
    sheetname = "SevOb_SchIMDGap_TSeries_R"
    empty_cols = []
    write_cell = "C3"
    col_order = ["severely_obese_prevalence"]

    return create_output(df, groups, col_order,
                         sheetname, empty_cols, write_cell,
                         output_file, df_part_ref,
                         filter_condition, replace_null_logic,
                         total, subtotal,
                         col1_row_subgroups,
                         sort_logic,
                         col1_row_order, col2_row_order,
                         disclosure_control,
                         la_lower_tier,
                         table_bmi_category,
                         participation,
                         write_variable)


def chart_severely_obese_school_imd_gender_R(df, output_file, df_part_ref):
    groups = ["GenderCode", "SchoolIndexOfMultiDeprivationD"]
    filter_condition = 'SchoolYear == "R"'
    replace_null_logic = {"SchoolIndexOfMultiDeprivationD": "Unknown"}
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = ["1", "2"]
    col2_row_order = ["1.0", "10.0"]
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    write_variable = False
    sheetname = "SevOb_SchIMDGap_Sex_TSeries_R"
    empty_cols = []
    write_cell = "D3"
    col_order = ["severely_obese_prevalence"]

    return create_output(df, groups, col_order,
                         sheetname, empty_cols, write_cell,
                         output_file, df_part_ref,
                         filter_condition, replace_null_logic,
                         total, subtotal,
                         col1_row_subgroups,
                         sort_logic,
                         col1_row_order, col2_row_order,
                         disclosure_control,
                         la_lower_tier,
                         table_bmi_category,
                         participation,
                         write_variable)


def chart_severely_obese_school_imd_6(df, output_file, df_part_ref):
    groups = ["SchoolIndexOfMultiDeprivationD"]
    filter_condition = 'SchoolYear == "6"'
    replace_null_logic = {"SchoolIndexOfMultiDeprivationD": "Unknown"}
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = ["1.0", "10.0"]
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    write_variable = False
    sheetname = "SevOb_SchIMDGap_TSeries_6"
    empty_cols = []
    write_cell = "C3"
    col_order = ["severely_obese_prevalence"]

    return create_output(df, groups, col_order,
                         sheetname, empty_cols, write_cell,
                         output_file, df_part_ref,
                         filter_condition, replace_null_logic,
                         total, subtotal,
                         col1_row_subgroups,
                         sort_logic,
                         col1_row_order, col2_row_order,
                         disclosure_control,
                         la_lower_tier,
                         table_bmi_category,
                         participation,
                         write_variable)


def chart_severely_obese_school_imd_gender_6(df, output_file, df_part_ref):
    groups = ["GenderCode", "SchoolIndexOfMultiDeprivationD"]
    filter_condition = 'SchoolYear == "6"'
    replace_null_logic = {"SchoolIndexOfMultiDeprivationD": "Unknown"}
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = ["1", "2"]
    col2_row_order = ["1.0", "10.0"]
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    write_variable = False
    sheetname = "SevOb_SchIMDGap_Sex_TSeries_6"
    empty_cols = []
    write_cell = "D3"
    col_order = ["severely_obese_prevalence"]

    return create_output(df, groups, col_order,
                         sheetname, empty_cols, write_cell,
                         output_file, df_part_ref,
                         filter_condition, replace_null_logic,
                         total, subtotal,
                         col1_row_subgroups,
                         sort_logic,
                         col1_row_order, col2_row_order,
                         disclosure_control,
                         la_lower_tier,
                         table_bmi_category,
                         participation,
                         write_variable)


def chart_region_R_multi(df, output_file, df_part_ref):
    groups = ['PupilRegionCode']
    filter_condition = 'SchoolYear == "R"'
    replace_null_logic = None
    total = True
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = ["grand_total", "E12000001", "E12000002", "E12000003",
                      "E12000004", "E12000005", "E12000006", "E12000007",
                      "E12000008", "E12000009"]
    col2_row_order = None
    sort_logic = None
    disclosure_control = True
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    write_variable = False
    empty_cols = []
    write_cell = "C3"

    chart_bmi_categories = ["underweight", "healthy_weight",
                            "overweight", "obese",
                            "severely_obese", "overweight_obese"]

    # outputs reception regional data for each bmi category onto  onto a separate tab
    for chart_bmi_category in chart_bmi_categories:
        sheetname = "Region_SchYr_" + chart_bmi_category
        col_order = [chart_bmi_category + "_prevalence",
                     chart_bmi_category + "_ci_lower",
                     chart_bmi_category + "_ci_upper"]

        create_output(df, groups, col_order,
                      sheetname, empty_cols, write_cell,
                      output_file, df_part_ref,
                      filter_condition, replace_null_logic,
                      total, subtotal,
                      col1_row_subgroups,
                      sort_logic,
                      col1_row_order, col2_row_order,
                      disclosure_control,
                      la_lower_tier,
                      table_bmi_category,
                      participation,
                      write_variable)


def chart_region_6_multi(df, output_file, df_part_ref):
    groups = ['PupilRegionCode']
    filter_condition = 'SchoolYear == "6"'
    replace_null_logic = None
    total = True
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = ["grand_total", "E12000001", "E12000002", "E12000003",
                      "E12000004", "E12000005", "E12000006", "E12000007",
                      "E12000008", "E12000009"]
    col2_row_order = None
    sort_logic = None
    disclosure_control = True
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    write_variable = False
    empty_cols = []
    write_cell = "C13"
    chart_bmi_categories = ["underweight", "healthy_weight",
                            "overweight", "obese",
                            "severely_obese", "overweight_obese"]

    # outputs year 6 regional data for each bmi category onto  onto a separate tab
    for chart_bmi_category in chart_bmi_categories:
        sheetname = "Region_SchYr_" + chart_bmi_category
        col_order = [chart_bmi_category + "_prevalence",
                     chart_bmi_category + "_ci_lower",
                     chart_bmi_category + "_ci_upper"]

        create_output(df, groups, col_order,
                      sheetname, empty_cols, write_cell,
                      output_file, df_part_ref,
                      filter_condition, replace_null_logic,
                      total, subtotal,
                      col1_row_subgroups,
                      sort_logic,
                      col1_row_order, col2_row_order,
                      disclosure_control,
                      la_lower_tier,
                      table_bmi_category,
                      participation,
                      write_variable)


def chart_rurality_R_multi(df, output_file, df_part_ref):
    groups = ["PupilUrbanRuralIndicator"]
    filter_condition = 'SchoolYear == "R"'
    replace_null_logic = {"PupilUrbanRuralIndicator": "Unknown"}
    total = True
    subtotal = False
    col1_row_subgroups = {"Urban": ["1", "5"], "Town_Fringe": ["2", "6"],
                          "Village_Hamlet_Isolated": ["3", "7", "4", "8"]}
    col1_row_order = ["grand_total", "Town_Fringe", "Urban", "Village_Hamlet_Isolated"]
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    write_variable = False
    empty_cols = []
    write_cell = "C3"
    chart_bmi_categories = ["underweight", "healthy_weight",
                            "overweight", "obese",
                            "severely_obese", "overweight_obese"]

    # outputs reception rurality data for each bmi category onto  onto a separate tab
    for chart_bmi_category in chart_bmi_categories:
        sheetname = "Rurality_SchYr_" + chart_bmi_category
        col_order = [chart_bmi_category + "_prevalence",
                     chart_bmi_category + "_ci_lower",
                     chart_bmi_category + "_ci_upper"]

        create_output(df, groups, col_order,
                      sheetname, empty_cols, write_cell,
                      output_file, df_part_ref,
                      filter_condition, replace_null_logic,
                      total, subtotal,
                      col1_row_subgroups,
                      sort_logic,
                      col1_row_order, col2_row_order,
                      disclosure_control,
                      la_lower_tier,
                      table_bmi_category,
                      participation,
                      write_variable)


def chart_rurality_6_multi(df, output_file, df_part_ref):
    groups = ["PupilUrbanRuralIndicator"]
    filter_condition = 'SchoolYear == "6"'
    replace_null_logic = {"PupilUrbanRuralIndicator": "Unknown"}
    total = True
    subtotal = False
    col1_row_subgroups = {"Urban": ["1", "5"], "Town_Fringe": ["2", "6"],
                          "Village_Hamlet_Isolated": ["3", "7", "4", "8"]}
    col1_row_order = ["grand_total", "Town_Fringe", "Urban", "Village_Hamlet_Isolated"]
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    write_variable = False
    empty_cols = []
    write_cell = "C7"
    chart_bmi_categories = ["underweight", "healthy_weight",
                            "overweight", "obese",
                            "severely_obese", "overweight_obese"]

    # outputs year 6 rurality data for each bmi category onto  onto a separate tab
    for chart_bmi_category in chart_bmi_categories:
        sheetname = "Rurality_SchYr_" + chart_bmi_category
        col_order = [chart_bmi_category + "_prevalence",
                     chart_bmi_category + "_ci_lower",
                     chart_bmi_category + "_ci_upper"]

        create_output(df, groups, col_order,
                      sheetname, empty_cols, write_cell,
                      output_file, df_part_ref,
                      filter_condition, replace_null_logic,
                      total, subtotal,
                      col1_row_subgroups,
                      sort_logic,
                      col1_row_order, col2_row_order,
                      disclosure_control,
                      la_lower_tier,
                      table_bmi_category,
                      participation,
                      write_variable)


def chart_ons_area_R_multi(df, output_file, df_part_ref):
    groups = ["PupilOnsSupergroupCode"]
    filter_condition = 'SchoolYear == "R"'
    replace_null_logic = {"PupilOnsSupergroupCode": "Unknown"}
    total = True
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = ["grand_total", "1", "6", "7", "5", "2", "3", "4"]
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    write_variable = False
    empty_cols = []
    write_cell = "C3"
    chart_bmi_categories = ["underweight", "healthy_weight",
                            "overweight", "obese",
                            "severely_obese", "overweight_obese"]

    # outputs reception ONS area data for each bmi category onto a separate tab
    for chart_bmi_category in chart_bmi_categories:
        sheetname = "ONSArea_SchYr_" + chart_bmi_category
        col_order = [chart_bmi_category + "_prevalence",
                     chart_bmi_category + "_ci_lower",
                     chart_bmi_category + "_ci_upper"]

        create_output(df, groups, col_order,
                      sheetname, empty_cols, write_cell,
                      output_file, df_part_ref,
                      filter_condition, replace_null_logic,
                      total, subtotal,
                      col1_row_subgroups,
                      sort_logic,
                      col1_row_order, col2_row_order,
                      disclosure_control,
                      la_lower_tier,
                      table_bmi_category,
                      participation,
                      write_variable)


def chart_ons_area_6_multi(df, output_file, df_part_ref):
    groups = ["PupilOnsSupergroupCode"]
    filter_condition = 'SchoolYear == "6"'
    replace_null_logic = {"PupilOnsSupergroupCode": "Unknown"}
    total = True
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = ["grand_total", "1", "6", "7", "5", "2", "3", "4"]
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    write_variable = False
    empty_cols = []
    write_cell = "C11"
    chart_bmi_categories = ["underweight", "healthy_weight",
                            "overweight", "obese",
                            "severely_obese", "overweight_obese"]

    # outputs year 6 ONS area data for each bmi category onto a separate tab
    for chart_bmi_category in chart_bmi_categories:
        sheetname = "ONSArea_SchYr_" + chart_bmi_category
        col_order = [chart_bmi_category + "_prevalence",
                     chart_bmi_category + "_ci_lower",
                     chart_bmi_category + "_ci_upper"]

        create_output(df, groups, col_order,
                      sheetname, empty_cols, write_cell,
                      output_file, df_part_ref,
                      filter_condition, replace_null_logic,
                      total, subtotal,
                      col1_row_subgroups,
                      sort_logic,
                      col1_row_order, col2_row_order,
                      disclosure_control,
                      la_lower_tier,
                      table_bmi_category,
                      participation,
                      write_variable)


def chart_ethnicity_R_multi(df, output_file, df_part_ref):
    groups = ["NcmpEthnicityCode"]
    filter_condition = 'SchoolYear == "R"'
    replace_null_logic = {"NcmpEthnicityCode": "Not stated"}
    total = True
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = ["grand_total", "Asian", "Black", "Chinese", "Mixed",
                      "White", "Any other ethnic group"]
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    write_variable = False
    empty_cols = []
    write_cell = "C3"
    chart_bmi_categories = ["underweight", "healthy_weight",
                            "overweight", "obese",
                            "severely_obese", "overweight_obese"]

    # outputs reception ethnicity data for each bmi category onto a separate tab
    for chart_bmi_category in chart_bmi_categories:
        sheetname = "Ethnic_SchYr_" + chart_bmi_category
        col_order = [chart_bmi_category + "_prevalence",
                     chart_bmi_category + "_ci_lower",
                     chart_bmi_category + "_ci_upper"]

        create_output(df, groups, col_order,
                      sheetname, empty_cols, write_cell,
                      output_file, df_part_ref,
                      filter_condition, replace_null_logic,
                      total, subtotal,
                      col1_row_subgroups,
                      sort_logic,
                      col1_row_order, col2_row_order,
                      disclosure_control,
                      la_lower_tier,
                      table_bmi_category,
                      participation,
                      write_variable)


def chart_ethnicity_6_multi(df, output_file, df_part_ref):
    groups = ["NcmpEthnicityCode"]
    filter_condition = 'SchoolYear == "6"'
    replace_null_logic = {"NcmpEthnicityCode": "Not stated"}
    total = True
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = ["grand_total", "Asian", "Black", "Chinese", "Mixed",
                      "White", "Any other ethnic group"]
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    write_variable = False
    empty_cols = []
    write_cell = "C10"
    chart_bmi_categories = ["underweight", "healthy_weight",
                            "overweight", "obese",
                            "severely_obese", "overweight_obese"]

    # outputs year 6 ethnicity data for each bmi category onto a separate tab
    for chart_bmi_category in chart_bmi_categories:
        sheetname = "Ethnic_SchYr_" + chart_bmi_category
        col_order = [chart_bmi_category + "_prevalence",
                     chart_bmi_category + "_ci_lower",
                     chart_bmi_category + "_ci_upper"]

        create_output(df, groups, col_order,
                      sheetname, empty_cols, write_cell,
                      output_file, df_part_ref,
                      filter_condition, replace_null_logic,
                      total, subtotal,
                      col1_row_subgroups,
                      sort_logic,
                      col1_row_order, col2_row_order,
                      disclosure_control,
                      la_lower_tier,
                      table_bmi_category,
                      participation,
                      write_variable)


def chart_obese_pupil_imd_R_ts(df, output_file, df_part_ref):
    groups = ["PupilIndexOfMultipleDeprivationD"]
    filter_condition = 'SchoolYear == "R"'
    replace_null_logic = {"PupilIndexOfMultipleDeprivationD": "Unknown"}
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = ["1.0", "10.0"]
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    write_variable = False
    sheetname = "Obese_PupIMDGap_TSeries_R"
    empty_cols = []
    write_cell = "C3"
    col_order = ["obese_prevalence"]

    return create_output(df, groups, col_order,
                         sheetname, empty_cols, write_cell,
                         output_file, df_part_ref,
                         filter_condition, replace_null_logic,
                         total, subtotal,
                         col1_row_subgroups,
                         sort_logic,
                         col1_row_order, col2_row_order,
                         disclosure_control,
                         la_lower_tier,
                         table_bmi_category,
                         participation,
                         write_variable)


def chart_obese_pupil_imd_gender_R_ts(df, output_file, df_part_ref):
    groups = ["GenderCode", "PupilIndexOfMultipleDeprivationD"]
    filter_condition = 'SchoolYear == "R"'
    replace_null_logic = {"PupilIndexOfMultipleDeprivationD": "Unknown"}
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = ["1", "2"]
    col2_row_order = ["1.0", "10.0"]
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    write_variable = False
    sheetname = "Obese_PupIMDGap_Sex_TSeries_R"
    empty_cols = []
    write_cell = "D3"
    col_order = ["obese_prevalence"]

    return create_output(df, groups, col_order,
                         sheetname, empty_cols, write_cell,
                         output_file, df_part_ref,
                         filter_condition, replace_null_logic,
                         total, subtotal,
                         col1_row_subgroups,
                         sort_logic,
                         col1_row_order, col2_row_order,
                         disclosure_control,
                         la_lower_tier,
                         table_bmi_category,
                         participation,
                         write_variable)


def chart_obese_pupil_imd_6_ts(df, output_file, df_part_ref):
    groups = ["PupilIndexOfMultipleDeprivationD"]
    filter_condition = 'SchoolYear == "6"'
    replace_null_logic = {"PupilIndexOfMultipleDeprivationD": "Unknown"}
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = ["1.0", "10.0"]
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    write_variable = False
    sheetname = "Obese_PupIMDGap_TSeries_6"
    empty_cols = []
    write_cell = "C3"
    col_order = ["obese_prevalence"]

    return create_output(df, groups, col_order,
                         sheetname, empty_cols, write_cell,
                         output_file, df_part_ref,
                         filter_condition, replace_null_logic,
                         total, subtotal,
                         col1_row_subgroups,
                         sort_logic,
                         col1_row_order, col2_row_order,
                         disclosure_control,
                         la_lower_tier,
                         table_bmi_category,
                         participation,
                         write_variable)


def chart_obese_pupil_imd_gender_6_ts(df, output_file, df_part_ref):
    groups = ["GenderCode", "PupilIndexOfMultipleDeprivationD"]
    filter_condition = 'SchoolYear == "6"'
    replace_null_logic = {"PupilIndexOfMultipleDeprivationD": "Unknown"}
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = ["1", "2"]
    col2_row_order = ["1.0", "10.0"]
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    write_variable = False
    sheetname = "Obese_PupIMDGap_Sex_TSeries_6"
    empty_cols = []
    write_cell = "D3"
    col_order = ["obese_prevalence"]

    return create_output(df, groups, col_order,
                         sheetname, empty_cols, write_cell,
                         output_file, df_part_ref,
                         filter_condition, replace_null_logic,
                         total, subtotal,
                         col1_row_subgroups,
                         sort_logic,
                         col1_row_order, col2_row_order,
                         disclosure_control,
                         la_lower_tier,
                         table_bmi_category,
                         participation,
                         write_variable)


def chart_severely_obese_pupil_imd_R_ts(df, output_file, df_part_ref):
    groups = ["PupilIndexOfMultipleDeprivationD"]
    filter_condition = 'SchoolYear == "R"'
    replace_null_logic = {"PupilIndexOfMultipleDeprivationD": "Unknown"}
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = ["1.0", "10.0"]
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    write_variable = False
    sheetname = "SevOb_PupIMDGap_TSeries_R"
    empty_cols = []
    write_cell = "C3"
    col_order = ["severely_obese_prevalence"]

    return create_output(df, groups, col_order,
                         sheetname, empty_cols, write_cell,
                         output_file, df_part_ref,
                         filter_condition, replace_null_logic,
                         total, subtotal,
                         col1_row_subgroups,
                         sort_logic,
                         col1_row_order, col2_row_order,
                         disclosure_control,
                         la_lower_tier,
                         table_bmi_category,
                         participation,
                         write_variable)


def chart_severely_obese_pupil_imd_gender_R_ts(df, output_file, df_part_ref):
    groups = ["GenderCode", "PupilIndexOfMultipleDeprivationD"]
    filter_condition = 'SchoolYear == "R"'
    replace_null_logic = {"PupilIndexOfMultipleDeprivationD": "Unknown"}
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = ["1", "2"]
    col2_row_order = ["1.0", "10.0"]
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    write_variable = False
    sheetname = "SevOb_PupIMDGap_Sex_TSeries_R"
    empty_cols = []
    write_cell = "D3"
    col_order = ["severely_obese_prevalence"]

    return create_output(df, groups, col_order,
                         sheetname, empty_cols, write_cell,
                         output_file, df_part_ref,
                         filter_condition, replace_null_logic,
                         total, subtotal,
                         col1_row_subgroups,
                         sort_logic,
                         col1_row_order, col2_row_order,
                         disclosure_control,
                         la_lower_tier,
                         table_bmi_category,
                         participation,
                         write_variable)


def chart_severely_obese_pupil_imd_6_ts(df, output_file, df_part_ref):
    groups = ["PupilIndexOfMultipleDeprivationD"]
    filter_condition = 'SchoolYear == "6"'
    replace_null_logic = {"PupilIndexOfMultipleDeprivationD": "Unknown"}
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = ["1.0", "10.0"]
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    write_variable = False
    sheetname = "SevOb_PupIMDGap_TSeries_6"
    empty_cols = []
    write_cell = "C3"
    col_order = ["severely_obese_prevalence"]

    return create_output(df, groups, col_order,
                         sheetname, empty_cols, write_cell,
                         output_file, df_part_ref,
                         filter_condition, replace_null_logic,
                         total, subtotal,
                         col1_row_subgroups,
                         sort_logic,
                         col1_row_order, col2_row_order,
                         disclosure_control,
                         la_lower_tier,
                         table_bmi_category,
                         participation,
                         write_variable)


def chart_severely_obese_pupil_imd_gender_6_ts(df, output_file, df_part_ref):
    groups = ["GenderCode", "PupilIndexOfMultipleDeprivationD"]
    filter_condition = 'SchoolYear == "6"'
    replace_null_logic = {"PupilIndexOfMultipleDeprivationD": "Unknown"}
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = ["1", "2"]
    col2_row_order = ["1.0", "10.0"]
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    write_variable = False
    sheetname = "SevOb_PupIMDGap_Sex_TSeries_6"
    empty_cols = []
    write_cell = "D3"
    col_order = ["severely_obese_prevalence"]

    return create_output(df, groups, col_order,
                         sheetname, empty_cols, write_cell,
                         output_file, df_part_ref,
                         filter_condition, replace_null_logic,
                         total, subtotal,
                         col1_row_subgroups,
                         sort_logic,
                         col1_row_order, col2_row_order,
                         disclosure_control,
                         la_lower_tier,
                         table_bmi_category,
                         participation,
                         write_variable)


def chart_dq_participation(df_la, df_la_dq, df_la_dq_sysbreachref,
                           la_dq_col_order, output_file):
    total_only = True
    col_order = ["derived_participation_R",
                 "derived_participation_6",
                 "derived_participation_overall"]
    sheetname = "DQ_Participation"
    empty_cols = []
    write_cell = "B4"
    write_type = "excel_static"

    return create_output_dq(df_la, df_la_dq, df_la_dq_sysbreachref,
                            la_dq_col_order, sheetname, empty_cols, write_cell,
                            output_file, write_type,
                            total_only, col_order)


def chart_dq_HW_whole_num(df_la, df_la_dq, df_la_dq_sysbreachref,
                          la_dq_col_order, output_file):
    total_only = True
    col_order = ["PercentageWholeNumberHeights",
                 "PercentageWholeNumberWeights"]
    sheetname = "DQ_Height_Weight_WholeNum"
    empty_cols = []
    write_cell = "B4"
    write_type = "excel_static"

    return create_output_dq(df_la, df_la_dq, df_la_dq_sysbreachref,
                            la_dq_col_order, sheetname, empty_cols, write_cell,
                            output_file, write_type,
                            total_only, col_order)


def chart_dq_missing(df_la, df_la_dq, df_la_dq_sysbreachref,
                     la_dq_col_order, output_file):

    total_only = True
    col_order = ["PercentageBlankPostcode",
                 "PercentageEthnicGroupUnknown",
                 "PercentageBlankNhsNumber"]
    sheetname = "DQ_Missing_Blank_Unknown"
    empty_cols = []
    write_cell = "B4"
    write_type = "excel_static"

    return create_output_dq(df_la, df_la_dq, df_la_dq_sysbreachref,
                            la_dq_col_order, sheetname, empty_cols, write_cell,
                            output_file, write_type,
                            total_only, col_order)
