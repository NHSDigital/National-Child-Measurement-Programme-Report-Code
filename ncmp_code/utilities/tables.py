from ncmp_code.utilities.processing_steps import create_output
from ncmp_code.utilities.processing_steps import create_output_dq


# Create the list of tables to produce
def get_tables():

    table_list = [table_schoolyear_gender,
                  table_academicyear_R,
                  table_academicyear_6,
                  table_ethnicity_category_R,
                  table_ethnicity_subcategory_R,
                  table_ethnicity_category_6,
                  table_ethnicity_subcategory_6,
                  table_org_submit_region_R,
                  table_org_submit_upper_tier_la_R,
                  table_org_submit_region_6,
                  table_org_submit_upper_tier_la_6,
                  table_org_school_upper_tier_region_R,
                  table_org_school_upper_tier_la_R,
                  table_org_school_upper_tier_region_6,
                  table_org_school_upper_tier_la_6,
                  table_org_school_lower_tier_region_R,
                  table_org_school_lower_tier_la_R,
                  table_org_school_lower_tier_region_6,
                  table_org_school_lower_tier_la_6,
                  table_org_pupil_region_R,
                  table_org_pupil_upper_tier_la_R,
                  table_org_pupil_region_6,
                  table_org_pupil_upper_tier_la_6,
                  table_org_pupil_lower_tier_region_R,
                  table_org_pupil_lower_tier_la_R,
                  table_org_pupil_lower_tier_region_6,
                  table_org_pupil_lower_tier_la_6,
                  table_rural_urban_pupil_6,
                  table_rural_urban_pupil_R,
                  table_rural_urban_school_6,
                  table_rural_urban_school_R,
                  table_imd_pupil_R,
                  table_imd_pupil_6,
                  table_imd_school_R,
                  table_imd_school_6,
                  table_imd_gender_obese_school_R,
                  table_imd_gender_obese_school_6,
                  table_imd_gender_severely_obese_school_R,
                  table_imd_gender_severely_obese_school_6,
                  table_ons_area_class_pupil_R,
                  table_ons_area_class_pupil_6,
                  table_ons_area_class_school_R,
                  table_ons_area_class_school_6,
                  table_imd_gender_obese_pupil_R,
                  table_imd_gender_obese_pupil_6,
                  table_imd_gender_severely_obese_pupil_R,
                  table_imd_gender_severely_obese_pupil_6]

    return table_list


def table_schoolyear_gender(df, col_order, output_file, df_part_ref):
    groups = ["SchoolYear", "GenderCode"]
    filter_condition = None
    replace_null_logic = None
    total = False
    subtotal = True
    col1_row_subgroups = None
    col1_row_order = ["R", "6"]
    col2_row_order = ["1", "2", "R - subtotal", "6 - subtotal"]
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    sheetname = "Table 1a"
    empty_cols = ["G", "L", "Q", "V", "AA", "AF"]
    write_cell = "C16"
    write_variable = False

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


# the following two functions output in the same worksheet in Excel
def table_academicyear_R(df, col_order, output_file, df_part_ref):
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
    sheetname = "Table 1b"
    empty_cols = ["F", "K", "P", "U", "Z", "AE"]
    write_cell = "B32" #UPDATE "write_cell" VALUE EACH YEAR
    write_variable = False

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


def table_academicyear_6(df, col_order, output_file, df_part_ref):
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
    sheetname = "Table 1b"
    empty_cols = ["F", "K", "P", "U", "Z", "AE"]
    write_cell = "B50" #UPDATE "write_cell" VALUE EACH YEAR
    write_variable = False

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


# the following two functions output in the same worksheet in Excel
def table_ethnicity_category_R(df, col_order, output_file, df_part_ref):
    groups = ["NcmpEthnicityCode"]
    filter_condition = 'SchoolYear == "R"'
    replace_null_logic = {"NcmpEthnicityCode": "Not stated"}
    total = True
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = ["grand_total", "Asian", "Black", "Chinese", "Mixed",
                      "White", "Any other ethnic group", "Not stated"]
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    sheetname = "Table 4_R"
    empty_cols = ["G", "L", "Q", "V", "AA"]
    write_cell = "C16"
    write_variable = False

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


def table_ethnicity_subcategory_R(df, col_order, output_file, df_part_ref):
    groups = ["NcmpEthnicityCode", "NhsEthnicityCode"]
    filter_condition = 'SchoolYear == "R"'
    replace_null_logic = {"NhsEthnicityCode": "Z",
                          "NcmpEthnicityCode": "Not stated"}
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = ["Asian", "Black", "Mixed", "White"]
    col2_row_order = ["K", "H", "J", "L",
                      "N", "M", "P",
                      "F", "E", "D", "G",
                      "A", "B", "C"]
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    sheetname = "Table 4_R"
    empty_cols = ["G", "L", "Q", "V", "AA"]
    write_cell = "C25"
    write_variable = False

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


# the following two functions output in the same worksheet in Excel
def table_ethnicity_category_6(df, col_order, output_file, df_part_ref):
    groups = ["NcmpEthnicityCode"]
    filter_condition = 'SchoolYear == "6"'
    replace_null_logic = {"NcmpEthnicityCode": "Not stated"}
    total = True
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = ["grand_total", "Asian", "Black", "Chinese", "Mixed",
                      "White", "Any other ethnic group", "Not stated"]
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    sheetname = "Table 4_6"
    empty_cols = ["G", "L", "Q", "V", "AA"]
    write_cell = "C16"
    write_variable = False

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


def table_ethnicity_subcategory_6(df, col_order, output_file, df_part_ref):
    groups = ["NcmpEthnicityCode", "NhsEthnicityCode"]
    filter_condition = 'SchoolYear == "6"'
    replace_null_logic = {"NhsEthnicityCode": "Z",
                          "NcmpEthnicityCode": "Not stated"}
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = ["Asian", "Black", "Mixed", "White"]
    col2_row_order = ["K", "H", "J", "L",
                      "N", "M", "P",
                      "F", "E", "D", "G",
                      "A", "B", "C"]
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    sheetname = "Table 4_6"
    empty_cols = ["G", "L", "Q", "V", "AA"]
    write_cell = "C25"
    write_variable = False

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


# the following two functions output in the same worksheet in Excel
def table_org_school_upper_tier_region_R(df, col_order, output_file, df_part_ref):
    groups = ["SchoolRegionCode"]
    filter_condition = 'SchoolYear == "R"'
    replace_null_logic = None
    total = True
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = None
    col2_row_order = None
    sort_logic = {"SchoolRegionCode": True}
    disclosure_control = True
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    sheetname = "Table 3a_R_UTLA"
    empty_cols = ["H", "M", "R", "W", "AB"]
    write_cell = "D16"
    write_variable = False

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


def table_org_school_upper_tier_la_R(df, col_order, output_file, df_part_ref):
    groups = ["SchoolTier1LocalAuthority",
              "SchoolTier1LocalAuthority_Name",
              "SchoolRegionCode",
              "SchoolRegionCode_Name"]
    filter_condition = 'SchoolYear == "R"'
    replace_null_logic = None
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = None
    col2_row_order = None
    sort_logic = {"SchoolRegionCode": True,
                  "SchoolTier1LocalAuthority_Name": True}
    disclosure_control = True
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    sheetname = "Table 3a_R_UTLA"
    empty_cols = ["H", "M", "R", "W", "AB"]
    write_cell = "A27"
    write_variable = True

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


# the following two functions output in the same worksheet in Excel
def table_org_school_upper_tier_region_6(df, col_order, output_file, df_part_ref):
    groups = ["SchoolRegionCode"]
    filter_condition = 'SchoolYear == "6"'
    replace_null_logic = None
    total = True
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = None
    col2_row_order = None
    sort_logic = {"SchoolRegionCode": True}
    disclosure_control = True
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    sheetname = "Table 3a_6_UTLA"
    empty_cols = ["H", "M", "R", "W", "AB"]
    write_cell = "D16"
    write_variable = False

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


def table_org_school_upper_tier_la_6(df, col_order, output_file, df_part_ref):
    groups = ["SchoolTier1LocalAuthority",
              "SchoolTier1LocalAuthority_Name",
              "SchoolRegionCode",
              "SchoolRegionCode_Name"]
    filter_condition = 'SchoolYear == "6"'
    replace_null_logic = None
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = None
    col2_row_order = None
    sort_logic = {"SchoolRegionCode": True,
                  "SchoolTier1LocalAuthority_Name": True}
    disclosure_control = True
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    sheetname = "Table 3a_6_UTLA"
    empty_cols = ["H", "M", "R", "W", "AB"]
    write_cell = "A27"
    write_variable = True

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


# the following two functions output in the same worksheet in Excel
def table_org_school_lower_tier_region_R(df, col_order, output_file, df_part_ref):
    groups = ["SchoolRegionCode"]
    filter_condition = 'SchoolYear == "R"'
    replace_null_logic = {"SchoolRegionCode": "Unknown"}
    total = True
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = None
    col2_row_order = None
    sort_logic = {"SchoolRegionCode": True}
    disclosure_control = True
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    sheetname = "Table 3a_R_LTLA"
    empty_cols = ["I", "N", "S", "X", "AC"]
    write_cell = "E16"
    write_variable = False

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


def table_org_school_lower_tier_la_R(df, col_order, output_file, df_part_ref):
    groups = ["SchoolTier2LocalAuthority",
              "SchoolTier2LocalAuthority_Name",
              "SchoolTier1LocalAuthority_Name",
              "SchoolRegionCode",
              "SchoolRegionCode_Name"]
    filter_condition = 'SchoolYear == "R"'
    replace_null_logic = None
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = None
    col2_row_order = None
    sort_logic = {"SchoolRegionCode": True,
                  "SchoolTier1LocalAuthority_Name": True,
                  "SchoolTier2LocalAuthority_Name": True}
    disclosure_control = True
    la_lower_tier = True
    table_bmi_category = None
    participation = None
    sheetname = "Table 3a_R_LTLA"
    empty_cols = ["I", "N", "S", "X", "AC"]
    write_cell = "A27"
    write_variable = True

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


# the following two functions output in the same worksheet in Excel
def table_org_school_lower_tier_region_6(df, col_order, output_file, df_part_ref):
    groups = ["SchoolRegionCode"]
    filter_condition = 'SchoolYear == "6"'
    replace_null_logic = {"SchoolRegionCode": "Unknown"}
    total = True
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = None
    col2_row_order = None
    sort_logic = {"SchoolRegionCode": True}
    disclosure_control = True
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    sheetname = "Table 3a_6_LTLA"
    empty_cols = ["I", "N", "S", "X", "AC"]
    write_cell = "E16"
    write_variable = False

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


def table_org_school_lower_tier_la_6(df, col_order, output_file, df_part_ref):
    groups = ["SchoolTier2LocalAuthority",
              "SchoolTier2LocalAuthority_Name",
              "SchoolTier1LocalAuthority_Name",
              "SchoolRegionCode",
              "SchoolRegionCode_Name"]
    filter_condition = 'SchoolYear == "6"'
    replace_null_logic = None
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = None
    col2_row_order = None
    sort_logic = {"SchoolRegionCode": True,
                  "SchoolTier1LocalAuthority_Name": True,
                  "SchoolTier2LocalAuthority_Name": True}
    disclosure_control = True
    la_lower_tier = True
    table_bmi_category = None
    participation = None
    sheetname = "Table 3a_6_LTLA"
    empty_cols = ["I", "N", "S", "X", "AC"]
    write_cell = "A27"
    write_variable = True

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


# the following two functions output in the same worksheet in Excel
def table_org_pupil_region_R(df, col_order, output_file, df_part_ref):
    groups = ["PupilRegionCode"]
    filter_condition = 'SchoolYear == "R"'
    replace_null_logic = {"PupilRegionCode": "Unknown"}
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
    sheetname = "Table 3b_R_UTLA"
    empty_cols = ["H", "M", "R", "W", "AB"]
    write_cell = "D16"
    write_variable = False

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


def table_org_pupil_upper_tier_la_R(df, col_order, output_file, df_part_ref):
    groups = ["PupilTier1LocalAuthority",
              "PupilTier1LocalAuthority_Name",
              "PupilRegionCode",
              "PupilRegionCode_Name"]
    filter_condition = 'SchoolYear == "R"'
    replace_null_logic = None
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = None
    col2_row_order = None
    sort_logic = {"PupilRegionCode": True,
                  "PupilTier1LocalAuthority_Name": True}
    disclosure_control = True
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    sheetname = "Table 3b_R_UTLA"
    empty_cols = ["H", "M", "R", "W", "AB"]
    write_cell = "A27"
    write_variable = True

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


# the following two functions output in the same worksheet in Excel
def table_org_pupil_region_6(df, col_order, output_file, df_part_ref):
    groups = ["PupilRegionCode"]
    filter_condition = 'SchoolYear == "6"'
    replace_null_logic = {"PupilRegionCode": "Unknown"}
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
    sheetname = "Table 3b_6_UTLA"
    empty_cols = ["H", "M", "R", "W", "AB"]
    write_cell = "D16"
    write_variable = False

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


def table_org_pupil_upper_tier_la_6(df, col_order, output_file, df_part_ref):
    groups = ["PupilTier1LocalAuthority",
              "PupilTier1LocalAuthority_Name",
              "PupilRegionCode",
              "PupilRegionCode_Name"]
    filter_condition = 'SchoolYear == "6"'
    replace_null_logic = None
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = None
    col2_row_order = None
    sort_logic = {"PupilRegionCode": True,
                  "PupilTier1LocalAuthority_Name": True}
    disclosure_control = True
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    sheetname = "Table 3b_6_UTLA"
    empty_cols = ["H", "M", "R", "W", "AB"]
    write_cell = "A27"
    write_variable = True

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


# the following two functions output in the same worksheet in Excel
def table_rural_urban_pupil_6(df, col_order, output_file, df_part_ref):
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
    sheetname = "Table 5a_6"
    empty_cols = ["F", "K", "P", "U", "Z"]
    write_cell = "B16"
    write_variable = False

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


def table_rural_urban_pupil_R(df, col_order, output_file, df_part_ref):
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
    sheetname = "Table 5a_R"
    empty_cols = ["F", "K", "P", "U", "Z"]
    write_cell = "B16"
    write_variable = False

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


def table_imd_pupil_R(df, col_order, output_file, df_part_ref):
    groups = ["PupilIndexOfMultipleDeprivationD"]
    filter_condition = 'SchoolYear == "R"'
    replace_null_logic = {"PupilIndexOfMultipleDeprivationD": "Unknown"}
    total = True
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = ["grand_total", "1.0", "2.0", "3.0", "4.0", "5.0", "6.0",
                      "7.0", "8.0", "9.0", "10.0"]
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    sheetname = "Table 6a_R"
    empty_cols = ["F", "K", "P", "U", "Z"]
    write_cell = "B16"
    write_variable = False

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


def table_imd_pupil_6(df, col_order, output_file, df_part_ref):
    groups = ["PupilIndexOfMultipleDeprivationD"]
    filter_condition = 'SchoolYear == "6"'
    replace_null_logic = {"PupilIndexOfMultipleDeprivationD": "Unknown"}
    total = True
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = ["grand_total", "1.0", "2.0", "3.0", "4.0", "5.0", "6.0",
                      "7.0", "8.0", "9.0", "10.0"]
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    sheetname = "Table 6a_6"
    empty_cols = ["F", "K", "P", "U", "Z"]
    write_cell = "B16"
    write_variable = False

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


# the following two functions output in the same worksheet in Excel
def table_imd_gender_obese_school_R(df, col_order, output_file, df_part_ref):
    groups = ["SchoolIndexOfMultiDeprivationD", "GenderCode"]
    filter_condition = 'SchoolYear == "R" & (SchoolIndexOfMultiDeprivationD == 1 | SchoolIndexOfMultiDeprivationD == 10)'
    replace_null_logic = {"SchoolIndexOfMultiDeprivationD": "Unknown"}
    total = False
    subtotal = True
    col1_row_subgroups = None
    col1_row_order = None
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = "obese"
    participation = None
    sheetname = "Table 6c"
    empty_cols = ["F", "K", "P", "U", "Z"]
    write_cell = "B32" #UPDATE "write_cell" VALUE EACH YEAR
    write_variable = False

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


def table_imd_gender_obese_school_6(df, col_order, output_file, df_part_ref):
    groups = ["SchoolIndexOfMultiDeprivationD", "GenderCode"]
    filter_condition = 'SchoolYear == "6" & (SchoolIndexOfMultiDeprivationD == 1 | SchoolIndexOfMultiDeprivationD == 10)'
    replace_null_logic = {"SchoolIndexOfMultiDeprivationD": "Unknown"}
    total = False
    subtotal = True
    col1_row_subgroups = None
    col1_row_order = None
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = "obese"
    participation = None
    sheetname = "Table 6c"
    empty_cols = ["F", "K", "P", "U", "Z"]
    write_cell = "B50" #UPDATE "write_cell" VALUE EACH YEAR
    write_variable = False

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


# the following two functions output in the same worksheet in Excel
def table_imd_gender_severely_obese_school_R(df, col_order, output_file, df_part_ref):
    groups = ["SchoolIndexOfMultiDeprivationD", "GenderCode"]
    filter_condition = 'SchoolYear == "R" & (SchoolIndexOfMultiDeprivationD == 1 | SchoolIndexOfMultiDeprivationD == 10)'
    replace_null_logic = {"SchoolIndexOfMultiDeprivationD": "Unknown"}
    total = False
    subtotal = True
    col1_row_subgroups = None
    col1_row_order = None
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = "severely_obese"
    participation = None
    sheetname = "Table 6d"
    empty_cols = ["F", "K", "P", "U", "Z"]
    write_cell = "B32" #UPDATE "write_cell" VALUE EACH YEAR
    write_variable = False

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


def table_imd_gender_severely_obese_school_6(df, col_order, output_file, df_part_ref):
    groups = ["SchoolIndexOfMultiDeprivationD", "GenderCode"]
    filter_condition = 'SchoolYear == "6" & (SchoolIndexOfMultiDeprivationD == 1 | SchoolIndexOfMultiDeprivationD == 10)'
    replace_null_logic = {"SchoolIndexOfMultiDeprivationD": "Unknown"}
    total = False
    subtotal = True
    col1_row_subgroups = None
    col1_row_order = None
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    participation = None
    table_bmi_category = "severely_obese"
    sheetname = "Table 6d"
    empty_cols = ["F", "K", "P", "U", "Z"]
    write_cell = "B50" #UPDATE "write_cell" VALUE EACH YEAR
    write_variable = False

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


def table_rural_urban_school_6(df, col_order, output_file, df_part_ref):
    groups = ["SchoolUrbanRuralIndicator"]
    filter_condition = 'SchoolYear == "6"'
    replace_null_logic = {"SchoolUrbanRuralIndicator": "Unknown"}
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
    sheetname = "Table 5b_6"
    empty_cols = ["F", "K", "P", "U", "Z"]
    write_cell = "B16"
    write_variable = False

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


def table_rural_urban_school_R(df, col_order, output_file, df_part_ref):
    groups = ["SchoolUrbanRuralIndicator"]
    filter_condition = 'SchoolYear == "R"'
    replace_null_logic = {"SchoolUrbanRuralIndicator": "Unknown"}
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
    sheetname = "Table 5b_R"
    empty_cols = ["F", "K", "P", "U", "Z"]
    write_cell = "B16"
    write_variable = False

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


def table_imd_school_R(df, col_order, output_file, df_part_ref):
    groups = ["SchoolIndexOfMultiDeprivationD"]
    filter_condition = 'SchoolYear == "R"'
    replace_null_logic = {"SchoolIndexOfMultiDeprivationD": "Unknown"}
    total = True
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = ["grand_total", "1.0", "2.0", "3.0", "4.0", "5.0", "6.0",
                      "7.0", "8.0", "9.0", "10.0"]
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    sheetname = "Table 6b_R"
    empty_cols = ["F", "K", "P", "U", "Z"]
    write_cell = "B16"
    write_variable = False

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


def table_imd_school_6(df, col_order, output_file, df_part_ref):
    groups = ["SchoolIndexOfMultiDeprivationD"]
    filter_condition = 'SchoolYear == "6"'
    replace_null_logic = {"SchoolIndexOfMultiDeprivationD": "Unknown"}
    total = True
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = ["grand_total", "1.0", "2.0", "3.0", "4.0", "5.0", "6.0",
                      "7.0", "8.0", "9.0", "10.0"]
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = None
    participation = None
    sheetname = "Table 6b_6"
    empty_cols = ["F", "K", "P", "U", "Z"]
    write_cell = "B16"
    write_variable = False

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


def table_ons_area_class_pupil_R(df, col_order, output_file, df_part_ref):
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
    sheetname = "Table 7a_R"
    empty_cols = ["F", "K", "P", "U", "Z"]
    write_cell = "B16"
    write_variable = False

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


def table_ons_area_class_pupil_6(df, col_order, output_file, df_part_ref):
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
    sheetname = "Table 7a_6"
    empty_cols = ["F", "K", "P", "U", "Z"]
    write_cell = "B16"
    write_variable = False

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


def table_ons_area_class_school_R(df, col_order, output_file, df_part_ref):
    groups = ["SchoolOnsSupergroupCode"]
    filter_condition = 'SchoolYear == "R"'
    replace_null_logic = {"SchoolOnsSupergroupCode": "Unknown"}
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
    sheetname = "Table 7b_R"
    empty_cols = ["F", "K", "P", "U", "Z"]
    write_cell = "B16"
    write_variable = False

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


def table_ons_area_class_school_6(df, col_order, output_file, df_part_ref):
    groups = ["SchoolOnsSupergroupCode"]
    filter_condition = 'SchoolYear == "6"'
    replace_null_logic = {"SchoolOnsSupergroupCode": "Unknown"}
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
    sheetname = "Table 7b_6"
    empty_cols = ["F", "K", "P", "U", "Z"]
    write_cell = "B16"
    write_variable = False

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


# the following two functions output in the same worksheet in Excel
def table_org_submit_region_R(df, col_order, output_file, df_part_ref):
    groups = ["RegionCode"]
    filter_condition = 'SchoolYear == "R"'
    replace_null_logic = None
    total = True
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = None
    col2_row_order = None
    sort_logic = {"RegionCode": True}
    disclosure_control = True
    la_lower_tier = False
    table_bmi_category = None
    participation = {"org_code_col": "RegionCode",
                     "participation_col": "derived_participation_R"}
    sheetname = "Table 2_R"
    empty_cols = ["H", "M", "R", "W", "AB"]
    write_cell = "D16"
    write_variable = False

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


def table_org_submit_upper_tier_la_R(df, col_order, output_file, df_part_ref):
    groups = ["OrgCode_ONS", "OrgCode_ONS_Name",
              "RegionCode", "RegionCode_Name"]
    filter_condition = 'SchoolYear == "R"'
    replace_null_logic = None
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = None
    col2_row_order = None
    sort_logic = {"RegionCode": True, "OrgCode_ONS_Name": True}
    disclosure_control = True
    la_lower_tier = False
    table_bmi_category = None
    participation = {"org_code_col": "OrgCode_ONS",
                     "participation_col": "derived_participation_R"}
    sheetname = "Table 2_R"
    empty_cols = ["H", "M", "R", "W", "AB"]
    write_cell = "A27"
    write_variable = True

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


# the following two functions output in the same worksheet in Excel
def table_org_submit_region_6(df, col_order, output_file, df_part_ref):
    groups = ["RegionCode"]
    filter_condition = 'SchoolYear == "6"'
    replace_null_logic = None
    total = True
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = None
    col2_row_order = None
    sort_logic = {"RegionCode": True}
    disclosure_control = True
    la_lower_tier = False
    table_bmi_category = None
    participation = {"org_code_col": "RegionCode",
                     "participation_col": "derived_participation_6"}
    sheetname = "Table 2_6"
    empty_cols = ["H", "M", "R", "W", "AB"]
    write_cell = "D16"
    write_variable = False

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


def table_org_submit_upper_tier_la_6(df, col_order, output_file, df_part_ref):
    groups = ["OrgCode_ONS", "OrgCode_ONS_Name",
              "RegionCode", "RegionCode_Name"]
    filter_condition = 'SchoolYear == "6"'
    replace_null_logic = None
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = None
    col2_row_order = None
    sort_logic = {"RegionCode": True, "OrgCode_ONS_Name": True}
    disclosure_control = True
    la_lower_tier = False
    table_bmi_category = None
    participation = {"org_code_col": "OrgCode_ONS",
                     "participation_col": "derived_participation_6"}
    sheetname = "Table 2_6"
    empty_cols = ["H", "M", "R", "W", "AB"]
    write_cell = "A27"
    write_variable = True

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


# the following two functions output in the same worksheet in Excel
def table_org_pupil_lower_tier_region_R(df, col_order, output_file, df_part_ref):
    groups = ["PupilRegionCode"]
    filter_condition = 'SchoolYear == "R"'
    replace_null_logic = {"PupilRegionCode": "Unknown"}
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
    sheetname = "Table 3b_R_LTLA"
    empty_cols = ["I", "N", "S", "X", "AC"]
    write_cell = "E16"
    write_variable = False

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


def table_org_pupil_lower_tier_la_R(df, col_order, output_file, df_part_ref):
    groups = ["PupilTier2LocalAuthority",
              "PupilTier2LocalAuthority_Name",
              "PupilTier1LocalAuthority_Name",
              "PupilRegionCode",
              "PupilRegionCode_Name"]
    filter_condition = 'SchoolYear == "R"'
    replace_null_logic = None
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = None
    col2_row_order = None
    sort_logic = {"PupilRegionCode": True,
                  "PupilTier1LocalAuthority_Name": True,
                  "PupilTier2LocalAuthority_Name": True}
    disclosure_control = True
    la_lower_tier = True
    table_bmi_category = None
    participation = None
    sheetname = "Table 3b_R_LTLA"
    empty_cols = ["I", "N", "S", "X", "AC"]
    write_cell = "A27"
    write_variable = True

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


# the following two functions output in the same worksheet in Excel
def table_org_pupil_lower_tier_region_6(df, col_order, output_file, df_part_ref):
    groups = ["PupilRegionCode"]
    filter_condition = 'SchoolYear == "6"'
    replace_null_logic = {"PupilRegionCode": "Unknown"}
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
    sheetname = "Table 3b_6_LTLA"
    empty_cols = ["I", "N", "S", "X", "AC"]
    write_cell = "E16"
    write_variable = False

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


def table_org_pupil_lower_tier_la_6(df, col_order, output_file, df_part_ref):
    groups = ["PupilTier2LocalAuthority",
              "PupilTier2LocalAuthority_Name",
              "PupilTier1LocalAuthority_Name",
              "PupilRegionCode",
              "PupilRegionCode_Name"]
    filter_condition = 'SchoolYear == "6"'
    replace_null_logic = None
    total = False
    subtotal = False
    col1_row_subgroups = None
    col1_row_order = None
    col2_row_order = None
    sort_logic = {"PupilRegionCode": True,
                  "PupilTier1LocalAuthority_Name": True,
                  "PupilTier2LocalAuthority_Name": True}
    disclosure_control = True
    la_lower_tier = True
    table_bmi_category = None
    participation = None
    sheetname = "Table 3b_6_LTLA"
    empty_cols = ["I", "N", "S", "X", "AC"]
    write_cell = "A27"
    write_variable = True

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


# the following two functions output in the same worksheet in Excel
def table_imd_gender_obese_pupil_R(df, col_order, output_file, df_part_ref):
    groups = ["PupilIndexOfMultipleDeprivationD", "GenderCode"]
    filter_condition = 'SchoolYear == "R" & (PupilIndexOfMultipleDeprivationD == 1 | PupilIndexOfMultipleDeprivationD == 10)'
    replace_null_logic = {"PupilIndexOfMultipleDeprivationD": "Unknown"}
    total = False
    subtotal = True
    col1_row_subgroups = None
    col1_row_order = None
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = "obese"
    participation = None
    sheetname = "Table 6e"
    empty_cols = ["F", "K", "P", "U", "Z"]
    write_cell = "B25" #UPDATE "write_cell" VALUE EACH YEAR
    write_variable = False

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


def table_imd_gender_obese_pupil_6(df, col_order, output_file, df_part_ref):
    groups = ["PupilIndexOfMultipleDeprivationD", "GenderCode"]
    filter_condition = 'SchoolYear == "6" & (PupilIndexOfMultipleDeprivationD == 1 | PupilIndexOfMultipleDeprivationD == 10)'
    replace_null_logic = {"PupilIndexOfMultipleDeprivationD": "Unknown"}
    total = False
    subtotal = True
    col1_row_subgroups = None
    col1_row_order = None
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = "obese"
    participation = None
    sheetname = "Table 6e"
    empty_cols = ["F", "K", "P", "U", "Z"]
    write_cell = "B36" #UPDATE "write_cell" VALUE EACH YEAR
    write_variable = False

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


# the following two functions output in the same worksheet in Excel
def table_imd_gender_severely_obese_pupil_R(df, col_order, output_file, df_part_ref):
    groups = ["PupilIndexOfMultipleDeprivationD", "GenderCode"]
    filter_condition = 'SchoolYear == "R" & (PupilIndexOfMultipleDeprivationD == 1 | PupilIndexOfMultipleDeprivationD == 10)'
    replace_null_logic = {"PupilIndexOfMultipleDeprivationD": "Unknown"}
    total = False
    subtotal = True
    col1_row_subgroups = None
    col1_row_order = None
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = "severely_obese"
    participation = None
    sheetname = "Table 6f"
    empty_cols = ["F", "K", "P", "U", "Z"]
    write_cell = "B25" #UPDATE "write_cell" VALUE EACH YEAR
    write_variable = False

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


def table_imd_gender_severely_obese_pupil_6(df, col_order, output_file, df_part_ref):
    groups = ["PupilIndexOfMultipleDeprivationD", "GenderCode"]
    filter_condition = 'SchoolYear == "6" & (PupilIndexOfMultipleDeprivationD == 1 | PupilIndexOfMultipleDeprivationD == 10)'
    replace_null_logic = {"PupilIndexOfMultipleDeprivationD": "Unknown"}
    total = False
    subtotal = True
    col1_row_subgroups = None
    col1_row_order = None
    col2_row_order = None
    sort_logic = None
    disclosure_control = False
    la_lower_tier = False
    table_bmi_category = "severely_obese"
    participation = None
    sheetname = "Table 6f"
    empty_cols = ["F", "K", "P", "U", "Z"]
    write_cell = "B36" #UPDATE "write_cell" VALUE EACH YEAR
    write_variable = False

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


def reporting_table_dq(df_la, df_la_dq, df_la_dq_sysbreachref,
                       la_dq_col_order, output_file):
    total_only = True
    col_order = ["derived_participation_R",
                 "derived_participation_6",
                 "derived_participation_overall",
                 "PercentageWholeNumberHeights",
                 "PercentageWholeNumberWeights",
                 "PercentageBlankPostcode",
                 "PercentageEthnicGroupUnknown",
                 "PercentageBlankNhsNumber"]
    sheetname = "Table A1 DQ indicators"
    empty_cols = []
    write_cell = "B4"
    write_type = "excel_static"

    return create_output_dq(df_la, df_la_dq, df_la_dq_sysbreachref,
                            la_dq_col_order, sheetname, empty_cols, write_cell,
                            output_file, write_type,
                            total_only, col_order)


def table_dq_measures(df_la, df_la_dq, df_la_dq_sysbreachref,
                      la_dq_col_order, output_file):
    total_only = False
    col_order = None
    sheetname = "Table 8"
    empty_cols = ["E", "H", "R", "V", "AC", "AE", "AV", "BA"]
    write_cell = "A29"
    write_type = "excel_variable"

    return create_output_dq(df_la, df_la_dq, df_la_dq_sysbreachref,
                            la_dq_col_order, sheetname, empty_cols, write_cell,
                            output_file, write_type,
                            total_only, col_order)
