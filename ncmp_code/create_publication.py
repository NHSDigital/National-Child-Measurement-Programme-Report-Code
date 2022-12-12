import time
import timeit
import logging
import pandas as pd
import xlwings as xw

from ncmp_code.utilities import logger_config
import ncmp_code.parameters as param
import ncmp_code.utilities.import_data as import_data
import ncmp_code.utilities.processing_steps as processing
import ncmp_code.utilities.tables as tables
import ncmp_code.utilities.charts as charts
import ncmp_code.utilities.post_deadline_validations as pd_vals
import ncmp_code.utilities.publication as publication
import ncmp_code.utilities.write_excel as write_excel
import ncmp_code.model.run_model as run_model


def main():

    # Load parameters
    server = param.SERVER
    database = param.DATABASE
    table = param.TABLE
    year = param.YEAR
    comp_year = param.COMP_YEAR
    bmi_categories = param.BMI_CATEGORIES
    col_order = param.COL_ORDER
    col_ref = param.LA_ORG_CODE_MAP_TO_NAMES
    ladq_file = param.LADQ_PATH
    la_dq_col_order = param.LA_DQ_COL_ORDER
    small_la_combine = param.SMALL_LA_COMBINE
    region_code_cols = param.REGION_CODE_COLS
    region_code_name_map = param.REGION_CODE_NAME_MAP
    ladq_sys_breach_file = param.LADQ_SYS_BREACH_REASONS_PATH

    # get raw data for this year from database
    df_raw = import_data.import_asset_data(server, database, table, year)

    # update and add org columns e.g. combine small LAs, add region names
    df_org_update = processing.update_org_columns(df_raw,
                                                  small_la_combine,
                                                  region_code_cols,
                                                  region_code_name_map)

    # get LA reference data
    df_la_ref = import_data.import_la_ref(server, database, year)

    # apply LA references
    df_la = processing.map_la_code_to_name(df_org_update, df_la_ref, col_ref)

    # get LA data quality data containing participation rates for each LA
    df_la_dq = import_data.import_la_dq_file(ladq_file)

    # map the participation data to ONS LA codes
    df_part_ref = processing.create_participation_reference(df_la, df_la_dq)

    # filter for NCMP schools with measurements
    df_ncmp_meas = df_la[(df_la["NcmpSchoolStatus"] == "NCMP") &
                         (df_la["Bmi"].notnull())].copy()

    # add weight category columns
    df_weight = processing.add_bmi_categories(df_ncmp_meas, bmi_categories).copy()

    # import LA DQ system standard breach reasons
    df_la_dq_sysbreachref = pd.read_csv(ladq_sys_breach_file)

    if param.RUN_PD_VALIDATIONS:
        # Run the post-deadline validations master file update process

        # import data for comparison year
        df_compyear_raw = import_data.import_asset_data(server, database, table,
                                                        comp_year)

        # filter comparison year data for NCMP schools with measurements
        df_compyear = df_compyear_raw[(df_compyear_raw["NcmpSchoolStatus"] == "NCMP") &
                                      (df_compyear_raw["Bmi"].notnull())].copy()

        # import school data for this year
        df_school_level = pd.read_csv(param.SCHOOL_DATA_PATH)

        # set output file for post deadline validations
        output_file = param.PD_VAL_OUTPUT_PATH

        # run all post-deadline validations
        validations_list_1 = pd_vals.get_validations_group1()
        validations_list_2 = pd_vals.get_validations_group2()
        validations_list_3 = pd_vals.get_validations_group3()

        for val in validations_list_1:
            val(df_compyear, df_weight, output_file)

        for val in validations_list_2:
            val(df_weight, output_file)

        for val in validations_list_3:
            val(df_school_level, df_la_dq, output_file)

        pd_vals.pd_validation_4_change_ethnicity(df_compyear_raw, df_la, output_file)
        pd_vals.pd_validation_8_num_measured_ncmpschoolstatus(df_la, output_file)

        # save and close the post deadline validations master file
        write_excel.save_close_output(output_file)

    if param.RUN_TABLES:
        # Run the data tables master file update process

        # Set output file for data tables
        output_file = param.TABLE_TEMPLATE

        # Run all tables
        table_list = tables.get_tables()

        for table in table_list:
            logging.info(f"Running {table}")
            table(df_weight, col_order, output_file, df_part_ref)

        logging.info("Running table_dq_measures")
        tables.table_dq_measures(df_la, df_la_dq, df_la_dq_sysbreachref,
                                 la_dq_col_order, output_file)

        # save and close the tables master file
        write_excel.save_close_output(output_file)

    if param.RUN_CHARTS:
        # Run the charts master file update process

        # Set output file for charts data
        output_file = param.CHART_TEMPLATE

        # Run all charts
        chart_list = charts.get_charts()
        chart_list_dq = charts.get_charts_dq()

        for chart in chart_list:
            logging.info(f"Running {chart}")
            chart(df_weight, output_file, df_part_ref)

        for chart in chart_list_dq:
            logging.info(f"Running {chart}")
            chart(df_la, df_la_dq, df_la_dq_sysbreachref,
                  la_dq_col_order, output_file)

        # Output report year to file for lookups

        # Create dataframe with report year for output
        df_year = pd.DataFrame({"ReportYear": [param.YEAR]})

        # Write to Excel output
        write_excel.write_to_excel_static(df_year, "ReportYears", None, "A2",
                                          output_file)

        # save and close the charts master file
        write_excel.save_close_output(output_file)

    if param.RUN_REPTABLES:

        output_file = param.REPTABLE_TEMPLATE

        logging.info("Running reporting_table_dq")
        tables.reporting_table_dq(df_la, df_la_dq, df_la_dq_sysbreachref,
                                  la_dq_col_order, output_file)

        write_excel.save_close_output(output_file)

    if param.RUN_DGMODEL:

        # Run the deprivation gap model master file update process
        output_file = param.DG_OUTPUT_PATH

        inputs = run_model.inputs

        for sheet in inputs:
            logging.info(f"Running gap analysis for {sheet}")
            df_raw = pd.read_excel(param.DG_INPUT_PATH, sheet_name=sheet,
                                   usecols='F:L', skiprows=2)
            DG_model = run_model.get_tidy_table(df_raw)
            result = run_model.run_OLS_model(DG_model)

            write_excel.write_to_excel_static(result, sheet, None, 'B2',
                                              output_file)

        # save and close the model master file
        write_excel.save_close_output(output_file)

    if param.RUN_PUBLICATION_OUTPUTS:
        # Run the publication outputs process
        # Creates final versions of master files and chart images for upload to web

        # Create year value without '/' for filenames
        reportyear = str(param.YEAR).replace("/", "-")

        # Data tables
        source_file = param.TABLE_TEMPLATE
        output_path = param.TAB_DIR
        save_name = "ncmp-chil-meas-prog-eng-" + reportyear + "-tab.xlsx"

        publication.save_masterfile(source_file, output_path, save_name)

        # Deprivation gap data model
        source_file = param.DG_OUTPUT_PATH
        output_path = param.DG_DIR
        save_name = "ncmp_deprivation_gap_model_" + reportyear + ".xlsx"

        publication.save_masterfile(source_file, output_path, save_name)

        # Report tables
        source_file = param.REPTABLE_TEMPLATE
        output_path = param.REPTAB_DIR
        save_name = "ncmp_reporttables_" + reportyear + ".xlsx"

        publication.save_masterfile(source_file, output_path, save_name)

        # Charts Excel file
        source_file = param.CHART_TEMPLATE
        output_path = param.CHART_DIR
        save_name = "ncmp_charts_" + reportyear + ".xlsx"

        publication.save_masterfile(source_file, output_path, save_name)

        # Chart images
        publication.save_charts_as_image(source_file, output_path, reportyear)

        # Close Excel
        xw.apps.active.api.Quit()


if __name__ == "__main__":

    # Setup logging
    formatted_time = time.strftime("%Y%m%d-%H%M%S")
    logger = logger_config.setup_logger(
        # Setup file & path for log, as_posix returns the path as a string
        file_name=(
            param.LOG_DIR / f"ncmp_create_pub_{formatted_time}.log"
        ).as_posix())

    start_time = timeit.default_timer()
    main()
    total_time = timeit.default_timer() - start_time
    logging.info(
        f"Running time of create_publication: {int(total_time / 60)} minutes and {round(total_time%60)} seconds.")
    logger_config.clean_up_handlers(logger)
