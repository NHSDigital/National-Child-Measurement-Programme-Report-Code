# Set the parameters for the project
import pathlib

# Sets the file paths for the project
BASE_DIR = pathlib.Path(r"projectfilepath")

INPUT_DIR = BASE_DIR / "Inputs"
OUTPUT_DIR = BASE_DIR / "Outputs"

# RAP > Inputs >
SCHOOL_DATA_DIR = INPUT_DIR / "SchoolLevelData"
LA_DQ_DATA_DIR = INPUT_DIR / "LADQData"
REF_DATA_DIR = INPUT_DIR / "RefData"

# RAP > Outputs >
MASTER_DIR = OUTPUT_DIR / "MasterFiles"
PUB_DIR = OUTPUT_DIR / "PublicationFiles"
LOG_DIR = OUTPUT_DIR / "Logs"

# RAP > Outputs > PublicationFiles >
TAB_DIR = PUB_DIR / "DataTables"
CHART_DIR = PUB_DIR / "Charts"
DG_DIR = PUB_DIR / "DataModel"
REPTAB_DIR = PUB_DIR / "ReportTables"

# Sets the name and path of the current school level data file to be imported
SCHOOL_DATA_FILE = "PHE_School_level.csv"
SCHOOL_DATA_PATH = SCHOOL_DATA_DIR / SCHOOL_DATA_FILE

# Sets the name and path of the current LA DQ data file to be imported
LADQ_DATA_FILE = "DQ_system indicators.xlsx"
LADQ_PATH = LA_DQ_DATA_DIR / LADQ_DATA_FILE

# Sets the name and path of the output file for post deadline validations
PD_VAL_OUTPUT_FILE = "ncmp_post_deadline_validations.xlsx"
PD_VAL_OUTPUT_PATH = MASTER_DIR / PD_VAL_OUTPUT_FILE

# Sets the name and path of the master file for data table outputs
TABLE_TEMPLATE_FILE = "ncmp_datatables.xlsx"
TABLE_TEMPLATE = MASTER_DIR / TABLE_TEMPLATE_FILE

# Sets the name and path of the master file for chart outputs
CHART_TEMPLATE_FILE = "ncmp_charts.xlsx"
CHART_TEMPLATE = MASTER_DIR / CHART_TEMPLATE_FILE

# Sets the name and path of the master file for report table outputs
REPTABLE_TEMPLATE_FILE = "ncmp_reporttables.xlsx"
REPTABLE_TEMPLATE = MASTER_DIR / REPTABLE_TEMPLATE_FILE

# Sets the name and path of input for modelling - default is master charts file
DG_INPUT_FILE = CHART_TEMPLATE_FILE
DG_INPUT_PATH = MASTER_DIR / DG_INPUT_FILE

# Sets the name and path of output for model results
DG_OUTPUT = "ncmp_deprivation_gap_model.xlsx"
DG_OUTPUT_PATH = MASTER_DIR / DG_OUTPUT

# Sets the name and path of standard LA DQ breach reasons available in NCMP system
LADQ_SYS_BREACH_REASONS_FILE = "DQ_Breach_Reasons.csv"
LADQ_SYS_BREACH_REASONS_PATH = REF_DATA_DIR / LADQ_SYS_BREACH_REASONS_FILE

# Sets the SQL database parameters
PROJECT_NAME = "name"
SERVER = "server"
DATABASE = "database"
TABLE = "table"

# Sets the reporting year
YEAR = "2021/22"

# Sets the comparison year for post deadline validations
COMP_YEAR = "2018/19"

# Sets which outputs should be updated as part of the create_publication process
# Set to True to run or False to not
# IMPORTANT write_cell values for time series tables in tables.py are different for each year
# Check they're pointing to the correct year row in master file before running
RUN_PD_VALIDATIONS = False  # Post deadline validations master file
RUN_TABLES = False  # Data tables master file
RUN_CHARTS = False  # Charts master file
RUN_REPTABLES = False  # Report tables master file
RUN_DGMODEL = False  # Deprivation gap model master file

# Sets whether the final publication outputs should be written as part of the pipeline
# Saves final versions of the master files and chart images ready for the web
RUN_PUBLICATION_OUTPUTS = False

# Small LAs (key of dictionary) to combine with larger LAs (value of dictionary)
# Reassigns City of London LA E09000001 to Hackney LA E09000012
# Reassigns Isles of Scilly LA E06000053 to Cornwall LA E06000052
SMALL_LA_COMBINE = {"E09000001": "E09000012", "E06000053": "E06000052"}

# region code columns we need to map to region names
REGION_CODE_COLS = ["RegionCode", "SchoolRegionCode", "PupilRegionCode"]

# region code-name mapping
# TODO - update process so it gets these names from corp ref data, like with LAs
# Should we be outputting region codes and names to master file, in case they change?
REGION_CODE_NAME_MAP = {"E12000001": "North East",
                        "E12000002": "North West",
                        "E12000003": "Yorkshire and the Humber",
                        "E12000004": "East Midlands",
                        "E12000005": "West Midlands",
                        "E12000006": "East of England",
                        "E12000007": "London",
                        "E12000008": "South East",
                        "E12000009": "South West"}

# Define org la code columns we want to map to la names
LA_ORG_CODE_MAP_TO_NAMES = ["OrgCode_ONS",
                            "PupilTier1LocalAuthority",
                            "PupilTier2LocalAuthority",
                            "SchoolTier1LocalAuthority",
                            "SchoolTier2LocalAuthority"]

# Defines the BMI category logic based on BmiPScore
BMI_CATEGORIES = {"Underweight": "BmiPScore <= 0.02",
                  "Healthy weight": "BmiPScore > 0.02 and BmiPScore < 0.85",
                  "Overweight": "BmiPScore >= 0.85 and BmiPScore < 0.95",
                  "Obese": "BmiPScore >= 0.95",
                  "Severely obese": "BmiPScore >= 0.996",
                  "Overweight and obese": "BmiPScore >= 0.85"}

# Defines the default column order for tables
COL_ORDER = ["underweight", "underweight_prevalence",
             "underweight_ci_lower",
             "underweight_ci_upper",
             "healthy_weight", "healthy_weight_prevalence",
             "healthy_weight_ci_lower",
             "healthy_weight_ci_upper",
             "overweight", "overweight_prevalence",
             "overweight_ci_lower",
             "overweight_ci_upper",
             "obese", "obese_prevalence",
             "obese_ci_lower",
             "obese_ci_upper",
             "severely_obese", "severely_obese_prevalence",
             "severely_obese_ci_lower",
             "severely_obese_ci_upper",
             "overweight_obese", "overweight_obese_prevalence",
             "overweight_obese_ci_lower",
             "overweight_obese_ci_upper",
             "total"]

# Defines the columns to use from the LA DQ file and their data types
LA_DQ_COLS_DTYPES = {"DQAugustDateOfMeasurement": "str",
                     "DQBlankNHSNumbers": "str",
                     "DQBlankPostcodes": "str",
                     "DQExtremeBMIs": "str",
                     "DQExtremeHeights": "str",
                     "DQExtremeWeights": "str",
                     "DQHalfHeights": "str",
                     "DQHalfWeights": "str",
                     "DQSameEthnicity-Asian": "str",
                     "DQSameEthnicity-Black": "str",
                     "DQSameEthnicity-Chinese": "str",
                     "DQSameEthnicity-Mixed": "str",
                     "DQSameEthnicity-White": "str",
                     "DQSameEthnicity-Any other ethnic group": "str",
                     "DQSameEthnicity-Not stated": "str",
                     "DQSamePupilSchoolPostcode": "str",
                     "DQSplitPercentageMaleYearR": "str",
                     "DQSplitPercentageMaleYear6": "str",
                     "DQSplitYearRYear6": "str",
                     "DQUnknownEthnicity": "str",
                     "DQWeekendDateOfMeasurement": "str",
                     "DQWholeHeights": "str",
                     "DQWholeWeights": "str",
                     "LocalAuthorityName": "str",
                     "OrgCode": "str",
                     "PercentageBlankNhsNumber": "float",
                     "PercentageBlankPostcode": "float",
                     "PercentageDateOfMeasurementAugust": "float",
                     "PercentageDateOfMeasurementWeekend": "float",
                     "PercentageEthnicGroupAsian": "float",
                     "PercentageEthnicGroupBlack": "float",
                     "PercentageEthnicGroupChinese": "float",
                     "PercentageEthnicGroupMixed": "float",
                     "PercentageEthnicGroupOther": "float",
                     "PercentageEthnicGroupUnknown": "float",
                     "PercentageEthnicGroupWhite": "float",
                     "PercentageExtremeBmi": "float",
                     "PercentageExtremeHeight": "float",
                     "PercentageExtremeWeight": "float",
                     "PercentageHalfNumberHeights": "float",
                     "PercentageHalfNumberWeights": "float",
                     "PercentageParticipationYearR": "float",
                     "PercentageParticipationYear6": "float",
                     "PercentagePostcodeSameAsSchool": "float",
                     "PercentageWholeNumberHeights": "float",
                     "PercentageWholeNumberWeights": "float",
                     "PercentageYearR": "float",
                     "PercentageYear6": "float",
                     "PercentageYearRMale": "float",
                     "PercentageYearRFemale": "float",
                     "PercentageYear6Male": "float",
                     "PercentageYear6Female": "float",
                     "TotalEligibleYearR": "int",
                     "TotalEligibleYear6": "int",
                     "TotalEligibleMeasuredYearR": "int",
                     "TotalEligibleMeasuredYear6": "int",
                     }

# Defines default column order, and contents, for LA DQ outputs
LA_DQ_COL_ORDER = ["OrgCode_ONS_Name",
                   "OrgCode_ONS",
                   "derived_participation_R",
                   "derived_participation_6",
                   "PercentageYearR",
                   "PercentageYear6",
                   "PercentageYearRMale",
                   "PercentageYearRFemale",
                   "PercentageYear6Male",
                   "PercentageYear6Female",
                   "PercentageBlankPostcode",
                   "PercentagePostcodeSameAsSchool",
                   "PercentageEthnicGroupUnknown",
                   "Percentage of records sharing the same ethnicity",
                   "PercentageBlankNhsNumber",
                   "PercentageExtremeHeight",
                   "PercentageExtremeWeight",
                   "PercentageExtremeBmi",
                   "PercentageWholeNumberHeights",
                   "PercentageWholeNumberWeights",
                   "PercentageHalfNumberHeights",
                   "PercentageHalfNumberWeights",
                   "PercentageDateOfMeasurementWeekend",
                   "PercentageDateOfMeasurementAugust",
                   "DQSplitYearRYear6",
                   "DQSplitPercentageMaleYearR",
                   "DQSplitPercentageMaleYear6",
                   "DQBlankPostcodes",
                   "DQSamePupilSchoolPostcode",
                   "DQUnknownEthnicity",
                   "DQSameEthnicity-Asian",
                   "DQSameEthnicity-Black",
                   "DQSameEthnicity-Chinese",
                   "DQSameEthnicity-Mixed",
                   "DQSameEthnicity-White",
                   "DQSameEthnicity-Any other ethnic group",
                   "DQSameEthnicity-Not stated",
                   "DQBlankNHSNumbers",
                   "DQExtremeHeights",
                   "DQExtremeWeights",
                   "DQExtremeBMIs",
                   "DQWholeHeights",
                   "DQWholeWeights",
                   "DQHalfHeights",
                   "DQHalfWeights",
                   "DQWeekendDateOfMeasurement",
                   "DQAugustDateOfMeasurement",
                   ]
