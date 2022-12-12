Warning - this repository is a snapshot of a repository internal to NHS Digital. This means that links to videos and some URLs may not work.***

***Repository owner: Analytical Services: Population Health, Clinical Audit and Specialist Care***

***Email: ncmp@nhs.net***

***To contact us raise an issue on Github or via email and we will respond promptly.***

# Getting Started

## Clone repository
To clone respositary, please see our [community of practice page](https://github.com/NHSDigital/rap-community-of-practice/blob/main/development-approach/02_using-git-collaboratively.md).

## Set up environment
There are two options to set up the python enviroment:
1. Pip using `requirements.txt`.
2. Conda using `environment.yml`.

Users would need to delete as appropriate which set they do not need. For details, please see our [virtual environments in the community of practice page](https://github.com/NHSDigital/rap-community-of-practice/blob/main/python/virtual-environments.md).

## Initial package set up

Run the following command in Terminal or VScode or Anaconda Prompt to set up the package
```
pip install --user -r requirements.txt
```

or if using conda environments:
```
conda env create -f environment.yml
```
# National Child Measurement Programme (NCMP) background

The National Child Measurement Programme (NCMP) is a key element of the
Government’s approach to tackling child obesity by annually measuring the
height and weight of children in Reception (aged 4–5 years) and 
Year 6 (aged 10–11 years) in mainstream state-maintained schools in England. 

Local Authorities (LAs) in England measure children during the school year
with the programme running between September and August each year to coincide
with the academic year.

The dataset is based on pupils and school level extract from the NCMP system after the submission period.

This project produces the required outputs - data tables and charts - for the annual NCMP publication.

# Directory structure:
```
national-child-measurement-programme-rap
│   environment.yml                         - Used to install the conda environment
│   LICENSE
│   README.md
│   requirements.txt                        - Used to install the python dependencies
│   setup.py                                - Used to install this pipeline as a package, not yet in use
│
├───ncmp_code                               - This is the main code directory for this project
│   │   create_publication.py               - This script runs the entire publication
│   │   parameters.py                       - Contains parameters that define the how the publication will run
│   │   __init__.py                           
│   │
│   ├───model                               - This folder contains the Python modelling code for OLS regression
|   |   |   run_model.py                    - Imports and cleans the data, and runs the OLS model using the statsmodel.api function
|   | 
│   ├───sql_code                            - This folder contains all the SQL queries used in the import data stage
│   │   │   query_asset.sql                 - Defines the SQL query to import pupil level data and only required columns
│   │   │   query_asset_full.sql            - Defines the SQL query to import pupil level data and all columns
│   │   │   query_laref.sql                 - Defines the SQL query to import LA reference data from corporate reference data
│   │
│   ├───utilities                           - This module contains all the main modules used to create the publication
│   │   │   charts.py                       - Defines the functions needed to create and export charts to Excel
│   │   │   data_connections.py             - Defines the df_from_sql function, used when importing SQL data
│   │   │   helpers.py                      - Contains generalised functions used within the project, that can be used across multiple projects
│   │   │   import_data.py                  - Contains functions for reading in the required data from .csv files and SQL tables
│   │   │   logger_config.py                - The configuration functions for the publication logger
│   │   │   post_deadline_validations.py    - Contains the functions used to perform the various post deadline validation checks on the data
│   │   │   processing_steps.py             - Defines the main functions used to manipulate data and produce outputs
│   │   │   publication.py                  - Contains functions used to create publication ready outputs and save in relevant folders
│   │   │   tables.py                       - Contains every output table defined as a function
│   │   │   write_excel.py                  - Contains functions needed for writing output to Excel
│   │   │   __init__.py
|
├───tests                               
│   │   __init__.py 
│   ├───unittests                           - Unit tests for all Python functions/modules
│   │   │   test_helpers.py
│   │   │   test_processing_steps.py
│   │   │   __init__.py
```


# Running the publication process

There are two main files that users running the process will need to interact with:

- [parameters.py](ncmp_code/parameters.py)

- [create_publication.py](ncmp_code/create_publication.py)

The file parameters.py contains all of the things that we expect to change from one publication
to the next. Indeed, if the methodology has not changed, then this should be the only file you need
to modify. However, at the moment for updating any timeseries table, it is expected that the 'write_cell' parameter in the function 
needs to be updated to avoid overriding previous years data. These changes are to be made in the [tables.py](ncmp_code/utilities/tables.py) for table_academicyear_R, table_academicyear_6, table_imd_gender_obese_school_R, table_imd_gender_obese_school_6, table_imd_gender_severely_obese_school_R, table_imd_gender_severely_obese_school_6, table_imd_gender_obese_pupil_R, table_imd_gender_obese_pupil_6, table_imd_gender_severely_obese_pupil_R, table_imd_gender_severely_obese_pupil_6.

This file specifies the input and output filepaths, the year filters,
LA exclusions/updates, etc., and also allows the user to control which
parts of the report (model, post validation, tables and(or) charts) are generated when the process is run. 

The publication process is run using the top-level script, create_publication.py. 
This script imports and runs all the required functions and from the sub-modules.


# Link to the publication
https://digital.nhs.uk/data-and-information/publications/statistical/national-child-measurement-programme/2021-22-school-year

# License
The NCMP publication codebase is released under the MIT License.

Copyright © 2022, Health and Social Care Information Centre. The Health and Social Care Information Centre is a non-departmental body created by statute, also known as NHS Digital.
________________________________________
You may re-use this document/publication (not including logos) free of charge in any format or medium, under the terms of the Open Government Licence v3.0.
Information Policy Team, The National Archives, Kew, Richmond, Surrey, TW9 4DU;
email: psi@nationalarchives.gsi.gov.uk
