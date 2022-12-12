import statsmodels.api as sm
import pandas as pd
import ncmp_code.utilities.write_excel as write_excel
import ncmp_code.parameters as param

# the sheets in the input file to be analysed
inputs = ["Obese_SchIMDGap_Sex_TSeries_R",
          "Obese_SchIMDGap_Sex_TSeries_6",
          "SevOb_SchIMDGap_Sex_TSeries_R",
          "SevOb_SchIMDGap_Sex_TSeries_6",
          "Obese_PupIMDGap_Sex_TSeries_6",
          "Obese_PupIMDGap_Sex_TSeries_R",
          "SevOb_PupIMDGap_Sex_TSeries_R",
          "SevOb_PupIMDGap_Sex_TSeries_6"]

# processing data to extract required cells


def get_tidy_table(df):
    """ Formats table by adding derived columns and transposing table to have
        deprivation gap values in one column.

    Parameters
    ----------
    df: pandas.DataFrame


    Returns
    -------
    df : pandas.DataFrame
    """

    # remove 2020/21 - weighted - year
    df = df[~df["School year"].isin(["2020/21*"])]

    # add new column that reflects the years to numbers
    df["Year_Num"] = list(range(1, len(df)+1))

    # transform data to the right shape
    df = pd.melt(df,
                 id_vars=["Year_Num"],
                 value_vars=["Deprivation gap", "Deprivation gap.1"],
                 var_name="Gender_Num",
                 value_name="Deprivation Gap")

    # replace value_vars to reflect Gender
    # 1 for boy and 2 for girl
    df = df.replace({"Deprivation gap": 1, "Deprivation gap.1": 2})

    # creates the interaction term
    df["Year_Gender"] = df["Year_Num"] * df["Gender_Num"]

    return df

# running the OLS model on the cleaned data


def run_OLS_model(cleaned_df):
    """ Runs OLS model, prints the inbuilt results summary
    and creates a dataframe of the results.

    Parameters
    ----------
    cleaned_df: pandas.DataFrame


    Returns
    -------
    df : pandas.DataFrame
    """

    # create a table for the independent varaibles
    X = cleaned_df[["Year_Gender", "Year_Num", "Gender_Num"]]

    # adds a constant
    X = sm.add_constant(X, prepend=False)

    # creates a table for the dependent variable
    y = cleaned_df[["Deprivation Gap"]]

    # run the OLS model
    results = (sm.OLS(y, X)).fit()

    '''take the result of an statsmodel results table and transforms it into a dataframe'''
    pvals = results.pvalues
    coeff = results.params
    conf_lower = results.conf_int()[0]
    conf_higher = results.conf_int()[1]

    results_df = pd.DataFrame({"pvals": pvals,
                               "coeff": coeff,
                               "conf_lower": conf_lower,
                               "conf_higher": conf_higher
                               })

    # Reordering...
    results_df = results_df[["coeff", "pvals", "conf_lower", "conf_higher"]]
    return results_df

