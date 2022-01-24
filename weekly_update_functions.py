import pandas as pd
import os
from extra_functions import to_raw, overwrite_cols


def asl_reporting(path, login, desktop, current_year):
    dnr = pd.read_excel(
        to_raw(os.path.join(path, login, desktop, "Hybrid Journals TA Article List " + str(current_year) +".xlsx")),
        "Do NOT Report")

    asl = pd.DataFrame(dnr.loc[(dnr["Monthly Reporting Status"] != "Do NOT Report")
                               | (dnr["Monthly Reporting Status"].isnull()), "DOI"])

    asl["Monthly Reporting Status"] = "Do NOT Report"

    return asl[["DOI", "Monthly Reporting Status"]]


def production_checks(path, login, desktop, journal_list_file, appr_this_year: pd.DataFrame,
                      appr_previously: pd.DataFrame, rejections: pd.DataFrame, unverified: pd.DataFrame,
                      optout: pd.DataFrame, final_del_cols: list):
    """Creates a Data Frame containing the DOIs to be investigated in Delilah from the updated hybrid articles list on Fridays

    Args:
        path, login, desktop, journal_list_file: imported from general_variables
        appr_this_year: 1st sheet of the hybrid articles file
        appr_previously: 2nd sheet of the hybrid articles file
        rejections: 3rd sheet of the hybrid articles file
        unverified: 4th sheet of the hybrid articles file
        optout: 5th sheet of the hybrid articles file
        final_del_cols: correct order of columns for delilah

    """

    ext_j_list = pd.read_excel(to_raw(os.path.join(path, login, desktop, journal_list_file)), "external")
    my_df = appr_this_year.append(appr_previously).append(rejections).append(unverified).append(optout)
    not_external = my_df.loc[~my_df["Journal ID"].isin(ext_j_list.ID)]

    final_df = not_external.loc[(not_external["Monthly Reporting Status"].isnull())]

    final_df["articledoi"] = final_df["DOI"]

    final_df = final_df.loc[
        ~final_df["Monthly Reporting Status"].str.contains("Article Reported as ", case=False, na=False)]

    return final_df[final_del_cols]


def create_deleted_list(jflux: pd.DataFrame, cancelled: pd.DataFrame,
                        cols_to_overwrite: list, cols_for_del_file: list):
    """merges production checks file with Delilah empty articles, creating a list of DOIs to be investigate in another Data Frame

    jflux: jflux delilah download
    pe_emails: Excel file containing the names of PEs by journal ID
    to_investigate_cols: list of columns needed for the deleted dois file from the jflux download
    to_investigate_cols_orl: list of columns needed for the deleted dois file from the orl
    cols_to_overwrite: list of columns to be overwritten in the merge between orl(z) and deleted list file
    cols_for_del_file: list of columns for the orl 4th sheet

    """

    deleted_dois = jflux[jflux["current_state"].isnull()]

    cancelled["DOI"] = cancelled["DOI"].str.strip()
    cancelled["New DOI"] = cancelled["New DOI"].str.strip()

    cancelled = cancelled.merge(deleted_dois, how="outer", on="DOI")

    overwrite_cols(cols_to_overwrite, cancelled)

    cancelled["Date"] = cancelled["Date"].fillna(pd.Timestamp.today().date())
    cancelled = cancelled

    return cancelled[cols_for_del_file].sort_values(by=["Date"])
