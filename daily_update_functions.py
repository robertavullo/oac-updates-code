import pandas as pd
import numpy as np
import os
from extra_functions import find_key, overwrite_cols, to_raw
from formatting_functions import format_date
from general_variables import *


def create_optouts(path, login, desktop, institution_list_compact, journal_list_file, palm_rep: pd.DataFrame,
                   optout_df_cols: list, palm_optout_cols: list, current_year, dict: dict) -> pd.DataFrame:
    """Filter PALM reconciliation report, change the names of the relevant columns and merge with the previous day's opt-outs + add missing information

    Args:
        path, login, desktop, current_year, institution_list_compact, journal_list_file: imported from general_variables
        palm_rep: reconciliation report downloaded from the palm dashboard
        palm_optout_cols: columns from the palm dashboard report that need to be merged (including DOI)
        optout_df_cols: correct order of columns in the 2nd sheet of the ORL
        palm_col_dict: column name changes for the palm dashboard report

    Returns:
        Pandas Data Frame with updated opt-out list
    """
    agreement_dict_bpid = \
        pd.read_excel(to_raw(os.path.join(path, login, desktop, institution_list_compact)),
                      "Institution List").set_index(
            "BPID")["Agreement"].to_dict()
    agreement_dict_name = \
        pd.read_excel(to_raw(os.path.join(path, login, desktop, institution_list_compact)),
                      "Institution List").set_index(
            "Institution name")["Agreement"].to_dict()
    journal_dict = \
        pd.read_excel(to_raw(os.path.join(path, login, desktop, journal_list_file)), "full").set_index("Journal ID")[
            "Journal Title"].to_dict()
    journal_license = \
        pd.read_excel(to_raw(os.path.join(path, login, desktop, journal_list_file)), "full").set_index("Journal ID")[
            "License"].to_dict()

    palm_rep = palm_rep.loc[(palm_rep["Author opt out"] == "y")
                            & (palm_rep["Date Rejected"].isnull())]

    palm_rep["Date Sent to AAS"] = pd.to_datetime(palm_rep["Date Sent to AAS"])
    optouts_df = pd.read_excel(
        to_raw(os.path.join(path, login, desktop, "Hybrid Journals TA Article List " + str(current_year) + ".xlsx")),
        "Opt-Outs")

    optouts_df = optouts_df[optout_df_cols].merge(palm_rep[palm_optout_cols], how="outer", on="DOI")

    optouts_df = overwrite_cols(palm_optout_cols[1:], optouts_df)

    optouts_df.loc[(optouts_df["Agreement"].isnull()), "Agreement"] = optouts_df[
        "Initial Institution (sent to) BPID"].apply(lambda x: agreement_dict_bpid.get(x, ""))
    optouts_df.loc[(optouts_df["Agreement"].isnull()), "Agreement"] = optouts_df[
        "Author Affiliation"].apply(lambda x: agreement_dict_name.get(x, ""))
    optouts_df["Author Email_suffix"] = optouts_df["Author Email"].str.slice(-2)
    optouts_df.loc[(optouts_df["Agreement"].isnull()) | (optouts_df["Agreement"] == ""), "Agreement"] = optouts_df[
        "Author Email_suffix"].apply(lambda x: dict.get(x, "Not found!"))

    optouts_df.loc[optouts_df["Journal ID"].isnull(), "Journal ID"] = \
        optouts_df["DOI"].str.split("/s", expand=True)[1].str.split("-", expand=True)[0].astype("float")
    optouts_df.loc[(optouts_df["Journal Title"].isnull())
                   & (~optouts_df["Journal ID"].isnull()),
                   "Journal Title"] = optouts_df["Journal ID"].apply(lambda x: journal_dict.get(x, ""))
    optouts_df.loc[(optouts_df["Journal ID"].isnull())
                   & (~optouts_df["Journal Title"].isnull()),
                   "Journal ID"] = optouts_df["Journal Title"].apply(find_key, in_dict=journal_dict)
    optouts_df["Journal License"] = optouts_df["Journal ID"].apply(lambda x: journal_license.get(x, ""))

    optouts_df = optouts_df.drop_duplicates(subset=["DOI"], keep="last")
    optouts_df["Eligibility Completed"] = pd.to_datetime(optouts_df["Eligibility Completed"],
                                                         format="%Y-%m-%d %H:%M:%S").dt.tz_localize(None)

    return optouts_df[optout_df_cols].sort_values(by=["Eligibility Completed"])


def create_oa_req_list(path, login, desktop, institution_list_compact, journal_list_file, palm_rep: pd.DataFrame,
                       aas_dashboard_cols: list, complete_oa_req_cols: list, palm_req_cols: list, palm_dates: list,
                       fr_3days: list, fr_5days: list, aas_rep_dict: dict, palm_col_dict: dict,
                       current_year, dict: dict) -> pd.DataFrame:
    """Filter PALM reconciliation report, change the names of the relevant columns and merge with the previous day's non-opt-out list

    Args:
        path, login, desktop, current_year, institution_list_compact, journal_list_file: imported from general_variables
        palm_rep: reconciliation report downloaded from the palm dashboard
        aas_dashboard_cols: columns from the AAS dashboard used in the update
        complete_oa_req_cols: correct order of columns in the 1st 4 sheets of the HAL (?) Excel file
        palm_req_cols: columns from the palm report needed for the update (non-opt-out)
        palm_dates: columns containing datetimes in the palm dashboard report
        fr_3days: list of countries which get a final reminder after 3 business days
        fr_5days: list of countries which get a final reminder after 5 business days
        palm_col_dict: column name changes for the palm dashboard report
        aas_rep_dict: dictionary to change names of the column from the AAS dashboard report

    Returns:
        Pandas Data Frame with updated list of hybrid articles (not filtered)
    """
    agreement_dict_bpid = \
        pd.read_excel(to_raw(os.path.join(path, login, desktop, institution_list_compact)),
                      "Institution List").set_index(
            "BPID")["Agreement"].to_dict()
    agreement_dict_name = \
        pd.read_excel(to_raw(os.path.join(path, login, desktop, institution_list_compact)),
                      "Institution List").set_index(
            "Institution name")["Agreement"].to_dict()
    approver_dict_hybrid = \
        pd.read_excel(to_raw(os.path.join(path, login, desktop, institution_list_compact)),
                      "Institution List").set_index(
            "BPID")["Approval Manager Emails"].to_dict()
    am_email_dict_hybrid = \
        pd.read_excel(to_raw(os.path.join(path, login, desktop, institution_list_compact)),
                      "Institution List").set_index(
            "Approval Manager Emails")["Agreement"].to_dict()
    licence_id_dict = \
        pd.read_excel(to_raw(os.path.join(path, login, desktop, institution_list_compact)),
                      "Institution List").set_index(
            "Agreement")["Licence ID"].to_dict()
    spain_licence_id_dict = \
        pd.read_excel(to_raw(os.path.join(path, login, desktop, institution_list_compact)),
                      "Institution List").set_index(
            "Institution name")["Licence ID"].to_dict()
    agreement_deal_list = [key for key in list(agreement_dict_bpid.keys()) if "DEAL" in agreement_dict_bpid[key]]
    agreement_bibsam_list = [key for key in list(agreement_dict_bpid.keys()) if "BIBSAM" in agreement_dict_bpid[key]]

    bpid_dict = \
        pd.read_excel(to_raw(os.path.join(path, login, desktop, institution_list_compact)),
                      "Institution List").set_index(
            "Institution name")["BPID"].to_dict()
    journal_dict = \
        pd.read_excel(to_raw(os.path.join(path, login, desktop, journal_list_file)), "full").set_index("Journal ID")[
            "Journal Title"].to_dict()
    natureOA_j = pd.read_excel(to_raw(os.path.join(path, login, desktop, journal_list_file)), "DEAL Nature OA")[
        "Journal ID"].tolist()
    journal_license = \
        pd.read_excel(to_raw(os.path.join(path, login, desktop, journal_list_file)), "full").set_index("Journal ID")[
            "License"].to_dict()

    palm_rep = palm_rep.rename(columns=palm_col_dict)
    palm_rep = palm_rep.loc[(palm_rep["Author opt out"] == "n")
                            & (palm_rep["Publishing Model Decision"] != "SubscriptionChosen")
                            & (~palm_rep["Date Sent to AAS"].isnull())]

    oa_req_list_appr = pd.read_excel(
        to_raw(os.path.join(path, login, desktop, "Hybrid Journals TA Article List " + str(current_year) + ".xlsx")),
        "Approved")
    oa_req_list_appr_prev = pd.read_excel(
        to_raw(os.path.join(path, login, desktop, "Hybrid Journals TA Article List " + str(current_year) + ".xlsx")),
        "Approved (prev. years)")
    oa_req_list_rej = pd.read_excel(
        to_raw(
            os.path.join(
                path, login, desktop, "Hybrid Journals TA Article List " + str(current_year) + ".xlsx")), "Rejected")
    oa_req_list_unver = pd.read_excel(
        to_raw(os.path.join(
            path, login, desktop, "Hybrid Journals TA Article List " + str(current_year) + ".xlsx")), "In Verification")

    oa_req_list = oa_req_list_appr.append(oa_req_list_appr_prev).append(oa_req_list_rej).append(oa_req_list_unver)

    aas_open = pd.read_excel(to_raw(os.path.join(path, login, desktop, "AAS_open.xlsx")))
    aas_approved = pd.read_excel(to_raw(os.path.join(path, login, desktop, "AAS_approved.xlsx")))
    aas_rejected = pd.read_excel(to_raw(os.path.join(path, login, desktop, "AAS_rejected.xlsx")))
    aas_report = aas_open.append(aas_approved).append(aas_rejected)
    aas_report = aas_report.rename(columns=aas_rep_dict)
    aas_report[["Author First Name", "Author Last Name"]] = aas_report["corresponding author name"].str.split(
        " ", n=1, expand=True)
    aas_report["Initial Institution (sent to) BPID"] = aas_report["membership institute"].apply(
        lambda x: bpid_dict.get(x, np.nan))
    aas_report = aas_report.loc[(aas_report["journal publishing model"] == "Hybrid") & (~aas_report["DOI"].isnull())]

    for col in palm_dates:
        palm_rep[col] = pd.to_datetime(palm_rep[col], format="%Y-%m-%d %H:%M:%S")

    for col in palm_dates[:-1]:
        aas_report[col] = pd.to_datetime(aas_report[col], format="%Y-%m-%d %H:%M:%S")

    palm_rep = palm_rep.sort_values(by=["Date Sent to AAS"]).drop_duplicates(subset="DOI", keep="last")
    aas_report = aas_report.sort_values(by=["Date Sent to AAS"]).drop_duplicates(subset="DOI", keep="last")

    oa_req_list = oa_req_list[complete_oa_req_cols].merge(palm_rep[palm_req_cols], how="outer", on="DOI")

    oa_req_list = overwrite_cols(palm_req_cols[1:], oa_req_list)

    oa_req_list = oa_req_list[complete_oa_req_cols].merge(aas_report[aas_dashboard_cols], how="outer", on="DOI")

    oa_req_list = overwrite_cols(aas_dashboard_cols[1:], oa_req_list)

    oa_req_list.loc[(oa_req_list["Journal Title"].isnull()) & (~oa_req_list["Journal ID"].isnull()),
                    "Journal Title"] = oa_req_list["Journal ID"].apply(lambda x: journal_dict.get(x, ""))
    oa_req_list.loc[oa_req_list["Journal ID"].isnull(), "Journal ID"] = \
        oa_req_list["DOI"].str.split("/s", expand=True)[1].str.split("-", expand=True)[0].astype("float")
    oa_req_list.loc[(oa_req_list["Journal ID"].isnull()) & (~oa_req_list["Journal Title"].isnull()),
                    "Journal ID"] = oa_req_list["Journal Title"].apply(find_key, in_dict=journal_dict)
    oa_req_list.loc[(oa_req_list["Agreement"].isnull()), "Agreement"] = oa_req_list[
        "Initial Institution (sent to) BPID"].apply(lambda x: agreement_dict_bpid.get(x, ""))
    oa_req_list.loc[(oa_req_list["Agreement"].isnull()), "Agreement"] = oa_req_list[
        "Approver Email"].apply(lambda x: am_email_dict_hybrid.get(x, ""))
    oa_req_list.loc[(oa_req_list["Agreement"].isnull()), "Agreement"] = oa_req_list[
        "Author Affiliation"].apply(lambda x: agreement_dict_name.get(x, ""))

    oa_req_list["Approver Email_suffix"] = oa_req_list["Approver Email"].str.slice(-2)
    oa_req_list.loc[(oa_req_list["Agreement"].isnull()) | (oa_req_list["Agreement"] == ""), "Agreement"] = oa_req_list[
        "Approver Email_suffix"].apply(lambda x: dict.get(x, "Not found!"))

    oa_req_list["Journal License"] = oa_req_list["Journal ID"].apply(lambda x: journal_license.get(x, ""))

    for col in palm_dates[:-1]:
        oa_req_list[col] = pd.to_datetime(oa_req_list[col], format="%Y-%m-%d %H:%M:%S")

    oa_req_list["Final Reminder to be Sent"] = oa_req_list["Final Reminder to be Sent"].fillna("N/A")
    oa_req_list["Final Reminder Sent?"] = oa_req_list["Final Reminder Sent?"].fillna("N/A")

    create_final_reminders(oa_req_list, fr_3days, fr_5days)
    oa_req_list["Final Reminder to be Sent"] = oa_req_list["Final Reminder to be Sent"].fillna("N/A")
    oa_req_list["Final Reminder Sent?"] = oa_req_list["Final Reminder Sent?"].fillna("N/A")

    sn_support = oa_req_list.loc[(oa_req_list["Approver Email"].str.contains("@springernature", case=False, na=False))
                                 | (oa_req_list["Approver Email"].str.contains("@springer", case=False, na=False))
                                 | (oa_req_list["Approver Email"].str.contains("@nature", case=False, na=False))
                                 | (oa_req_list["Approver Email"].str.contains("@biomedcentral", case=False, na=False))]
    oacs_approvers = list(sn_support.DOI)
    oa_req_list.loc[oa_req_list["DOI"].isin(oacs_approvers), "Approver Email"] = "Springer Nature Support"

    deal_nat_oa_list = ["2021", "2022", "2023", "2024"]
    bibsam_nat_oa_list = ["2022", "2023", "2024"]

    oa_req_list = oa_req_list.dropna(subset=["Article Title"]).drop_duplicates(subset=["DOI"], keep="last")
    oa_req_list["Date Sent to AAS_year"] = oa_req_list["Date Sent to AAS"].dt.year
    oa_req_list.loc[(oa_req_list["Journal ID"].isin(natureOA_j))
                    & (oa_req_list["Initial Institution (sent to) BPID"].isin(agreement_deal_list))
                    & (oa_req_list["Date Sent to AAS_year"].isin(deal_nat_oa_list)), "Agreement"] = "DEAL NATURE OA"
    oa_req_list.loc[(~oa_req_list["Journal ID"].isin(natureOA_j))
                    & (oa_req_list["Initial Institution (sent to) BPID"].isin(
                       agreement_deal_list)), "Agreement"] = "DEAL"
    oa_req_list.loc[(oa_req_list["Journal ID"].isin(natureOA_j))
                    & (oa_req_list["Initial Institution (sent to) BPID"].isin(agreement_bibsam_list))
                    & (oa_req_list["Date Sent to AAS_year"].isin(bibsam_nat_oa_list)), "Agreement"] = "BIBSAM NATURE OA"
    oa_req_list.loc[(~oa_req_list["Journal ID"].isin(natureOA_j))
                    & (oa_req_list["Initial Institution (sent to) BPID"].isin(
                       agreement_bibsam_list)), "Agreement"] = "BIBSAM"

    oa_req_list.loc[(oa_req_list["Approver Email"].isnull()), "Approver Email"] = oa_req_list[
        "Initial Institution (sent to) BPID"].apply(
        lambda x: approver_dict_hybrid.get(x, "No AM found, check master list"))

    licence_id_list = ["BIBSAM", "BIBSAM NATURE OA", "CRUI-CARE", "CSAL", "DEAL", "DEAL NATURE OA", "FINELIB", "HAS",
                       "ICM", "IREL", "JISC", "KEMÖ", "MANI", "QNL", "UKB", "SIKT", "CNR", "CAUL", "FSLN", "CDL"]
    oa_req_list.loc[
        ((oa_req_list["Agreement"].isin(licence_id_list)) & (oa_req_list["Licence ID"].isnull())), "Licence ID"] = \
        oa_req_list["Agreement"].apply(lambda x: licence_id_dict.get(x, ""))
    oa_req_list.loc[
        ((oa_req_list["Agreement"] == "CRUE-CSIC") & (oa_req_list["Licence ID"].isnull())), "Licence ID"] = \
        oa_req_list["Initial Institution (sent to) BPID"].apply(lambda x: spain_licence_id_dict.get(x, ""))

    oa_req_list = oa_req_list.fillna(value={"OASiS status": "N/A", "OASiS status date": "N/A", "Author Share": "N/A",
                                            "Institutional Share": "N/A", "Reason for Full Coverage": "N/A"})

    return oa_req_list[complete_oa_req_cols].sort_values(by=["Date Sent to AAS"])


def find_correct_verification(current_year, oa_req_list: pd.DataFrame, complete_oa_req_cols):
    """filter DOIs with more than one verification status and define which verification status is the correct one

    Args:
        current_year: imported from general_variables
        oa_req_list: Data Frame of all the hybrid articles (HAL, AAS, PALM), unfiltered at first as returned from the function create_oa_req_list()
        complete_oa_req_cols: correct order of columns in the 1st 4 sheets of the HAL (?) Excel file

    Returns:
        Pandas Data Frame with updated hybrid articles request, with an extra column for which sheet they should go into"""

    my_df = oa_req_list.loc[(~oa_req_list["Date Approved"].isnull()) & (~oa_req_list["Date Rejected"].isnull())]
    appr_dois = my_df.loc[(my_df["Date Approved"] > my_df["Date Rejected"]), "DOI"].tolist()
    rej_dois = my_df.loc[(my_df["Date Approved"] < my_df["Date Rejected"]), "DOI"].tolist()
    oa_req_list.loc[(oa_req_list["DOI"].isin(appr_dois)), "Date Rejected"] = np.nan
    oa_req_list.loc[(oa_req_list["DOI"].isin(rej_dois)), "Date Approved"] = np.nan

    oa_req_list.loc[(oa_req_list["OASiS Status"] == "CANCELLED")
                    | (oa_req_list["Monthly Reporting Status"].str.contains("do not report", case=False, na=False))
                    | (oa_req_list["Comments"].str.contains("deleted", case=False, na=False))
                    | (oa_req_list["Comments"].str.contains("do not report", case=False, na=False)),
                    "sheet in file"] = "6"
    oa_req_list.loc[(oa_req_list["OASiS Status"] != "CANCELLED") & (oa_req_list["Date Approved"].isnull()) & (
        oa_req_list["Date Rejected"].isnull()),
                    "sheet in file"] = "4"
    oa_req_list.loc[(oa_req_list["OASiS Status"] != "CANCELLED") & (~oa_req_list["Date Rejected"].isnull()),
                    "sheet in file"] = "3"
    oa_req_list.loc[(oa_req_list["OASiS Status"] != "CANCELLED") & (~oa_req_list["Date Approved"].isnull()) & (
            pd.DatetimeIndex(oa_req_list["Date Approved"]).year != current_year),
                    "sheet in file"] = "2"
    oa_req_list.loc[(oa_req_list["OASiS Status"] != "CANCELLED") & (~oa_req_list["Date Approved"].isnull()) & (
            pd.DatetimeIndex(oa_req_list["Date Approved"]).year == current_year),
                    "sheet in file"] = "1"

    complete_oa_req_cols.append("sheet in file")

    return oa_req_list[complete_oa_req_cols].sort_values(by=["Date Sent to AAS"])


def create_final_reminders(df: pd.DataFrame, fr_3days: list, fr_5days: list) -> pd.DataFrame:
    """Filters the updated ORL sheet to onlly select the articles which need a final reminder and adds the day the reminder needs to be sent to the ORL

    Args:
        df: the updated 1st sheet of the ORL
        fr_3days: list of countries which get a final reminder after 3 business days
        fr_5days: list of countries which get a final reminder after 5 business days

    Returns:
        updated ORL as a Pandas Data Frame
    """
    unverified = df.loc[(df["Date Approved"].isnull())
                        & (df["Date Rejected"].isnull())
                        & (df["Date Forwarded"].isnull())
                        & (~df["Date Sent to AAS"].isnull())
                        & (df["Final Reminder to be Sent"] == "N/A")
                        & (~df["Comments"].str.contains("deleted", case=False, na=False))
                        & (~df["Comments"].str.contains("flagged", case=False, na=False))
                        & (~df["DOI"].str.contains("10.1057", na=False))]

    countries_3b = unverified.loc[(unverified["Agreement"].isin(fr_3days))]
    countries_5b = unverified.loc[(unverified["Agreement"].isin(fr_5days))]

    countries_3b["day to email"] = pd.to_datetime(countries_3b["Date Sent to AAS"]).dt.date + pd.offsets.BDay(3)
    countries_5b["day to email"] = pd.to_datetime(countries_5b["Date Sent to AAS"]).dt.date + pd.offsets.BDay(5)

    df2 = countries_3b.append(countries_5b)

    unverified2 = df2[(pd.to_datetime(df2["day to email"]).dt.date <= pd.Timestamp.today().date())]

    df.loc[(df["DOI"].isin(list(unverified2["DOI"])))
           & (~df["DOI"].isnull()), "Final Reminder to be Sent"] = pd.Timestamp.today().date()

    return df


def create_foa(path, login, desktop, institution_list_fullyOA, journal_list_file, filename: pd.DataFrame,
               aas_report: pd.DataFrame, recon_col: list, aas_rep_cols: list, sheet_cols: list,
               recon_update_cols: list, aas_update_cols: list, date_cols: list, fr_3days: list, fr_5days: list,
               standalone: list, recrep_column_dict: dict, AAS_rep_dict_foal: dict,
               articletype_dict, current_year) -> pd.DataFrame:
    """Filter s150 csv file, shuffle columns and merge with the previous day's FOAL

    Args:
        path, login, desktop, current_year, institution_list_fullyOA, journal_list_file:imported from general_variables
        filename: reconciliation report download as csv file (full)
        aas_report: report downloaded from the super user AAS dashboard
        recon_col: list of columns in the reconciliation report csv that need to be copied (includes UAI)
        aas_rep_cols: list of columns in the AAS report csv that need to be copied
        sheet_cols: correct order of columns in the 1st sheet of the FOAL
        recon_update_cols: columns in the FOAL that need to be replaced with the info from the recon report csv (no UAI)
        aas_update_cols: columns in the FOAL that need to be replaced with the info from the AAS report (no DOI)
        date_cols: columns containing datetimes in the FOAL
        fr_3days: list of countries which get a final reminder after 3 business days
        fr_5days: list of countries which get a final reminder after 5 business days
        standalone: list of standalone agreements acronyms
        inst_country_dict: dictionary to add country code using BPIDs
        sel_inst_dict: dictionary to add Selected Institution
        AAS_rep_dict_foal: column name changes for the AAS report
        recrep_column_dict: dict of column names for the reconciliation report
        articletype_dict: dictionary of article types to match the AAS report to the FOAL

    Returns:
        Pandas Data Frame with updated ORL sheet
    """
    country = \
        pd.read_excel(to_raw(os.path.join(path, login, desktop, institution_list_fullyOA)), "AM emails").set_index(
            "Approval Manager Email")["Agreement"].to_dict()
    inst_country_dict = \
        pd.read_excel(to_raw(os.path.join(path, login, desktop, institution_list_fullyOA)),
                      "Institution List").set_index(
            "BPID")["Agreement"].to_dict()
    licence_id_dict = \
        pd.read_excel(to_raw(os.path.join(path, login, desktop, institution_list_fullyOA)),
                      "Institution List").set_index(
            "Agreement")["Licence ID"].to_dict()
    sel_inst_dict = \
        pd.read_excel(to_raw(os.path.join(path, login, desktop, institution_list_fullyOA)),
                      "Institution List").set_index(
            "BPID")["Institution name"].to_dict()
    inst_dict_approver_inst = \
        pd.read_excel(to_raw(os.path.join(path, login, desktop, institution_list_fullyOA)),
                      "Institution List").set_index(
            "Institution name")["Agreement"].to_dict()
    fully_spon_list = pd.read_excel(
        to_raw(os.path.join(path, login, desktop, journal_list_file)), "fully sponsored")["Journal Title"].to_list()

    filename = filename.rename(columns=recrep_column_dict)
    aas_report = aas_report.rename(columns=AAS_rep_dict_foal)
    aas_report = aas_report.loc[(aas_report["journal publishing model"] == "Fully_Open_Access")
                                & (~aas_report["DOI"].isnull())]

    aas_report["Article Type"] = aas_report["Article Type"].replace(articletype_dict)
    aas_report[["Author First Name", "Author Last Name"]] = aas_report["Corresponding Author"].str.split(
        " ", 1, expand=True)

    recon = filename.dropna(subset=["Date Sent to AAS"])
    recon["Selected Affiliation"] = recon["Selected Affiliation BPID"].apply(lambda x: sel_inst_dict.get(x, np.nan))
    df = pd.read_excel(to_raw(os.path.join(
        path, login, desktop, "FOA Journals TA Article List " + str(current_year) + ".xlsx")), "Articles in AAS")

    foal_today = df[sheet_cols].merge(recon[recon_col], how="outer", on="Unique Article ID")

    foal_today = overwrite_cols(recon_update_cols, foal_today)

    foal_today = foal_today[sheet_cols].merge(aas_report[aas_rep_cols], how="outer", on="DOI")

    foal_today = overwrite_cols(aas_update_cols, foal_today)

    my_df = foal_today.loc[(~foal_today["Date Approved"].isnull()) & (~foal_today["Date Rejected"].isnull())]
    format_date(my_df, ["Date Approved", "Date Rejected"])
    appr_dois = my_df.loc[(my_df["Date Approved"] > my_df["Date Rejected"]), "DOI"].tolist()
    rej_dois = my_df.loc[(my_df["Date Approved"] < my_df["Date Rejected"]), "DOI"].tolist()
    foal_today.loc[(foal_today["DOI"].isin(appr_dois)), "Date Rejected"] = np.nan
    foal_today.loc[(foal_today["DOI"].isin(rej_dois)), "Date Approved"] = np.nan

    foal_today.loc[(foal_today["Agreement"].isnull())
                   | (foal_today["Agreement"].str.contains("Not found!", case=False, na=False)),
                   "Agreement"] = foal_today["Approver Email"].apply(lambda x: country.get(x, np.nan))
    foal_today.loc[(foal_today["Agreement"].isnull())
                   | (foal_today["Agreement"].str.contains("Not found!", case=False, na=False)),
                   "Agreement"] = foal_today["Approving Institution"].apply(
        lambda x: inst_dict_approver_inst.get(x, "Not found!"))
    foal_today.loc[(foal_today["Agreement"].isnull())
                   | (foal_today["Agreement"].str.contains("Not found!", case=False, na=False)),
                   "Agreement"] = foal_today["Selected Affiliation BPID"].apply(
        lambda x: inst_country_dict.get(x, "Not found!"))

    format_date(foal_today, date_cols)

    foal_today["Final Reminder to be Sent"] = foal_today["Final Reminder to be Sent"].fillna("N/A")
    foal_today["Final Reminder Sent?"] = foal_today["Final Reminder Sent?"].fillna("N/A")

    format_date(foal_today, date_cols)

    create_final_reminders(foal_today, fr_3days, fr_5days)

    sn_support = foal_today.loc[foal_today["Approver Email"].str.contains("@springernature", case=False, na=False)]
    oacs_approvers = list(sn_support.DOI)
    foal_today.loc[foal_today["DOI"].isin(oacs_approvers), "Approver Email"] = "Springer Nature Support"

    foal_today = foal_today.fillna(
        value={"Author Share": "N/A", "Institutional Share": "N/A", "Reason for Full Coverage": "N/A"})
    foal_today.loc[foal_today["OC Membership State"].isnull(), "OC Membership State"] = "MEMBERSHIP_FOUND"
    foal_today.loc[foal_today["Waiver Type"].isnull(), "Waiver Type"] = "WAIVER_INSTITUTION"

    licence_id_list = ["BIBSAM", "CAS", "CDL", "DEAL", "HAS", "NIH", "QNL", "SU", "UDR", "UMS", "VIU", "VMU", "WIS"]
    foal_today.loc[
        ((foal_today["Agreement"].isin(licence_id_list)) & (foal_today["Licence ID"].isnull())), "Licence ID"] = \
        foal_today["Agreement"].apply(lambda x: licence_id_dict.get(x, ""))

    foal_today = foal_today.fillna(value={"OASiS status": "N/A", "OASiS status date": "N/A", "Author Share": "N/A",
                                          "Institutional Share": "N/A", "Reason for Full Coverage": "N/A"})
    foal_today.loc[foal_today["Institutional Share"] != "N/A", "Author Share"] == 0.0

    foal_today.loc[(foal_today["Agreement"].isin([standalone]) & foal_today[
        "Agreement Type"].isnull()), "Agreement Type"] = "Standalone"
    foal_today.loc[(~foal_today["Agreement"].isin([standalone]) & foal_today[
        "Agreement Type"].isnull()), "Agreement Type"] = "Fully OA"

    foal_today.loc[(foal_today["Journal Title"].isin(fully_spon_list)), "Agreement"] = "Fully sponsored"

    foal_today_1 = foal_today[foal_today["DOI"].isnull()]
    foal_today_2 = foal_today[~foal_today["DOI"].isnull()].drop_duplicates(subset=["DOI"], keep="last")
    foal_today = foal_today_1.append(foal_today_2)

    foal_today_1 = foal_today[foal_today["Unique Article ID"].isnull()]
    foal_today_2 = foal_today[~foal_today["Unique Article ID"].isnull()].drop_duplicates(
        subset=["Unique Article ID"], keep="last")
    foal_today = foal_today_1.append(foal_today_2)

    foal_today_1 = foal_today[foal_today["PRS Article ID"].isnull()]
    foal_today_2 = foal_today[~foal_today["PRS Article ID"].isnull()].drop_duplicates(
        subset=["PRS Article ID"], keep="last")
    foal_today = foal_today_1.append(foal_today_2)

    for col in date_cols:
        foal_today[col] = pd.to_datetime(foal_today[col])

    return foal_today[sheet_cols].sort_values(by=["Date Sent to AAS"])
