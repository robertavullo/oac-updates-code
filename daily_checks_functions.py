import pandas as pd
import numpy as np
import os
from extra_functions import to_raw


def find_inel_articles(df: pd.DataFrame, col_name: str, inst_list: list, eligibilitytypes: list, col_list: list) \
        -> pd.DataFrame:
    """finds articles with ineligible article/license types and returns them as a new Data Frame

    Args:
        df: Pandas Data Frame containing the updated orl/foal
        col_name: column name to be checked
        inst_list: Consortium whose institutions have certain article/license types
        eligibilitytypes: list of eligible article/license types
        col_list: column list to be returned
    Returns:
        Filtered df

    """
    my_df = df.loc[(df["Author Affiliation"].isin(inst_list))
                   & (~df[col_name].isin(eligibilitytypes))
                   & (~df["Comments"].str.contains("deleted", case=False, na=False))
                   & (~df["Comments"].str.contains("flagged", case=False, na=False)),
                   col_list]

    return my_df


def find_fullyoa_inst(df: pd.DataFrame, inst_list: list, col_list: list):
    """finds articles whose institution should not be in the Hybrid OA file

    Args:
        df: Pandas Data Frame containing the updated orl
        inst_list: institution list that only has fully OA
        col_list: column list to be returned
    """
    my_df = df.loc[df["Author Affiliation"].isin(inst_list), col_list]

    return my_df


def find_must_oa_optouts(df: pd.DataFrame, inst_list: list, col_list: list):
    """finds optout articles where the institution is part of the default OA list

    Args:
        df: Pandas Data Frame containing the updated orl
        inst_list: institution list which have default OA
        col_list: column list to be returned
    """

    my_df = df.loc[(~df["Comments"].str.contains("Flagged", case=False, na=False))
                   & (~df["Comments"].str.contains("deleted", case=False, na=False))
                   & (df["Author Affiliation"].isin(inst_list)),
                   col_list]

    return my_df


def find_inel_ATs_foal(df: pd.DataFrame, agr_list: list, articletypes: list, col_list: list):
    """finds articles with ineligible article/license types and returns them as a new Data Frame

    Args:
        df: Pandas Data Frame containing the updated orl/foal
        agr_list: Consortium list whose institutions have certain article types
        articletypes: list of eligible article/license types
        col_list: column list to be returned

    """
    my_df = df.loc[(df["Agreement"].isin(agr_list))
                   & (~df["Article Type"].isin(articletypes))
                   & (~df["Comments"].str.contains("flagged", case=False, na=False)),
                   col_list]

    return my_df


def add_action_cols(df: pd.DataFrame, issue: str, action: str):
    """Adds columns with issues and action to take for each QA check problem

    Args:
        df: Pandas Data Frame containing the list of issues
        issue: problem found in the QA functions
        action: action to take
    """

    df["Issue"] = issue
    df["Action to take"] = action

    return df


def qa_checks_hybrid(path, login, desktop, institution_list_hybrid, aas: pd.DataFrame, optout: pd.DataFrame,
                     fullyoa_only: list, articletypes_reg: list, articletypes_deal: list,
                     articletypes_ita: list, col_list: list):
    """ calls all the orl QA functions and appends the Data Frames to return a single QA Data Frame

    Args:
        path, login, desktop, institution_list_compact: imported from general_variables
        aas: updated oa req list (non-opt-outs)
        optout: updated oa req list (opt-outs)
        no_palgrave_list: agreement list which have no Palgrave articles
        fullyoa_only: list of BIBSAM institutions which only have fully OA
        articletypes_reg: normal article types list
        articletypes_deal: DEAL article types
        articletypes_ita: CRUI-CARE and CNR article types
        licensetypes_reg: normal license types list
        licensetypes_jisc: JISC license types
        col_list: columns for QA checks

    """
    inst_list_orl = pd.read_excel(to_raw(os.path.join(path, login, desktop, institution_list_hybrid)))
    optout_all = inst_list_orl.loc[inst_list_orl["opt out"] == "No", "Institution name"].tolist()
    optout_jisc = inst_list_orl.loc[(inst_list_orl["opt out"] == "Yes")
                                    & (inst_list_orl["Agreement"] == "JISC"), "Institution name"].tolist()
    optout_nojisc_inst = inst_list_orl.loc[(inst_list_orl["opt out"] == "Yes")
                                           & (inst_list_orl["Agreement"] != "JISC"), "Institution name"].tolist()
    optout_reg_inst = inst_list_orl.loc[(inst_list_orl["opt out"] == "Yes")
                                        & (inst_list_orl["Agreement"] != "DEAL"), "Institution name"].tolist()
    normal_AT_inst_list = inst_list_orl.loc[
        ~inst_list_orl["Agreement"].isin(["DEAL", "CARE", "CNR"]), "Institution name"].tolist()
    deal_inst_list = inst_list_orl.loc[inst_list_orl["Agreement"] == "DEAL", "Institution name"].tolist()
    italian_inst_list = inst_list_orl.loc[inst_list_orl["Agreement"].isin(["CARE", "CNR"]), "Institution name"].tolist()
    normal_LT_inst_list = inst_list_orl.loc[inst_list_orl["Agreement"] != "JISC", "Institution name"].tolist()
    jisc_inst_list = inst_list_orl.loc[inst_list_orl["Agreement"] == "JISC", "Institution name"].tolist()

    wrongAT_regular = find_inel_articles(aas, "Article Type", normal_AT_inst_list, articletypes_reg, col_list)
    wrongAT_DEAL = find_inel_articles(aas, "Article Type", deal_inst_list, articletypes_deal, col_list)
    wrongAT_ITA = find_inel_articles(aas, "Article Type", italian_inst_list, articletypes_ita, col_list)
    wrongAT_optout = find_inel_articles(optout, "Article Type", optout_reg_inst, articletypes_reg, col_list)
    wrongAT_DEAL_optout = find_inel_articles(optout, "Article Type", deal_inst_list, articletypes_deal, col_list)
    wrongAT_ITA_optout = find_inel_articles(optout, "Article Type", italian_inst_list, articletypes_ita, col_list)

    fullyOA_bibsam = find_fullyoa_inst(aas, fullyoa_only, col_list)
    defaultOA_optout = find_must_oa_optouts(optout, optout_all, col_list)

    article_type_issue = wrongAT_regular.append(wrongAT_DEAL).append(wrongAT_ITA).append(wrongAT_optout).append(
        wrongAT_DEAL_optout).append(wrongAT_ITA_optout)

    article_type_issue = add_action_cols(article_type_issue, "Ineligible article type",
                                         "Message #ask-oasis channel and tag relevant OAC")
    fullyOA_bibsam = add_action_cols(fullyOA_bibsam, "Fully OA-only institution not eligible for Compact",
                                     "Message #ask-oasis channel and tag relevant OAC")
    defaultOA_optout = add_action_cols(defaultOA_optout,
                                       "Institution has default OA but author could select to opt out",
                                       "Message #ask-oasis channel and tag relevant OAC")

    all_errors_orl = article_type_issue.append(fullyOA_bibsam).append(defaultOA_optout)

    return all_errors_orl[col_list]


def qa_checks_foa(foal: pd.DataFrame, articletypes_reg: list, articletypes_deal_stand: list,
                  normal_AT_agr_list: list, col_list: list):
    """ calls all the foal QA functions and appends the Data Frames to return a single QA Data Frame

    Args:
        path, login, desktop, journal_list_file: imported from general_variables
        foal: updated foal's 1st sheet
        recrep: today's reconciliation report
        approved_mship: list of eligible membership states for approved articles
        rejected_mship: list of eligible membership states for rejected articles
        waiver_app: list of eligible waiver types for approved articles
        no_aas_mship: list of membership states for articles which have not reached the aas
        aas_off_status: list of membership states for articles where AAS is off
        articletypes_reg: list of regular article types
        articletypes_deal_stand: list of DEAL and standalone article types
        normal_AT_agr_list: list of agreements with regular article types
        col_list: columns for QA checks
        recrep_rename_dict: columns to be renamed
    """
    inel_type = find_inel_ATs_foal(foal, normal_AT_agr_list, articletypes_reg, col_list)
    inel_type_other = find_inel_ATs_foal(foal, ["DEAL"], articletypes_deal_stand, col_list)

    inel_article_type = inel_type.append(inel_type_other)

    inel_article_type = add_action_cols(inel_article_type,
                                        "Article type is not eligible but the article entered the AAS",
                                        "Flag to OASiS team on Slack (@here and @RS)")

    all_errors_foal = inel_article_type

    return all_errors_foal[col_list]


def approve_unverified(df_hybrid: pd.DataFrame, df_foa: pd.DataFrame, app2: list, app3: list, app4: list, app5: list):
    """Filter the articles which have had final reminders sent and need to be approved manually by OACs

    Args:
        df_hybrid: unverified hybrid articles list
        df_foa: fully oa articles
        app2: list of countries whose articles need to be approved 2 business days after the final reminder is sent
        app3: list of countries whose articles need to be approved 3 business days after the final reminder is sent
        app4: list of countries whose articles need to be approved 4 business days after the final reminder is sent
        app5: list of countries whose articles need to be approved 5 business days after the final reminder is sent

    Returns:
        Pandas Data Frame with articles that need to be approved
    """
    unverified_hybrid = df_hybrid.loc[(df_hybrid["Final Reminder Sent?"] == "Yes")]

    unverified_foa = df_foa.loc[(df_foa["Date Approved"].isnull())
                                & (df_foa["Date Rejected"].isnull())
                                & (df_foa["Date Forwarded"].isnull())
                                & (df_foa["Final Reminder Sent?"] == "Yes")]

    unverified = unverified_hybrid.append(unverified_foa)

    df_2BDays_pre = unverified[unverified["Final Reminder to be Sent"] == pd.Timestamp.today().date() - pd.offsets.BDay(2)]
    df_3BDays_pre = unverified[unverified["Final Reminder to be Sent"] == pd.Timestamp.today().date() - pd.offsets.BDay(3)]
    df_4BDays_pre = unverified[unverified["Final Reminder to be Sent"] == pd.Timestamp.today().date() - pd.offsets.BDay(4)]
    df_5BDays_pre = unverified[unverified["Final Reminder to be Sent"] == pd.Timestamp.today().date() - pd.offsets.BDay(5)]

    df_2 = df_2BDays_pre.loc[df_2BDays_pre["Agreement"].isin(app2)]
    df_3 = df_3BDays_pre.loc[df_3BDays_pre["Agreement"].isin(app3)]
    df_4 = df_4BDays_pre.loc[df_4BDays_pre["Agreement"].isin(app4)]
    df_5 = df_5BDays_pre.loc[df_5BDays_pre["Agreement"].isin(app5)]

    to_be_approved = df_2.append(df_3).append(df_4).append(df_5)

    return to_be_approved[["DOI", "Unique Article ID", "Article Title", "Agreement"]]
