import pandas as pd
from general_variables import pie_colours, simple_bar_chart_colours
from master_dicts import (palm_col_dict, recrep_column_dict, aas_rep_dict, AAS_rep_dict_foal,
    articletype_dict, jflux_dict, month_name, doi_dict_1, doi_dict_2, doi_dict_3, doi_dict_4,
    doi_dict_5, doi_dict_6, error_trends_dict_2, error_trends_dict_3, email_country_dict)
from master_lists import (optout_df_cols, palm_optout_cols, aas_dashboard_cols, approved_cols,
    rejected_cols, in_verification_cols, recon_report_columns, complete_oa_req_cols, FOAL_columns, palm_req_cols,
    cancelled_cols, aas_rep_update, foal_update_cols, hybrid_errors_cols, foal_errors_cols, all_errors_cols,
    AAS_rep_columns_foal, aas_rep_update_foal, standalone_list, del_cols_to_overwrite, files_to_delete,
    horizontal_bar_chart_colours, old_new_list, reason_del_list, req_cols, palm_dates, appr_time_cols, rej_time_cols,
    unver_time_cols, optout_time_cols, dnr_time_cols, col_times_foal, final_del_cols, normal_article_types,
    deal_article_types, fullyoa_bibsam, reg_at, fr_in_3bdays, fr_in_5bdays, approve_after2,
    approve_after3, approve_after4, approve_after5)

master_dict_of_cols = {0: optout_df_cols,
                       1: palm_optout_cols,
                       2: aas_dashboard_cols,

                       3: col_times_foal,
                       4: col_times_foal[:6],

                       5: complete_oa_req_cols,
                       6: palm_req_cols,
                       7: palm_dates,

                       9: cancelled_cols,
                       10: appr_time_cols,

                       11: recon_report_columns,

                       12: approved_cols,
                       13: rejected_cols,
                       14: in_verification_cols,

                       15: FOAL_columns,

                       16: rej_time_cols,
                       17: unver_time_cols,

                       18: foal_errors_cols,
                       19: final_del_cols,

                       20: optout_time_cols,

                       21: AAS_rep_columns_foal,

                       22: dnr_time_cols,
                       23: del_cols_to_overwrite,
                       25: hybrid_errors_cols,
                       26: all_errors_cols
                       }

master_dict_of_dicts = {0: palm_col_dict,

                        1: recrep_column_dict,
                        2: aas_rep_dict,
                        3: email_country_dict,

                        5: AAS_rep_dict_foal,
                        6: articletype_dict
                        }

master_dict_for_updates = {0: fr_in_3bdays,
                           1: fr_in_5bdays,
                           2: fr_in_5bdays[1:],

                           6: aas_rep_update,

                           9: foal_update_cols,
                           11: aas_rep_update_foal,
                           12: standalone_list
                           }

master_dict_for_checks = {0: normal_article_types,
                          1: deal_article_types,
                          2: normal_article_types[0:2],

                          11: [x.upper() for x in normal_article_types],
                          12: [x.upper() for x in normal_article_types][0:3],
                          14: fullyoa_bibsam,
                          15: reg_at,
                          16: approve_after2,
                          17: approve_after3,
                          18: approve_after4,
                          19: approve_after5,
                          20: approve_after3[:-1]
                          }

master_dict_for_deleted_dois = {0: jflux_dict,

                                1: palm_dates,

                                4: appr_time_cols,
                                5: rej_time_cols,
                                6: unver_time_cols,

                                8: dnr_time_cols,
                                9: optout_time_cols
                                }

master_dict_for_error_trends = {0: doi_dict_1,
                                1: doi_dict_2,
                                2: doi_dict_3,
                                3: doi_dict_4,
                                4: doi_dict_5,
                                5: doi_dict_6,
                                6: pie_colours,
                                7: simple_bar_chart_colours,
                                8: horizontal_bar_chart_colours,

                                10: error_trends_dict_2,
                                11: error_trends_dict_3,
                                12: files_to_delete,
                                13: req_cols,
                                14: month_name,
                                15: reason_del_list,
                                16: old_new_list
                                }

read_me_hybrid = pd.DataFrame({
    "Springer":
        ["",
         "",
         "",
         "Purpose",
         "Sources",
         "Data Restrictions"
         ],

    "Nature":
        ["",
         "",
         "",
         "This file contains a list of all articles in hybrid journals for which the authors have requested the"
         " APC to be covered by their institution."
         " The articles are split in different sheets based on the verification status and date."
         " All article cancellations can be found in the 'Do NOT Report' tab."
         " Articles are ordered by Eligibility Requested or Date Sent to AAS."
         " This file is the basis of the customer monthly reports, the end-of-year countdown of APCs,"
         " and it is used to manually track edge cases as well as non-automated processes (e.g. final reminders).",
         "The main source for this file is the PALM reconciliation report."
         " Older article data was sourced from AqApp/s150 reports (decommissioned on 31 Oct 2021)."
         " Article data is further being supplement by AAS reports data using SuperUser privileges."
         " Where data is present in more than one source, the hierarchy for filling in the fields is:"
         " PALM report -> AAS report -> s150."
         " In the 'Do NOT Report' sheet, Dates are filled in with today's date every Friday where the Eligibility Date"
         " is not available."
         " All comments are filled in by the Open Access Coordinators;"
         " Final Reminder Sent information is handled by the OA Verification Team.",
         "As this file is updated daily using scripts that read certain columns, and because this file is used for"
         " monthly reports and KPI analyses, please make sure to observe the following rules:"
         " All articles that are deleted from the Production system MUST have a comment that contains the word"
         " 'deleted'. All articles that have a 'Do NOT Report' status must have a reason/comment. All agreements"
         " acronyms must be consistent with previous entries. Comments should be reserved only for cases where the"
         " information cannot be obtained from any other field/column."
         ]
})

read_me_foa = pd.DataFrame({
    "Springer":
        ["",
         "",
         "",
         "Purpose",
         "Sources",
         "Data Restrictions"
         ],

    "Nature":
        ["",
         "",
         "",
         "This file contains a list of all articles in fully Open Access journals for which the authors have requested"
         " the APC to be covered by their institution."
         " Articles are ordered by Eligibility Requested or Date Sent to AAS."
         " This file is the basis of the customer monthly reports, the end-of-year countdown of APCs,"
         " and it is used to manually track edge cases as well as non-automated processes (e.g. final reminders).",
         "The main source for this file is the deal interceptor reconciliation report."
         " Article data is further being supplement by AAS reports data using SuperUser privileges."
         " Where data is present in more than one source, the hierarchy for filling in the fields is:"
         " Deal interceptor reconciliation report -> AAS reports."
         " All comments are filled in by the Open Access Coordinators;"
         " Final Reminder Sent information is handled by the OA Verification Team.",
         "As this file is updated daily using scripts that read certain columns, and because this file is used for"
         " monthly reports and KPI analyses, please make sure to observe the following rules:"
         " All articles that have a 'Do NOT Report' status must have a reason/comment. All agreements"
         " acronyms must be consistent with previous entries. Comments should be reserved only for cases where the"
         " information cannot be obtained from any other field/column."
         ]
})

df0 = pd.DataFrame({
    "Reason for Deletion":
        ["N/A",
         "AM verified by mistake",
         "Author mistake in OA request forms",
         "Author wanted to opt out",
         "CDL - opt-out due to funding issues",
         "Change in CA to make article eligible",
         "Other",
         "Production mistake",
         "Revised manuscript",
         "Tech issue",
         "Under investigation",
         "Withdrawn from pub."
         ],
    "Explanation":
        ["Article not deleted - Do not report for other reasons (e.g. cancellations, sponsorships)",
         "The Approval Manager has approved or rejected the article by mistake.",
         "The author has entered the wrong information (incorrect institution spelling, similar institution names,"
         " personal email address, and similar) when filling in the OA request forms.",
         "The author was recognised as eligible for a TA but did not want to publish Open Access.",
         "The author was recognised as eligible for the TA under the CDL agreement but had to opt-out because they"
         "were not able to cover the APC.",
         "The author is not recognised as eligible but would like to publish OA using their institution's Compact"
         "agreement (includes cases where a group of authors - some of which are affiliated with a Compact-eligible"
         " institution - is using the main researcher's email address as the corresponding author, but the main"
         " researcher is not affiliated with any Compact-eligible institution; an author has been recognised by IP"
         " address when they're vising at another institution that they're not affiliated with and the Approval Manager"
         " has rejected the article).",
         "Any other reason, including DOIs deleted due to sponsorship agreements covering the article.",
         "DOI deleted because of a mistake made by the Production team (e.g. accidental deletion or duplicate article"
         " creation).",
         "The author or Editor provided a heavily revised manuscript.",
         "Springer Nature technical error, OASiS workflow issue, or wider bug.",
         "DOI currently being investigated by the OAC with Production Editors to find out the reason for deletion and"
         " any potential new DOIs to replace the deleted ones.",
         "Article temporarily or permanently withdrawn from publication."
         ]})
