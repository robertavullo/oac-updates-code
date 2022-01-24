import pandas as pd
import os
import calendar
from datetime import date
from timeit import default_timer as timer
import warnings
from daily_update_functions import create_optouts, create_oa_req_list, find_correct_verification, create_foa
from weekly_update_functions import production_checks
from daily_checks_functions import qa_checks_hybrid, qa_checks_foa, approve_unverified
from formatting_functions import format_hybrid, format_foa
from extra_functions import to_raw, overwrite_cols, create_excel_file, rename_files
from master_dicts_of_lists import master_dict_of_cols, master_dict_of_dicts, master_dict_for_updates, \
    master_dict_for_checks
from general_variables import path, desktop, timestr, current_year, quarter_dict, current_month, \
    institution_list_hybrid, institution_list_fullyOA, journal_list_file

warnings.filterwarnings("ignore")

def master_function_to_rule_them_all(login):
    """Filter PALM dashboard, reconciliation, and AAS reports, shuffle columns and merge with the\
       previous day's HAL (?) and FOAL

    Args:
        login: login code
    """

    before = timer()
    print("Reading the AAS reports...")
    print("(You better get a coffee, this might take a while)")
    AAS_open = pd.read_excel(to_raw(os.path.join(path, login, desktop, "AAS_open.xlsx")))
    AAS_approved = pd.read_excel(to_raw(os.path.join(path, login, desktop, "AAS_approved.xlsx")))
    AAS_rejected = pd.read_excel(to_raw(os.path.join(path, login, desktop, "AAS_rejected.xlsx")))
    AAS_report = AAS_open.append(AAS_approved).append(AAS_rejected)
    print("... reading the PALM report...")
    print("(Seriously, get a drink)")
    print("('tis a big report)")
    palm_rep = pd.read_excel(to_raw(os.path.join(path, login, desktop, "palm_report.xlsx")))
    palm_rep = palm_rep.rename(columns=master_dict_of_dicts[0])
    print("... reading the fully OA reconciliation report...")
    recon_report = pd.read_csv(to_raw(os.path.join(path, login, desktop, "recon_report.csv")), encoding="utf-8")
    print("... and even MORE files!")
    dnr = pd.read_excel(
        to_raw(os.path.join(path, login, desktop, "Hybrid Journals TA Article List " + str(current_year) + ".xlsx")),
        "Do NOT Report")

    print("Working on the opt-outs...")
    optouts = create_optouts(path, login, desktop, institution_list_hybrid, journal_list_file, palm_rep,
                             master_dict_of_cols[0], master_dict_of_cols[1], current_year, master_dict_of_dicts[3])
    print("Opt-out sheet complete!")

    print("Now working on the non-opt-out hybrid articles...")
    oa_req_list = create_oa_req_list(path, login, desktop, institution_list_hybrid, journal_list_file, palm_rep,
                                     master_dict_of_cols[2], master_dict_of_cols[5], master_dict_of_cols[6],
                                     master_dict_of_cols[7], master_dict_for_updates[0], master_dict_for_updates[1],
                                     master_dict_of_dicts[2], master_dict_of_dicts[0], current_year,
                                     master_dict_of_dicts[3])
    print("Time to split the articles in the different tabs!")
    oa_req_list = find_correct_verification(current_year, oa_req_list, master_dict_of_cols[5])

    appr_this_year = oa_req_list.loc[oa_req_list["sheet in file"] == "1", master_dict_of_cols[12]]
    print(f"Sheet with {current_year} approvals done")
    appr_previously = oa_req_list.loc[oa_req_list["sheet in file"] == "2", master_dict_of_cols[12]]
    print("Sheet with previous years approvals done")
    rejected_articles = oa_req_list.loc[oa_req_list["sheet in file"] == "3", master_dict_of_cols[13]]
    print("Sheet with rejections done")
    unverified_articles = oa_req_list.loc[oa_req_list["sheet in file"] == "4", master_dict_of_cols[14]]
    print("Sheet with unverified articles done")

    print("Time to fix the cancellations sheet")
    oa_req_list["Date"] = oa_req_list["Date Sent to AAS"]

    cancellations = oa_req_list.loc[oa_req_list["sheet in file"] == "6", master_dict_of_cols[9]]
    cancellations["Date"] = cancellations["Date"].fillna(pd.Timestamp.today().date())
    dnr["Date"] = pd.to_datetime(dnr["Date"], errors="coerce")

    cancellations = cancellations.merge(dnr, how="outer", on="DOI")
    overwrite_cols(master_dict_of_cols[23], cancellations)
    cancellations = cancellations[master_dict_of_cols[9]].sort_values(by=["Date"])

    print("Sheet with cancellations done")

    foal_df = create_foa(path, login, desktop, institution_list_fullyOA, journal_list_file, recon_report, AAS_report,
                         master_dict_of_cols[11], master_dict_of_cols[21], master_dict_of_cols[15],
                         master_dict_for_updates[9], master_dict_for_updates[11], master_dict_of_cols[4],
                         master_dict_for_updates[0], master_dict_for_updates[2], master_dict_for_updates[12],
                         master_dict_of_dicts[1], master_dict_of_dicts[5], master_dict_of_dicts[6], current_year)
    print("FOAL update done")

    print("Formatting all the files now...")
    cancellations["Date"] = pd.to_datetime(cancellations["Date"], errors="coerce")

    format_hybrid(path, login, desktop, appr_this_year, appr_previously, rejected_articles,
                  unverified_articles, optouts, cancellations, master_dict_of_cols[10], master_dict_of_cols[16],
                  master_dict_of_cols[17], master_dict_of_cols[20], master_dict_of_cols[22], current_year)
    format_foa(path, login, desktop, foal_df, master_dict_of_cols[3], current_year)

    print("Looking for errors and manual approvals...")
    all_errors_hybrid = qa_checks_hybrid(path, login, desktop, institution_list_hybrid, oa_req_list, optouts,
                                         master_dict_for_checks[14], master_dict_for_checks[0],
                                         master_dict_for_checks[1], master_dict_for_checks[2], master_dict_of_cols[25])
    all_errors_hybrid["sheet in file"] = all_errors_hybrid["sheet in file"].fillna("5")

    all_errors_foa = qa_checks_foa(foal_df, master_dict_for_checks[11], master_dict_for_checks[12],
                                   master_dict_for_checks[15], master_dict_of_cols[18])

    approve = approve_unverified(unverified_articles, foal_df, master_dict_for_checks[16], master_dict_for_checks[17],
                                 master_dict_for_checks[18], master_dict_for_checks[19])

    all_errors = all_errors_hybrid.append(all_errors_foa)[master_dict_of_cols[26]].rename(
        columns={"sheet in file": "what sheet to find this article in"}).replace(
        {"what sheet to find this article in": {
            "1": "Approved this year",
            "2": "Approved (prev. years)",
            "3": "Rejected",
            "4": "In Verification",
            "5": "Opt-Out",
            "6": "Do NOT Report"
        }})
    all_errors["what sheet to find this article in"] = all_errors["what sheet to find this article in"].fillna(
        "Fully OA file")

    print("You're almost there!")

    AAS_report.to_excel(to_raw(os.path.join(path, login, desktop, "AAS_report" + timestr + ".xlsx")),
                        index=False)
    os.remove(to_raw(os.path.join(path, login, desktop, "AAS_open.xlsx")))
    os.remove(to_raw(os.path.join(path, login, desktop, "AAS_approved.xlsx")))
    os.remove(to_raw(os.path.join(path, login, desktop, "AAS_rejected.xlsx")))

    rename_files(path, login, desktop, timestr, "recon_report.csv", "recon_report", ".csv")
    rename_files(path, login, desktop, timestr, "palm_report.xlsx", "palm_report", ".xlsx")

    if len(all_errors) == 0:
        print("No errors today! *dancing panda*")
    else:
        print("sad panda :(")
        create_excel_file(path, login, desktop, timestr, all_errors, "QA checks", ".xlsx")
        print("Remember to add comments to the daily update files with the issue of the day (make sure to include the "
              "word 'flagged' to avoid the same articles being flagged every day) and who you flagged the issue to!")
        print("Now wait for approvals...")

    if len(approve) == 0:
        print("Nothing to approve today!")
    else:
        print("Creating the approvals file...")
        create_excel_file(path, login, desktop, timestr, approve, "Manual approvals", ".xlsx")

    if calendar.day_name[date.today().weekday()] == "Friday":
        print("Creating the upload for Delilah-JFlux")

        for_delilah = production_checks(path, login, desktop, journal_list_file, appr_this_year, appr_previously,
                                        rejected_articles, unverified_articles, optouts, master_dict_of_cols[19])

        for_delilah.to_excel(to_raw(os.path.join(path, login, desktop, "for_delilah.xlsx")), index=False)
        print("You're almost there!")

        after = timer()
        m, s = divmod(after - before, 60)
        print(f"Daily update update completed in {round(m)} minutes {round(s)} seconds")
        print("Done!\nNow check the Agreement column and add any missing institutions, then do the extended weekly "
              "checks.")
    else:
        after = timer()
        m, s = divmod(after - before, 60)
        print(f"Daily update update completed in {round(m)} minutes {round(s)} seconds")
        print("Done!\nNow check the Agreement column and add any missing institutions")
