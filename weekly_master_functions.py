import pandas as pd
import os
from timeit import default_timer as timer
from matplotlib import pyplot as plt
from weekly_update_functions import create_deleted_list
from formatting_functions import format_hybrid, format_date
from extra_functions import to_raw, rename_files, overwrite_cols, create_groupby_df, delete_excel_files, \
    create_percentage_as_number, create_percentage_string, overwrite_num_cols, \
    create_percentage_str_number, convert
from master_dicts_of_lists import master_dict_for_deleted_dois, df0, master_dict_for_error_trends, master_dict_of_cols
from general_variables import path, desktop, timestr, current_year, quarter_dict, current_month, \
    institution_list_hybrid, institution_list_fullyOA, journal_list_file


def press_the_big_red_button(login):
    """
    Args:
        login: login code
    """
    before = timer()
    print("Reading the all the sheets in all the files...")
    appr_this_year = pd.read_excel(
        to_raw(os.path.join(path, login, desktop, "Hybrid Journals TA Article List " + str(current_year) + ".xlsx")),
        "Approved")
    print("... the many...")
    appr_previously = pd.read_excel(
        to_raw(os.path.join(path, login, desktop, "Hybrid Journals TA Article List " + str(current_year) + ".xlsx")),
        "Approved (prev. years)")
    rejected_articles = pd.read_excel(
        to_raw(os.path.join(path, login, desktop, "Hybrid Journals TA Article List " + str(current_year) + ".xlsx")),
        "Rejected")
    print("... many...")
    unverified_articles = pd.read_excel(
        to_raw(os.path.join(path, login, desktop, "Hybrid Journals TA Article List " + str(current_year) + ".xlsx")),
        "In Verification")
    optouts = pd.read_excel(
        to_raw(os.path.join(path, login, desktop, "Hybrid Journals TA Article List " + str(current_year) + ".xlsx")),
        "Opt-Outs")
    print("... sheets")
    cancellations = pd.read_excel(
        to_raw(os.path.join(path, login, desktop, "Hybrid Journals TA Article List " + str(current_year) + ".xlsx")),
        "Do NOT Report")
    print("Now onto the Delilah report")
    jflux_report = pd.read_excel(to_raw(os.path.join(path, login, desktop, "jflux.xlsx")), "Results1")

    jflux_report = jflux_report.rename(columns=master_dict_for_deleted_dois[0])

    print("Updating the Do NOT Report sheet")
    cancellations = create_deleted_list(jflux_report, cancellations, master_dict_of_cols[23],
                                        master_dict_for_deleted_dois[3])

    col_list = ["Final Reminder to be Sent", "Institutional Share", "Author Share"]

    for col in col_list:
        appr_this_year[col] = appr_this_year[col].fillna("N/A")
        appr_previously[col] = appr_previously[col].fillna("N/A")
        rejected_articles[col] = rejected_articles[col].fillna("N/A")
        unverified_articles[col] = unverified_articles[col].fillna("N/A")

    print("Formatting the files again...")
    format_hybrid(path, login, desktop, appr_this_year, appr_previously, rejected_articles, unverified_articles,
                  optouts, cancellations, master_dict_for_deleted_dois[4], master_dict_for_deleted_dois[5],
                  master_dict_for_deleted_dois[6], master_dict_for_deleted_dois[9], master_dict_for_deleted_dois[8],
                  current_year)

    os.remove(to_raw(os.path.join(path, login, desktop, "for_delilah.xlsx")))
    rename_files(path, login, desktop, timestr, "jflux.xlsx", "Delilah-JFlux_report", ".xlsx")

    after = timer()
    m, s = divmod(after - before, 60)
    print(f"Deleted DOI investigations completed in {round(m)} minutes {round(s)} seconds")
    print("Done!\nNow send the update email to the team and follow the guidelines for the error trend analysis.")


def create_error_trends(login):
    """
        Args:
            login: login code
        """

    before = timer()
    print("Reading the ORL...")
    orl_aas = pd.read_excel(to_raw(os.path.join(path, login, desktop, "overview request " + str(
        current_year) + "_" + quarter_dict.get(current_month) + ".xlsx")), "Articles in AAS")
    orl_optout = pd.read_excel(to_raw(os.path.join(path, login, desktop, "overview request " + str(
        current_year) + "_" + quarter_dict.get(current_month) + ".xlsx")), "Opt-Out Articles")
    print("... and this week's deleted DOIs...")
    orl_deleted = pd.read_excel(to_raw(os.path.join(path, login, desktop, "overview request " + str(
        current_year) + "_" + quarter_dict.get(current_month) + ".xlsx")), "Articles to Delete from AAS")

    format_date(orl_aas, ["Date Approved", "Date Rejected"])
    format_date(orl_deleted, ["Date"])

    orl_deleted["Month"] = pd.DatetimeIndex(orl_deleted["Date"]).month
    orl_deleted["Year"] = pd.DatetimeIndex(orl_deleted["Date"]).year

    create_del_doi_bar_chart(login, orl_deleted)
    """create_yes_no_percentage(login, orl_deleted)
    create_yes_no_stacked(login, orl_deleted)"""
    result1 = create_old_vs_new_doi(login, orl_aas, orl_optout, orl_deleted)
    create_del_explanations(login, orl_deleted)
    result2 = create_table_by_country(login, orl_deleted)
    result3 = create_graph_by_journal(login, orl_deleted)
    create_graph_by_pe(login, orl_deleted)

    create_error_trends_excel_file(login)

    after = timer()
    m, s = divmod(after - before, 60)
    print(f"Deleted DOI error trend analysis completed in {round(m)} minutes {round(s)} seconds")
    print("Done!\nNow send the update email to the team with screenshots of the saved images and of the tables below.")

    return result1, result2, result3


def create_error_trends_excel_file(login):
    writer = pd.ExcelWriter(to_raw(os.path.join(path, login, desktop, "Deleted_DOI_error_trends" + timestr + ".xlsx")),
                            engine="xlsxwriter")

    df1 = pd.read_excel(to_raw(os.path.join(path, login, desktop, "Total no. of deleted DOIs.xlsx")))
    """df2 = pd.read_excel(to_raw(os.path.join(path, login, desktop, "Found through checks.xlsx")))
    df3 = pd.read_excel(to_raw(os.path.join(path, login, desktop, "Found through checks progress.xlsx")))"""
    df4 = pd.read_excel(to_raw(os.path.join(path, login, desktop, "Old vs. New DOI.xlsx")))
    df5 = pd.read_excel(to_raw(os.path.join(path, login, desktop, "Reason for deletion.xlsx")))
    df6 = pd.read_excel(to_raw(os.path.join(path, login, desktop, "Deletion by Agreement.xlsx")))
    df7 = pd.read_excel(to_raw(os.path.join(path, login, desktop, "Deletion by Journal.xlsx")))
    df8 = pd.read_excel(to_raw(os.path.join(path, login, desktop, "Deletion by PE.xlsx")))

    df0.to_excel(writer, sheet_name="ReadMe – Deletion explanations", index=False)
    df1.to_excel(writer, sheet_name="Total no. of deleted DOIs", index=False)
    """df2.to_excel(writer, sheet_name="Found through checks", index=False)
    df3.to_excel(writer, sheet_name="Found through checks progress", index=False)"""
    df4.to_excel(writer, sheet_name="Old vs. New DOI", index=False)
    df5.to_excel(writer, sheet_name="Reason for deletion", index=False)
    df6.to_excel(writer, sheet_name="Deletion by Agreement", index=False)
    df7.to_excel(writer, sheet_name="Deletion by Journal", index=False)
    df8.to_excel(writer, sheet_name="Deletion by PE", index=False)

    ### formats
    workbook = writer.book
    header_format = workbook.add_format({"align": "center",
                                         "valign": "top",
                                         "bold": True,
                                         "border": True
                                         })

    header_2 = workbook.add_format({"text_wrap": True,
                                    "align": "center",
                                    "valign": "top",
                                    "bold": True,
                                    "border": True,
                                    "bg_color": "#D7D7D7"
                                    })

    wrap_text = workbook.add_format({"text_wrap": True,
                                     "valign": "top"
                                     })

    all_border_cells = workbook.add_format({"valign": "top",
                                            "border": True
                                            })

    ## creating worksheets

    print("Formatting the Excel file for the error trend team update email")
    # 1st Excel sheet
    worksheet1 = writer.sheets["ReadMe – Deletion explanations"]

    worksheet1.set_zoom(85)
    worksheet1.set_row(0, 15)
    worksheet1.set_column(0, 0, 45)
    worksheet1.set_column(1, 1, 120)

    for col_num, value in enumerate(df0.columns.values):
        worksheet1.write(0, col_num, value, header_format)

    for row_num, data in enumerate(df0["Reason for Deletion – Category"]):
        worksheet1.write(row_num + 1, 0, data, wrap_text)

    for row_num, data in enumerate(df0["Explanation"]):
        worksheet1.write(row_num + 1, 1, data, wrap_text)

    # 2nd Excel sheet
    worksheet2 = writer.sheets["Total no. of deleted DOIs"]

    worksheet2.freeze_panes(1, 0)

    worksheet2.set_zoom(85)
    worksheet2.set_column(0, 1, 12)

    for col_num, value in enumerate(df1.columns.values):
        worksheet2.write(0, col_num, value, header_format)

    """# 3rd Excel sheet
    worksheet3 = writer.sheets["Found through checks"]

    worksheet3.set_column(0, 0, 33)
    worksheet3.set_column(1, 1, 11)

    for col_num, value in enumerate(df2.columns.values):
        worksheet3.write(0, col_num, value, header_2)

    for row_num, data in enumerate(df2["Found through deleted DOI checks?"]):
        worksheet3.write(row_num + 1, 0, data, all_border_cells)

    for row_num, data in enumerate(df2["Yes/No"]):
        worksheet3.write(row_num + 1, 1, data, all_border_cells)

    # 4th Excel sheet
    worksheet4 = writer.sheets["Found through checks progress"]

    worksheet4.set_zoom(85)
    worksheet4.freeze_panes(1, 0)
    worksheet4.set_column(0, 0, 12)
    worksheet4.set_column(1, 1, 32)
    worksheet4.set_column(2, 2, 12)

    for col_num, value in enumerate(df3.columns.values):
        worksheet4.write(0, col_num, value, header_format)"""

    # 5th Excel sheet
    worksheet5 = writer.sheets["Old vs. New DOI"]

    worksheet5.set_zoom(90)
    worksheet5.set_row(0, 15)
    worksheet5.set_column(0, 0, 20)
    worksheet5.set_column(1, 1, 32)
    worksheet5.set_column(2, 2, 15)

    for col_num, value in enumerate(df4.columns.values):
        worksheet5.write(0, col_num, value, header_format)

    # 6th Excel sheet
    worksheet6 = writer.sheets["Reason for deletion"]

    worksheet6.set_column(0, 0, 45)
    worksheet6.set_column(1, 2, 12)

    for col_num, value in enumerate(df5.columns.values):
        worksheet6.write(0, col_num, value, header_format)

    # 7th Excel sheet
    worksheet7 = writer.sheets["Deletion by Agreement"]

    worksheet7.set_row(0, 15)
    worksheet7.freeze_panes(1, 0)
    worksheet7.set_zoom(80)
    worksheet7.set_column(0, 0, 12)
    worksheet7.set_column(1, 1, 48)
    worksheet7.set_column(2, 2, 12)

    for col_num, value in enumerate(df6.columns.values):
        worksheet7.write(0, col_num, value, header_format)

    # 8th Excel sheet
    worksheet8 = writer.sheets["Deletion by Journal"]

    worksheet8.set_row(0, 30)
    worksheet8.freeze_panes(1, 1)
    worksheet8.set_zoom(80)
    worksheet8.set_column(0, 12, 12)

    for col_num, value in enumerate(df7.columns.values):
        worksheet8.write(0, col_num, value, header_format)

    # 9th Excel sheet
    worksheet9 = writer.sheets["Deletion by PE"]

    worksheet9.set_row(0, 30)
    worksheet9.freeze_panes(1, 1)
    worksheet9.set_zoom(80)
    worksheet9.set_column(0, 12, 12)

    for col_num, value in enumerate(df8.columns.values):
        worksheet9.write(0, col_num, value, header_format)

    # delete useless files and save

    writer.save()

    delete_excel_files(login, master_dict_for_error_trends[12])


def create_del_doi_bar_chart(login, orl_deleted):
    """
            Args:
                login: login code
            """

    df1 = create_groupby_df(orl_deleted, ["Year", "Month"]).agg({"Old DOI": "count"}).rename(
        columns=master_dict_for_error_trends[0]).replace({"Month": master_dict_for_error_trends[14]})
    df1["Check month"] = df1["Month"].astype(str) + " " + df1["Year"].astype(str)

    plt.figure(figsize=(20, 10))
    plt.bar(df1["Check month"], df1["Total DOIs"],
            color=master_dict_for_error_trends[7])

    plt.grid(True, axis="y")
    plt.xticks(rotation=90, fontsize=14)
    plt.yticks(fontsize=18)
    plt.ylabel("Total DOIs", fontsize=22)
    plt.xlabel("Check month", fontsize=22)
    plt.title("Monthly deleted DOI deletion rate (Last 12 months)", fontsize=25)

    df1[["Check month", "Total DOIs"]].to_excel(
        to_raw(os.path.join(path, login, desktop, "Total no. of deleted DOIs.xlsx")), index=False)
    plt.savefig(to_raw(os.path.join(path, login, desktop, "bar_chart_1.jpeg")), optimize=True, transparent=True,
                bbox_inches="tight")
    print("DOI deletion rate graph created!")


def create_yes_no_percentage(login, orl_deleted):
    my_df = create_groupby_df(orl_deleted, ["Found through deleted DOI checks"])["Old DOI"].count().rename(
        columns=master_dict_for_error_trends[1])

    create_percentage_as_number(my_df, "Yes/No", "% as number")
    create_percentage_string(my_df, "Yes/No", "%")
    create_percentage_str_number(my_df, "Yes/No", "%", "Yes/No + %")

    df = my_df[["Found through deleted DOI checks", "Yes/No + %"]].rename(columns=master_dict_for_error_trends[9])
    df.to_excel(to_raw(os.path.join(path, login, desktop, "Found through checks.xlsx")), index=False)


def create_yes_no_stacked(login, orl_deleted):
    my_df = create_groupby_df(orl_deleted, ["Year", "Month", "Found through deleted DOI checks"])[
        "Old DOI"].count().rename(columns=master_dict_for_error_trends[1])

    sums_by_month = create_groupby_df(my_df, ["Year", "Month"])["Yes/No"].sum().rename(
        columns=master_dict_for_error_trends[2])

    my_df = my_df.merge(sums_by_month, on=["Year", "Month"], how="outer").replace(
        {"Month": master_dict_for_error_trends[14]})

    my_df["monthly %"] = (100. * my_df["Yes/No"] / my_df["monthly count"]).round(1).astype(str) + "%"
    my_df["% as number"] = (100. * my_df["Yes/No"] / my_df["monthly count"]).round(1)

    my_df_no = my_df.loc[(my_df["% as number"] == 100.0) & (my_df["Found through deleted DOI checks"] == "No")]
    my_df_yes = my_df.loc[(my_df["% as number"] == 100.0) & (my_df["Found through deleted DOI checks"] == "Yes")]

    my_df_empty_corrected_no = my_df_no.append(my_df_no.replace(master_dict_for_error_trends[10]))
    my_df_empty_corrected_yes = my_df_yes.append(my_df_yes.replace(master_dict_for_error_trends[10]))

    overwrite_num_cols(my_df_empty_corrected_no, "Yes/No", "% as number")
    overwrite_num_cols(my_df_empty_corrected_yes, "Yes/No", "% as number")
    overwrite_num_cols(my_df_empty_corrected_no, "monthly count", "% as number")
    overwrite_num_cols(my_df_empty_corrected_yes, "monthly count", "% as number")

    my_df_empty_corrected = my_df_empty_corrected_no.append(my_df_empty_corrected_yes)

    my_df_completed = my_df.append(my_df_empty_corrected)

    my_df_completed["Check month"] = my_df_completed["Month"].astype(str) + " " + my_df_completed["Year"].astype(str)
    my_df_completed = my_df_completed.drop_duplicates(subset=["Year", "Month", "Found through deleted DOI checks"])

    my_df_completed["Check month_date"] = pd.to_datetime(
        my_df_completed["Month"].astype(str) + "-" + my_df_completed["Year"].astype(str))

    my_df_completed = my_df_completed.sort_values(by=["Check month_date", "Found through deleted DOI checks"])

    labels = convert(my_df_completed["Check month"].unique().tolist())

    df = my_df_completed.pivot(index="Check month_date", columns="Found through deleted DOI checks",
                               values="% as number")

    ax = df.plot.bar(figsize=(20, 8), stacked=True, color=[(0.76, 0.11, 0.16), (0.2, 0.54, 0.75)])

    plt.grid(True, axis="y")
    plt.xticks(fontsize=18)
    ax.set_xticklabels(labels)
    plt.yticks(fontsize=18)
    plt.ylabel("Monthly %", fontsize=22)
    plt.xlabel("Month", fontsize=22)
    plt.title("Deleted DOIs identified through weekly checks?", fontsize=25)
    plt.legend(fontsize=16, loc="upper right")

    plt.savefig(to_raw(os.path.join(path, login, desktop, "bar_chart_2.jpeg")), optimize=True, transparent=True,
                bbox_inches="tight")

    my_df_completed[["Check month", "Found through deleted DOI checks", "% as number"]].to_excel(
        to_raw(os.path.join(path, login, desktop, "Found through checks progress.xlsx")), index=False)
    print("Found through weekly checks graph created!")


def create_old_vs_new_doi(login, orl_aas, orl_optout, orl_deleted):
    under_inv = orl_deleted.loc[
        orl_deleted["Reason for Deletion – Category"] == "Under Investigation", "Old DOI"].to_list()

    approved = orl_aas.loc[(orl_aas["DOI"].isin(under_inv)) & (~orl_aas["Date Approved"].isnull())]
    rejected = orl_aas.loc[(orl_aas["DOI"].isin(under_inv)) & (~orl_aas["Date Rejected"].isnull())]
    not_verified = orl_aas.loc[
        (orl_aas["DOI"].isin(under_inv)) & (orl_aas["Date Approved"].isnull()) & (orl_aas["Date Rejected"].isnull())]
    opt_out = orl_optout[orl_optout["DOI"].isin(under_inv)]

    approved["Old DOI Verification Status"] = "Approved"
    rejected["Old DOI Verification Status"] = "Rejected"
    not_verified["Old DOI Verification Status"] = "Not verified"
    opt_out["Old DOI Verification Status"] = "Opt-Out"

    under_inv_df = approved.append(rejected).append(not_verified).append(opt_out)[master_dict_for_error_trends[16]]

    orl_delete_renamed = orl_deleted.rename(columns=master_dict_for_deleted_dois[1])

    my_df = orl_delete_renamed.merge(under_inv_df, how="outer", on="DOI")

    overwrite_cols(["Old DOI Verification Status"], my_df)

    my_df = my_df[["DOI", "Old DOI Verification Status", "New DOI Verification Status"]]
    my_df = create_groupby_df(my_df, ["Old DOI Verification Status", "New DOI Verification Status"]).agg(
        master_dict_for_error_trends[3]).rename(columns=master_dict_for_error_trends[11])

    df_3 = my_df.pivot(index="Old verification status", columns="New verification status", values="Total DOIs")

    missing_cols = [x for x in master_dict_for_error_trends[13] if x not in list(df_3.columns)]

    for col in missing_cols:
        df_3[col] = 0.0

    for new_col in master_dict_for_error_trends[13]:
        df_3[new_col] = df_3[new_col].fillna(0.0)

    df = df_3[master_dict_for_error_trends[13]]

    df.plot.barh(figsize=(20, 8), stacked=True, color=master_dict_for_error_trends[8])

    plt.grid(True, axis="x")
    plt.xticks(fontsize=18)
    plt.yticks(fontsize=18)
    plt.ylabel("Old verification status", fontsize=22)
    plt.xlabel("New DOI", fontsize=22)
    plt.title("Old vs new DOIs verification statuses (Last 12 months)", fontsize=25)
    plt.legend(fontsize=16)

    plt.savefig(to_raw(os.path.join(path, login, desktop, "bar_chart_3.jpeg")), optimize=True, transparent=True,
                bbox_inches="tight")
    my_df.to_excel(
        to_raw(os.path.join(path, login, desktop, "Old vs. New DOI.xlsx")), index=False)
    print("Old vs. New DOI verification status graph created!")

    return df_3


def create_del_explanations(login, orl_deleted):
    my_df = create_groupby_df(orl_deleted, ["Reason for Deletion – Category"]).agg(
        master_dict_for_error_trends[5]).rename(columns=master_dict_for_error_trends[0])

    create_percentage_as_number(my_df, "Total DOIs", "% as number")
    create_percentage_string(my_df, "Total DOIs", "%")
    create_percentage_str_number(my_df, "Total DOIs", "%", "Total DOIs + %")

    df = my_df[master_dict_for_error_trends[15]].sort_values(by=["% as number"], ascending=False)

    fig1, ax1 = plt.subplots(figsize=(12, 12))

    ax1.pie(df["% as number"].to_list(), explode=(0.125, 0.1, 0.1, 0.1, 0.1, 0.1, 0.1, 0.1, 0.1, 0.1, 0.1, 0.1),
            colors=master_dict_for_error_trends[6], autopct="%1.1f%%", pctdistance=1.1)

    plt.legend(df["Reason for Deletion – Category"].to_list(), fontsize=12, bbox_to_anchor=(1, 0, 0.5, 1))
    plt.title("Deleted DOIs explanations/reasons", fontsize=25)

    plt.savefig(to_raw(os.path.join(path, login, desktop, "pie_chart_1.jpeg")), optimize=True, transparent=True,
                bbox_inches="tight")
    df.rename(columns=master_dict_for_error_trends[4]).sort_values(by="percentage", ascending=False).to_excel(
        to_raw(os.path.join(path, login, desktop, "Reason for deletion.xlsx")), index=False)
    print("Reason for deletion pie chart created!")


def create_table_by_country(login, orl_deleted):
    my_df = create_groupby_df(orl_deleted, ["Agreement", "Reason for Deletion – Category"]).agg(
        master_dict_for_error_trends[5]).rename(columns=master_dict_for_error_trends[0]).pivot(
        index="Agreement", columns="Reason for Deletion – Category", values="Total DOIs").fillna(0.0).astype(int)

    my_df["TOTAL"] = my_df.sum(axis=1)
    my_df = my_df.sort_values(by=["Author mistake in OA request forms", "Author wanted to opt out",
                                  "Change in CA to make Compact-eligible", "TOTAL"], ascending=False)

    df = create_groupby_df(orl_deleted, ["Agreement", "Reason for Deletion – Category"]).agg(
        master_dict_for_error_trends[5]).rename(columns=master_dict_for_error_trends[0]).sort_values(
        by=["Total DOIs", "Agreement"], ascending=False)

    df.to_excel(to_raw(os.path.join(path, login, desktop, "Deletion by Agreement.xlsx")), index=False)

    return my_df


def create_graph_by_journal(login, orl_deleted):
    my_df = create_groupby_df(orl_deleted, ["Journal Title", "Reason for Deletion – Category"]).agg(
        master_dict_for_error_trends[5]).rename(columns=master_dict_for_error_trends[0]).pivot(
        index="Journal Title", columns="Reason for Deletion – Category", values="Total DOIs").fillna(0.0).astype(int)

    my_df["TOTAL"] = my_df.sum(axis=1)
    my_df = my_df.sort_values(by=["TOTAL", "Production mistake"], ascending=False)

    my_df.to_excel(to_raw(os.path.join(path, login, desktop, "Deletion by Journal.xlsx")), index=False)
    print("Deletions by Journal analysis done!")

    return my_df.head(13)


def create_graph_by_pe(login, orl_deleted):
    my_df = create_groupby_df(orl_deleted, ["Production Editor", "Reason for Deletion – Category"]).agg(
        master_dict_for_error_trends[5]).rename(columns=master_dict_for_error_trends[0]).pivot(
        index="Production Editor", columns="Reason for Deletion – Category", values="Total DOIs").fillna(0.0).astype(
        int)

    my_df["TOTAL"] = my_df.sum(axis=1)
    my_df = my_df.sort_values(by=["Production mistake", "TOTAL"], ascending=False)

    my_df.to_excel(to_raw(os.path.join(path, login, desktop, "Deletion by PE.xlsx")), index=False)
    print("Analysis of deleted DOIs error trends split by PE done!")

    return my_df.head(13)
