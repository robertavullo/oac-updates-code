import pandas as pd
from extra_functions import to_raw
import os
from general_variables import *
from master_dicts_of_lists import read_me_hybrid, read_me_foa


def format_date(df: pd.DataFrame, col_list: list):
    """Format certain columns in DataFrame to datetime and fill the empty cells with a blank string"""
    for col in col_list:
        df[col] = pd.to_datetime(df[col]).fillna("")


def format_approved_this_year(input_worksheet, df1: pd.DataFrame, header_format, normaldate, currency_usd):
    """
    formats the first sheet of the newly-created hybrids article list

    Args:
        df1: Pandas Data Frame containing the non-formatted hybrid articles list (but already updated with today's data)
        input_worksheet: worksheet1
    """

    input_worksheet.set_tab_color("#92D050")

    for i in range(2, len(df1) + 2):
        input_worksheet.data_validation(f"V{i}", {"validate": "list",
                                                  "source": ["N/A", "Yes", "No"]})

    input_worksheet.freeze_panes(1, 1)
    input_worksheet.set_zoom(80)
    input_worksheet.set_row(0, 30)
    input_worksheet.set_column(0, 0, 30)
    input_worksheet.set_column(1, 2, 20)
    input_worksheet.set_column(3, 3, 12)
    input_worksheet.set_column(4, 4, 20)
    input_worksheet.set_column(5, 7, 15)
    input_worksheet.set_column(8, 10, 20)
    input_worksheet.set_column(11, 13, 15)
    input_worksheet.set_column(14, 14, 30)
    input_worksheet.set_column(15, 15, 20)
    input_worksheet.set_column(16, 19, 12)
    input_worksheet.set_column(20, 24, 15)
    input_worksheet.set_column(25, 25, 30)

    for col_num, value in enumerate(df1.columns.values):
        input_worksheet.write(0, col_num, value, header_format)

    for row_num, data in enumerate(df1["Date Sent to AAS"]):
        input_worksheet.write(row_num + 1, 11, data, normaldate)

    for row_num, data in enumerate(df1["Date Approved"]):
        input_worksheet.write(row_num + 1, 12, data, normaldate)

    for row_num, data in enumerate(df1["Date Forwarded"]):
        input_worksheet.write(row_num + 1, 13, data, normaldate)

    for row_num, data in enumerate(df1["Final Reminder to be Sent"]):
        input_worksheet.write(row_num + 1, 20, data, normaldate)

    for row_num, data in enumerate(df1["Author Share"]):
        input_worksheet.write(row_num + 1, 22, data, currency_usd)

    for row_num, data in enumerate(df1["Institutional Share"]):
        input_worksheet.write(row_num + 1, 23, data, currency_usd)

    input_worksheet.autofilter("A1:Z1")


def format_approved_previously(input_worksheet, df2: pd.DataFrame, header_format, normaldate, currency_usd):
    """
    formats the second sheet of the hybrid articles list

    Args:
        df2: Pandas Data Frame containing the non-formatted hybrid articles list (but already updated with today's data)
        input_worksheet: worksheet2
    """

    input_worksheet.set_tab_color("#00B050")

    for i in range(2, len(df2) + 2):
        input_worksheet.data_validation(f"V{i}", {"validate": "list",
                                                  "source": ["N/A", "Yes", "No"]})

    input_worksheet.freeze_panes(1, 1)
    input_worksheet.set_zoom(80)
    input_worksheet.set_row(0, 30)
    input_worksheet.set_column(0, 0, 30)
    input_worksheet.set_column(1, 2, 20)
    input_worksheet.set_column(3, 3, 12)
    input_worksheet.set_column(4, 4, 20)
    input_worksheet.set_column(5, 7, 15)
    input_worksheet.set_column(8, 10, 20)
    input_worksheet.set_column(11, 13, 15)
    input_worksheet.set_column(14, 14, 30)
    input_worksheet.set_column(15, 15, 20)
    input_worksheet.set_column(16, 19, 12)
    input_worksheet.set_column(20, 24, 15)
    input_worksheet.set_column(25, 25, 30)

    for col_num, value in enumerate(df2.columns.values):
        input_worksheet.write(0, col_num, value, header_format)

    for row_num, data in enumerate(df2["Date Sent to AAS"]):
        input_worksheet.write(row_num + 1, 11, data, normaldate)

    for row_num, data in enumerate(df2["Date Approved"]):
        input_worksheet.write(row_num + 1, 12, data, normaldate)

    for row_num, data in enumerate(df2["Date Forwarded"]):
        input_worksheet.write(row_num + 1, 13, data, normaldate)

    for row_num, data in enumerate(df2["Final Reminder to be Sent"]):
        input_worksheet.write(row_num + 1, 20, data, normaldate)

    for row_num, data in enumerate(df2["Author Share"]):
        input_worksheet.write(row_num + 1, 22, data, currency_usd)

    for row_num, data in enumerate(df2["Institutional Share"]):
        input_worksheet.write(row_num + 1, 23, data, currency_usd)

    input_worksheet.autofilter("A1:Z1")


def format_rejected(input_worksheet, df3: pd.DataFrame, header_format, normaldate, currency_usd):
    """
    formats the third sheet of the hybrid articles list

    Args:
        df3: Pandas Data Frame containing the non-formatted hybrid articles list (but already updated with today's data)
        input_worksheet: worksheet3
    """

    input_worksheet.set_tab_color("#FF0000")

    for i in range(2, len(df3) + 2):
        input_worksheet.data_validation(f"V{i}", {"validate": "list",
                                                  "source": ["N/A", "Yes", "No"]})

    input_worksheet.freeze_panes(1, 1)
    input_worksheet.set_zoom(80)
    input_worksheet.set_row(0, 30)
    input_worksheet.set_column(0, 0, 30)
    input_worksheet.set_column(1, 2, 20)
    input_worksheet.set_column(3, 3, 12)
    input_worksheet.set_column(4, 4, 20)
    input_worksheet.set_column(5, 7, 15)
    input_worksheet.set_column(8, 10, 20)
    input_worksheet.set_column(11, 13, 15)
    input_worksheet.set_column(14, 14, 30)
    input_worksheet.set_column(15, 15, 20)
    input_worksheet.set_column(16, 19, 12)
    input_worksheet.set_column(20, 24, 15)
    input_worksheet.set_column(25, 25, 30)

    for col_num, value in enumerate(df3.columns.values):
        input_worksheet.write(0, col_num, value, header_format)

    for row_num, data in enumerate(df3["Date Sent to AAS"]):
        input_worksheet.write(row_num + 1, 11, data, normaldate)

    for row_num, data in enumerate(df3["Date Rejected"]):
        input_worksheet.write(row_num + 1, 12, data, normaldate)

    for row_num, data in enumerate(df3["Date Forwarded"]):
        input_worksheet.write(row_num + 1, 13, data, normaldate)

    for row_num, data in enumerate(df3["Final Reminder to be Sent"]):
        input_worksheet.write(row_num + 1, 20, data, normaldate)

    for row_num, data in enumerate(df3["Author Share"]):
        input_worksheet.write(row_num + 1, 22, data, currency_usd)

    for row_num, data in enumerate(df3["Institutional Share"]):
        input_worksheet.write(row_num + 1, 23, data, currency_usd)

    input_worksheet.autofilter("A1:Z1")


def format_unverified(input_worksheet, df4: pd.DataFrame, header_format, normaldate, currency_usd):
    """
    formats the fourth sheet of the hybrid articles list

    Args:
        df4: Pandas Data Frame containing the non-formatted hybrid articles list (but already updated with today's data)
        input_worksheet: worksheet4
    """

    input_worksheet.set_tab_color("#BFBFBF")

    for i in range(2, len(df4) + 2):
        input_worksheet.data_validation(f"U{i}", {"validate": "list",
                                                  "source": ["N/A", "Yes", "No"]})

    input_worksheet.freeze_panes(1, 1)
    input_worksheet.set_zoom(80)
    input_worksheet.set_row(0, 30)
    input_worksheet.set_column(0, 0, 30)
    input_worksheet.set_column(1, 2, 20)
    input_worksheet.set_column(3, 3, 12)
    input_worksheet.set_column(4, 4, 20)
    input_worksheet.set_column(5, 7, 15)
    input_worksheet.set_column(8, 10, 20)
    input_worksheet.set_column(11, 12, 15)
    input_worksheet.set_column(13, 13, 30)
    input_worksheet.set_column(14, 14, 20)
    input_worksheet.set_column(15, 18, 12)
    input_worksheet.set_column(19, 23, 15)
    input_worksheet.set_column(24, 24, 30)

    for col_num, value in enumerate(df4.columns.values):
        input_worksheet.write(0, col_num, value, header_format)

    for row_num, data in enumerate(df4["Date Sent to AAS"]):
        input_worksheet.write(row_num + 1, 11, data, normaldate)

    for row_num, data in enumerate(df4["Date Forwarded"]):
        input_worksheet.write(row_num + 1, 12, data, normaldate)

    for row_num, data in enumerate(df4["Final Reminder to be Sent"]):
        input_worksheet.write(row_num + 1, 19, data, normaldate)

    for row_num, data in enumerate(df4["Author Share"]):
        input_worksheet.write(row_num + 1, 21, data, currency_usd)

    for row_num, data in enumerate(df4["Institutional Share"]):
        input_worksheet.write(row_num + 1, 22, data, currency_usd)

    input_worksheet.autofilter("A1:Y1")


def format_optout(input_worksheet, df5: pd.DataFrame, header_format, normaldate):
    """
        formats the fifth sheet of the hybrid articles list (opt-out sheet)

        Args:
            df5: Pandas Data Frame containing the non-formatted hybrid articles list (but already updated with today's data)
            input_worksheet: worksheet5
        """

    input_worksheet.set_tab_color("#00B0F0")
    input_worksheet.freeze_panes(1, 1)
    input_worksheet.set_zoom(80)
    input_worksheet.set_row(0, 30)
    input_worksheet.set_column(0, 0, 30)
    input_worksheet.set_column(1, 2, 20)
    input_worksheet.set_column(3, 3, 12)
    input_worksheet.set_column(4, 4, 20)
    input_worksheet.set_column(5, 7, 15)
    input_worksheet.set_column(8, 10, 20)
    input_worksheet.set_column(11, 11, 15)
    input_worksheet.set_column(12, 12, 30)
    input_worksheet.set_column(13, 16, 12)
    input_worksheet.set_column(17, 17, 15)
    input_worksheet.set_column(18, 18, 30)

    for col_num, value in enumerate(df5.columns.values):
        input_worksheet.write(0, col_num, value, header_format)

    for row_num, data in enumerate(df5["Eligibility Completed"]):
        input_worksheet.write(row_num + 1, 11, data, normaldate)

    input_worksheet.autofilter("A1:S1")


def format_cancel(input_worksheet, df6: pd.DataFrame, header_format, normaldate):
    """
        formats the last sheet of the hybrid articles sheet (deleted/cancelled DOIs and do not report)

        Args:
            df6: Pandas Data Frame containing the non-formatted hybrid articles list (deleted dois sheet)
            input_worksheet: worksheet5
        """

    input_worksheet.set_tab_color("#FFC000")

    for i in range(2, len(df6) + 2):
        input_worksheet.data_validation(f"G{i}", {"validate": "list",
                                                  "source": ["N/A", "AM verified by mistake",
                                                             "Author mistake in OA request forms",
                                                             "Author wanted to opt out",
                                                             "CDL - opt-out due to funding issues",
                                                             "Change in CA to make article eligible", "Other",
                                                             "Production mistake", "Revised manuscript", "Tech issue",
                                                             "Under investigation", "Withdrawn from pub."]})
        input_worksheet.data_validation(f"C{i}", {"validate": "list",
                                                  "source": ["Approved", "Rejected", "Not verified", "Opt-Out"]})
        input_worksheet.data_validation(f"J{i}", {"validate": "list",
                                                  "source": ["Approved", "Rejected", "Not verified", "Opt-Out",
                                                             "Ineligible for TAs",
                                                             "No new DOI", "Waiting for Author to Complete OASiS"]})

    input_worksheet.freeze_panes(1, 2)
    input_worksheet.set_zoom(80)
    input_worksheet.set_row(0, 30)
    input_worksheet.set_column(0, 0, 30)
    input_worksheet.set_column(1, 1, 20)
    input_worksheet.set_column(2, 4, 15)
    input_worksheet.set_column(5, 8, 30)
    input_worksheet.set_column(9, 11, 15)
    input_worksheet.set_column(12, 12, 30)

    for col_num, value in enumerate(df6.columns.values):
        input_worksheet.write(0, col_num, value, header_format)

    for row_num, data in enumerate(df6["Date"]):
        input_worksheet.write(row_num + 1, 11, data, normaldate)

    input_worksheet.autofilter("A1:M1")


def format_foa_aas(input_worksheet, df1: pd.DataFrame, header_format, normaldate, currency_usd, green, pink, blue):
    """
    formats the first sheet of the newly-created foal

    Args:
        df1: Pandas Data Frame containing the non-formatted foal (but already updated with today's data)
        input_worksheet: worksheet1
    """
    input_worksheet.set_tab_color("#92D050")

    input_worksheet.freeze_panes(1, 3)
    input_worksheet.set_zoom(80)
    input_worksheet.set_row(0, 30)
    input_worksheet.set_column(0, 2, 20)
    input_worksheet.set_column(3, 3, 15)
    input_worksheet.set_column(4, 4, 18)
    input_worksheet.set_column(5, 15, 15)
    input_worksheet.set_column(16, 16, 30)
    input_worksheet.set_column(17, 29, 15)
    input_worksheet.set_column(30, 37, 18)
    input_worksheet.set_column(38, 38, 24)

    for col_num, value in enumerate(df1.columns.values):
        input_worksheet.write(0, col_num, value, header_format)

    for row_num, data in enumerate(df1["Date Sent to AAS"]):
        input_worksheet.write(row_num + 1, 12, data, normaldate)

    for row_num, data in enumerate(df1["Date Approved"]):
        input_worksheet.write(row_num + 1, 13, data, green)

    for row_num, data in enumerate(df1["Date Rejected"]):
        input_worksheet.write(row_num + 1, 14, data, pink)

    for row_num, data in enumerate(df1["Date Forwarded"]):
        input_worksheet.write(row_num + 1, 15, data, blue)

    for row_num, data in enumerate(df1["Final Reminder to be Sent"]):
        input_worksheet.write(row_num + 1, 33, data, normaldate)

    for row_num, data in enumerate(df1["Author Share"]):
        input_worksheet.write(row_num + 1, 35, data, currency_usd)

    for row_num, data in enumerate(df1["Institutional Share"]):
        input_worksheet.write(row_num + 1, 36, data, currency_usd)

    input_worksheet.autofilter("A1:AM1")


def format_hybrid(path, login, desktop, appr_this_year, appr_previously, rejections, unverified,
                  optout, cancellations, appr_time_cols, rej_time_cols, unver_time_cols, optout_time_cols,
                  dnr_time_cols, current_year):
    """formats all sheets of the hybrid article list by calling previous functions and defining formats

    Args:
        path, login, desktop, current_year, quarter_dict, current_month: imported from general_variables
        col_times_aas, col_times_opt, col_times_del: columns with datetimes in the whole file
        appr_this_year: first sheet
        appr_previously: second sheet
        rejections: third sheet
        unverified: fourth sheet
        optout: fifth sheet
        cancellations: last sheet

    """

    format_date(appr_this_year, appr_time_cols)
    format_date(appr_previously, appr_time_cols)
    format_date(rejections, rej_time_cols)
    format_date(unverified, unver_time_cols)
    format_date(optout, optout_time_cols)
    format_date(cancellations, dnr_time_cols)

    filepath_with_r = to_raw(os.path.join(path, login, desktop, "Hybrid Journals TA Article List " + str(current_year) +".xlsx"))

    writer = pd.ExcelWriter(filepath_with_r, engine="xlsxwriter")

    read_me_hybrid.to_excel(writer, sheet_name="ReadMe", index=False)
    appr_this_year.to_excel(writer, sheet_name="Approved", index=False)
    appr_previously.to_excel(writer, sheet_name="Approved (prev. years)", index=False)
    rejections.to_excel(writer, sheet_name="Rejected", index=False)
    unverified.to_excel(writer, sheet_name="In Verification", index=False)
    optout.to_excel(writer, sheet_name="Opt-Outs", index=False)
    cancellations["New DOI"] = cancellations["New DOI"].fillna("N/A")
    cancellations.to_excel(writer, sheet_name="Do NOT Report", index=False)

    workbook = writer.book

    header_format = workbook.add_format({"align": "left",
                                         "text_wrap": True,
                                         "valign": "top",
                                         "bold": True})

    merge_format = workbook.add_format({"valign": "center",
                                        "bold": True,
                                        "font_size": 21})

    springer_blue = workbook.add_format({"valign": "center",
                                         "bold": True,
                                         "font_size": 21,
                                         "font_color": "#1F497D"})

    nature_red = workbook.add_format({"valign": "center",
                                      "bold": True,
                                      "font_size": 21,
                                      "font_color": "#C00000"})

    normaldate = workbook.add_format({"num_format": "dd/mm/yyyy",
                                      "align": "right"})

    currency_usd = workbook.add_format({"num_format": "$ #,###.00",
                                        "align": "right"})

    wrap_text = workbook.add_format({"text_wrap": True,
                                     "valign": "top"
                                     })

    worksheet0 = writer.sheets["ReadMe"]
    worksheet1 = writer.sheets["Approved"]
    worksheet2 = writer.sheets["Approved (prev. years)"]
    worksheet3 = writer.sheets["Rejected"]
    worksheet4 = writer.sheets["In Verification"]
    worksheet5 = writer.sheets["Opt-Outs"]
    worksheet6 = writer.sheets["Do NOT Report"]

    # sorting out the formatting of the ReadMe tab

    worksheet0.set_zoom(85)
    worksheet0.set_row(0, 30)
    worksheet0.set_column(0, 0, 15)
    worksheet0.set_column(1, 1, 120)

    for row_num, data in enumerate(read_me_hybrid["Springer"]):
        worksheet0.write(row_num + 2, 0, data, header_format)
    for row_num, data in enumerate(read_me_hybrid["Nature"]):
        worksheet0.write(row_num + 2, 1, data, wrap_text)

    worksheet0.merge_range("A1:B1", "SpringerNature", merge_format)
    worksheet0.write_rich_string("A1",
                                 springer_blue, "Springer",
                                 nature_red, "Nature",
                                 merge_format)

    format_approved_this_year(worksheet1, appr_this_year, header_format, normaldate, currency_usd)
    format_approved_previously(worksheet2, appr_previously, header_format, normaldate, currency_usd)
    format_rejected(worksheet3, rejections, header_format, normaldate, currency_usd)
    format_unverified(worksheet4, unverified, header_format, normaldate, currency_usd)
    format_optout(worksheet5, optout, header_format, normaldate)
    format_cancel(worksheet6, cancellations, header_format, normaldate)

    writer.save()


def format_foa(path, login, desktop, foa_aas, col_times_foal, current_year):
    """formats all sheets of the hybrid articles list by calling previous functions and defining formats

    Args:
        path, login, desktop, current_year: imported from general_variables
        foal_aas: first sheet of the foal
        col_times_foal: columns with datetimes in all sheets of the foal
        filepath_with_r: filepath to the foal, including the "r" at the beginning

    """

    format_date(foa_aas, col_times_foal)

    filepath_with_r = to_raw(os.path.join(path, login, desktop, "FOA Journals TA Article List " + str(current_year) + ".xlsx"))

    writer = pd.ExcelWriter(filepath_with_r, engine="xlsxwriter")

    read_me_foa.to_excel(writer, sheet_name="ReadMe", index=False)
    foa_aas.to_excel(writer, sheet_name="Articles in AAS", index=False)

    workbook = writer.book

    header_format = workbook.add_format({"align": "left",
                                         "text_wrap": True,
                                         "valign": "top",
                                         "bold": True})

    merge_format = workbook.add_format({"valign": "center",
                                        "bold": True,
                                        "font_size": 21})

    springer_blue = workbook.add_format({"valign": "center",
                                         "bold": True,
                                         "font_size": 21,
                                         "font_color": "#1F497D"})

    nature_red = workbook.add_format({"valign": "center",
                                      "bold": True,
                                      "font_size": 21,
                                      "font_color": "#C00000"})

    normaldate = workbook.add_format({"num_format": "dd/mm/yyyy",
                                      "align": "right"})

    green = workbook.add_format({"bg_color": "#92D050",
                                 "num_format": "dd/mm/yyyy",
                                 "bold": True,
                                 "align": "right"})

    pink = workbook.add_format({"bg_color": "#E6B8B7",
                                "num_format": "dd/mm/yyyy",
                                "bold": True,
                                "align": "right"})

    blue = workbook.add_format({"bg_color": "#C5D9F1",
                                "num_format": "dd/mm/yyyy",
                                "bold": True,
                                "align": "right"})

    currency_usd = workbook.add_format({"num_format": "$ #,###.00",
                                        "align": "right"})

    wrap_text = workbook.add_format({"text_wrap": True,
                                     "valign": "top"
                                     })

    worksheet0 = writer.sheets["ReadMe"]
    worksheet1 = writer.sheets["Articles in AAS"]

    worksheet0.set_zoom(85)
    worksheet0.set_row(0, 30)
    worksheet0.set_column(0, 0, 15)
    worksheet0.set_column(1, 1, 120)

    for row_num, data in enumerate(read_me_foa["Springer"]):
        worksheet0.write(row_num + 2, 0, data, header_format)
    for row_num, data in enumerate(read_me_foa["Nature"]):
        worksheet0.write(row_num + 2, 1, data, wrap_text)

    worksheet0.merge_range("A1:B1", "SpringerNature", merge_format)
    worksheet0.write_rich_string("A1",
                                 springer_blue, "Springer",
                                 nature_red, "Nature",
                                 merge_format)

    format_foa_aas(worksheet1, foa_aas, header_format, normaldate, currency_usd, green, pink, blue)

    writer.save()
