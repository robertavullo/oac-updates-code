from datetime import datetime
import time

timestr = time.strftime("_%d.%m.%Y")
path = "C:\\Users"
desktop = "Desktop"
institution_list_hybrid = "QAchecks_files\\institution_list_hybrid.xlsx"
institution_list_fullyOA = "QAchecks_files\\institution_list_fullyOA.xlsx"
journal_list_file = "QAchecks_files\\Journal_lists.xlsx"
current_month = datetime.now().month
current_year = datetime.now().year

quarter_dict = {1: "Q1",
                2: "Q1",
                3: "Q1",
                4: "Q2",
                5: "Q2",
                6: "Q2",
                7: "Q3",
                8: "Q3",
                9: "Q3",
                10: "Q4",
                11: "Q4",
                12: "Q4"
                }

pie_colours = ("#E60505", "#FF82C3", "#D778FF", "#915AFF", "#5A96FF", "#00CDFF", "#11E5A1", "#1ED241", "#A0FF00",
               "#FFFF00", "#FFDC00", "#FFAA00"
               )

simple_bar_chart_colours = ("#9AAFBE", "#9AAFBE", "#9AAFBE",
                            "#348ABE", "#348ABE", "#348ABE",
                            "#C21D29", "#C21D29", "#C21D29",
                            "#142F39", "#142F39", "#142F39"
                            )