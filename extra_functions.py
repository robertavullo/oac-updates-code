import pandas as pd
import numpy as np
import os
from bs4 import BeautifulSoup
import requests
from general_variables import path, desktop


def check_missing_cols(filename: pd.DataFrame, col_list: list):
    """Checking missing columns from the input files

    Args:
        filename: Pandas data frame to be checked
        col_list: List of expected columns
    """
    missing_cols = [col for col in col_list if col not in filename.columns]
    if len(missing_cols) > 0:
        print(f"The following columns are missing: {missing_cols}")
    else:
        print("All a-OK")


def find_key(value, in_dict):
    keys = [name for name in in_dict.keys() if in_dict[name] == value]

    if len(keys) == 1:
        return keys[0]
    elif len(keys) == 0:
        return "Not found"
    else:
        print(f"Can't find a unique key for value {value}; instead get {keys}")
    return "No unique ID"


def overwrite_cols(new_col_list: list, df: pd.DataFrame):
    """Overwrite data in the ORL columns that need to be overwritten with s150 data

    Args:
        new_col_list: list of columns that need to be overwritten
        df: merged dataframe with duplicates (_x and _y columns)

    Returns:
        Pandas Data Frame
    """
    for col in new_col_list:
        df[col] = np.where(df[col + "_y"].isnull(), df[col + "_x"], df[col + "_y"])

    return df


def to_raw(string):
    return fr"{string}"


def rename_files(path, login, desktop, timestr, filename_with_ext: str, filename_no_ext: str, file_ext: str):
    os.rename(to_raw(os.path.join(path, login, desktop, filename_with_ext)),
              to_raw(os.path.join(path, login, desktop, filename_no_ext + timestr + file_ext)))


def foalDOI(uai):
    url = "https://deal-interceptor.springernature.app/internal/manuscript-aggregate/" + uai[1:]
    page_response = requests.get(url)
    html_doc = BeautifulSoup(page_response.content, "html.parser")

    if html_doc is None:
        return {"Unique Article ID": uai, "DOI": np.nan}

    elif html_doc.h1.get_text() == "Page not found":
        return {"Unique Article ID": uai, "DOI": np.nan}

    else:
        found_doi = []
        data = html_doc.find_all("td")
        for i in data:
            if "ArticleDoi" in i.get_text():
                found_doi.append(i.get_text()[11:-1])
        if found_doi == []:
            return {"Unique Article ID": uai, "DOI": np.nan}
        else:
            return {"Unique Article ID": uai, "DOI": found_doi[0]}


def convert(list):
    return (*list,)


def create_groupby_df(df: pd.DataFrame, col_list: list):
    result = df.groupby(col_list, as_index=False)

    return result


def overwrite_num_cols(df, col1, col2):
    df[col1] = np.where(df[col2] == 0.0,
                        0,
                        df[col1]
                        )
    return df


def delete_excel_files(login, filename_list):
    for name in filename_list:
        os.remove(to_raw(os.path.join(path, login, desktop, name + ".xlsx")))


def create_percentage_as_number(df: pd.DataFrame, origin_col: str, result_col: str):
    df[result_col] = (100. * df[origin_col] / df[origin_col].sum()).round(1)
    return df[result_col]


def create_percentage_string(df: pd.DataFrame, origin_col: str, result_col: str):
    df[result_col] = (100. * df[origin_col] / df[origin_col].sum()).round(1).astype(str) + "%"
    return df[result_col]


def create_percentage_str_number(df: pd.DataFrame, origin_col: str, percent_col: str, result_col: str):
    df[result_col] = df[origin_col].astype(str) + " (" + df[percent_col] + ")"
    return df[result_col]


def create_excel_file(path, login, desktop, timestr, df: pd.DataFrame, filename: str, file_ext: str):
    df.to_excel(to_raw(os.path.join(path, login, desktop, filename + timestr + file_ext)), index=False)
