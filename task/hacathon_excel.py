# from pprint import pprint
# from openpyxl import Workbook, load_workbook
from datetime import datetime
# from openpyxl.worksheet.table import Table  # , TableStyleInfo
from google_sheets import get_from_google_sheet
import os
import pandas as pd
# from dataclasses import dataclass


today = datetime.today().strftime("%d.%m.%Y %H:%M:%S")


def check_unique_numbers(df_old: pd.DataFrame, df_new: pd.DataFrame):
    new_unique_numbers: set = set(df_new["Уникальный номер размещения"]). \
        symmetric_difference(set(df_old["Уникальный номер размещения"]))
    if len(new_unique_numbers) > 0:
        for new_number in sorted(list(new_unique_numbers)):
            df_ = df_new[df_new['Уникальный номер размещения'] == new_number]
            df_ = df_.drop('Месяц учета оказания услуг', axis=1).rename(
                columns={'Дата учета оказания услуг': today})
            date = pd.read_excel('date.xlsx')
            date = pd.concat([date, df_], axis=0)
            date.to_excel("date.xlsx")

            month = pd.read_excel('month.xlsx')
            month = pd.concat([month, df_], axis=0)
            month.to_excel("month.xlsx")
        return True
    return False


def check_dates(df_old: pd.DataFrame, df_new: pd.DataFrame):
    idx = df_old.shape[0]
    are_dates_equal = df_new[:idx]["Дата учета оказания услуг"] != df_old["Дата учета оказания услуг"]
    new_dates = df_new[:idx][are_dates_equal]
    new_dates = new_dates["Дата учета оказания услуг"].rename(today)
    date = pd.read_excel('date.xlsx')
    date.loc[:idx-1, today] = new_dates
    date.to_excel("date.xlsx")


def check_months(df_old: pd.DataFrame, df_new: pd.DataFrame):
    idx = df_old.shape[0]
    are_months_equal = df_new[:idx]["Месяц учета оказания услуг"] != df_old["Месяц учета оказания услуг"]
    new_months = df_new[:idx][are_months_equal]
    new_months = new_months["Месяц учета оказания услуг"].rename(today)
    month = pd.read_excel('month.xlsx')
    month.loc[:idx-1, today] = new_months
    month.to_excel("date.xlsx")


def check_values_on_changes(df_new: pd.DataFrame):
    df_old = pd.read_excel('old_table.xlsx')
    check_unique_numbers(df_old, df_new)
    check_dates(df_old, df_new)
    check_months(df_old, df_new)


if __name__ == "__main__":
    curr_table: dict = get_from_google_sheet(cred_file_name='creds.json',
                                             sheet_id='1DNKTyIuRqVPm4vsgMpkDkRwlQo8LVBgOO7cmtCGjIhY')
    columns = ['ФИО/Название\nподрядчика', 'Уникальный номер размещения',
               'Дата учета оказания услуг', 'Месяц учета оказания услуг']
    df = pd.concat([pd.Series(name=column, data=curr_table.get(column)) for column in columns],
                   axis=1)
    if os.path.isfile('old_table.xlsx'):
        check_values_on_changes(df)
        df.to_excel("old_table.xlsx")
    else:
        df.to_excel("old_table.xlsx")
        df.drop('Дата учета оказания услуг', axis=1).rename(columns={'Месяц учета оказания услуг': today}).to_excel(
            "month.xlsx")
        df.drop('Месяц учета оказания услуг', axis=1).rename(columns={'Дата учета оказания услуг': today}).to_excel(
            "date.xlsx")
