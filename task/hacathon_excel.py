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
    column = "Дата учета оказания услуг"
    idx = df_old.shape[0]
    are_dates_equal = df_new[column][:idx] != df_old[column]  # True or False
    new_dates = df_new[are_dates_equal][:idx]  # from df_new select values which are not equal to old (with index)
    new_dates = new_dates[column].rename(today)  # rename column "Дата учета оказания услуг"
    write_values(values=new_dates, file_name='date', idx=idx)


def check_months(df_old: pd.DataFrame, df_new: pd.DataFrame):
    column = "Месяц учета оказания услуг"
    idx = df_old.shape[0]
    are_months_equal = df_new[column][:idx] != df_old[column]  # True or False
    new_months = df_new[are_months_equal][:idx]  # from df_new select values which are not equal to old (with index)
    new_months = new_months[column].rename(today)  # rename column "Месяц учета оказания услуг"
    write_values(values=new_months, file_name='month', idx=idx)


def write_values(values, file_name: str, idx: int):
    df_ = pd.read_excel(f"{file_name}.xlsx")
    df_.loc[:idx - 1, today] = values
    df_.to_excel(f"{file_name}.xlsx")


def check_values_on_changes(df_new: pd.DataFrame):
    df_old = pd.read_excel("old_table.xlsx")
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
