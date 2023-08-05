# from pprint import pprint
# from openpyxl import Workbook, load_workbook
from datetime import datetime
# from openpyxl.worksheet.table import Table  # , TableStyleInfo
from google_sheets import get_from_google_sheet
import os
import pandas as pd
from settings import SHEET_ID
# from dataclasses import dataclass


today = datetime.today().strftime("%d.%m.%Y %H:%M:%S")


def check_unique_numbers(df_old: pd.DataFrame, df_new: pd.DataFrame):
    new_unique_numbers: set = set(df_new["Уникальный номер размещения"]). \
        symmetric_difference(set(df_old["Уникальный номер размещения"]))
    if len(new_unique_numbers) > 0:
        for new_number in sorted(list(new_unique_numbers)):
            df_ = df_new[df_new['Уникальный номер размещения'] == new_number]
            df_dates = df_.drop('Месяц учета оказания услуг', axis=1).rename(
                columns={'Дата учета оказания услуг': today})
            df_months = df_.drop('Дата учета оказания услуг', axis=1).rename(
                columns={'Месяц учета оказания услуг': today})
            date = pd.read_excel('Table.xlsx', engine='openpyxl', sheet_name='date', index_col=0)
            month = pd.read_excel('Table.xlsx', engine='openpyxl', sheet_name='month', index_col=0)

            date = pd.concat([date, df_dates], axis=0)
            month = pd.concat([month, df_months], axis=0)

            with pd.ExcelWriter("Table.xlsx", engine='openpyxl', mode='a', if_sheet_exists='replace') as writer_:
                date.to_excel(writer_, sheet_name="date")
                month.to_excel(writer_, sheet_name="month")
        return True
    return False


def check_dates(df_old: pd.DataFrame, df_new: pd.DataFrame):
    column = "Дата учета оказания услуг"
    idx = df_old.shape[0]
    are_dates_equal = df_new[column][:idx] != df_old[column]  # True or False
    new_dates = df_new[:idx][are_dates_equal]  # from df_new select values which are not equal to old (with index)
    new_dates = new_dates[column].rename(today)  # rename column "Дата учета оказания услуг"
    write_values(values=new_dates, sheet='date', idx=idx)


def check_months(df_old: pd.DataFrame, df_new: pd.DataFrame):
    column = "Месяц учета оказания услуг"
    idx = df_old.shape[0]
    are_months_equal = df_new[column][:idx] != df_old[column]  # True or False
    new_months = df_new[:idx][are_months_equal]  # from df_new select values which are not equal to old (with index)
    new_months = new_months[column].rename(today)  # rename column "Месяц учета оказания услуг"
    write_values(values=new_months, sheet='month', idx=idx)


def write_values(values, sheet: str, idx: int):
    df_ = pd.read_excel("Table.xlsx", engine='openpyxl', sheet_name=f"{sheet}", index_col=0)
    df_.loc[:idx - 1, today] = values
    with pd.ExcelWriter("Table.xlsx", engine='openpyxl', mode='a', if_sheet_exists='replace') as writer_:
        df_.to_excel(writer_, sheet_name=f"{sheet}")


def check_values_on_changes(df_new: pd.DataFrame):
    df_old = pd.read_excel("Table_old.xlsx", engine='openpyxl', index_col=0)
    check_unique_numbers(df_old, df_new)
    check_dates(df_old, df_new)
    check_months(df_old, df_new)


if __name__ == "__main__":
    curr_table: dict = get_from_google_sheet(cred_file_name='creds.json',
                                             sheet_id=SHEET_ID)
    columns = ['ФИО/Название\nподрядчика', 'Уникальный номер размещения',
               'Дата учета оказания услуг', 'Месяц учета оказания услуг']
    df = pd.concat([pd.Series(name=column, data=curr_table.get(column)) for column in columns],
                   axis=1)
    if os.path.isfile('Table_old.xlsx'):
        check_values_on_changes(df)
        df.to_excel("Table_old.xlsx", engine='openpyxl')
    else:
        df.to_excel("Table_old.xlsx", engine='openpyxl')
        with pd.ExcelWriter("Table.xlsx", engine='openpyxl') as writer:
            df.drop('Месяц учета оказания услуг', axis=1).rename(columns={'Дата учета оказания услуг': today}).to_excel(
                writer, sheet_name="date")
            df.drop('Дата учета оказания услуг', axis=1).rename(columns={'Месяц учета оказания услуг': today}).to_excel(
                writer, sheet_name="month")
