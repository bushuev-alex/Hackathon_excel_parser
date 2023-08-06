from datetime import datetime
from google_sheets import get_from_google_sheet
import os
import pandas as pd
from settings import SHEET_ID, DATE_ACCURACY, ENGINE


class :
    today = datetime.today().strftime(DATE_ACCURACY)

    def __init__(self):
        self.sheet_id = SHEET_ID
        self.table_name = "Table.xlsx"
        self.engine = ENGINE

    def rename_and_write_rows(self, data: pd.DataFrame, type_: str, sheet: str):
        df_cleared = data.drop(f'{type_} учета оказания услуг', axis=1).rename(
            columns={f"{'Месяц' if type_ == 'Дата' else 'Дата'} учета оказания услуг": self.today})
        df_from_sheet = pd.read_excel(f"{self.table_name}", engine=self.engine, sheet_name=sheet, index_col=0)
        df_to_write = pd.concat([df_from_sheet, df_cleared], axis=0)
        with pd.ExcelWriter(f"{self.table_name}", engine='openpyxl', mode='a', if_sheet_exists='replace') as writer_:
            df_to_write.to_excel(writer_, sheet_name=sheet)

    def check_unique_numbers(self, df_old: pd.DataFrame, df_new: pd.DataFrame) -> None:
        new_unique_numbers: set = set(df_new["Уникальный номер размещения"]). \
            symmetric_difference(set(df_old["Уникальный номер размещения"]))
        if len(new_unique_numbers) > 0:
            for new_number in sorted(list(new_unique_numbers)):
                df_ = df_new[df_new['Уникальный номер размещения'] == new_number]
                self.rename_and_write_rows(df_, 'Дата', 'month')
                self.rename_and_write_rows(df_, 'Месяц', 'date')

    def check_dates(self, df_old: pd.DataFrame, df_new: pd.DataFrame) -> None:
        column = "Дата учета оказания услуг"
        idx = df_old.shape[0]
        are_dates_equal = df_new[column][:idx] != df_old[column]  # True or False
        new_dates = df_new[:idx][are_dates_equal]  # from df_new select values which are not equal to old (with index)
        new_dates = new_dates[column].rename(self.today)  # rename column "Дата учета оказания услуг"
        self.write_values(values=new_dates, sheet='date', idx=idx)

    def check_months(self, df_old: pd.DataFrame, df_new: pd.DataFrame) -> None:
        column = "Месяц учета оказания услуг"
        idx = df_old.shape[0]
        are_months_equal = df_new[column][:idx] != df_old[column]  # True or False
        new_months = df_new[:idx][are_months_equal]  # from df_new select values which are not equal to old (with index)
        new_months = new_months[column].rename(self.today)  # rename column "Месяц учета оказания услуг"
        self.write_values(values=new_months, sheet='month', idx=idx)

    def write_values(self, values: pd.Series, sheet: str, idx: int) -> None:
        df_ = pd.read_excel(f"{self.table_name}", engine='openpyxl', sheet_name=f"{sheet}", index_col=0)
        df_.loc[:idx - 1, self.today] = values
        with pd.ExcelWriter(f"{self.table_name}", engine='openpyxl', mode='a', if_sheet_exists='replace') as writer_:
            df_.to_excel(writer_, sheet_name=f"{sheet}")

    def check_values_on_changes(self, df_new: pd.DataFrame) -> None:
        df_old = pd.read_excel("Table_old.xlsx", engine=ENGINE, index_col=0)
        self.check_unique_numbers(df_old, df_new)
        self.check_dates(df_old, df_new)
        self.check_months(df_old, df_new)

    def get_google_data(self) -> pd.DataFrame:
        curr_table: dict = get_from_google_sheet(cred_file_name='creds.json',
                                                 sheet_id=self.sheet_id)
        columns = ['ФИО/Название\nподрядчика', 'Уникальный номер размещения',
                   'Дата учета оказания услуг', 'Месяц учета оказания услуг']
        df_ = pd.concat([pd.Series(name=column, data=curr_table.get(column)) for column in columns],
                        axis=1)
        return df_

    def write_first_table_to_compare(self, df_: pd.DataFrame) -> None:
        df_.to_excel("Table_old.xlsx", engine=ENGINE)
        with pd.ExcelWriter(f"{self.table_name}", engine='openpyxl') as writer:
            df_.drop('Месяц учета оказания услуг', axis=1).rename(
                columns={'Дата учета оказания услуг': self.today}).to_excel(writer, sheet_name="date")
            df_.drop('Дата учета оказания услуг', axis=1).rename(
                columns={'Месяц учета оказания услуг': self.today}).to_excel(writer, sheet_name="month")


def main():
    parser = ExcelParser()
    df = parser.get_google_data()

    if os.path.isfile('Table_old.xlsx'):
        parser.check_values_on_changes(df)
        df.to_excel("Table_old.xlsx", engine=ENGINE)
    else:
        parser.write_first_table_to_compare(df)


if __name__ == "__main__":
    main()
