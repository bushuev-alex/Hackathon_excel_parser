from datetime import datetime
from google_sheets import get_from_google_sheet
import os
import pandas as pd
from settings import SHEET_ID, DATE_ACCURACY, ENGINE


class ExcelParser:
    today = datetime.today().strftime(DATE_ACCURACY)

    def __init__(self):
        self.sheet_id = SHEET_ID
        self.table_name = "Table.xlsx"
        self.engine = ENGINE

    def rename_and_write_rows(self, data: pd.DataFrame, type_: str, sheet: str):
        df_cleared = data.drop(type_ + ' учета оказания услуг', axis=1).rename(
            columns={f"{'Месяц' if type_ == 'Дата' else 'Дата'} учета оказания услуг": self.today})
        df_from_sheet = pd.read_excel(self.table_name, engine=self.engine, sheet_name=sheet, index_col=0)
        df_to_write = pd.concat([df_from_sheet, df_cleared], axis=0)
        with pd.ExcelWriter(self.table_name, engine=self.engine, mode='a', if_sheet_exists='replace') as writer_:
            df_to_write.to_excel(writer_, sheet_name=sheet)

    def check_unique_numbers(self, df_old: pd.DataFrame, df_new: pd.DataFrame) -> None:
        new_unique_numbers: set = set(df_new["Уникальный номер размещения"]). \
            symmetric_difference(set(df_old["Уникальный номер размещения"]))
        if new_unique_numbers:
            df_ = df_new[df_new['Уникальный номер размещения'].isin(list(new_unique_numbers))]  # rows with new numbers
            self.rename_and_write_rows(df_, 'Дата', 'month')
            self.rename_and_write_rows(df_, 'Месяц', 'date')

    def check_dates(self, df_old: pd.DataFrame, df_new: pd.DataFrame, type_: str, sheet: str) -> None:
        column = type_ + " учета оказания услуг"
        idx = df_old.shape[0]
        are_dates_equal = df_new[column][:idx] != df_old[column]  # True or False
        new_dates = df_new[:idx][are_dates_equal]  # from df_new select values which are not equal to old (with index)
        new_dates = new_dates[column].rename(self.today)  # rename column "... учета оказания услуг"
        self.write_values(values=new_dates, sheet=sheet, idx=idx)

    def write_values(self, values: pd.Series, sheet: str, idx: int) -> None:
        df_ = pd.read_excel(self.table_name, engine=self.engine, sheet_name=sheet, index_col=0)
        df_.loc[:idx - 1, self.today] = values
        with pd.ExcelWriter(self.table_name, engine=self.engine, mode='a', if_sheet_exists='replace') as writer_:
            df_.to_excel(writer_, sheet_name=sheet)

    def check_values_on_changes(self, df_new: pd.DataFrame) -> None:
        df_old = pd.read_excel("Table_old.xlsx", engine=self.engine, index_col=0)
        self.check_unique_numbers(df_old, df_new)
        self.check_dates(df_old, df_new, "Дата", "date")
        self.check_dates(df_old, df_new, "Месяц", "month")

    def get_google_data(self) -> pd.DataFrame:
        table: dict = get_from_google_sheet(cred_file_name='creds.json',  # easy to make pandas.DataFrame from dict
                                            sheet_id=self.sheet_id)
        columns = ['ФИО/Название\nподрядчика', 'Уникальный номер размещения',  # cols needed from table
                   'Дата учета оказания услуг', 'Месяц учета оказания услуг']
        # df_ = pd.concat([pd.Series(name=column, data=table.get(column)) for column in columns],
        #                 axis=1)  # concat because of some cells in cols can be empty
        df = pd.DataFrame(data=table, columns=columns)
        return df

    def write_date_month_tbls(self, df_: pd.DataFrame) -> None:
        with pd.ExcelWriter(self.table_name, engine=self.engine) as writer:
            df_.drop('Месяц учета оказания услуг', axis=1).rename(
                columns={'Дата учета оказания услуг': self.today}).to_excel(writer, sheet_name="date")
            df_.drop('Дата учета оказания услуг', axis=1).rename(
                columns={'Месяц учета оказания услуг': self.today}).to_excel(writer, sheet_name="month")


def main():
    parser = ExcelParser()
    df = parser.get_google_data()

    if os.path.isfile('Table_old.xlsx'):  # True if Table_old.xlsx exist, so we can compare new data with old data
        parser.check_values_on_changes(df)
    else:
        parser.write_date_month_tbls(df)
    df.to_excel("Table_old.xlsx", engine=parser.engine)  # rewrite old data with new


if __name__ == "__main__":
    main()
