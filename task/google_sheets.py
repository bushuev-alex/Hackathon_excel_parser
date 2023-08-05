from pprint import pprint
import httplib2
import apiclient
from settings import SHEET_ID
from oauth2client.service_account import ServiceAccountCredentials


def make_dict_from_list(data_list: list) -> dict:
    dict_from_list = {item[0]: item[1:] for item in data_list}
    return dict_from_list


def get_from_google_sheet(cred_file_name: str, sheet_id: str) -> dict:
    # Авторизуемся и получаем service — экземпляр доступа к API
    credentials = ServiceAccountCredentials.from_json_keyfile_name(cred_file_name,
                                                                   ['https://www.googleapis.com/auth/spreadsheets',
                                                                    'https://www.googleapis.com/auth/drive'])
    http_auth = credentials.authorize(httplib2.Http())
    service = apiclient.discovery.build('sheets', 'v4', http=http_auth)
    values = service.spreadsheets().values().get(spreadsheetId=sheet_id,  # Пример чтения файла
                                                 range='A:AQ',
                                                 majorDimension='COLUMNS').execute()
    values_dict = make_dict_from_list(values.get('values'))
    return values_dict


if __name__ == "__main__":
    values_ = get_from_google_sheet(cred_file_name='creds.json',
                                    sheet_id=SHEET_ID)
    pprint(values_)
