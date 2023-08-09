# Hackathon_excel_parser
Excel parser needed by SkillFactory management team
## Description
Script should be launched from command line.\
Before you start:
* insert your excel google sheet id into SHEET_ID variable at settings.py.
* put your credentials json file creds.json to ./task folder\

To launch script, you should go to ./task folder \
and launch hacathon_excel.py with command: \
**python hacathon_excel.py**

At first launch two Excel files will arise in current folder:
* **Table_old.xlsx** \
  File contains data you should compare with and search for differences when you launch next time.
* **Table.xlsx** \
  Output file. Contains two sheets:
  * **date** - with differences in column "Дата учета оказания услуг"
  * **month** - with differences in column "Месяц учета оказания услуг"

Each next launch will update both tables with new data.
  
