# spreadsheet-boy
A simple Python package with an executable script that serves for uploading CSV files
to Google Spreadsheets. 

This is useful when you have different cron jobs and different processes that generate several kinds of reports, on daily/weekly basis. Use this as a standalone tool in you cron jobs that will upload your reports to Google Spreadsheets upon generation. Each time `spreadsheet-boy` uploads a doc, it adds a new sheet to a destination spreadsheet with name of current date, like "06-04-2016". This means that default implementation is limited to at least **daily** uploads. Further uploads during the day will **rewrite the sheet**. 

1) Configure a `spreadsheets.cfg` config file: set up spreadsheets that program should know of.
   Set up authentication and logging. See `spreadsheets.cfg.example` to get an idea of configuration format.
   
   Each document is basically a key to Google document and file path to file itself. Formats suppported: CSV
   
   For a custom path of config file, use `--config <path>` command line arg for script.
   
2a) To fire upload of all configured items, run: 

    $ python upload_spreadsheet.py

2b) To upload only specific items, for example sspreadsheet `report` and `expenses` (names must match names set in `app -> spreadsheets` key of config), run:

    $ python upload_spreadsheet.py --doc report expenses
