## Automate payroll record-keeping and generate invoice

- The working directory is like so:

```bash
 invoices
 timesheet.db
 timesheets
 token.json
 helper.py
 main.py
 sheet_email_manager.py
 sqlite.py
 teamwork_request.py
 .env
```

#### Simply run `python main.py -d DD-MM-YYYY` 
- This now has been streamlined to be one `main.py` file that has several modules
- The date argument is used to grab the proper date range. The invoices correspond to bi-weekly schedule so it will always either grab the 1st to the 15th of the month or the 16th to the end of the month
- Might have to delete `token.json` to freshly generate one in the browser
- The database table name is RCH_TIMESHEET


#### Flow of script:
- this python script requires the `.env` file that stores the environmental variables. Update as needed
- use of the teamwork API to fetch the proper data and return a json payload
- python loops through payload and uploads to the sqlite3 database, `DB_NAME` which is part of the `.env` content.
- Use of Google Sheets API to update several cells in a specific sheet
- Download the sheet locally and also attach it to a draft email using Google Gmail API 
