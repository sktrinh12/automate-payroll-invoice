## Automate payroll record-keeping and generate invoice

- The working directory is like so:

```bash
 draft_gmail.py
 invoices
 timesheet.db
 timesheets
 token.json
 upload_timesheet_to_db.py
 .env
```

#### 1. First download the exported xlsx file to the appropriate directory `timesheets`
#### 1a. Might have to delete `token.json` to freshly generate one in the browser

#### 2. In the terminal run the following in order to upload the data to a sqlite3 database:

```bash
python upload_timesheet_to_db.py timesheets/2024/{MONTH}/{EXPORTED_FILE_NAME}
```

#### 3. Then draft the gmail:
- this python script requires the `.env` file that stores the environmental variables. Update as needed.
- also requires the date passed as an argument so it can look up the sqlite3 database for the proper date range.

```bash
python draft_gmail.py -d MM-DD-YYYY
```
