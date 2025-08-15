from helper import datetime, parse_arguments, pprint, load_env_file
from teamwork_request import fetch_timelogs
from sqlite import upload_to_sqlite
from sheet_email_manager import draft_email
from sqlite import extract_data_from_db

if __name__ == "__main__":
    load_env_file(".env")

    day, month, year, boolean_email, boolean_sqlite3 = parse_arguments()

    if day <= 15:
        from_date = f"{year}-{month:02d}-01"
        to_date = f"{year}-{month:02d}-15"
    else:
        import calendar
        last_day = calendar.monthrange(year, month)[1]
        from_date = f"{year}-{month:02d}-16"
        to_date = f"{year}-{month:02d}-{last_day:02d}"

    pprint(f"From date: {from_date}, To date: {to_date}")


    # Upload timelogs to SQLite & fetch timelogs from the API
    if boolean_sqlite3:
        request_url, timelogs = fetch_timelogs(from_date, to_date)
        upload_to_sqlite(timelogs, request_url)

    total_hours = extract_data_from_db(from_date, to_date)

    # Draft an email
    if boolean_email:
        draft_email(from_date, to_date, total_hours)
