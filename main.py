import os
import calendar
from functions import datetime, parse_arguments, pprint, load_env_file
from api_request import fetch_timelogs
from sqlite_uploader import upload_to_sqlite
from draft_gmail import draft_email


if __name__ == "__main__":
    # Load environment variables
    load_env_file(".env")

    day, month, year = parse_arguments()

    if day <= 15:
        from_date = f"{year}-{month:02d}-01"
        to_date = f"{year}-{month:02d}-15"
    else:
        last_day = calendar.monthrange(year, month)[1]
        from_date = f"{year}-{month:02d}-16"
        to_date = f"{year}-{month:02d}-{last_day:02d}"

    pprint(f"From date: {from_date}, To date: {to_date}")

    # Fetch timelogs from the API
    request_url, timelogs = fetch_timelogs(from_date, to_date)

    # Upload timelogs to SQLite
    upload_to_sqlite(timelogs, request_url)

    # Draft an email
    draft_email(from_date, to_date)
