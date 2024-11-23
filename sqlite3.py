import sqlite3
from os import getenv
from functions import pprint


def upload_to_sqlite(timelogs, request_url):
    from_date = None
    to_date = None
    conn = sqlite3.connect(getenv('DB_NAME'))
    cursor = conn.cursor()

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS RCH_TIMESHEET (
            ID INTEGER PRIMARY KEY AUTOINCREMENT,
            DATE_TIME TEXT,
            DECIMAL_HOURS REAL,
            REQUEST_URL TEXT
        )
    """
    )

    last_data = len(timelogs)
    for i, timelog in enumerate(timelogs):
        if i == 0:
            from_date = timelog["dateCreated"]
        if i == last_data-1:
            to_date = timelog["dateCreated"]
        record_id = timelog["id"]
        date_time = timelog["dateCreated"]
        decimal_hours = round(timelog["minutes"] / 60, 2)
        try:
            cursor.execute(
                """
                INSERT INTO RCH_TIMESHEET (ID, DATE_TIME, DECIMAL_HOURS, REQUEST_URL)
                VALUES (?, ?, ?, ?)
            """,
                (record_id, date_time, decimal_hours, request_url),
            )
        except sqlite3.IntegrityError as e:
            print(f"Primary key error on row {i}: {row.to_dict()}")
            print(f"Error message: {e}")

    conn.commit()
    conn.close()
    print(f"Data from {from_date} to {to_date} uploaded successfully.")


def extract_data_from_db(from_date, to_date):
    conn = sqlite3.connect(getenv('DB_NAME'))
    cursor = conn.cursor()

    query = """
        SELECT DECIMAL_HOURS FROM RCH_TIMESHEET
        WHERE DATE(DATE_TIME) BETWEEN ? AND ?
    """
    cursor.execute(query, (from_date, to_date))
    rows = cursor.fetchall()

    total_hours = sum(row[0] for row in rows)
    pprint(f'Total hours: {total_hours}')
    conn.close()

    return total_hours
