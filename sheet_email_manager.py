import os.path
import sqlite3
import calendar
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import glob
import base64
import argparse


SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/gmail.compose",
]

INVOICE_COUNT = 0
OUTPUT_FILE = ""
CURRENT_DATE = datetime.now()
MONTH = CURRENT_DATE.strftime("%B")
MONTH_NUM = CURRENT_DATE.strftime("%m")
YEAR = CURRENT_DATE.year


def load_env_variables(env_file=".env"):
    env_vars = {}
    with open(env_file) as file:
        for line in file:
            line = line.strip()
            if line and not line.startswith("#"):
                key, value = line.split("=", 1)
                env_vars[key.strip()] = value.strip()
    return env_vars


env_vars = load_env_variables()
SPREADSHEET_ID = env_vars.get("SPREADSHEET_ID")
TOKEN_PATH = env_vars.get("TOKEN_PATH")
BASE_DIR = env_vars.get("BASE_DIR")
SENDER_NAME = env_vars.get("SENDER_NAME")
RECIEVER_NAME=env_vars.get("RECIEVER_NAME")
RECIEVER_EMAIL=env_vars.get("RECIEVER_EMAIL")
PREFIX=env_vars.get("PREFIX")
SENDER_COMPANY=env_vars.get("SENDER_COMPANY")


def pprint(message, separator="=", width=50):
    print(f"\n{separator * width}\n{message}\n{separator * width}\n")


def count_invoices(base_directory):
    global INVOICE_COUNT
    pattern = os.path.join(base_directory, "**", f"{PREFIX}_*.xlsx")
    invoice_files = glob.glob(pattern, recursive=True)
    INVOICE_COUNT = len(invoice_files) + 1
    pprint(f"invoice count: {INVOICE_COUNT}")


def extract_data_from_db(day=None, month=None, year=None):
    if day <= 15:
        from_date = f"{year}-{month:02d}-01"
        to_date = f"{year}-{month:02d}-15"
    else:
        last_day = calendar.monthrange(year, month)[1]
        from_date = f"{year}-{month:02d}-16"
        to_date = f"{year}-{month:02d}-{last_day:02d}"

    pprint(f"From date: {from_date}, To date: {to_date}")

    conn = sqlite3.connect("timesheet.db")
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

    return from_date, to_date, total_hours


def parse_arguments():
    parser = argparse.ArgumentParser(description='Process a date string in the format DD-MM-YYYY or MM-DD-YYYY.')
    parser.add_argument('-d', '--date', type=str, help='Required date string (e.g., 15-10-2024 or 10-15-2024)')
    args = parser.parse_args()

    try:
        parsed_date = datetime.strptime(args.date, '%d-%m-%Y')
    except ValueError:
        try:
            parsed_date = datetime.strptime(args.date, '%m-%d-%Y')
        except ValueError:
            parser.error("Date must be in the format DD-MM-YYYY or MM-DD-YYYY.")

    return parsed_date.day, parsed_date.month, parsed_date.year


def update_sheet(service, from_date, to_date, e22_value):
    global OUTPUT_FILE
    formatted_date = CURRENT_DATE.strftime("%m/%d/%Y")

    b22_value = f"{from_date} to {to_date} paid hours worked for {SENDER_COMPANY}"
    count_invoices(BASE_DIR)
    formatted_count = f"{INVOICE_COUNT:03}"
    g3_value = f"{formatted_count}"
    # need to reformat date because the excel export is Y-m-d
    # file naming in project directory is m-d-Y
    date_obj = datetime.strptime(to_date, "%Y-%m-%d")
    reformatted_date = date_obj.strftime("%m_%d_%Y")
    OUTPUT_FILE = f"{PREFIX}_{reformatted_date}_{formatted_count}.xlsx"
    range_g3 = "Invoice template!G3"

    # Update G3 Invoice number
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=range_g3,
        valueInputOption="RAW",
        body={"values": [[g3_value]]},
    ).execute()

    # Update G6 Date
    range_g6 = "Invoice template!G6"
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=range_g6,
        valueInputOption="RAW",
        body={"values": [[f"Sent date: {formatted_date}"]]},
    ).execute()

    # Update B22 text box
    range_b22 = "Invoice template!B22"
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=range_b22,
        valueInputOption="RAW",
        body={"values": [[b22_value]]},
    ).execute()

    # Update E22 time hours
    range_e22 = "Invoice template!E22"
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=range_e22,
        valueInputOption="RAW",
        body={"values": [[e22_value]]},
    ).execute()


def create_gmail_draft(service_gmail, from_date, to_date, file_path):
    subject = f"Invoice_{INVOICE_COUNT:03} from {SENDER_NAME} - {from_date} to {to_date}"
    body_html = f"""
<html>
<body>
    <p>Hi {RECIEVER_NAME} Payroll Department,</p>
    <p></p>
    <p>
        Please see attached current invoice number {INVOICE_COUNT:03} for
        <b>{from_date}</b> to <b>{to_date}</b> in .xlsx file format.
    </p>
    <p></p>
    <p>Please let me know if you have any questions. Thank you.</p>
    <p></p>
    <p></p>
    <p></p>
    <p>Best regards, </p>
    <p></p>
    <p>{SENDER_NAME}</p>
</body>
</html>
"""

    message = MIMEMultipart()
    message["to"] = RECIEVER_EMAIL
    message["subject"] = subject
    message.attach(MIMEText(body_text, "plain"))

    with open(file_path, "rb") as file:
        mime_base = MIMEBase("application", "octet-stream")
        mime_base.set_payload(file.read())
        encoders.encode_base64(mime_base)
        mime_base.add_header(
            "Content-Disposition", f"attachment; filename={os.path.basename(file_path)}"
        )
        message.attach(mime_base)

    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()

    draft = (
        service_gmail.users()
        .drafts()
        .create(userId="me", body={"message": {"raw": raw_message}})
        .execute()
    )

    pprint(f"Draft email created ({subject}) with ID: {draft['id']}")


def main():
    global MONTH
    global MONTH_NUM
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(TOKEN_PATH, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        service_sheets = build("sheets", "v4", credentials=creds)
        service_drive = build("drive", "v3", credentials=creds)
        service_gmail = build("gmail", "v1", credentials=creds)
        day, month, year = parse_arguments()
        from_date, to_date, total_hours = extract_data_from_db(day, month, year)
        update_sheet(service_sheets, from_date, to_date, total_hours)
        request = service_drive.files().export_media(
            fileId=SPREADSHEET_ID,
            mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        recorded_month = datetime.strptime(to_date, "%Y-%m-%d")
        recorded_month_num = recorded_month.month
        recorded_month_name = recorded_month.strftime("%B")
        pprint(f"recorded month name: {recorded_month_name}")
        if recorded_month != MONTH:
            MONTH = recorded_month_name
            MONTH_NUM = str(recorded_month_num).zfill(2)
        full_file_path = f"{BASE_DIR}/{YEAR}/{MONTH_NUM}_{MONTH}/{OUTPUT_FILE}"

        with open(full_file_path, "wb") as f:
            f.write(request.execute())

        pprint(f"File downloaded successfully as {full_file_path}")
        create_gmail_draft(service_gmail, from_date, to_date, full_file_path)
    except HttpError as err:
        print(err)


if __name__ == "__main__":
    main()
