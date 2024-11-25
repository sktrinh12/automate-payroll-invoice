from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from helper import pprint, count_invoices, datetime
import os.path
import base64


SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/gmail.compose",
]

INVOICE_COUNT = 0
OUTPUT_FILE = ""


def update_sheet(service, from_date, to_date, total_hours):
    global INVOICE_COUNT
    global OUTPUT_FILE
    todays_date = datetime.now().strftime("%m/%d/%Y")

    b22_value = f"{from_date} to {to_date} paid hours worked for {os.getenv('SENDER_COMPANY')}"
    INVOICE_COUNT = count_invoices(os.getenv("BASE_DIR"))
    g3_value = f"{INVOICE_COUNT:03}"
    # need to reformat date because the excel export is Y-m-d
    # file naming in project directory is m-d-Y
    date_obj = datetime.strptime(to_date, "%Y-%m-%d")
    reformatted_date = date_obj.strftime("%m_%d_%Y")
    OUTPUT_FILE = f"{os.getenv('PREFIX')}_{reformatted_date}_{INVOICE_COUNT:03}.xlsx"
    range_g3 = "Invoice template!G3"

    # Update G3 Invoice number
    service.spreadsheets().values().update(
        spreadsheetId=os.getenv('SPREADSHEET_ID'),
        range=range_g3,
        valueInputOption="RAW",
        body={"values": [[g3_value]]},
    ).execute()

    # Update G6 Date
    range_g6 = "Invoice template!G6"
    service.spreadsheets().values().update(
        spreadsheetId=os.getenv('SPREADSHEET_ID'),
        range=range_g6,
        valueInputOption="RAW",
        body={"values": [[f"Sent date: {todays_date}"]]},
    ).execute()

    # Update B22 text box
    range_b22 = "Invoice template!B22"
    service.spreadsheets().values().update(
        spreadsheetId=os.getenv('SPREADSHEET_ID'),
        range=range_b22,
        valueInputOption="RAW",
        body={"values": [[b22_value]]},
    ).execute()

    # Update E22 time hours
    range_e22 = "Invoice template!E22"
    service.spreadsheets().values().update(
        spreadsheetId=os.getenv('SPREADSHEET_ID'),
        range=range_e22,
        valueInputOption="RAW",
        body={"values": [[total_hours]]},
    ).execute()


def create_gmail_draft(service_gmail, from_date, to_date, file_path):
    subject = f"Invoice_{INVOICE_COUNT:03} from {os.getenv('SENDER_NAME')} - {from_date} to {to_date}"
    body_html = f"""
<html>
<body>
    <p>Hi {os.getenv('RECIEVER_NAME')} Payroll Department,</p>
    <br>
    <p>
        Please see attached current invoice number <b>{INVOICE_COUNT:03}</b> for
        <b>{from_date}</b> to <b>{to_date}</b> in <span style="font-family: monospace; background-color: #f4f4f4; padding: 2px 4px; border-radius: 3px; border: 1px solid #ddd;">.xlsx</span> file format.
    </p>
    <br>
    <br>
    <p>Please let me know if you have any questions. Thank you.</p>
    <br>
    <br>
    <br>
    <p>Best regards, </p>
    <br>
    <p>{os.getenv('SENDER_NAME')}</p>
</body>
</html>
"""

    message = MIMEMultipart()
    message["to"] = os.getenv("RECIEVER_EMAIL")
    message["subject"] = subject
    message.attach(MIMEText(body_html, "html"))

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


def draft_email(from_date, to_date, total_hours):
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
            flow = InstalledAppFlow.from_client_secrets_file(os.getenv('TOKEN_PATH'), SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        service_sheets = build("sheets", "v4", credentials=creds)
        service_drive = build("drive", "v3", credentials=creds)
        service_gmail = build("gmail", "v1", credentials=creds)
        update_sheet(service_sheets, from_date, to_date, total_hours)
        request = service_drive.files().export_media(
            fileId=os.getenv("SPREADSHEET_ID"),
            mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        to_date_obj = datetime.strptime(to_date, "%Y-%m-%d")
        MONTH = to_date_obj.strftime("%B")
        MONTH_NUM = str(to_date_obj.month).zfill(2)
        YEAR = to_date_obj.year
        full_file_path = (
            f"{os.getenv('BASE_DIR')}/{YEAR}/{MONTH_NUM}_{MONTH}/{OUTPUT_FILE}"
        )

        with open(full_file_path, "wb") as f:
            f.write(request.execute())

        pprint(f"File downloaded successfully as {full_file_path}")
        create_gmail_draft(service_gmail, from_date, to_date, full_file_path)
    except HttpError as err:
        print(err)
