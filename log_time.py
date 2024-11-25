import requests
import base64
import json
import os
import pytz
import argparse
from datetime import datetime, date
from helper import load_env_file

load_env_file(".env")

api_key = os.getenv("API_KEY")
password = os.getenv("PASS")

credentials = f"{api_key}:{password}"
encoded_credentials = base64.b64encode(credentials.encode("utf-8")).decode("utf-8")

task_id_dct = {
    'dm':os.getenv('TASK_DM_SUPPORT_ID'),
    '1on1':os.getenv('TASK_1ON1_GENARO_ID'),
    'meeting':os.getenv('TASK_INTERNAL_TEAM_MEETING_ID')
 }

def valid_task(task_key):
    if task_key not in task_id_dct:
        raise argparse.ArgumentTypeError(
            f"Invalid task '{task_key}'. Valid options are: {', '.join(task_id_dct.keys())}"
        )
    return task_key

def valid_time(time_string):
    try:
        return datetime.strptime(time_string, "%H:%M").time()
    except ValueError:
        raise argparse.ArgumentTypeError(
            f"Invalid time format: '{time_string}'. Expected format: HH:MM"
        )

def convert_to_utc(est_time):
    est = pytz.timezone('US/Eastern')
    today = datetime.today().date()
    local_time = datetime.combine(today, est_time)
    # Localize the time to EST timezone
    local_time = est.localize(local_time)
    # convert to utc
    utc_time = local_time.astimezone(pytz.utc)

    return utc_time.strftime("%H:%M:%S")

headers = {
    "Content-Type": "application/json",
    "Authorization": f"Basic {encoded_credentials}"
}


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Log time for particular task in teamwork')
    parser.add_argument('-t', '--task', type=valid_task, required=True, help=f"Required task name to log time. Valid tasks: {', '.join(task_id_dct.keys())}")
    parser.add_argument('-d', '--description', type=str, default='', help='Optional description for the task')
    parser.add_argument('-s', '--start-time', type=valid_time, required=True, help='Start time in HH:MM format')
    parser.add_argument('-m', '--minutes', type=int, help='Time in minutes elapsed for task')
    parser.add_argument('-hr', '--hours', type=int, help='Time in hours elapsed for task (will be converted to minutes)')

    args = parser.parse_args()

    if args.hours:
        args.minutes = args.hours*60

    if args.minutes is None:
        raise ValueError("You must specify either -m for minutes or -h for hours")

    current_date = date.today()
    task_id = task_id_dct[args.task]
    url = f"https://{os.getenv("SITE_NAME")}.teamwork.com/projects/api/v3/tasks/{task_id}/time.json"

    payload = {
        "timelog": {
            "date": current_date.isoformat(),
            "time": convert_to_utc(args.start_time),
            "isUtc": True,
            "description": args.description,
            "isBillable": True,
            "minutes": args.minutes,
            "projectId": int(os.getenv('PROJECT_ID')),
            "taskId": int(task_id),
            "userId": int(os.getenv("USER_ID")),
        },
        "timelogOptions": {
            "markTaskComplete": False
        }
    }

    # pprint(json.dumps(payload))
    response = requests.post(url, headers=headers, data=json.dumps(payload))
    if response.status_code == 201:
        print(response.text)
    else:
        print(f'>>>ERROR {response.status_code}')
        print(response.text)
