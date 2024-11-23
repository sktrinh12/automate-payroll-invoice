import os.path
import requests
import base64


def fetch_timelogs(start_date, end_date):
    api_key = os.getenv("API_KEY")
    pwd = os.getenv("PASS")
    site_name = os.getenv("SITE_NAME")
    project_id = os.getenv("PROJECT_ID")
    if not api_key or not site_name or not project_id or not pwd:
        raise ValueError("Missing required environment variables.")

    credentials = f"{api_key}:{pwd}"
    encoded_credentials = base64.b64encode(credentials.encode("utf-8")).decode("utf-8")
    fields = "id,dateCreated,minutes,description"
    params = {"startDate": start_date, "endDate": end_date, "fields[timelogs]": fields}
    url = f"https://{site_name}.teamwork.com/projects/api/v3/projects/{project_id}/time.json"
    headers = {
    "Authorization": f"Basic {encoded_credentials}"
    }

    response = requests.get(url, headers=headers, params=params)
    response.raise_for_status()  # Raise an error for bad responses

    data = response.json()
    return response.request.url, data["timelogs"]
