from dotenv import load_dotenv
import logging
import pandas as pd
import numpy as np
import requests
import os

 
from conenctions.redmine import get_locations_from_db
from conenctions.redmine import get_from_api
from datetime import datetime, timedelta

def get_events_from_api(locations: dict):
    # dotenv_path = os.path.join(os.path.dirname(__file__), '.env')
    
    # if os.path.exists(dotenv_path):
    #     load_dotenv(dotenv_path)

    # key = os.environ['REDMINE_KEY']
    key = "54bd095f52e85f988860c6e2def8bf91b9177c7b"
    
    try:
        logging.debug('try to get data')
        url = "https://sd.talantiuspeh.ru/issues.json?project_id=9&status_id=*&offset=0&limit=1000&per_page=1000;"
        data = get_from_api(url=url, redmine_key=key)

    except Exception as ex:
        logging.debug(ex)

    if data["issues"]:
        issues = data["issues"]
        issues_list = []

        for row in issues:
            issue = {
                "id": row["id"],
                "subject": row["subject"],
                "status_id": row["status"]["id"],
                "start_time": "",
                "finish_time": "",
                "duration": "",
                "location": "None"
                }

            for custom_field in row["custom_fields"]:
                if custom_field["id"] == 24:
                    issue["start_time"] = custom_field["value"]
                
                if custom_field["id"] == 110:
                    issue["location"] = locations[int(custom_field["value"])]

                if custom_field["id"] == 115 and custom_field["value"]:
                    issue["duration"] = int(custom_field["value"])
                    # issue["finish_time"] = issue["start_time"] + timedelta(minutes=issue["duration"])
                    issue["finish_time"] = str(datetime.strptime(str(issue["start_time"]), "%Y-%m-%d %H:%M:%S") + timedelta(minutes=issue["duration"]))
            
            if issue["finish_time"]:
                issues_list.append(issue)
                  
        df = pd.DataFrame(issues_list, columns=["id", "subject", "status_id", "start_time", "finish_time","duration", "location"])

        return df
    else:
        logging.debug("there is no issues")
        logging.info("shutdown")
        exit(1)

def get_dataframed_events():
    locations_from_bd = get_locations_from_db()
    # accidents = get_accidents_from_bd()
    events = get_events_from_api(locations_from_bd)

    events.dropna(subset=['finish_time'], inplace=True)

    return events
