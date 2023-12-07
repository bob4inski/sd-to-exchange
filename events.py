from dotenv import load_dotenv
import logging
import pandas as pd
import numpy as np
import requests
import os

 
from conenctions.redmine import get_locations_from_db
from conenctions.redmine import get_from_api
from datetime import datetime, timedelta
from redminelib import Redmine

def get_events_from_api(locations: dict):
    dotenv_path = os.path.join(os.path.dirname(__file__), '.env')
    if os.path.exists(dotenv_path):
        load_dotenv(dotenv_path)

    key = os.environ['REDMINE_KEY']
    
    try:
        logging.debug('try to get data')
        redmine = Redmine("https://sd.talantiuspeh.ru/", key=key)
        issues = redmine.issue.filter(
                    project_id=9,
                    limit=1000
                )
    except Exception as ex:
        logging.debug(ex)

    if issues:
        issues_list = []

        for issue in issues:
            issue_dict = {
                "id": issue.id,
                "subject": issue.subject,
                "status": issue.status,
                "start_time": "",
                "finish_time": "",
                "duration": "",
                "close_code": "",
                "location": "None"
                }

            for custom_field in issue.custom_fields:
                if custom_field.id  == 24:
                    issue_dict["start_time"] = custom_field.value

                if custom_field.id  == 81:
                    issue_dict["close_code"] = custom_field.value

                if custom_field.id == 110:
                    issue_dict["location"] = locations[int(custom_field.value)]

                if custom_field.id == 115 and custom_field.value:
                    issue_dict["duration"] = int(custom_field.value)
                    issue_dict["finish_time"] = str(datetime.strptime(str(issue_dict["start_time"]), "%Y-%m-%d %H:%M:%S") + timedelta(minutes=issue_dict["duration"]))
            
            if issue_dict["finish_time"]:
                issues_list.append(issue_dict)
                  
        df = pd.DataFrame(issues_list, columns=["id", "subject", "status_id", "start_time", "finish_time","duration","close_code", "location"])

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
