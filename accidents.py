from dotenv import load_dotenv
import logging
import pandas as pd
import numpy as np
import requests
import os

from conenctions.redmine import get_locations_from_db
from conenctions.redmine import get_from_api

def get_accidents_from_api(locations: dict):
    dotenv_path = os.path.join(os.path.dirname(__file__), '.env')
    
    if os.path.exists(dotenv_path):
        load_dotenv(dotenv_path)

    url = os.environ['REDMINE_URL']
    key = os.environ['REDMINE_KEY']
    
    try:
        logging.debug('try to get data')
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
                "start": "",
                "finish": "",
                "object": "None"
                }

            for custom_field in row["custom_fields"]:
                if custom_field["id"] == 146:
                    issue["start"] = custom_field["value"]

                if custom_field["id"] == 147:
                    issue["finish"] = custom_field["value"]

                if custom_field["id"] == 110:
                    issue["object"] = locations[int(custom_field["value"])]

            issues_list.append(issue)

        df = pd.DataFrame(issues_list, columns=["id", "subject", "status_id", "start", "finish", "object"])

        return df
    else:
        logging.debug("there is no issues")
        logging.info("shutdown")
        exit(1)

def normalize():
    locations_from_bd = get_locations_from_db()

    # accidents = get_accidents_from_bd()
    accidents = get_accidents_from_api(locations=locations_from_bd)

    accidents["start"].replace("", np.nan, inplace=True)
    accidents.dropna(subset=['start'], inplace=True)

    return accidents

