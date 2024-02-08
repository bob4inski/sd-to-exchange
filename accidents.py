from conenctions.redmine import get_locations_from_db
from datetime import datetime, timedelta
from redminelib import Redmine
from dotenv import load_dotenv
import logging
import pandas as pd
import numpy as np
import requests
import os

def get_accidents_from_api(locations: dict):
    dotenv_path = os.path.join(os.path.dirname(__file__), '.env')
    
    if os.path.exists(dotenv_path):
        load_dotenv(dotenv_path)

    url = os.environ['REDMINE_URL']
    key = os.environ['REDMINE_KEY']
    
    try:
        logging.debug('try to get data')
        redmine = Redmine("https://sd.talantiuspeh.ru/", key=key)
        issues = redmine.issue.filter(
                    project_id=29,
                    status_id = "*",
                    created_on='><2023-12-01'   
        )
    except Exception as ex:
        logging.debug(ex)
        exit(1)

    if issues:
        issues_list = []

        for issue in issues:
            try:
                issue_dict = {
                    "id": issue.id,
                    "subject": issue.subject,
                    "status": issue.status,
                    "start_time": "",
                    "finish_time": "",
                    "location": "None"
                        }

                for custom_field in issue.custom_fields:
                    if custom_field.id == 146:
                        try:
                            # issue_dict["start_time"] = str(datetime.strptime(custom_field.value , "%Y-%m-%d %H:%M:%S") + timedelta(hours=3))
                            issue_dict["start_time"] = custom_field.value
                        except Exception as ex:
                            logging.critical(f'issue {issue.id} doesnt have start_time')
                            logging.debug(ex)
                            continue
                    if custom_field.id == 147:
                        issue_dict["finish_time"] = custom_field.value

                    if custom_field.id == 110:
                        issue_dict["location"] = locations[int(custom_field.value)]

                if issue_dict["finish_time"] == "":
                    issue_dict["finish_time"] = str(datetime.strptime(issue_dict["start_time"], "%Y-%m-%d %H:%M:%S") + timedelta(hours=1))
                # else:
                #     issue_dict["finish_time"] = str(datetime.strptime(issue_dict["finish_time"], "%Y-%m-%d %H:%M:%S") + timedelta(hours=3))
                issues_list.append(issue_dict)
            except Exception as ex:
                logging.critical("ошибка при создании issue")
                logging.critical(ex)

        df = pd.DataFrame(issues_list, columns=["id", "subject", "status_id", "start_time", "finish_time", "location"])

        return df
    
    else:
        logging.debug("there is no issues")
        logging.info("shutdown")
        exit(1)

def normalize():
    locations_from_bd = get_locations_from_db()
    accidents = get_accidents_from_api(locations=locations_from_bd)

    accidents["start_time"].replace("", np.nan, inplace=True)
    accidents.dropna(subset=['start_time'], inplace=True)

    logging.info("dataframe loaded successfully")

    return accidents

