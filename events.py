from dotenv import load_dotenv
import mysql.connector
import pandas as pd
import numpy as np
import requests
import os


def get_events_from_bd():

    mydb =  mysql.connector.MySQLConnection(
        user=os.environ["DB_USER"] or "user",
        password=os.environ["DB_PASSWORD"] or "password",
        host=os.environ["DB_HOST"] or "localhost",
        port=os.environ["DB_PORT"] or 5432,
        database=os.environ["DB_DATABASE"] or "redmine_prod"
    )

    cur = mydb.cursor()
    cur.execute(
        '''
        SELECT i.id, 
               i.subject, 
               i.status_id,
               cv_1.value as start,  
               cv_0.value as finish
        FROM issues AS i
            join custom_values AS cv_1 ON cv_1.customized_id = i.id and cv_1.custom_field_id = 146
            join custom_values AS cv_0 ON cv_0.customized_id = i.id and cv_0.custom_field_id = 147
                
        WHERE i.project_id = 29
        ''')

    df = pd.DataFrame(cur, columns=["id", "subject", "status_id", "start", "finish"])

    return df


def get_object_from_bd(object_id: str):

    mydb = mysql.connector.MySQLConnection(
        user=os.environ["DB_USER"] or "user",
        password=os.environ["DB_PASSWORD"] or "password",
        host=os.environ["DB_HOST"] or "localhost",
        port=os.environ["DB_PORT"] or 5432,
        database=os.environ["DB_DATABASE"] or "redmine_prod"
    )

    cur = mydb.cursor()
    querry = """
                SELECT name FROM 
                    custom_field_enumerations AS cfe 
                WHERE id = %s 
            """

    cur.execute(querry, [object_id])
    response = cur.fetchall()

    return response[0][0]

def get_from_api(url: str, redmine_key: str):

    headers = {
        'Content-Type': 'application/json',
        'X-Redmine-API-Key': redmine_key,
    }

    response = requests.get(url, headers=headers)

    return response.json()


def get_events_from_api():
    dotenv_path = os.path.join(os.path.dirname(__file__), '.env')
    
    if os.path.exists(dotenv_path):
        load_dotenv(dotenv_path)

    url = os.environ['REDMINE_URL']
    key = os.environ['REDMINE_KEY']

    data = get_from_api(url=url, redmine_key=key)
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
                issue["object"] = get_object_from_bd(custom_field["value"])

        issues_list.append(issue)

    df = pd.DataFrame(issues_list, columns=["id", "subject", "status_id", "start", "finish", "object"])
    return df


def normalize():
    
    # events = get_events_from_bd()
    events = get_events_from_api()

    events["start"].replace("", np.nan, inplace=True)
    events.dropna(subset=['start'], inplace=True)

    return events

