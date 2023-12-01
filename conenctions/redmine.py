import mysql.connector
import requests
import os

def get_locations_from_db():

    mydb = mysql.connector.MySQLConnection(
        user=os.environ["DB_USER"] or "user",
        password=os.environ["DB_PASSWORD"] or "password",
        host=os.environ["DB_HOST"] or "localhost",
        port=os.environ["DB_PORT"] or 5432,
        database=os.environ["DB_DATABASE"] or "redmine_prod"
    )

    cur = mydb.cursor()
    querry = """
                SELECT id, name FROM 
                    custom_field_enumerations AS cfe 
                where custom_field_id = 110
            """

    cur.execute(querry)
    response = cur.fetchall()

    response_dict = {object_id: value for object_id, value in response}
    return response_dict


def get_from_api(url: str, redmine_key: str):

    headers = {
        'Content-Type': 'application/json',
        'X-Redmine-API-Key': redmine_key,
    }

    response = requests.get(url, headers=headers)

    return response.json()


