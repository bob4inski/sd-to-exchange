import os
import logging
import win32com.client
from redis import Redis
from events import normalize
from dotenv import load_dotenv
from datetime import datetime, timedelta



class Calendar():

    def __init__(self, account: str, calendar_name: str, host: str, port: int, password: str):

        outlook = win32com.client.Dispatch("Outlook.Application")
        self.namespace = outlook.GetNamespace("MAPI")
        self.calendar_name = calendar_name
        self.calendar = self.namespace.Folders.Item(account).Folders.Item("Календарь").Folders.Item(calendar_name)

        self.db_connection = Redis(
                                host= host or "10.9.44.12",
                                port= port or 6379,
                                password= password or "biba",
                                decode_responses=True
                            )

    def get_event(self, event_id):
        event = self.namespace.GetItemFromId(event_id)
        logging.debug(f"get event {event_id}")
        return event

    def new_event(self, subject: str, body: str, location: str, id: str,  time_start, time_finish):
        new_event = self.calendar.Items.Add(1)

        new_event.Location = location
        new_event.body = f"https://sd.talantiuspeh.ru/issues/{body}"
        new_event.Subject = subject
        new_event.Start = time_start
        new_event.End = time_finish
        new_event.Save()

        self.db_connection.set(id, new_event.EntryId)
        logging.info(f"event {id}  added to calendar {self.calendar_name}")

    def update_event(self, subject: str, body: str, location: str, event_id: str,  time_start, time_finish):

        event = self.get_event(event_id)
        event.Location = location
        event.body = f"https://sd.talantiuspeh.ru/issues/{body}"
        event.Subject = subject
        event.Start = time_start
        event.End = time_finish
        event.Save()

        event.Save()
        logging.info(f"event {subject} updated")

    def delete_all(self):
        events = self.calendar.Items
        while events.count > 0:
            events.Remove(1)

        logging.info("everything deleted")


def upload(calendar):
    events = normalize()

    for index, row in events.iterrows():
        start_date = datetime.strptime(row["start"], "%Y-%m-%d %H:%M:%S")
        # start_date = row["start"]

        if row["finish"] == "":
            finish_date = start_date + timedelta(hours=1)
        else:
            finish_date = row["finish"]

        logging.debug(f'{row["id"]}, {start_date}, {finish_date}')

        try:
            event_id = calendar.db_connection.get(row["id"])
            if event_id:
                calendar.update_event(subject=row["subject"],
                                   location=row["object"],
                                   body=row["id"],
                                   event_id=event_id,
                                   time_start=start_date,
                                   time_finish=finish_date)
            else:
                calendar.new_event(subject=row["subject"],
                                      location=row["object"],
                                      body=row["id"],
                                      id=row["id"],
                                      time_start=start_date,
                                      time_finish=finish_date)
        except Exception as ex:
            logging.critical(ex)


def main():

    dotenv_path = os.path.join(os.path.dirname(__file__), '.env')
    
    if os.path.exists(dotenv_path):
        load_dotenv(dotenv_path)
    else:
        logging.CRITICAL('There is no .env file!')
        return False

    now = datetime.now().strftime('%Y-%m-%d-%H-%M-%S')
    filename = f'{str(now)}.log'

    logging.basicConfig(filename=filename, level=logging.INFO)
    logging.info('Started')

    user_email = os.environ['USER_EMAIL']
    calendar_name = os.environ['CALENDAR_NAME']
    host = os.environ['REDIS_HOST']
    port = os.environ['REDIS_PORT']
    passwd = os.environ['REDIS_PASSWD']

    cal = Calendar(account=user_email, calendar_name=calendar_name, host=host, port=port, password=passwd)

    # cal.delete_all()
    upload(cal)

    logging.info('Finished')

if __name__ == '__main__':
    main()