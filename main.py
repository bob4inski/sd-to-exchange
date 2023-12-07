import os
import logging
import win32com.client
from redis import Redis
from dotenv import load_dotenv
from datetime import datetime, timedelta

from accidents import normalize
from events import get_dataframed_events

class Calendar():

    categories = {}
    def __init__(self, account: str, calendar_name: str, host: str, port: int, password: str):

        outlook = win32com.client.Dispatch("Outlook.Application")
        self.namespace = outlook.GetNamespace("MAPI")
        self.calendar_name = calendar_name
        print(calendar_name)
        try:
            self.calendar = self.namespace.Folders.Item(account).Folders.Item("Календарь").Folders.Item(calendar_name)
        except Exception as ex:
            print(ex)
            exit(1)

        self.db_connection = Redis(
                                host= host,
                                port= port,
                                password= password,
                                decode_responses=True
                            )

    def get_categories(self):
        categories = {}
        categories_outlook = self.namespace.Categories
        for category in categories_outlook:
            categories[category.Name] = category.Color

        return categories,categories_outlook
    
    def update_categories(self, accidents):
        actual_categories, categories_outlook  = self.get_categories() # получаем те категории, которые уже существуют в календаре
        all_categories_from_sd = set(accidents["object"]) # получаем список категорий, которые есть в списке задач
        new_categories = list(all_categories_from_sd - set(actual_categories))

        all_colors = list(range(len(actual_categories) + 1)) 
        free_colors = list(set(all_colors) - set(actual_categories.values()))

        if len(free_colors) == 0:
            logging.info("Нет доступных цветов")
        else:
            if new_categories:
                k = 1
                for i in new_categories:
                    logging.debug(f'Добавляю категорю {i} с цветом {k}')
                    try:
                        new_category = categories_outlook.Add(i,free_colors[k])
                        logging.debug(f'Добавил категорю {i} с цветом {k}')
                    except Exception as ex:
                        logging.CRITICAL(ex)
                    k += 1
                else:
                    logging.debug("Все итерации прошли успешно")
            else:
                logging.debug("Все категории уже существуют")

    def delete_by_id(self, issue_id):

        exchange_id = self.db_connection.get(issue_id)
        accident = self.namespace.GetItemFromId(exchange_id)
        accident.Delete()
        logging.debug(f"issue {issue_id} deleted from exchange")

        self.db_connection.delete(issue_id)
        logging.debug(f"issue {issue_id} deleted from redis")


    def get_accident(self, accident_id):
        accident = self.namespace.GetItemFromId(accident_id)
        logging.debug(f"get issue {accident_id}")
        return accident

    def new_accident(self, subject: str, body: str, location: str, id: str,  time_start, time_finish):
        new_accident = self.calendar.Items.Add(1)

        new_accident.Location = location
        new_accident.body = f"https://sd.talantiuspeh.ru/issues/{body}"
        new_accident.Subject = subject
        new_accident.Categories = location
        new_accident.Start = time_start + timedelta(hours=3)
        new_accident.End = time_finish + timedelta(hours=3)
        try:
            new_accident.Save()
            self.db_connection.set(id, new_accident.EntryId)
            logging.info(f"accident {id}  added to calendar {self.calendar_name}")
        except Exception as ex:
            logging.CRITICAL(f"accident {body} cant be updated")
            logging.DEBUG(f"accident {subject} details: subject: {subject}, body: {body}, location: {location}, time_start: {time_start}, time_finish: {time_finish} ")

    def update_accident(self, subject: str, body: str, location: str, accident_id: str,  time_start, time_finish):

        accident = self.get_accident(accident_id)
        accident.Location = location
        accident.body = f"https://sd.talantiuspeh.ru/issues/{body}"
        accident.Subject = subject
        accident.Categories = location
        accident.Start = time_start
        accident.End = time_finish
        try:
            accident.Save()
            logging.info(f"accident {subject} updated")
        except Exception as ex:
            logging.CRITICAL(f"accident {subject} cant be updated")
            logging.DEBUG(f"accident {subject} details: subject: {subject}, body: {body}, location: {location}, time_start: {time_start}, time_finish: {time_finish} ")

    def new_event(self, subject: str, body: str, location: str, id: str,  time_start: str , time_finish: str):
        new_event = self.calendar.Items.Add(1)
        new_event.Location = location
        new_event.body = f"https://sd.talantiuspeh.ru/issues/{body}"
        new_event.Subject = subject
        new_event.Categories = location
        new_event.Start = time_start 
        new_event.End = time_finish 
        try:
            new_event.Save()

            self.db_connection.set(id, new_event.EntryId)

            logging.info(f"event {id}  added to calendar {self.calendar_name}")
        except Exception as ex:
            print(ex)
            logging.CRITICAL(f"event {body} cant be updated")
            logging.DEBUG(f"event {subject} details: subject: {subject}, body: {body}, location: {location}, time_start: {time_start}, time_finish: {time_finish} ")

    def update_event(self, subject: str, body: str, location: str, event_id: str,  time_start: str , time_finish: str):

        event = self.get_accident(event_id)
        event.Location = location
        event.body = f"https://sd.talantiuspeh.ru/issues/{body}"
        event.Subject = subject
        event.Categories = location
        event.Start = time_start
        event.End = time_finish
        try:
            event.Save()
            logging.info(f"event {subject} updated")
        except Exception as ex:
            print(ex)
            logging.CRITICAL(f"event {subject} cant be updated")
            logging.DEBUG(f"event {subject} details: subject: {subject}, body: {body}, location: {location}, time_start: {time_start}, time_finish: {time_finish} ")

    def delete_all(self):
        accidents = self.calendar.Items
        while accidents.count > 0:
            accidents.Remove(1)

        logging.info("everything deleted")


def upload_accidents(calendar, accidents):
    
    for index, row in accidents.iterrows():
        start_date = datetime.strptime(row["start"], "%Y-%m-%d %H:%M:%S") + timedelta(hours=3)

        if row["finish"] == "":
            finish_date = start_date + timedelta(hours=1) + timedelta(hours=3)
        else:
            finish_date = datetime.strptime(row["finish"], "%Y-%m-%d %H:%M:%S") + timedelta(hours=3)

        logging.debug(f'{row["id"]}, {start_date}, {finish_date}')

        try:
            accident_id = calendar.db_connection.get(row["id"])
            if accident_id:
                calendar.update_accident(subject=row["subject"],
                                   location=row["object"],
                                   body=row["id"],
                                   accident_id=accident_id,
                                   time_start=start_date,
                                   time_finish=finish_date)
            else:
                calendar.new_accident(subject=row["subject"],
                                      location=row["object"],
                                      body=row["id"],
                                      id=row["id"],
                                      time_start=start_date,
                                      time_finish=finish_date)
        except Exception as ex:
            logging.critical(ex)      
    else:
        logging.info("upload finished")


def upload_events(calendar, events):
    for index, row in events.iterrows():
        logging.debug(f'{row["id"]}, {row["start_time"]}, {row["finish_time"]}')

        event_id = calendar.db_connection.get(row["id"])
        if event_id:
            if row["close_code"] == "МР-8":
                logging.info(f'мероприятие {row["id"]} будет удалено  ')
                calendar.delete_by_id(row["id"])
                logging.info(f'мероприятие {row["id"]} удалено  ')
            else:
                calendar.update_event(subject=row["subject"],
                            location=row["location"],
                            body=row["id"],
                            event_id=event_id,
                            time_start=row["start_time"],
                            time_finish=row["finish_time"])
        else:
            if row["close_code"] == "МР-8":
                logging.info(f'мероприятие {row["id"]} не будет создано')
            else:
                calendar.new_event(subject=row["subject"],
                                location=row["location"],
                                body=row["id"],
                                id=row["id"],
                                time_start=row["start_time"],
                                time_finish=row["finish_time"])
    else:
        logging.info("upload finished")



def main():
    dotenv_path = os.path.join(os.path.dirname(__file__), '.env')
    
    if os.path.exists(dotenv_path):
        load_dotenv(dotenv_path)
    else:
        logging.CRITICAL('There is no .env file!')
        exit(1)

    now = datetime.now().strftime('%Y-%m-%d-%H-%M-%S')
    filename = f'logs/{str(now)}.log'

    logging.basicConfig(filename=filename, level=logging.DEBUG)
    logging.info('Started')

    user_email = os.environ['USER_EMAIL']

    accidents_calendar = os.environ['ACCIDENTS_CALENDAR']
    events_calendar = os.environ['EVENTS_CALENDAR']

    redis_host = os.environ['REDIS_HOST']
    redis_port = os.environ['REDIS_PORT']
    redis_passwd = os.environ['REDIS_PASSWD']

    # try:
    #     accident_calendar = Calendar(account=user_email, calendar_name=events_calendar, host=host, port=port, password=passwd)
    #     accidents = normalize() #получение таблицы всех заявок с полями 
    #     accident_calendar.update_categories(accidents=accidents) # обновление категорий в календаре чтобы отражать по цветам
    #     # cal.delete_all() #выстрел себе в колено
    #     upload_accidents(accident_calendar,accidents) #загрузка всех заявок в календарь 
    # except Exception as ex:
    #     print(ex)

    try:
        events_calendar = Calendar(account=user_email, calendar_name=events_calendar, host=redis_host, port=redis_port, password=redis_passwd)
        events = get_dataframed_events() #получение таблицы всех заявок с полями 
        # cal.delete_all() #выстрел себе в колено
        upload_events(calendar=events_calendar, events=events) #загрузка всех заявок в календарь 

        logging.info("Okaaay")
    except Exception as ex:
        print(ex)
    

if __name__ == '__main__':
    main()