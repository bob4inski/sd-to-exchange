import os
import logging
import win32com.client
from redis import Redis
from events import normalize
from dotenv import load_dotenv
from datetime import datetime, timedelta



class Calendar():

    categories = {}
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

    def get_categories(self):
        categories = {}

        categories_outlook = self.namespace.Categories

        for category in categories_outlook:
            categories[category.Name] = category.Color

        return categories,categories_outlook
    
    def update_categories(self, events):
        actual_categories, categories_outlook  = self.get_categories() # получаем те категории, которые уже существуют в календаре
        all_categories_from_sd = set(events["object"]) # получаем список категорий, которые есть в списке задач
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


    def get_event(self, event_id):
        event = self.namespace.GetItemFromId(event_id)
        logging.debug(f"get event {event_id}")
        return event

    def new_event(self, subject: str, body: str, location: str, id: str,  time_start, time_finish):
        new_event = self.calendar.Items.Add(1)

        new_event.Location = location
        new_event.body = f"https://sd.talantiuspeh.ru/issues/{body}"
        new_event.Subject = subject
        new_event.Categories = location
        new_event.Start = time_start + timedelta(hours=3)
        new_event.End = time_finish + timedelta(hours=3)
        try:
            new_event.Save()
            self.db_connection.set(id, new_event.EntryId)
            logging.info(f"event {id}  added to calendar {self.calendar_name}")
        except Exception as ex:
            logging.CRITICAL(f"event {body} cant be updated")
            logging.DEBUG(f"event {subject} details: subject: {subject}, body: {body}, location: {location}, time_start: {time_start}, time_finish: {time_finish} ")


        

    def update_event(self, subject: str, body: str, location: str, event_id: str,  time_start, time_finish):

        event = self.get_event(event_id)
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
            logging.CRITICAL(f"event {subject} cant be updated")
            logging.DEBUG(f"event {subject} details: subject: {subject}, body: {body}, location: {location}, time_start: {time_start}, time_finish: {time_finish} ")


        

    def delete_all(self):
        events = self.calendar.Items
        while events.count > 0:
            events.Remove(1)

        logging.info("everything deleted")


def upload(calendar, events):
    
    for index, row in events.iterrows():
        start_date = datetime.strptime(row["start"], "%Y-%m-%d %H:%M:%S") + timedelta(hours=3)
        # start_date = row["start"]

        if row["finish"] == "":
            finish_date = start_date + timedelta(hours=1) + timedelta(hours=3)
        else:
            finish_date = datetime.strptime(row["finish"], "%Y-%m-%d %H:%M:%S") + timedelta(hours=3)

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
    filename = f'logs/{str(now)}.log'

    logging.basicConfig(filename=filename, level=logging.INFO)
    logging.info('Started')

    user_email = os.environ['USER_EMAIL']
    calendar_name = os.environ['CALENDAR_NAME']
    host = os.environ['REDIS_HOST']
    port = os.environ['REDIS_PORT']
    passwd = os.environ['REDIS_PASSWD']

    cal = Calendar(account=user_email, calendar_name=calendar_name, host=host, port=port, password=passwd)

    events = normalize() #получение таблицы всех заявок с полями 

    cal.update_categories(events=events) # обновление категорий в календаре чтобы отражать по цветам

    # cal.delete_all() #выстрел себе в колено

    upload(cal,events) #загрузка всех заявок в календарь 

    logging.info('Finished')

if __name__ == '__main__':
    main()