import os
import logging
import win32com.client
from redis import Redis
from dotenv import load_dotenv
from datetime import datetime, timedelta

from accidents import normalize

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


    def get_accident(self, accident_id):
        accident = self.namespace.GetItemFromId(accident_id)
        logging.debug(f"get accident {accident_id}")
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


    def delete_all(self):
        accidents = self.calendar.Items
        while accidents.count > 0:
            accidents.Remove(1)

        logging.info("everything deleted")


def upload_calendar(calendar, accidents):
    
    for index, row in accidents.iterrows():
        start_date = datetime.strptime(row["start"], "%Y-%m-%d %H:%M:%S") + timedelta(hours=3)
        # start_date = row["start"]

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


def main():

    dotenv_path = os.path.join(os.path.dirname(__file__), '.env')
    
    if os.path.exists(dotenv_path):
        load_dotenv(dotenv_path)
    else:
        logging.CRITICAL('There is no .env file!')
        return False

    now = datetime.now().strftime('%Y-%m-%d-%H-%M-%S')
    filename = f'logs/{str(now)}.log'

    logging.basicConfig(filename=filename, level=logging.DEBUG)
    logging.info('Started')

    user_email = os.environ['USER_EMAIL']
    calendar_name = os.environ['CALENDAR_NAME']
    host = os.environ['REDIS_HOST']
    port = os.environ['REDIS_PORT']
    passwd = os.environ['REDIS_PASSWD']

    accident_calendar = Calendar(account=user_email, calendar_name=calendar_name, host=host, port=port, password=passwd)
    accidents = normalize() #получение таблицы всех заявок с полями 
    accident_calendar.update_categories(accidents=accidents) # обновление категорий в календаре чтобы отражать по цветам

    # cal.delete_all() #выстрел себе в колено

    upload_calendar(accident_calendar,accidents) #загрузка всех заявок в календарь 

if __name__ == '__main__':
    main()