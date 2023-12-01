# sd-to-exchange
MIni script that transfer data from Redmine and push it to exchange calendar




# how to

Что делает этот скрипт:

1. У нас есть Redmine в котором есть разные мероприятия и заявки. 

2. Через API Redmine получаем даныне по заявкам и формируем из них Dataframe.

3. Далее используя win32com и установленный Outlook на компьютере заходим под нужной учетной записью и создаем календарь
    ```python
    upload_calendar
    ```
4. Идем по строкам в Dataframe с заявками и создаем новые события в кадендарь Outlook
