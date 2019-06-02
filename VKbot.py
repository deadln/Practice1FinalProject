import vk_api
from vk_api.longpoll import VkLongPoll, VkEventType
from vk_api.utils import get_random_id
from vk_api.keyboard import VkKeyboard, VkKeyboardColor
import requests
from bs4 import BeautifulSoup
import xlrd
import pickle
import re
from datetime import datetime
from datetime import timedelta
import json

instruction = "Команды для работы с ботом:\n" 
instruction += "XXXX-XX-XX - сохранить номер своей группы\n"
instruction += "Бот - вывести расписание\n"
instruction += "Погода - вывести погоду\n"
instruction += "Помощь - вывести эту инструкцию снова"
instruction += "Бот день_недели - вывести расписание дня недели\n"
instruction += "Бот XXXX-XX-XX - вывести расписание данной группы\n"
instruction += "Бот день_недели XXXX-XX-XX - вывести расписание данной группы на данный день недели\n"

kurs = {
    "18": 1,
    "17": 2,
    "16": 3,
    "15": 4
    }

mesyacy = {
    1 : "января",
    2 : "февраля",
    3 : "марта",
    4 : "апреля",
    5 : "мая",
    6 : "июня",
    7 : "июля",
    8 : "августа",
    9 : "сентября",
    10 : "октября",
    11 : "ноября",
    12 : "декабря",
    }

days_of_week = {
    0 : "Понедельник",
    1 : "Вторник",
    2 : "Среда",
    3 : "Четверг",
    4 : "Пятница",
    5 : "Суббота",
    6 : "Воскресенье"

    }

weather_type = {
    "Thunderstorm" : "гроза",
    "Drizzle" : "изморось",
    "Rain" : "дождь",
    "Snow" : "снег",
    "Mist" : "туман",
    "Smoke" : "дымка",
    "Haze" : "туман",
    "Fog" : "туман",
    "Sand" : "песчаный вихрь",
    "Dust" : "пыльная буря",
    "Ash" : "вулканический пепел",
    "Squall" : "шквал",
    "Tornado" : "ураган",
    "Clear" : "ясно",
    "Clouds" : "облачно",

    }

weather_description = {
   "thunderstorm with light rain" : "Гроза с легким дождем", "thunderstorm with rain" : "Гроза с дождём", 
   "thunderstorm with heavy rain" : "Гроза с сильным дождем", "light thunderstorm" : "Небольшая гроза", 
   "thunderstorm" : "Гроза", "heavy thunderstorm" : "Сильная гроза", 
   "ragged thunderstorm" : "Рваная гроза", "thunderstorm with light drizzle" : "Гроза с легкой изморосью", 
   "thunderstorm with drizzle" : "Гроза с изморосью", "thunderstorm with heavy drizzle" : "Гроза с сильной изморосью", 
   "light intensity drizzle" : "Легкая изморось", "drizzle" : "Изморось", 
   "heavy intensity drizzle" : "Сильная изморось", "light intensity drizzle rain" : "Легкий моросящий дождь", 
   "drizzle rain" : "Моросящий дождь", "heavy intensity drizzle rain" : "Сильный моросящий дождь", 
   "shower rain and drizzle" : "Ливень и изморось", "heavy shower rain and drizzle" : "Сильный ливень и изморось", 
   "shower drizzle" : "Ливневая изморось", "light rain" : "Легкий дождь", 
   "moderate rain" : "Умеренный дождь", "heavy intensity rain" : "Интенсивный дождь", 
   "very heavy rain" : "Очень сильный дождь", "extreme rain" : "Сильный дождь", 
   "freezing rain" : "Ледяной дождь", "light intensity shower rain" : "Легкий ливневый дождь", 
   "shower rain" : "Ливневый дождь", "heavy intensity shower rain" : "Очень сильный ливневый дождь", 
   "ragged shower rain" : "Рваный ливневый дождь", "light snow	" : "Небольшой снегопад", 
   "Snow" : "Снегопад", "Heavy snow" : "Сильный снегопад", 
   "Sleet" : "Дождь со снегом", "Light shower sleet" : "Небольшой ливневый дождь со снегом", 
   "Shower sleet" : "Ливневый дождь со снегом", "Light rain and snow" : "Небольшой дождь и снег",
   "Rain and snow" : "Дождь и снег", "Light shower snow" : "Небольшой ливень снега", 
   "Shower snow" : "Ливень снега", "Heavy shower snow" : "Сильный ливень снега", 
   "mist" : "ТУман", "Smoke" : "Дымка", 
   "Haze" : "Туман", "sand/ dust whirls	" : "Песчаные / пыльные вихри", 
   "fog" : "Туман", "sand" : "Песчаный вихрь", 
   "dust" : "Пыльная буря", "volcanic ash" : "Вулканический пепел", 
   "squalls" : "Шквал", "tornado" : "Ураган", 
   "clear sky" : "Ясное небо", "few clouds: 11-25%" : "Мало облаков", 
   "scattered clouds: 25-50%" : "Рассеянные облака", "broken clouds: 51-84%" : "Рваные облака", 
   "overcast clouds: 85-100%" : "Пасмурные облака"
}

def get_bofort(speed):
    if speed <= 0.2:
        return "штиль"
    if speed <= 1.5:
        return "тихий"
    if speed <= 3.3:
        return "легкий"
    if speed <= 5.4:
        return "слабый"
    if speed <= 7.9:
        return "умеренный"
    if speed <= 10.7:
        return "свежий"
    if speed <= 13.8:
        return "сильный"
    if speed <= 17.1:
        return "крепкий"
    if speed <= 20.7:
        return "очень крепкий"
    if speed <= 24.4:
        return "шторм"
    if speed <= 28.4:
        return "сильный шторм(буря)"
    if speed <= 32.6:
        return "жестокий шторм"
    if speed > 32.6:
        return "ураган"

def get_rumb(angle):
    if angle <= 22.50:
        return "северный"
    if angle <= 67.50:
        return "северо-восточный"
    if angle <= 112.50:
        return "восточный"
    if angle <= 157.50:
        return "юго-восточный"
    if angle <= 202.50:
        return "южный"
    if angle <= 247.50:
        return "юго-западный"
    if angle <= 292.50:
        return "западный"
    if angle <= 337.50:
        return "северо-западный"
    if angle > 337.50:
        return "северный"


def main():
    #Получение файлов расписания
    page = requests.get("https://www.mirea.ru/education/schedule-main/schedule/")
    soup = BeautifulSoup(page.text, "html.parser")

    result = soup.find("div", {"id" : "toggle-3"}).findAll("a", {"class" : "xls"})

    for x in result:
        if "IIT" in str(x) and "vesna" in str(x):
            if "1k" in str(x):
                f = open("1kurs.xlsx", "wb")
                y = requests.get(x["href"])
                f.write(y.content)
            elif "2k" in str(x):
                f = open("2kurs.xlsx", "wb")
                y = requests.get(x["href"])
                f.write(y.content)
            elif "3k" in str(x):
                f = open("3kurs.xlsx", "wb")
                y = requests.get(x["href"])
                f.write(y.content)
            elif "4k" in str(x):
                f = open("4kurs.xlsx", "wb")
                y = requests.get(x["href"])
                f.write(y.content)

    #Инициализация бота
    vk_session = vk_api.VkApi(token = "44c700bc6db1238b1a3086147e2c55fb21e8a69d4441d1e45bcacc06fe16eb861643c34edc0aff39c2a1d")
    vk = vk_session.get_api()
    longpoll = VkLongPoll(vk_session)
    
    keyboard = VkKeyboard(one_time = True)
    keyboard.add_button("на сегодня", color = VkKeyboardColor.POSITIVE)
    keyboard.add_button("на завтра", color = VkKeyboardColor.NEGATIVE)
    keyboard.add_line()
    keyboard.add_button("на эту неделю", color = VkKeyboardColor.PRIMARY)
    keyboard.add_button("на следующую неделю", color = VkKeyboardColor.PRIMARY)
    keyboard.add_line()
    keyboard.add_button("какая неделя?", color = VkKeyboardColor.DEFAULT)
    keyboard.add_button("какая группа?", color = VkKeyboardColor.DEFAULT)

    try:
        with open("users.pickle", "rb") as f:
            userlist = pickle.load(f)
    except FileNotFoundError:
        userlist = {}

    userlist2 = {}
    custom_mark = {}
    for key in userlist.keys():
        custom_mark[key] = False
    
    for event in longpoll.listen():
        #Первое использование бота
        if event.type == VkEventType.MESSAGE_NEW and event.to_me and event.user_id not in userlist:
            vk.messages.send(
            user_id = event.user_id,
            random_id = get_random_id(),
            message = instruction
            ) 
            userlist[event.user_id] = ""
            userlist2[event.user_id] = ""
            with open("users.pickle", "wb") as f:
                pickle.dump(userlist, f)

        elif event.type == VkEventType.MESSAGE_NEW and event.to_me and event.text.lower() == "users":
            vk.messages.send(user_id = event.user_id,
                             random_id = get_random_id(),
                             message = str(userlist)
                             )

        elif event.type == VkEventType.MESSAGE_NEW and event.to_me and event.text.lower() == "помощь":
            vk.messages.send(user_id = event.user_id,
                             random_id = get_random_id(),
                             message = instruction
                             )

        elif event.type == VkEventType.MESSAGE_NEW and event.to_me and re.match(r"\w\w\w\w-\d\d-\d\d" ,event.text):
            vk.messages.send(
            user_id = event.user_id,
            random_id = get_random_id(),
            message = "Я запомнил, что ты из группы " + event.text.upper()
            )
            userlist[event.user_id] = event.text.upper()
            userlist2[event.user_id] = event.text.upper()
            custom_mark[event.user_id] = False
            with open("users.pickle", "wb") as f:
                pickle.dump(userlist, f)

        elif event.type == VkEventType.MESSAGE_NEW and event.to_me and event.text.lower() == "бот":
            vk.messages.send(user_id = event.user_id,
                             random_id = get_random_id(),
                             message = "Показать расписание...",
                             keyboard = keyboard.get_keyboard()
                             )

        elif event.type == VkEventType.MESSAGE_NEW and event.to_me and event.text.lower() == "на сегодня":
            if datetime.today().weekday() == 6:
                vk.messages.send(user_id = event.user_id,
                         random_id = get_random_id(),
                         message = str(datetime.today().day) + " " + mesyacy[datetime.today().month]+ " занятий нет"
                         )
            else:
                try:
                    if custom_mark[event.user_id]:
                        book = xlrd.open_workbook(str(kurs[userlist2[event.user_id][-2:]]) + "kurs.xlsx")
                    else:
                        book = xlrd.open_workbook(str(kurs[userlist[event.user_id][-2:]]) + "kurs.xlsx")
                    sheet = book.sheet_by_index(0)
                    num_cols = sheet.ncols
                    num_rows = sheet.nrows
                    found = False
                    ind = 0

                    for col_index in range(num_cols):
                        group_cell = str(sheet.cell(1,col_index).value)
                        if (custom_mark[event.user_id] and group_cell == userlist2[event.user_id]) or (not custom_mark[event.user_id] and group_cell == userlist[event.user_id]):
                            ind = col_index
                            found = True
                            break

                    if found:
                        schedule = "Расписание на " + str(datetime.today().day) + " " + mesyacy[datetime.today().month] 
                        schedule += "\n"
                        start_row = 3 + 12 * datetime.today().weekday()
                        if ((datetime.today() - datetime(2019,2,11)).days // 7 + 1) % 2 == 0:
                            start_row += 1
                        print(start_row)
                        for i in range(6):
                            if sheet.cell(start_row, ind).value == "":
                                schedule += str(i + 1) + ") -"
                            else:
                                schedule += str(i + 1) + ") " + sheet.cell(start_row, ind).value
                                if sheet.cell(start_row, ind + 1).value != "":
                                    schedule += ", " + sheet.cell(start_row, ind + 1).value
                                else:
                                    schedule += ", - "
                                if sheet.cell(start_row, ind + 2).value != "":
                                    schedule += ", " + sheet.cell(start_row, ind + 2).value
                                else:
                                    schedule += ", - "
                                if sheet.cell(start_row, ind + 3).value != "":
                                    schedule += ", " + sheet.cell(start_row, ind + 3).value
                                else:
                                    schedule += ", - "
                            schedule += "\n"
                            start_row += 2
                        vk.messages.send(user_id = event.user_id,
                                     random_id = get_random_id(),
                                         message = schedule
                            )

                    else:
                        if custom_mark[event.user_id]:
                            vk.messages.send(user_id = event.user_id,
                                random_id = get_random_id(),
                                message = "Некорректный номер группы - " + userlist2[event.user_id]
                                )
                        else:
                            vk.messages.send(user_id = event.user_id,
                                random_id = get_random_id(),
                                message = "Некорректный номер группы - " + userlist[event.user_id]
                                )


                except KeyError:
                    if custom_mark[event.user_id]:
                            vk.messages.send(user_id = event.user_id,
                                random_id = get_random_id(),
                                message = "Некорректный номер группы - " + userlist2[event.user_id]
                                )
                    else:
                        vk.messages.send(user_id = event.user_id,
                            random_id = get_random_id(),
                            message = "Некорректный номер группы - " + userlist[event.user_id]
                            )
                custom_mark[event.user_id] = False

        elif event.type == VkEventType.MESSAGE_NEW and event.to_me and event.text.lower() == "на завтра":
            tomorrow = datetime.today() + timedelta(days = 1)
            if tomorrow.weekday() == 6:
                vk.messages.send(user_id = event.user_id,
                             random_id = get_random_id(),
                             message = str(tomorrow.day) + " " + mesyacy[tomorrow.month]+ " занятий нет"
                             )
            else:
                try:
                    if custom_mark[event.user_id]:
                        book = xlrd.open_workbook(str(kurs[userlist2[event.user_id][-2:]]) + "kurs.xlsx")
                    else:
                        book = xlrd.open_workbook(str(kurs[userlist[event.user_id][-2:]]) + "kurs.xlsx")
                    sheet = book.sheet_by_index(0)
                    num_cols = sheet.ncols
                    num_rows = sheet.nrows
                    found = False
                    ind = 0

                    for col_index in range(num_cols):
                        group_cell = str(sheet.cell(1,col_index).value)
                        if (custom_mark[event.user_id] and group_cell == userlist2[event.user_id]) or (not custom_mark[event.user_id] and group_cell == userlist[event.user_id]):
                            ind = col_index
                            found = True
                            break

                    if found:
                        schedule = "Расписание на " + str(tomorrow.day) + " " + mesyacy[tomorrow.month] 
                        schedule += "\n"
                        start_row = 3 + 12 * tomorrow.weekday()
                        if ((tomorrow - datetime(2019,2,11)).days // 7 + 1) % 2 == 0:
                            start_row += 1
                        print(start_row)
                        for i in range(6):
                            if sheet.cell(start_row, ind).value == "":
                                schedule += str(i + 1) + ") -"
                            else:
                                schedule += str(i + 1) + ") " + sheet.cell(start_row, ind).value
                                if sheet.cell(start_row, ind + 1).value != "":
                                    schedule += ", " + sheet.cell(start_row, ind + 1).value
                                else:
                                    schedule += ", - "
                                if sheet.cell(start_row, ind + 2).value != "":
                                    schedule += ", " + sheet.cell(start_row, ind + 2).value
                                else:
                                    schedule += ", - "
                                if sheet.cell(start_row, ind + 3).value != "":
                                    schedule += ", " + sheet.cell(start_row, ind + 3).value
                                else:
                                    schedule += ", - "
                            schedule += "\n"
                            start_row += 2
                        vk.messages.send(user_id = event.user_id,
                                     random_id = get_random_id(),
                                         message = schedule
                            )

                    else:
                        if custom_mark[event.user_id]:
                            vk.messages.send(user_id = event.user_id,
                                random_id = get_random_id(),
                                message = "Некорректный номер группы - " + userlist2[event.user_id]
                                )
                        else:
                            vk.messages.send(user_id = event.user_id,
                                random_id = get_random_id(),
                                message = "Некорректный номер группы - " + userlist[event.user_id]
                                )


                except KeyError:
                    if custom_mark[event.user_id]:
                            vk.messages.send(user_id = event.user_id,
                                random_id = get_random_id(),
                                message = "Некорректный номер группы - " + userlist2[event.user_id]
                                )
                    else:
                        vk.messages.send(user_id = event.user_id,
                            random_id = get_random_id(),
                            message = "Некорректный номер группы - " + userlist[event.user_id]
                            )
                custom_mark[event.user_id] = False

        elif event.type == VkEventType.MESSAGE_NEW and event.to_me and event.text.lower() == "на эту неделю":
            try:
                    if custom_mark[event.user_id]:
                        book = xlrd.open_workbook(str(kurs[userlist2[event.user_id][-2:]]) + "kurs.xlsx")
                    else:
                        book = xlrd.open_workbook(str(kurs[userlist[event.user_id][-2:]]) + "kurs.xlsx")
                    sheet = book.sheet_by_index(0)
                    num_cols = sheet.ncols
                    num_rows = sheet.nrows
                    found = False
                    ind = 0

                    for col_index in range(num_cols):
                        group_cell = str(sheet.cell(1,col_index).value)
                        if (custom_mark[event.user_id] and group_cell == userlist2[event.user_id]) or (not custom_mark[event.user_id] and group_cell == userlist[event.user_id]):
                            ind = col_index
                            found = True
                            break

                    if found:
                        schedule = "Расписание на " + str((datetime.today() - datetime(2019,2,11)).days // 7 + 1)
                        schedule += " неделю\n"
                        start_row = 3
                        if ((datetime.today() - datetime(2019,2,11)).days // 7 + 1) % 2 == 0:
                            start_row += 1
                        print(start_row)
                        for j in range(6):
                            schedule += days_of_week[j] + "\n"
                            for i in range(6):
                                if sheet.cell(start_row, ind).value == "":
                                    schedule += str(i + 1) + ") -"
                                else:
                                    schedule += str(i + 1) + ") " + sheet.cell(start_row, ind).value
                                    if sheet.cell(start_row, ind + 1).value != "":
                                        schedule += ", " + sheet.cell(start_row, ind + 1).value
                                    else:
                                        schedule += ", - "
                                    if sheet.cell(start_row, ind + 2).value != "":
                                        schedule += ", " + sheet.cell(start_row, ind + 2).value
                                    else:
                                        schedule += ", - "
                                    if sheet.cell(start_row, ind + 3).value != "":
                                        schedule += ", " + sheet.cell(start_row, ind + 3).value
                                    else:
                                        schedule += ", - "
                                schedule += "\n"
                                start_row += 2
                            schedule += "\n"
                        vk.messages.send(user_id = event.user_id,
                                     random_id = get_random_id(),
                                         message = schedule
                            )
                    else:
                        if custom_mark[event.user_id]:
                            vk.messages.send(user_id = event.user_id,
                                random_id = get_random_id(),
                                message = "Некорректный номер группы - " + userlist2[event.user_id]
                                )
                        else:
                            vk.messages.send(user_id = event.user_id,
                                random_id = get_random_id(),
                                message = "Некорректный номер группы - " + userlist[event.user_id]
                                )
            except KeyError:
                if custom_mark[event.user_id]:
                            vk.messages.send(user_id = event.user_id,
                                random_id = get_random_id(),
                                message = "Некорректный номер группы - " + userlist2[event.user_id]
                                )
                else:
                    vk.messages.send(user_id = event.user_id,
                        random_id = get_random_id(),
                        message = "Некорректный номер группы - " + userlist[event.user_id]
                        )
            custom_mark[event.user_id] = False

        elif event.type == VkEventType.MESSAGE_NEW and event.to_me and event.text.lower() == "на следующую неделю":
            try:
                    if custom_mark[event.user_id]:
                        book = xlrd.open_workbook(str(kurs[userlist2[event.user_id][-2:]]) + "kurs.xlsx")
                    else:
                        book = xlrd.open_workbook(str(kurs[userlist[event.user_id][-2:]]) + "kurs.xlsx")
                    sheet = book.sheet_by_index(0)
                    num_cols = sheet.ncols
                    num_rows = sheet.nrows
                    found = False
                    ind = 0

                    for col_index in range(num_cols):
                        group_cell = str(sheet.cell(1,col_index).value)
                        if (custom_mark[event.user_id] and group_cell == userlist2[event.user_id]) or (not custom_mark[event.user_id] and group_cell == userlist[event.user_id]):
                            ind = col_index
                            found = True
                            break

                    if found:
                        schedule = "Расписание на " + str((datetime.today() - datetime(2019,2,11)).days // 7 + 2)
                        schedule += " неделю\n"
                        start_row = 3
                        if ((datetime.today() - datetime(2019,2,11)).days // 7 + 2) % 2 == 0:
                            start_row += 1
                        print(start_row)
                        for j in range(6):
                            schedule += days_of_week[j] + "\n"
                            for i in range(6):
                                if sheet.cell(start_row, ind).value == "":
                                    schedule += str(i + 1) + ") -"
                                else:
                                    schedule += str(i + 1) + ") " + sheet.cell(start_row, ind).value
                                    if sheet.cell(start_row, ind + 1).value != "":
                                        schedule += ", " + sheet.cell(start_row, ind + 1).value
                                    else:
                                        schedule += ", - "
                                    if sheet.cell(start_row, ind + 2).value != "":
                                        schedule += ", " + sheet.cell(start_row, ind + 2).value
                                    else:
                                        schedule += ", - "
                                    if sheet.cell(start_row, ind + 3).value != "":
                                        schedule += ", " + sheet.cell(start_row, ind + 3).value
                                    else:
                                        schedule += ", - "
                                schedule += "\n"
                                start_row += 2
                            schedule += "\n"
                        vk.messages.send(user_id = event.user_id,
                                     random_id = get_random_id(),
                                         message = schedule
                            )
                    else:
                        if custom_mark[event.user_id]:
                            vk.messages.send(user_id = event.user_id,
                                random_id = get_random_id(),
                                message = "Некорректный номер группы - " + userlist2[event.user_id]
                                )
                        else:
                            vk.messages.send(user_id = event.user_id,
                                random_id = get_random_id(),
                                message = "Некорректный номер группы - " + userlist[event.user_id]
                                )
            except KeyError:
                if custom_mark[event.user_id]:
                            vk.messages.send(user_id = event.user_id,
                                random_id = get_random_id(),
                                message = "Некорректный номер группы - " + userlist2[event.user_id]
                                )
                else:
                    vk.messages.send(user_id = event.user_id,
                        random_id = get_random_id(),
                        message = "Некорректный номер группы - " + userlist[event.user_id]
                        )
            custom_mark[event.user_id] = False

        elif event.type == VkEventType.MESSAGE_NEW and event.to_me and event.text.lower() == "какая неделя?":
            week = (datetime.today() - datetime(2019,2,11)).days // 7 + 1
            vk.messages.send(
            user_id = event.user_id,
            random_id = get_random_id(),
            message = "Идёт " + str(week) + " неделя"
            )

        elif event.type == VkEventType.MESSAGE_NEW and event.to_me and event.text.lower() == "какая группа?":
            vk.messages.send(
            user_id = event.user_id,
            random_id = get_random_id(),
            message = "Показываю расписание группы " + userlist[event.user_id]
            )

        elif event.type == VkEventType.MESSAGE_NEW and event.to_me and re.match(r"бот \w\w\w\w-\d\d-\d\d" ,event.text.lower()):
            userlist2[event.user_id] = re.findall(r"\w\w\w\w-\d\d-\d\d", event.text.lower())[0].upper()
            custom_mark[event.user_id] = True
            vk.messages.send(user_id = event.user_id,
                             random_id = get_random_id(),
                             message = "Показать расписание группы " + userlist2[event.user_id],
                             keyboard = keyboard.get_keyboard()
                             )

        elif event.type == VkEventType.MESSAGE_NEW and event.to_me and re.match(r"бот .+ \w\w\w\w-\d\d-\d\d" , event.text.lower()):
            dow = {}
            for key, value in days_of_week.items():
                dow[value.lower()] = key
            if re.findall(r"\w+",event.text.lower())[1] not in dow:
                 vk.messages.send(
                user_id = event.user_id,
                random_id = get_random_id(),
                message = "Неизвестная команда"
                )
            else:
                group = re.findall(r"\w\w\w\w-\d\d-\d\d",event.text.upper())[0]
                day = dow[re.findall(r"\w+",event.text.lower())[1]]
                if day == 6:
                    vk.messages.send(user_id = event.user_id,
                            random_id = get_random_id(),
                            message = "В воскресенье занятий нет"
                            )
                else:
                    try:
                        book = xlrd.open_workbook(str(kurs[group[-2:]]) + "kurs.xlsx")
                        sheet = book.sheet_by_index(0)
                        num_cols = sheet.ncols
                        num_rows = sheet.nrows
                        found = False
                        ind = 0

                        for col_index in range(num_cols):
                            group_cell = str(sheet.cell(1,col_index).value)
                            if group_cell == group:
                                ind = col_index
                                found = True
                                break

                        if found:
                            schedule = "Расписание на " + re.findall(r"\w+",event.text.lower())[1] 
                            schedule += "\n"
                            start_row = 3 + 12 * day
                            print(start_row)
                            schedule += "Нечётный\n"
                            for i in range(6):
                                if sheet.cell(start_row, ind).value == "":
                                    schedule += str(i + 1) + ") -"
                                else:
                                    schedule += str(i + 1) + ") " + sheet.cell(start_row, ind).value
                                    if sheet.cell(start_row, ind + 1).value != "":
                                        schedule += ", " + sheet.cell(start_row, ind + 1).value
                                    else:
                                        schedule += ", - "
                                    if sheet.cell(start_row, ind + 2).value != "":
                                        schedule += ", " + sheet.cell(start_row, ind + 2).value
                                    else:
                                        schedule += ", - "
                                    if sheet.cell(start_row, ind + 3).value != "":
                                        schedule += ", " + sheet.cell(start_row, ind + 3).value
                                    else:
                                        schedule += ", - "
                                schedule += "\n"
                                start_row += 2

                            start_row = 3 + 12 * day + 1
                            print(start_row)
                            schedule += "Чётный\n"
                            for i in range(6):
                                if sheet.cell(start_row, ind).value == "":
                                    schedule += str(i + 1) + ") -"
                                else:
                                    schedule += str(i + 1) + ") " + sheet.cell(start_row, ind).value
                                    if sheet.cell(start_row, ind + 1).value != "":
                                        schedule += ", " + sheet.cell(start_row, ind + 1).value
                                    else:
                                        schedule += ", - "
                                    if sheet.cell(start_row, ind + 2).value != "":
                                        schedule += ", " + sheet.cell(start_row, ind + 2).value
                                    else:
                                        schedule += ", - "
                                    if sheet.cell(start_row, ind + 3).value != "":
                                        schedule += ", " + sheet.cell(start_row, ind + 3).value
                                    else:
                                        schedule += ", - "
                                schedule += "\n"
                                start_row += 2
                            vk.messages.send(user_id = event.user_id,
                                     random_id = get_random_id(),
                                         message = schedule
                            )

                        else:
                            vk.messages.send(user_id = event.user_id,
                                random_id = get_random_id(),
                                message = "Некорректный номер группы - " + group
                                )


                    except KeyError:
                        vk.messages.send(user_id = event.user_id,
                            random_id = get_random_id(),
                            message = "Некорректный номер группы - " + group
                            )

        elif event.type == VkEventType.MESSAGE_NEW and event.to_me and re.match(r"бот .+" , event.text.lower()):
            dow = {}
            for key, value in days_of_week.items():
                dow[value.lower()] = key
            if re.findall(r"\w+",event.text.lower())[1] not in dow:
                 vk.messages.send(
                user_id = event.user_id,
                random_id = get_random_id(),
                message = "Неизвестная команда"
                )
            else:
                day = dow[re.findall(r"\w+",event.text.lower())[1]]
                if day == 6:
                    vk.messages.send(user_id = event.user_id,
                            random_id = get_random_id(),
                            message = "В воскресенье занятий нет"
                            )
                else:
                    try:
                        book = xlrd.open_workbook(str(kurs[userlist[event.user_id][-2:]]) + "kurs.xlsx")
                        sheet = book.sheet_by_index(0)
                        num_cols = sheet.ncols
                        num_rows = sheet.nrows
                        found = False
                        ind = 0

                        for col_index in range(num_cols):
                            group_cell = str(sheet.cell(1,col_index).value)
                            if group_cell == userlist[event.user_id]:
                                ind = col_index
                                found = True
                                break

                        if found:
                            schedule = "Расписание на " + re.findall(r"\w+",event.text.lower())[1] 
                            schedule += "\n"
                            start_row = 3 + 12 * day
                            print(start_row)
                            schedule += "Нечётный\n"
                            for i in range(6):
                                if sheet.cell(start_row, ind).value == "":
                                    schedule += str(i + 1) + ") -"
                                else:
                                    schedule += str(i + 1) + ") " + sheet.cell(start_row, ind).value
                                    if sheet.cell(start_row, ind + 1).value != "":
                                        schedule += ", " + sheet.cell(start_row, ind + 1).value
                                    else:
                                        schedule += ", - "
                                    if sheet.cell(start_row, ind + 2).value != "":
                                        schedule += ", " + sheet.cell(start_row, ind + 2).value
                                    else:
                                        schedule += ", - "
                                    if sheet.cell(start_row, ind + 3).value != "":
                                        schedule += ", " + sheet.cell(start_row, ind + 3).value
                                    else:
                                        schedule += ", - "
                                schedule += "\n"
                                start_row += 2

                            start_row = 3 + 12 * day + 1
                            print(start_row)
                            schedule += "Чётный\n"
                            for i in range(6):
                                if sheet.cell(start_row, ind).value == "":
                                    schedule += str(i + 1) + ") -"
                                else:
                                    schedule += str(i + 1) + ") " + sheet.cell(start_row, ind).value
                                    if sheet.cell(start_row, ind + 1).value != "":
                                        schedule += ", " + sheet.cell(start_row, ind + 1).value
                                    else:
                                        schedule += ", - "
                                    if sheet.cell(start_row, ind + 2).value != "":
                                        schedule += ", " + sheet.cell(start_row, ind + 2).value
                                    else:
                                        schedule += ", - "
                                    if sheet.cell(start_row, ind + 3).value != "":
                                        schedule += ", " + sheet.cell(start_row, ind + 3).value
                                    else:
                                        schedule += ", - "
                                schedule += "\n"
                                start_row += 2
                            vk.messages.send(user_id = event.user_id,
                                     random_id = get_random_id(),
                                         message = schedule
                            )

                        else:
                            vk.messages.send(user_id = event.user_id,
                                random_id = get_random_id(),
                                message = "Некорректный номер группы - " + userlist[event.user_id]
                                )


                    except KeyError:
                        vk.messages.send(user_id = event.user_id,
                            random_id = get_random_id(),
                            message = "Некорректный номер группы - " + userlist[event.user_id]
                            )

        elif event.type == VkEventType.MESSAGE_NEW and event.to_me and event.text.lower() == "погода":
            w = requests.get("http://api.openweathermap.org/data/2.5/weather?q=moscow&appid=9b39cc1cf3f2f9d9556efd5a0acbb4f2&units=metric")
            weather = w.json()
            output = "Погода в Москве: " + weather_type[weather["weather"][0]["main"]] + "\n"
            output += weather_description[weather["weather"][0]["description"]] + ", температура: "
            output += str(int(weather["main"]["temp_min"])) + "-" + str(int(weather["main"]["temp_max"])) + "˚C\n"
            output += "Давление: " + str(weather["main"]["pressure"]) + " мм рт. ст., влажность: "
            output += str(weather["main"]["humidity"]) + "%\n"
            output += "Ветер: " + get_bofort(weather["wind"]["speed"]) + ", " + str(weather["wind"]["speed"]) + "м/с, "
            output += get_rumb(weather["wind"]["deg"])
            vk.messages.send(
            user_id = event.user_id,
            random_id = get_random_id(),
            message = output
            )

        elif event.type == VkEventType.MESSAGE_NEW and event.to_me:
            vk.messages.send(
            user_id = event.user_id,
            random_id = get_random_id(),
            message = "Неизвестная команда"
            ) 

if __name__ == "__main__":
    main()

