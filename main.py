import sqlite3
import time
import sqlite3
from xlsxwriter.workbook import Workbook
import telebot

API_TOKEN = '6184698340:AAEvoX_JTrFBjNW27EEcTYY7h5lif98f-Ww'
DB_FILE_NAME = 'db.sqlite'


def create_connect():
    return sqlite3.connect(DB_FILE_NAME)


def init_db():
    # Создание базы и таблицы
    with create_connect() as connect:
        connect.execute('''
            CREATE TABLE IF NOT EXISTS Message (
                id      INTEGER  PRIMARY KEY,
                user_id INTEGER  NOT NULL,
                name TEXT NOT NULL,
                age INTEGER NOT NULL,
                type TEXT NOT NULL,
                text    TEXT  NOT NULL
            );
        ''')

        connect.commit()


def add_message(user_id, name, age, type, message):
    with create_connect() as connect:
        connect.execute(
            'INSERT INTO Message (user_id, name, age, type, text) VALUES (?, ?, ?, ?, ?)',
            (user_id, name, age, type, message)
        )
        connect.commit()


init_db()

bot = telebot.TeleBot(API_TOKEN)


@bot.message_handler(commands=['start'])
def start(message):
    bot.send_message(message.chat.id,
                     text="Здравствуйте, {0.first_name}! Я бот-помощник для сбора заявок в motivity. Введите через "
                          "пробел следующие данные: свою фамилию, возраст, тип обращения(заявка на "
                          "найм, консультацию или вопрос) и сопроводительный текст".format(
                         message.from_user))


@bot.message_handler(content_types=['text'])
def work(message):
    ans = message.text
    mas = ans.split()
    user_id = message.from_user.id

    name = mas[0]
    age = int(mas[1])
    type = mas[2]
    txt = mas[3]

    add_message(user_id, name, age, type, txt)
    stop_command(message)


def stop_command(message):
    bot.send_message(message.chat.id,
                     text="Спасибо за заявку, скоро мы с Вами свяжемся!".format(
                         message.from_user))
    export()
    bot.stop_polling()


def export():
    workbook = Workbook('base.xlsx')
    worksheet = workbook.add_worksheet()

    conn = sqlite3.connect('db.sqlite')
    c = conn.cursor()
    c.execute("select * from Message")
    mysel = c.execute("select * from Message")
    for i, row in enumerate(mysel):
        for j, value in enumerate(row):
            worksheet.write(i, j, value)
    workbook.close()


if __name__ == '__main__':
    while True:
        try:
            bot.polling(none_stop=True)
        except Exception as e:
            time.sleep(2)
            print(e)
