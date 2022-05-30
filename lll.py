import telebot
from telebot import types
import random
import xlrd
from xlutils.copy import copy
import xlwt

loc = ('data/oyuun.xls')
bot = telebot.TeleBot("")

@bot.message_handler(commands=['start'])
def send_welcome(message):

	markup = types.ReplyKeyboardMarkup(resize_keyboard=True)

	bot.reply_to(message, "Добрый день, {0.first_name}\nДанный бот предназначен для нахождения палиндромов в xls файле \nДля начала, нажмите СТАРТ и введите слово Палиндром \nВыполнен для нахождения палиндромов произведения Ойуун Туулэ".format(message.from_user),parse_mode='html',reply_markup=markup)


def check_palindrome(string):
    l = len(string)
    textlowerer = string.lower()
    first = 0
    last = l - 1
    isPalindrome = True
    
    while first < last:
        if textlowerer[first] == textlowerer[last]:
                first = first + 1
                last = last - 1
        else:
                isPalindrome = False
                break
        
    return isPalindrome


@bot.message_handler(content_types=["text"])
def handle_text(message):
    rb = xlrd.open_workbook(loc)
    r_sheet = rb.sheet_by_index(0)
    wb = copy(rb)
    w_sheet = wb.get_sheet(0)
    colrow = r_sheet.nrows
    if message.text == "Палиндром":
            for i in range(colrow):
                    isPalindrome = check_palindrome(r_sheet.cell(i,0).value)
                    if isPalindrome:
                            w_sheet.write(i, 3, r_sheet.cell(i,0).value)
                            bot.send_message(message.chat.id, 'Полиндром найден: ' + r_sheet.cell(i,0).value)               

            bot.send_message(message.chat.id, 'Процесс окончен, сохраняем файл')
            wb.save(loc)
    else:
            bot.send_message(message.chat.id, 'Неизвестная команда')


bot.polling(none_stop=True)
