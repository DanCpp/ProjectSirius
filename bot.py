import telebot
from telebot import TeleBot
from telebot import types
from openpyexcel import load_workbook

Book = load_workbook("Train.xlsx")

rev_sheet = Book["Rev"]

revs = []
revs_ton = []

count = 0

for item in rev_sheet:
    count += 1


for i in range(2, count + 1):
    text = ''.join(rev_sheet["B" + str(i)].value.replace(u'\xa0', u' '))
    if len(text) > 512:
        text = text[:(512-len(text))]
    revs.append(text)
    revs_ton.append(''.join(rev_sheet["C" + str(i)].value))

bot = TeleBot("")

idx = 0


@bot.message_handler(content_types='text')
def message_reply(message):
    bot.send_message(message.chat.id, "Я занят")

@bot.message_handler(commands = ["review"])
def review(message):
    idx = 0
    markup = types.InlineKeyboardMarkup(row_width=2)
    left = types.InlineKeyboardButton("<", callback_data="Left")
    right = types.InlineKeyboardButton(">", callback_data="Right")
    markup.add(left, right)
    bot.send_message(message.chat.id, "".join([revs[idx], revs_ton[idx]]), reply_markup=markup)

@bot.callback_query_handler(func = lambda call : True)
def callback(call):
    if(call.message):
        markup = types.InlineKeyboardMarkup(row_width=2)
        left = types.InlineKeyboardButton("<", callback_data="Left")
        right = types.InlineKeyboardButton(">", callback_data="Right")
        markup.add(left, right)
        if (call.data == "Left" and idx > 0):
            idx -= 1
            bot.send_message(call.message.chat.id, "".join([revs[idx], revs_ton[idx]]), reply_markup=markup)
        elif(call.data == "Right" and idx < len(revs)):
            idx+=1
            bot.send_message(call.message.chat.id, "".join([revs[idx], revs_ton[idx]]), reply_markup=markup)

bot.polling()