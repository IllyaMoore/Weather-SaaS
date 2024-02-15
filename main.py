import os
import telegram
from telegram.ext import CommandHandler, MessageHandler, filters, ConversationHandler, Updater
import openpyxl

# Змінна для зберігання даних стану розмови
START, CHOOSE_FILE, READ_FILE = range(3)


def start(update, context):
    update.message.reply_text("Привіт! Я Telegram бот для роботи з Excel таблицями. "
                              "Відправте /open щоб відкрити файл Excel.")

    return START


def open_excel(update, context):
    update.message.reply_text("Будь ласка, відправте Excel-файл, щоб його відкрити.")

    return CHOOSE_FILE


def read_excel(update, context):
    file = context.user_data.get("file")
    if file:
        workbook = openpyxl.load_workbook(file)
        sheet = workbook.active

        # Отримуємо дані з таблиці і відправляємо їх користувачеві
        data = "\n".join(["\t".join([str(cell.value) for cell in row]) for row in sheet.iter_rows()])
        update.message.reply_text(f"Зміст Excel-файлу:\n{data}")
    else:
        update.message.reply_text("Файл Excel не був обраний. Відправте /open щоб відкрити файл знову.")

    return ConversationHandler.END


def main():
    TOKEN = 'YOUR_BOT_TOKEN'
    updater = Updater(token=TOKEN, use_context=True)
    dispatcher = updater.dispatcher

    # Додаємо обробники команд
    start_handler = CommandHandler('start', start)
    open_excel_handler = CommandHandler('open', open_excel)
    read_excel_handler = MessageHandler(
        filters.document.mime_type("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"), read_excel)

    dispatcher.add_handler(start_handler)
    dispatcher.add_handler(open_excel_handler)
    dispatcher.add_handler(ConversationHandler(
        entry_points=[open_excel_handler],
        states={
            CHOOSE_FILE: [read_excel_handler]
        },
        fallbacks=[]
    ))

    updater.start_polling()
    updater.idle()


if __name__ == '__main__':
    main()
