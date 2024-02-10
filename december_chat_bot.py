import logging
import openpyxl
from telegram import Update
from telegram.ext import Updater, CommandHandler, CallbackContext

# Telegram bot token
TELEGRAM_TOKEN = '5925649372:AAH6Bw7NdVrpb8iSdTcJpiUwPTkhXCOIZSw'
ALERT_GROUP_CHAT_ID = 'https://t.me/+5PfZ6ysH7CM5ZTRk'

# Excel file and sheet
EXCEL_FILE_PATH = r'C:\Users\philip.otieno\Desktop\Mappings_Nairobi\December_Data.xlsx'
EXCEL_SHEET_NAME = 'December 1st'

# Configure logging
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)

def start(update: Update, context: CallbackContext) -> None:
    update.message.reply_text('Hello! I am your December Data Bot. Type /get_data to retrieve the data.')

def get_data(update: Update, context: CallbackContext) -> None:
    try:
        logging.info('Attempting to load Excel sheet')
        # Load Excel sheet
        workbook = openpyxl.load_workbook(EXCEL_FILE_PATH)
        sheet = workbook[EXCEL_SHEET_NAME]

        # Extract data from the specific range
        data = []
        data.append('YEAR\tRETAIL RIDES\tGROWTH')
        for row in sheet['A2:C25']:
            data.append(f'{row[0].value}\t{row[1].value}\t{row[2].value}')

        # Send data to the Telegram chat
        logging.info('Sending data to the Telegram chat')
        context.bot.send_message(chat_id=update.effective_chat.id, text='\n'.join(data))
        logging.info('Data sent successfully')

        # Send the same data to the alert group
        context.bot.send_message(chat_id=ALERT_GROUP_CHAT_ID, text='\n'.join(data))

    except Exception as e:
        error_message = f'An error occurred: {str(e)}'
        logging.error(error_message)
        context.bot.send_message(chat_id=update.effective_chat.id, text=error_message)

        # Send the error message to the alert group
        context.bot.send_message(chat_id=ALERT_GROUP_CHAT_ID, text=error_message)

def main() -> None:
    updater = Updater(TELEGRAM_TOKEN, use_context=True)
    dispatcher = updater.dispatcher

    # Register handlers
    dispatcher.add_handler(CommandHandler('start', start))
    dispatcher.add_handler(CommandHandler('get_data', get_data))

    # Start the Bot
    updater.start_polling()
    updater.idle()

if __name__ == '__main__':
    main()
