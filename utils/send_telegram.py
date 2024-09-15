from aiogram import Bot, Dispatcher
from aiogram.exceptions import TelegramAPIError
from telegram.error import TelegramError
import asyncio
from utils.set_mod import load_config


config = load_config()
token_bot = config['token_bot']
bot = Bot(token=token_bot)

def send_message(mod, message_text):
    if mod == 'dev':
        chat_id = config['chat_id_dev']
    elif mod == 'user':
        chat_id = config['chat_id_user']
    else:
        print('mod send tin nhắn bị sai !')
        return    
    try:
        bot.send_message(chat_id=chat_id, text=message_text)
    except TelegramError as e:
        print(f"Failed to send message: {e}")
# Sử dụng cú pháp 'async with' để đảm bảo phiên làm việc với bot được đóng khi hoàn thành
async def send_message_async(mod, message_text, delay=0):
    config = load_config()
    if mod == 'dev':
        chat_id = config['chat_id_dev']
    elif mod == 'user':
        chat_id = config['chat_id_user']
    else:
        print('Mod gửi tin nhắn bị sai!')
        return

    try:
        async with Bot(token=token_bot) as bot:
            await bot.send_message(chat_id=chat_id, text=message_text, parse_mode='HTML')
            if delay > 0:
                await asyncio.sleep(delay)  # Thời gian chờ giữa các lần gửi tin nhắn
    except TelegramAPIError as e:
        if e.error_code == 429:  # Mã lỗi cho Flood Control
            retry_after = e.parameters.get('retry_after', 60)  # Thời gian Telegram yêu cầu chờ trước khi thử lại
            print(f"Flood control exceeded. Retrying after {retry_after} seconds...")
            await asyncio.sleep(retry_after)
            await send_message_async(mod, message_text, delay)  # Thử lại sau khi chờ
        else:
            print(f"Failed to send message: {e}")
    except Exception as e:
        print(f"Failed to send message: {e}")