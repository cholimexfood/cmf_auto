#main.py
import asyncio
import keyboard
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes, Application
from utils.send_telegram import load_config, send_message, send_message_async
from utils.app_status_order import run_app_check_order
from utils.app_download import run_main_download_async
from utils.app_copy import run_main_copy_async
from utils.app_import import run_main_import_async
from utils.app_runscripts import run_stored_procedure
import nest_asyncio
from datetime import datetime, timedelta
import warnings
from utils.set_mod import load_config 
from colorama import Fore, Back, Style, init
init(autoreset=True) 
# Bỏ qua cảnh báo không cần thiết
warnings.filterwarnings("ignore", category=UserWarning, module='telegram.ext._applicationbuilder')

info_bot = '[Bot Ver2.3]:[Bé Na \U0001F40D]'
config = load_config()
number = int(config['number_repeat_order'])

async def auto_status_order(max_check):
    try:
        await send_message_async('dev', '--- [Auto]:[Stt_Order]:[Start] -',delay=2)
        kq = await run_app_check_order(max_check)
        await send_message_async('dev', '--- [Auto]:[Stt_Order]:[End] -',delay=2)
    except Exception as e:
        
        await send_message_async('dev', f'[ERROR] [Auto]:[Stt_Order] - {e}',delay=2)
    return kq  # Giả định

async def auto_sql():
    try:
        await send_message_async('dev', '--- [Auto]:[Download]:[Start] -',delay=2)
        await run_main_download_async()
        await send_message_async('dev', '--- [Auto]:[Download]:[End] -',delay=2)

        await send_message_async('dev', '--- [Auto]:[Form]:[Start] -',delay=2)
        await run_main_copy_async()
        await send_message_async('dev', '--- [Auto]:[Form]:[End] -',delay=2)

        await send_message_async('dev', '--- [Auto]:[Import]:[Start] -',delay=2)
        await run_main_import_async()
        await send_message_async('dev', '--- [Auto]:[Import]:[End] -',delay=2)
    except Exception as e:
        await send_message_async('dev', f'[ERROR] [Auto]:[SQL] - {e}',delay=2)

async def manual_sql(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if auto_running:
        await update.message.reply_text(f'{info_bot}\n\U0001F4E2 "Auto" đang chạy. Không nhận lệnh')
        return    
    try:
        await update.message.reply_text(f'{info_bot}\n- Khởi chạy App: Manual xong !')
        await send_message_async('dev', '--- [Manual]:[Download]:[Start] -',delay=2)
        await run_main_download_async()
        await send_message_async('dev', '--- [Manual]:[Download]:[End] -',delay=2)

        await send_message_async('dev', '--- [Manual]:[Form]:[Start] -',delay=2)
        await run_main_copy_async()
        await send_message_async('dev', '--- [Manual]:[Form]:[End] -',delay=2)

        await send_message_async('dev', '--- [Manual]:[Import]:[Start] -',delay=2)
        await run_main_import_async()
        await send_message_async('dev', '--- [Manual]:[Import]:[End] -',delay=2)
    except Exception as e:
        await update.message.reply_text(f'{info_bot}\n[ERROR] - {e}')


async def auto_tool():
    global auto_running
    try:
        auto_running = True
        print(Fore.BLUE + "Nhấn 'N' trong 3 giây để bỏ qua Auto (N):", end='', flush=True)

        key_pressed = False

        def check_key_event():
            nonlocal key_pressed
            if keyboard.is_pressed('n'):
                key_pressed = True
                # Hiển thị ký tự 'N' sau dòng thông báo
                print(Fore.GREEN + ' N', end='', flush=True)

        # Trong khi chờ 3 giây, kiểm tra xem người dùng có nhấn 'N' không
        for _ in range(30):  # 3 giây chia thành 30 lần lặp
            await asyncio.sleep(0.1)
            check_key_event()
            if key_pressed:
                print(Fore.GREEN + "\nĐã bỏ qua chế độ Auto !")
                return

        # Nếu không nhấn 'N', tiếp tục chế độ Auto
        print(Fore.GREEN + "\nTiếp tục chế độ Auto !")

        await send_message_async('dev', '\U00002708\U00002708\U00002708 - Begin "Auto" - \U00002708\U00002708\U00002708',delay=2)

        result = await auto_status_order(number)

        target_time_str = config['time_auto']
        target_time = datetime.strptime(target_time_str, '%H:%M').time()
        now = datetime.now().time()

        if now > target_time:
            target_datetime = datetime.combine(datetime.today() + timedelta(days=1), target_time)
        else:
            target_datetime = datetime.combine(datetime.today(), target_time)

        dem_nguoc = (target_datetime - datetime.now()).total_seconds()
        dem_nguoc_end = dem_nguoc + (30*60)
        dem_nguoc_str = f"{int(dem_nguoc // 3600):02} tiếng {int((dem_nguoc % 3600) // 60):02} phút"
        dem_nguoc_end_str = f"{int(dem_nguoc_end // 3600):02} tiếng {int((dem_nguoc_end % 3600) // 60):02} phút"

        if result == 1:
            await send_message_async('dev', f'\U0001F4E2\U0001F4E2\U0001F4E2 - Alert "Auto" - \U0001F4E2\U0001F4E2\U0001F4E2\nĐã duyệt xong đơn - Tiếp tục "Auto" sau: {dem_nguoc_str}',delay=2)
            await asyncio.sleep(dem_nguoc)
            await auto_sql()
            await send_message_async('dev', '\U0001F6A9\U0001F6A9\U0001F6A9 - Finish "Auto" - \U0001F6A9\U0001F6A9\U0001F6A9',delay=2)
        else:
            if dem_nguoc_end > (60*60):
                await send_message_async('dev', '\U0001F4E2\U0001F4E2\U0001F4E2 - Caution "Auto" - \U0001F4E2\U0001F4E2\U0001F4E2\nĐơn hàng vẫn chưa duyệt hết - hãy kiểm tra',delay=2)
                await send_message_async('dev', '[ERROR]\nRecommended: Duyệt hết đơn -> Khởi chạy Manual\n\U0001F6A9\U0001F6A9\U0001F6A9 - End "Auto" - \U0001F6A9\U0001F6A9\U0001F6A9',delay=2)
                return
            await send_message_async('dev', f'\U0001F4E2\U0001F4E2\U0001F4E2 - Alert "Auto" - \U0001F4E2\U0001F4E2\U0001F4E2\nKiểm lại đơn hàng "lần cuối" sau: {dem_nguoc_end_str}',delay=2)
            await asyncio.sleep(dem_nguoc_end)
            result = await auto_status_order(1)
            if result == 1:
                await auto_sql()
                await send_message_async('dev', '\U0001F6A9\U0001F6A9\U0001F6A9 - Finish "Auto" - \U0001F6A9\U0001F6A9\U0001F6A9',delay=2)
            else:
                await send_message_async('dev', '\U0001F4E2\U0001F4E2\U0001F4E2 - Caution "Auto" - \U0001F4E2\U0001F4E2\U0001F4E2\nĐơn hàng vẫn chưa duyệt hết - hãy kiểm tra',delay=2)
                await send_message_async('dev', '[ERROR]\nRecommended: Duyệt hết đơn -> Khởi chạy Manual\n\U0001F6A9\U0001F6A9\U0001F6A9 - End "Auto" - \U0001F6A9\U0001F6A9\U0001F6A9',delay=2)
    except Exception as e:
        await send_message_async('dev', f'[ERROR] [Auto Tool] - {e}',delay=2)
    finally:
        auto_running = False

async def ktttdh_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if auto_running:
        await update.message.reply_text(f'{info_bot}\n\U0001F4E2 "Auto" đang chạy. Không nhận lệnh')
        return
    try:
        await update.message.reply_text(f'{info_bot}\n- Khởi chạy App: Stt_Order xong !')
        await send_message_async('dev', '--- [Manual]:[Stt_Order]:[Start] -',delay=2)
        await run_app_check_order(1)
        await send_message_async('dev', '--- [Manual]:[Stt_Order]:[End] -',delay=2)
    except Exception as e:
        await update.message.reply_text(f'{info_bot}\n[ERROR] - {e}')

async def downrp_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if auto_running:
        await update.message.reply_text(f'{info_bot}\n\U0001F4E2 "Auto" đang chạy. Không nhận lệnh')
        return
    try:
        await update.message.reply_text(f'{info_bot}\n- Khởi chạy App: Download xong !')
        await send_message_async('dev', '--- [Manual]:[Download]:[Start] -',delay=2)
        await run_main_download_async()
        await send_message_async('dev', '--- [Manual]:[Download]:[End] -',delay=2)
    except Exception as e:
        await update.message.reply_text(f'{info_bot}\n[ERROR] - {e}')

async def copytoform_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if auto_running:
        await update.message.reply_text(f'{info_bot}\n\U0001F4E2 "Auto" đang chạy. Không nhận lệnh')
        return    
    try:
        await update.message.reply_text(f'{info_bot}\n- Khởi chạy App: Form xong !')
        await send_message_async('dev', '--- [Manual]:[Form]:[Start] -',delay=2)
        await run_main_copy_async()
        await send_message_async('dev', '--- [Manual]:[Form]:[End] -',delay=2)
    except Exception as e:
        await update.message.reply_text(f'{info_bot}\n[ERROR] - {e}')

async def importtosql_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if auto_running:
        await update.message.reply_text(f'{info_bot}\n\U0001F4E2 "Auto" đang chạy. Không nhận lệnh')
        return
    try:
        await update.message.reply_text(f'{info_bot}\n- Khởi chạy App: Import xong !')
        await send_message_async('dev', '--- [Manual]:[Import]:[Start] -',delay=2)
        await run_main_import_async()
        await send_message_async('dev', '--- [Manual]:[Import]:[End] -',delay=2)
    except Exception as e:
        await update.message.reply_text(f'{info_bot}\n[ERROR] - {e}')

async def runscript_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if auto_running:
        await update.message.reply_text(f'{info_bot}\n\U0001F4E2 "Auto" đang chạy. Không nhận lệnh')
        return
    try:
        await update.message.reply_text(f'{info_bot}\n- Khởi chạy App: RunScript xong !')
        await send_message_async('dev', '--- [Manual]:[Script]:[Start] -',delay=2)
        await run_stored_procedure()
        await send_message_async('dev', '--- [Manual]:[Script]:[End] -',delay=2)
    except Exception as e:
        await update.message.reply_text(f'{info_bot}\n[ERROR] - {e}')


async def hi_start_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if auto_running:
        await update.message.reply_text(f'{info_bot}\n\U0001F4E2 "Auto" đang chạy. Không nhận lệnh')
        return
    await update.message.reply_text(f'{info_bot}\nXin chào ! \U0000270C\nTôi là Bé Na, tôi được tạo bởi Mr Trọng.\nTính năng của tôi ở nút Menu nhé !')

def add_handlers(application: Application):
    command_mapping = {
        'start': hi_start_command,
        'hi': hi_start_command,
        'ktttdh': ktttdh_command,
        'downrp': downrp_command,
        'copytoform': copytoform_command,
        'importtosql': importtosql_command,
        'runscript':runscript_command,
        'manualsql':manual_sql,
    }

    for command, handler in command_mapping.items():
        application.add_handler(CommandHandler(command, handler))

async def main():
    print('Tool đang chạy...')
    time_now = datetime.now()
    time_now_str = time_now.strftime("%H:%M %d/%M/%Y")
    await send_message_async('dev', f'{info_bot}\n[{time_now_str}]\nBé Na, Xin chào ! \U0000270C\nTôi sẽ bắt đầu nhiệm vụ "Auto" hôm nay.',delay=2)

    try:
        config = load_config()
        token_bot = config['token_bot']
        application = ApplicationBuilder().token(token_bot).build()

        add_handlers(application)
    
        # Khởi chạy công việc tự động bằng asyncio.create_task
        run_auto = asyncio.create_task(auto_tool())
 
        # Sử dụng asyncio.gather để gom các công việc bất đồng bộ
        await asyncio.gather(
            application.run_polling(),
            run_auto
        )
   
        '''
        await auto_tool()
        await application.run_polling()
        '''
    except KeyboardInterrupt:
        print("Chương trình bị dừng bởi người dùng (Ctrl+C). Đang tắt bot...")

    except Exception as e:
        print(f"Đã xảy ra lỗi khi khởi động bot: {e}")

if __name__ == "__main__":
    nest_asyncio.apply()
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("Chương trình đã dừng lại thành công.")