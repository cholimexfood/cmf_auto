from termcolor import colored
import time
import pyodbc
import asyncio
from utils.send_telegram import send_message,send_message_async
import utils.set_mod
from colorama import Fore, Back, Style, init
init(autoreset=True) 
async def stored_procedure():
    conn_str = utils.set_mod.conn_str
    procedure_name = utils.set_mod.procedure_name
    print(Fore.MAGENTA+"--- Bước 3: Chạy Stored procedure:")
    print(Fore.BLUE+"Tự động chạy sau 5 giây.")
    await asyncio.sleep(5)
    try:
        print(Fore.BLUE+f"Đang chạy Stored procedure: {procedure_name}...")
        #asyncio.run_coroutine_threadsafe(send_message('dev',f'Đang chạy Script: {procedure_name}...'),loop)
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        cursor.execute(f"EXEC {procedure_name}")
        conn.commit()   
        print(Fore.GREEN+f"\rDone! Stored procedure: {procedure_name}!")
        await send_message_async('dev',f'\U0001F3C3 Script: {procedure_name} \U00002714',delay=2)
    except pyodbc.Error as e:
        print(Fore.RED+f"Lỗi khi chạy stored procedure: {e}")
        await send_message_async('dev',f'Lỗi khi chạy Script: {e}',delay=2)
    finally:
        cursor.close()
        conn.close()

async def run_stored_procedure():
    print(Style.BRIGHT + Fore.YELLOW +'---------- START APP RUN SCRIPTS ----------')
    await stored_procedure()
    print(Style.BRIGHT + Fore.YELLOW +'---------- STOP  APP RUN SCRIPTS ----------')
