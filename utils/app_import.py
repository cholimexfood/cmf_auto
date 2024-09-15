import pandas as pd
import pyodbc
import os
import numpy as np
import time,asyncio
from termcolor import colored
import utils.set_mod
from utils.send_telegram import  send_message, send_message_async
#import app_setup #chạy chương trình lẻ
from colorama import Fore, Back, Style, init
init(autoreset=True) 
output_directory = os.path.join(utils.set_mod.current_directory, 'data_import')
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

def replace_nan_with_null(df):
    """Thay thế tất cả giá trị NaN bằng None (NULL trong SQL)."""
    return df.replace({np.nan: None})

def format_datetime(df, columns):
    """Định dạng các cột ngày tháng năm giờ cho phù hợp với kiểu smalldatetime."""
    for column in columns:
        if column in df.columns:
            # Chuyển đổi định dạng ngày giờ
            df[column] = pd.to_datetime(df[column], format='%d/%m/%Y %H:%M:%S').dt.strftime('%Y/%m/%d')
    return df

async def excel_to_csv(excel_files):
    print(Fore.MAGENTA+"--- Bước 1: Convert File import:")
    csv_files = []
    for excel_file in excel_files:
        csv_file = os.path.splitext(excel_file)[0] + '.csv' 
        csv_name = os.path.basename(os.path.splitext(excel_file)[0] + '.csv')
        """
        # Kiểm tra xem file CSV đã tồn tại chưa
        if os.path.exists(csv_file):
            print(colored(f"File CSV {csv_file} đã tồn tại, không cần tạo lại.", "red"))
            csv_files.append(csv_file)
            continue
        """
        print(Fore.BLUE+f"Đang tạo file CSV {csv_name}...")
        #asyncio.run_coroutine_threadsafe(send_message('dev',f'Đang tạo CSV từ {excel_file}...'),loop)
        df = pd.read_excel(excel_file)
        df = replace_nan_with_null(df)
        datetime_columns = ['NGAY_DAT', 'NGAY_DUYET', 'NGAY_GIAO']  # Cập nhật tên cột tương ứng
        df = format_datetime(df, datetime_columns)
        df.to_csv(csv_file, index=False) #to_csv mặc định sẽ ghi đè lên file cũ
        print(Fore.GREEN+f"\rDone! Lưu File: {csv_name} thành công!")
        await send_message_async('dev',f'\U0001F4DD {csv_name} \U00002714',delay=2)
        csv_files.append(csv_file)
    return csv_files

def get_column_names(cursor, table_name):
    cursor.execute(f"SELECT TOP 0 * FROM {table_name}")
    return [desc[0] for desc in cursor.description]

async def import_csv_to_sql(csv_files, conn_str):
    print(Fore.MAGENTA+"--- Bước 2: Xóa dữ liệu và import:")
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()
    cursor.fast_executemany = True
    for csv_file in csv_files:
        success_count = 0
        table_name = os.path.splitext(os.path.basename(csv_file))[0]
        csv_name = table_name + '.csv'
        print(Fore.BLUE+f"Đang xóa bảng {table_name}...")
        try:
 
            cursor.execute(f"TRUNCATE TABLE {table_name}")
            conn.commit()
            print(Fore.GREEN+f"\rĐã xóa bảng {table_name} thành công!")
    
        except pyodbc.Error as e:
            print(Fore.RED+f"Lỗi khi xóa bảng {table_name}: {e}")
            continue

        print(Fore.BLUE+f"Đang đọc file {csv_name}...")
        df = pd.read_csv(csv_file, low_memory=False)
        df = replace_nan_with_null(df)
        print(Fore.GREEN+f"\rĐọc xong file {csv_name} thành công!")
        
        sql_columns = get_column_names(cursor, table_name)
        df = df[[col for col in df.columns if col in sql_columns]]

        columns = ', '.join(df.columns)
        placeholders = ', '.join(['?'] * len(df.columns))
        insert_query = f"INSERT INTO {table_name} ({columns}) VALUES ({placeholders})"

        print(Fore.BLUE+f"\rĐang import file {csv_name} vào bảng {table_name}...")
        #asyncio.run_coroutine_threadsafe(send_message('dev',f'Đang import {table_name}...'),loop)
        data = df.values.tolist()

        try:
        
            batch_size = 7000  # Tăng kích thước batch
            for i in range(0, len(data), batch_size):
                batch = data[i:i + batch_size]
                cursor.executemany(insert_query, batch)
                conn.commit()
                success_count += len(batch)
        except pyodbc.Error as e:
            print(Fore.RED+f"Lỗi khi nhập dữ liệu: {e}")
            print(Fore.RED+f"Dữ liệu lỗi: {batch}")
            await send_message_async('dev',f'Error: Lỗi import {table_name}: {e}',delay=2)
            return       
        print(Fore.GREEN+f"\rImport file {csv_name} vào {table_name} thành công! (Tổng số dòng import:{success_count})")
        print(Fore.YELLOW +'---')
        await send_message_async('dev',f'\U00002705 {table_name}:{success_count}',delay=2)
    cursor.close()
    conn.close()

async def run_stored_procedure(conn_str, procedure_name):
    print(Fore.MAGENTA+"--- Bước 3: Chạy Stored procedure:")
    print(Fore.BLUE+"Tự động chạy sau 5 giây.")
    await asyncio.sleep(5)
    """ 
    user_input = input(colored("\nBạn có muốn chạy Stored procedure? (Y/N): ", 'cyan')).strip().upper()
    if user_input != 'Y':
        print(colored("Đã bỏ qua việc chạy stored procedure.", "yellow"))
        return
    """
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
        #sys.exit()

async def importtosql():
    excel_files = [
        os.path.join(output_directory, 'tbl_tuyen.xlsx'),
        os.path.join(output_directory, 'tbl_khachhang.xlsx'),
        os.path.join(output_directory, 'tbl_donhang_import.xlsx')
    ]
    missing_files = [file for file in excel_files if not os.path.exists(file)]
    if missing_files:
        print(Fore.RED+"Các file sau không tồn tại trong thư mục dự án:")
        await send_message_async('dev',f'Không tồn tại File Excel:',delay=2)
        for file in missing_files:
            print(Fore.RED+f" {file}")
            await send_message_async('dev',f'-{file}',delay=2)
            await asyncio.sleep(2)
        return
    else:
        csvs_files = await excel_to_csv(excel_files)
        await import_csv_to_sql(csvs_files, utils.set_mod.conn_str)
        await run_stored_procedure(utils.set_mod.conn_str, utils.set_mod.procedure_name)

async def run_main_import_async():
    print(Style.BRIGHT + Fore.YELLOW +'---------- START APP CONVERT AND IMPORT TO SQL ----------')
    await importtosql()
    print(Style.BRIGHT + Fore.YELLOW +'---------- STOP  APP CONVERT AND IMPORT TO SQL ----------')

if __name__ == "__main__":
    asyncio.run(run_main_import_async())

