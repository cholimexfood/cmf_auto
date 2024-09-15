#app_status_order.py
from termcolor import colored
from datetime import datetime, timedelta
import os, sys,warnings,time
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import pandas as pd
from pathlib import Path
import utils.send_telegram
import utils.set_mod
from utils.send_telegram import  send_message,load_config, send_message_async
import asyncio
from colorama import Fore, Back, Style, init
init(autoreset=True) 
# Sử dụng driver từ set_mod.py
driver = utils.set_mod.get_driver()

# so_ngay_check = 0 nghĩa là kiểm tra 1 ngày trước đó
so_ngay_check = 6 
tt_app ='Thông tin App:'
nd_app ='Kiểm tra trạng thái đơn hàng:'
thoigian_tai = ''

alert_mod_dev = utils.set_mod.alert_mod_dev
alert_mod_user = utils.set_mod.alert_mod_user
info_app_dev = utils.set_mod.info_app_dev
info_app_user = utils.set_mod.info_app_user
info_app_error = utils.set_mod.info_app_error

delay_minute = 20
'''
nhap = '2024-08-21 11:57:26.888950'
datetime_object = datetime.strptime(nhap, '%Y-%m-%d %H:%M:%S.%f')
to_date = datetime_object -timedelta(days=1)
'''
to_date = datetime.now() -timedelta(days=1)
from_date = to_date - timedelta(days=so_ngay_check)
from_date_str = from_date.strftime('%d/%m/%Y')
to_date_str = to_date.strftime('%d/%m/%Y')


config = load_config()

def check_and_remove_old_files(statuses):
    for filename in os.listdir(utils.set_mod.download_directory):
        if any(filename.startswith(status) for status in statuses) and filename.endswith(".xlsx"):
            file_path = os.path.join(utils.set_mod.download_directory, filename)
            os.remove(file_path)
            #print(f"\nĐã xóa file {file_path}")
        else:
            #print('\nKhông có file để xóa !')
            return

def download_and_rename_file(expected_name, download_directory):
    wait_time = 300  # Thời gian chờ tối đa (giây)
    start_time = time.time()
    # Lưu danh sách các tệp trước khi tải xuống
    before_download_files = set(os.listdir(download_directory))
    while True:
        # Lấy danh sách các tệp trong thư mục tải xuống
        files = set(os.listdir(download_directory))
        # Tìm tệp mới được tải về
        new_files = files - before_download_files
        for filename in new_files:
            if filename.endswith(".xlsx"):
                old_file_path = os.path.join(download_directory, filename)
                new_file_path = os.path.join(download_directory, f"{expected_name}.xlsx")
                try:
                    # Đổi tên tệp
                    os.rename(old_file_path, new_file_path)
                    #print(f"\nFile đã được đổi tên thành {expected_name}.xlsx")
                    return 1
                except Exception as e:
                    print( Fore.RED + f"Lỗi khi đổi tên tệp: {e}")
                    return 0
        # Kiểm tra thời gian chờ
        if time.time() - start_time > wait_time:
            print(Fore.RED +"Lỗi! Hết thời gian chờ tải")
            return 0
        # Chờ một chút trước khi kiểm tra lại
        time.sleep(1)

async def check_status_order(trangthai):
    co_kiem_tra = 1
    co_tai_ve = 1
    try:
        print(Fore.BLUE +f'Bắt đầu kiểm tra {trangthai}...')        # Mở trang web và đăng nhập
        driver.get(config['web'])
        driver.implicitly_wait(1)

        username = driver.find_element(By.ID, "username")
        password = driver.find_element(By.ID, "password")
        username.send_keys(config['username'])
        password.send_keys(config['password'])
        password.send_keys(Keys.RETURN)
        driver.implicitly_wait(1)

        # Thao tác với các menu trên trang web
        banhang = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="424"]/a'))
            )
        banhang.click()

        donhang = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="child_922"]'))
            )
        donhang.click()

        time.sleep(10)

        dropdown = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'span.ui-dropdownchecklist-text'))
        )
        dropdown.click()

        dropdown_items = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div.ui-dropdownchecklist-item'))
        )
        values_to_select = ['IN', 'SO']
        for item in dropdown_items:
            checkbox = item.find_element(By.CSS_SELECTOR, 'input[type="checkbox"]')
            value = checkbox.get_attribute('value')
            if value in values_to_select:
                if not checkbox.is_selected():
                    checkbox.click()

        dropdown.click()

        select_element = driver.find_element(By.ID, "orderStatus")
        select = Select(select_element)
        select.select_by_visible_text(trangthai)

        # Chọn ngày
        date_input = driver.find_element(By.ID, 'fromDate')
        date_input.click()
        date_input.send_keys(Keys.DELETE)
        date_input.send_keys(from_date_str)
        date_input.send_keys(Keys.TAB)
        #time.sleep(1)
        date_output = driver.find_element(By.ID, 'toDate')
        date_output.send_keys(Keys.DELETE)
        date_output.send_keys(to_date_str)
        date_output.send_keys(Keys.TAB)

        pagination_info_element = driver.find_element(By.CLASS_NAME, "pagination-info")
        old_pagination_text = pagination_info_element.text

        search_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="btnSearch"]'))
            )
        search_button.click()
        
        time.sleep(10)

        WebDriverWait(driver, 10).until(
            lambda driver: pagination_info_element.text != old_pagination_text
        )

        new_pagination_info_element = driver.find_element(By.CLASS_NAME, "pagination-info")
        pagination_text = new_pagination_info_element.text
    
        if pagination_text == "Xem 0 đến 0 của 0 dòng":
            print(Fore.GREEN +f"{trangthai} không có dữ liệu để xuất!")
        else:
            export_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="btnExportExcel"]'))
            )
            export_button.click()
            tai_ve = download_and_rename_file(trangthai, utils.set_mod.download_directory) #trả vê 1 là tải về thành công trả về 0 là tải về bị lỗi
            if tai_ve == 1:
                co_tai_ve =1
            else:
                co_tai_ve= 0
    except Exception as e:
        print(Fore.RED +f"Đã xảy ra lỗi khi kiểm tra {trangthai}: {e}")
        await send_message_async(alert_mod_dev,f'Error: Lỗi Web khi đang kiểm tra {trangthai}',delay=2)
        co_kiem_tra = 0

    finally:
        driver.refresh()  # Tắt popup tải tuyến
        try:
            menu_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//span[@class='HideText' and text()='Menu']"))
            )
            menu_button.click()
            logout_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//a[@href='/commons/logout' and text()='Thoát']"))
            )
            logout_button.click()
            driver.delete_all_cookies()
        except Exception as e:
            print(Fore.RED +f"Đã xảy ra lỗi khi thoát: {e}")            
            await send_message_async(alert_mod_dev,f'Error: Lỗi Web khi thoát trang {trangthai}',delay=2)
            #sys.exit()
    time_check_str = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    return (time_check_str,co_tai_ve,co_kiem_tra)

def combine_and_process_excel_files(download_directory, statuses):
    warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')
    combined_data = pd.DataFrame()
    files_found = False
    
    for status in statuses:
        file_path = os.path.join(download_directory, f"{status}.xlsx")
        if os.path.exists(file_path):
            files_found = True
            #print(f"\nĐã xử lý file: {status}.xlsx")
            # Đọc dữ liệu từ dòng thứ 5 trở đi mà không có tiêu đề cột
            df = pd.read_excel(file_path, skiprows=4, header=None, engine='openpyxl')
            # Giữ lại cột thứ 4 (cột D) và cột thứ 6 (cột F)
            if df.shape[1] >= 6:  # Đảm bảo có ít nhất 6 cột
                df_filtered = df.iloc[:, [0,1, 3, 5]]  # Cột D (4) và cột F (6) có chỉ số 3 và 5
                combined_data = pd.concat([combined_data, df_filtered], ignore_index=True)
            else:
                print(Fore.RED +f"File {file_path} không đủ số cột cần thiết.")
    
    if not files_found:
        #print("Không tạo file Kết quả do đã duyệt hết đơn hàng")
        return 0  # Dừng hàm nếu không tìm thấy file nào
    
    # Loại bỏ giá trị trùng lặp
    combined_data.drop_duplicates(inplace=True)
    # Đổi tên cột
    combined_data.columns = ['Miền', 'Vùng' ,'Tên NPP', 'Số đơn']

    # Tạo bảng tổng hợp
    pivot_data = combined_data.pivot_table(index=['Miền', 'Vùng' ,'Tên NPP'], values='Số đơn', aggfunc='count').reset_index()
    pivot_data.columns = ['Miền', 'Vùng' ,'Tên NPP', 'Số đơn']

    # Lưu kết quả vào file
    output_file = os.path.join(download_directory, 'Kết quả chưa duyệt đơn hàng.xlsx')
    pivot_data.to_excel(output_file, index=False)
    #print("Đã lưu File: Kết quả chưa duyệt đơn hàng.xlsx")
    return 1

def extract_after_dash(text):
    """Lấy phần sau dấu '-' trong chuỗi, nếu có."""
    if '-' in text:
        return text.split('-')[1].strip()
    return text

def read_and_print_excel(download_directory, file_name):
    file_path = os.path.join(download_directory, file_name)
    # Đọc dữ liệu từ file Excel
    df = pd.read_excel(file_path,engine='openpyxl')
    
    # Chuỗi để lưu tất cả các dòng
    send_text = ""
    
    # Nhóm dữ liệu theo cột Miền
    miền_grouped = df.groupby('Miền')
    
    # Duyệt qua từng nhóm Miền
    for miền, miền_data in miền_grouped:
        send_text += f"* {miền}:\n"
        
        # Nhóm dữ liệu theo cột Vùng trong từng Miền
        vùng_grouped = miền_data.groupby('Vùng')
        
        # Duyệt qua từng nhóm Vùng
        for vùng, vùng_data in vùng_grouped:
            send_text += f"  -- {vùng}:\n"
            
            # Tạo danh sách các nhà phân phối và số đơn
            npp_grouped = vùng_data.groupby('Tên NPP')
            
            # Duyệt qua từng nhóm Tên NPP
            for ten_npp, npp_data in npp_grouped:
                ten_npp_formatted = extract_after_dash(ten_npp)  # Xử lý tên NPP
                for _, row in npp_data.iterrows():
                    so_don = str(row['Số đơn']).strip().replace('\n', '')
                    send_text += f"    --- {ten_npp_formatted} : {so_don} đơn\n"
    send_text = send_text
    #print(send_text)
    return send_text

#thông số chung lấy thời gian check đơn hàng đã duyệt hay chưa

async def check_order():
    statuses = ['Đã xác nhận đơn hàng','Chưa xác nhận đơn hàng']
    delete_file = ['Chưa xác nhận đơn hàng','Đã xác nhận đơn hàng','Kết quả chưa duyệt đơn hàng','Đã trả']
    check_and_remove_old_files(delete_file)
              
    for status in statuses:
        while True:
            thoigian_tai = await check_status_order(status)
            co_tai_ve, co_kiem_tra = thoigian_tai[1], thoigian_tai[2]
            if co_tai_ve == 1 and co_kiem_tra == 1:
                break  # Thoát vòng lặp while để tiếp tục với trạng thái tiếp theo
            else:
                time.sleep(2)  # Ví dụ: chờ 5 giây trước khi kiểm tra lại

    yes_no = combine_and_process_excel_files(utils.set_mod.download_directory, statuses) # 0 đã duyệt hêt đơn, 1 là chưa duyệt hết
    if yes_no == 0:
        return [0,thoigian_tai[0]] # trả về 0 là duyệt hết đơn
    else:
        danhsachdonconlai_text = read_and_print_excel(utils.set_mod.download_directory,'Kết quả chưa duyệt đơn hàng.xlsx')
        return [danhsachdonconlai_text,thoigian_tai[0]] # trả về danh sách đơn

async def app_check_order(max_checks):
    ngay_hien_tai = datetime.now()
    ngay_hien_tai_str = ngay_hien_tai.strftime('%d/%m/%Y %H:%M')
    await send_message_async(alert_mod_dev,f'▶️ {ngay_hien_tai_str}',delay=2)
    else_count = 0
    while else_count < max_checks:
        ngay_hien_tai = datetime.now()
        ngay_hien_tai_str = ngay_hien_tai.strftime('%d/%m/%Y %H:%M')
        result = await check_order()
        if result[0] == 0:
            ngay_hien_tai = datetime.now()
            ngay_hien_tai_str = ngay_hien_tai.strftime('%d/%m/%Y %H:%M')

            await send_message_async(alert_mod_dev,f'⏳ Hoàn thành',delay=2)
            time.sleep(2)
            await send_message_async(alert_mod_dev,f'⏹️ {ngay_hien_tai_str}',delay=2)
            time.sleep(2)
            await send_message_async(alert_mod_user,f'{info_app_user}\n{nd_app}\n\U000023F0 {result[1]}\n\U0001F4C5 {from_date_str} - {to_date_str}\n⏹️ Đã duyệt',delay=2)
            print(Fore.GREEN +f'Done! Đơn hàng đã duyệt xong !')
            return 1
        else:
            await send_message_async(alert_mod_user,f'{info_app_user}\n{nd_app}\n\U000023F0 {result[1]}\n\U0001F4C5 {from_date_str} - {to_date_str}\n⏹️ Chưa duyệt\n\U0001F4EE Chi tiết:\n{result[0]}',delay=2)
            time.sleep(2)
            else_count+=1
            if else_count == max_checks:
                await send_message_async(alert_mod_dev,f'⏳ Đã check {max_checks} lần, Admin hãy kiểm tra',delay=2)
                time.sleep(2)

                await send_message_async(alert_mod_dev,f'⏹️ {ngay_hien_tai_str}',delay=2)
                time.sleep(2)
                await send_message_async(alert_mod_user,f'{info_app_user}\n{tt_app}\n⏳ Đã check {max_checks} lần, Admin hãy kiểm tra\n⏹️ Kết thúc',delay=2)
                return 0
            print(Fore.BLUE +f'Done! Waitting {delay_minute} minute & recheck...:')
            await send_message_async(alert_mod_dev,f'⏳ Chờ {delay_minute} phút...',delay=2)
            time.sleep(delay_minute*60)  # Chờ 15 phút

async def run_app_check_order(max_checks):
    print(Style.BRIGHT + Fore.YELLOW +'---------- START APP STATUS ORDER ----------')
    result = await app_check_order(max_checks)  # Gọi trực tiếp hàm đồng bộ
    print(Style.BRIGHT + Fore.YELLOW +'---------- STOP  APP STATUS ORDER ----------')
    return result

if __name__ == "__main__":
    run_app_check_order(8)
    sys.exit()