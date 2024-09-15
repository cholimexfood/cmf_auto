from termcolor import colored
from datetime import datetime,timedelta
import pyodbc
import os
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException
import time,asyncio
from pathlib import Path
import utils.set_mod
from utils.send_telegram import  send_message,send_message_async
from colorama import Fore, Back, Style, init
init(autoreset=True) 
###############################################################################################################
alert_mod_dev = utils.set_mod.alert_mod_dev
alert_mod_user = utils.set_mod.alert_mod_user
info_app_dev = utils.set_mod.info_app_dev
info_app_user = utils.set_mod.info_app_user
info_app_error = utils.set_mod.info_app_error


driver = utils.set_mod.get_driver()

def is_valid_date(date_str):
    try:
        datetime.strptime(date_str, "%d/%m/%Y")
        return True
    except ValueError:
        return False

def get_dates():
    # Kết nối đến SQL Server
    conn = pyodbc.connect(utils.set_mod.conn_str)
    cursor = conn.cursor()
    query = "SELECT Ngày_Update from view_ReportVBA_NgayUpdate"
    cursor.execute(query)
    row = cursor.fetchone()
    cursor.close()
    conn.close()

    value = row[0]  # Lấy giá trị từ kết quả truy vấn
    #value = '25/08/2024'
    # Đảm bảo rằng giá trị là chuỗi trước khi chuyển đổi
    if isinstance(value, datetime):
        value = value.strftime('%d/%m/%Y')
    
    # Chuyển đổi giá trị sang đối tượng datetime để tính toán
    value_mode = datetime.strptime(value, '%d/%m/%Y')
    tu_ngay = value_mode - timedelta(days=4)
    den_ngay = datetime.now() - timedelta(days=1)
    
    # Chuyển đổi ngày thành chuỗi
    tu_ngay_str = tu_ngay.strftime('%d/%m/%Y')
    den_ngay_str = den_ngay.strftime('%d/%m/%Y')
    print(Fore.BLUE +f"Chọn khoảng thời xuất Danh_sách_đơn_hàng (Lưu ý: Đơn hàng trong DB mới nhất đến {value}):")
    print(Fore.BLUE+ f'Khuyến nghị bạn nhập Ngày bắt đầu: {tu_ngay_str} đến ngày {den_ngay_str} để đơn hàng được tải nhanh và mới nhất')
    print(Fore.BLUE+ f"Tự động chọn thời gian từ ngày {tu_ngay_str} đến ngày {den_ngay_str} để xuất Danh_sách_đơn_hàng")
    
    # Trả về hai giá trị chuỗi
    return [tu_ngay, den_ngay]

def display_menu(menu_options):
    print("Bạn chọn tải báo cáo nào:")
    for i, menu in enumerate(menu_options, start=1):
        print(f"    {i}. {menu}")

def get_user_choice(menu_options):
    print(Fore.BLUE+ 'Khuyến nghị: Bạn nên nhập số 4 để tải mới tất cả báo cáo!')
    print(Fore.BLUE +"Tự động chọn: 4")
    return 4


async def wait_for_download_and_rename(menu_name, download_directory, timeout):
    start_time = time.time()
    initial_files = os.listdir(download_directory)

    while True:
        files = os.listdir(download_directory)
        new_files = [f for f in files if f not in initial_files and f.endswith('.xlsx')]

        if new_files:
            downloaded_file = new_files[0]
            
            # Kiểm tra và xóa file trùng tên với menu_name nếu đã tồn tại
            existing_file_path = os.path.join(download_directory, f"{menu_name}.xlsx")
            if os.path.exists(existing_file_path):
                os.remove(existing_file_path)
            
            # Đổi tên file đã tải xuống
            new_file_path = os.path.join(download_directory, f"{menu_name}.xlsx")
            os.rename(os.path.join(download_directory, downloaded_file), new_file_path)
            
            print(Fore.GREEN +f"File đã được tải về và lưu với tên: {menu_name}")
            await send_message_async('dev', f'\U0001F4E5 {menu_name} \U00002714',delay=2)
            return 1  # Thành công

        current_time = time.time()
        elapsed_time = current_time - start_time
        if elapsed_time >= timeout:
            print(Fore.RED +f'Đã quá {timeout} thời gian chờ!')
            await send_message_async('dev', f'Tải\n{menu_name}\nThất bại - hết giờ',delay=2)
            return 0  # Thất bại
        
        #time.sleep(3)
        await asyncio.sleep(3)
    
async def download_report_tuyen(menu_name):
    check_download = 0
    try:
        print(Fore.BLUE +f'Bắt đầu tải {menu_name}...')
        # 1. Mở trang web
        driver.get("https://cholimexfood.dmsone.vn")
        driver.implicitly_wait(1)

        # 2. Đăng nhập
        username = driver.find_element(By.ID, "username")
        password = driver.find_element(By.ID, "password")
        username.send_keys("GADMIN")
        password.send_keys("Ch0oL!iM3ex#2024")
        password.send_keys(Keys.RETURN)

        driver.implicitly_wait(1)

        # 3. Thao tác với các menu trên trang web
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "1259"))
        )
        actions = ActionChains(driver)
        actions.move_to_element(element).perform()

        element = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "child_1302"))
        )
        element.click()

        element = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "child_Child_1303"))
        )
        element.click()

        element = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "btnExportExcel"))
        )
        element.click()

        element = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "btnSearchStyle2EasyUIUpdate"))
        )
        element.click()
        timeout = 5*60
        check_download = await wait_for_download_and_rename(menu_name, utils.set_mod.download_directory, timeout)
        #tắt popup tải tuyến
        driver.refresh()
        #vào nút thoát và thoát trang
        element = driver.find_element(By.XPATH, "//span[@class='HideText' and text()='Menu']")
        element.click()
        element = driver.find_element(By.XPATH, "//a[@href='/commons/logout' and text()='Thoát']")
        element.click()
        driver.delete_all_cookies() 
    except Exception as e:
        print(Fore.RED +f"Đã xảy ra lỗi khi kiểm tra {menu_name}: {e}")
        await send_message_async(alert_mod_dev,f'Error: Lỗi Web khi đang kiểm tra {menu_name}',delay=2)
        check_download = 0

    if  check_download == 1:
        return 1
    else:
        return 0


async def download_report_kh(menu_name):
    check_download=0
    try:
        print(Fore.BLUE+ 'Bắt đầu tải {menu_name}...')
        # 1. Mở trang web
        driver.get("https://cholimexfood.dmsone.vn")
        driver.implicitly_wait(1)

        # 2. Đăng nhập
        username = driver.find_element(By.ID, "username")
        password = driver.find_element(By.ID, "password")
        username.send_keys("GADMIN")
        password.send_keys("Ch0oL!iM3ex#2024")
        password.send_keys(Keys.RETURN)

        driver.implicitly_wait(1)

        # 3. Thao tác với các menu trên trang web
        
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "398"))
        )
        actions = ActionChains( driver)
        actions.move_to_element(element).perform()

        element = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "child_882"))
        )
        element.click()

        element = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "child_Child_760"))
        )
        element.click()
        #time.sleep(10)
        await asyncio.sleep(10)
        element = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "btnExportExcelDrop"))
        )
        element.click()

        element = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//a[@onclick="return CustomerCatalog.exportExcelCustomer();"]'))
        )
        element.click()

        timeout = 5*60
        check_download = await wait_for_download_and_rename(menu_name, utils.set_mod.download_directory, timeout)
        #tắt popup tải tuyến
        driver.refresh()
        #vào nút thoát và thoát trang
        element = driver.find_element(By.XPATH, "//span[@class='HideText' and text()='Menu']")
        element.click()
        element = driver.find_element(By.XPATH, "//a[@href='/commons/logout' and text()='Thoát']")
        element.click()
        driver.delete_all_cookies() 
    except Exception as e:
        print(Fore.RED+f"Đã xảy ra lỗi khi kiểm tra {menu_name}: {e}")
        await send_message_async(alert_mod_dev,f'Error: Lỗi Web khi đang kiểm tra {menu_name}',delay=2)
        check_download = 0
    
    if  check_download == 1:
        return 1
    else:
        return 0

async def download_report_dh(menu_name,from_date,to_date):
    check_download=0
    try:
        # 1. Mở trang web
        print(Fore.BLUE+f'Bắt đầu tải {menu_name}...')
        driver.get("https://cholimexfood.dmsone.vn")
        driver.implicitly_wait(1)

        # 2. Đăng nhập
        username = driver.find_element(By.ID, "username")
        password = driver.find_element(By.ID, "password")
        username.send_keys("GADMIN")
        password.send_keys("Ch0oL!iM3ex#2024")
        password.send_keys(Keys.RETURN)

        driver.implicitly_wait(1)

        # 3. Thao tác với các menu trên trang web
        
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "424"))
        )
        actions = ActionChains( driver)
        actions.move_to_element(element).perform()

        element = WebDriverWait( driver, 10).until(
            EC.element_to_be_clickable((By.ID, "child_922"))
        )
        element.click()

        # Mở dropdown
        dropdown = WebDriverWait( driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'span.ui-dropdownchecklist-text'))
        )
        dropdown.click()

        # Chờ cho các tùy chọn xuất hiện
        dropdown_items = WebDriverWait( driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div.ui-dropdownchecklist-item'))
        )

        # Chọn các giá trị mong muốn
        values_to_select = ['IN', 'SO']
        for item in dropdown_items:
            checkbox = item.find_element(By.CSS_SELECTOR, 'input[type="checkbox"]')
            value = checkbox.get_attribute('value')
            if value in values_to_select:
                if not checkbox.is_selected():
                    checkbox.click()

        dropdown = WebDriverWait( driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'span.ui-dropdownchecklist-text'))
        )
        dropdown.click()

        select_element =  driver.find_element(By.ID, "orderStatus")
        select = Select(select_element)
        select.select_by_visible_text("Đã duyệt và cập nhật kho")

        from_date_str = from_date.strftime('%d/%m/%Y')
        to_date_str = to_date.strftime('%d/%m/%Y')

        # Chuyển đổi ngày thành định dạng dd/mm/yyyy
        date_input =  driver.find_element(By.ID, 'fromDate')
        date_input.click()
        date_input.send_keys(Keys.DELETE) 
        date_input.send_keys(from_date_str)
        date_input.send_keys(Keys.TAB)
        #time.sleep(2)
        await asyncio.sleep(2)
        date_output =  driver.find_element(By.ID, 'toDate')
        date_output.send_keys(Keys.DELETE) 
        date_output.send_keys(to_date_str)
        date_output.send_keys(Keys.TAB)

        element = WebDriverWait( driver, 10).until(
            EC.element_to_be_clickable((By.ID, "btnExportExcel"))
        )
        element.click()
        timeout = 10*60
        check_download = await wait_for_download_and_rename(menu_name, utils.set_mod.download_directory, timeout)
        #tắt popup tải tuyến
        driver.refresh()
        #vào nút thoát và thoát trang
        element = driver.find_element(By.XPATH, "//span[@class='HideText' and text()='Menu']")
        element.click()
        element = driver.find_element(By.XPATH, "//a[@href='/commons/logout' and text()='Thoát']")
        element.click()
        driver.delete_all_cookies() 

    except Exception as e:
        print(Fore.RED+f"Đã xảy ra lỗi khi kiểm tra {menu_name}: {e}")
        await send_message_async(alert_mod_dev,f'Error: Lỗi Web khi đang kiểm tra {menu_name}',delay=2)
        check_download =0 

    if  check_download == 1:
        return 1
    else:
        return 0

async def download():
    menu_options = ["Tuyến bán hàng", "Danh sách khách hàng", "Danh sách đơn hàng", "Tất cả"]
    display_menu(menu_options)
    choice_value = get_user_choice(menu_options)

    if choice_value == 1:
        await download_report_tuyen(menu_options[choice_value - 1].replace(" ", "_"))
    elif choice_value == 2:
        await download_report_kh(menu_options[choice_value - 1].replace(" ", "_"))
    elif choice_value == 3:
        dates = get_dates()
        await download_report_dh(menu_options[choice_value - 1].replace(" ", "_"), dates[0], dates[1])
    else:
        dates = get_dates()

        # Tạo danh sách các hàm và tham số tương ứng
        download_functions = [
            (download_report_tuyen, [menu_options[choice_value - 4].replace(" ", "_")], "Tuyến bán hàng"),
            (download_report_kh, [menu_options[choice_value - 3].replace(" ", "_")], "Danh sách khách hàng"),
            (download_report_dh, [menu_options[choice_value - 2].replace(" ", "_"), dates[0], dates[1]], "Danh sách đơn hàng"),
        ]

        # Duyệt qua từng mục trong danh sách để tải báo cáo
        for func, args, report_name in download_functions:
            attempts = 0
            while attempts < 3:
                if await func(*args) == 1:  # Sử dụng await để gọi hàm bất đồng bộ
                    break
                else:
                    attempts += 1
                    print(Fore.RED+f'Lỗi: Thử tải lại {report_name} ({attempts}/3)')
                    await asyncio.sleep(2)  # Chờ 2 giây trước khi thử lại

    driver.quit()

async def  run_main_download_async():
    print(Style.BRIGHT + Fore.YELLOW +'---------- START APP DOWNLOAD REPORT ----------')
    await download()
    print(Style.BRIGHT + Fore.YELLOW +'---------- STOP  APP DOWNLOAD REPORT ----------')

if __name__ == "__main__":
    asyncio.run(run_main_download_async())


