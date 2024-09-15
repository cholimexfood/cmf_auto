#set_mod.py
import sys,os, threading, time
import json
import platform
from pathlib import Path
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException

# xác định đường dẫn của dự án
current_directory = os.path.dirname(os.path.abspath(__file__))
if getattr(sys, 'frozen', False):
    current_directory = os.path.dirname(sys.executable)  # Đường dẫn khi đóng gói
else:
    current_directory = os.path.dirname(os.path.abspath(__file__))  # Đường dẫn khi chạy code

# Thiết lập thư mục tải về nếu chưa có thư mục thì tạo
download_directory = os.path.join(current_directory, 'downloads')
if not os.path.exists(download_directory):
    os.makedirs(download_directory)

# Xác nhận đường dẫn cấu hình chrome
def get_chromedriver_path():
    system = platform.system()
    if system == "Darwin":  # macOS
        driver_name = "chromedriver_mac"
    elif system == "Windows":  # Windows
        driver_name = "chromedriver_win.exe"
    elif system == "Linux":  # Linux
        driver_name = "chromedriver_linux"
    else:
        raise Exception("Unsupported OS")

    if getattr(sys, 'frozen', False):
        # Ứng dụng đang chạy trong chế độ đóng gói
        base_path = Path(sys._MEIPASS)
        driver_path = base_path / "utils" / "drivers" / driver_name
    else:
        # Ứng dụng đang chạy trong chế độ phát triển
        base_path = Path(__file__).resolve().parent
        driver_path = base_path / "drivers" / driver_name
    return str(driver_path)
 
chrome_options = Options()
#chrome_options.add_argument("--headless")  # Chạy ở chế độ headless
chrome_options.add_argument("--disable-gpu")  # Vô hiệu hóa GPU
chrome_options.add_argument("--window-size=1920x1080")  # Đặt kích thước cửa sổ ảo
chrome_options.add_argument('--log-level=3') 
chrome_options.add_argument('--disable-notifications')  # Tắt thông báo
chrome_options.add_argument('--disable-popup-blocking')  # Tắt chặn cửa sổ pop-up

chrome_prefs = {
    "download.default_directory": download_directory,
    "profile.default_content_settings.popups": 0,
    "directory_upgrade": True
}
chromedriver_path = get_chromedriver_path()
chrome_options.add_experimental_option("prefs", chrome_prefs)
service = Service(chromedriver_path)
driver = webdriver.Chrome(service=service, options=chrome_options)

# Hàm để lấy driver
def get_driver():
    return driver
'''
procedure_name = 'RunUpdateTblTuyen'
server = '118.69.76.220,2403'
database = 'DMS_Report'
username = 'sa'
password = 'Ve$Mau@Trai!Tim'
'''
'''
procedure_name = 'RunSelectTest'
server = '10.211.55.3,1433'
database = 'SQLTEST'
username = 'sa'
password = '1'
'''
'''
    "token_bot": '7495350405:AAF7GoT513arnQgWBSDxyyng77QbRLhheGg',
    "chat_id_dev": '-4582491435',
    "chat_id_user": '-4553197920',
    "procedure_name" : 'RunSelectTest',
    "server" : '10.211.55.3,1433',
    "database" : 'SQLTEST',
    "usernameSQL" : 'sa',
    "passwordSQL" : '1',
'''
'''
    "token_bot": '7503563708:AAHTKkdefkN6POU8CRVOAX93K3HvkKabhN0',
    "chat_id_dev": '-1002152000157',
    "chat_id_user": '-1002174879429',
    "procedure_name" : 'RunUpdateTblTuyen',
    "server" : '118.69.76.220,2403',
    "database" : 'DMS_Report',
    "usernameSQL" : 'sa',
    "passwordSQL" : 'Ve$Mau@Trai!Tim',
'''
#dev -1002152000157
#user -1002174879429
#bot 7503563708:AAHTKkdefkN6POU8CRVOAX93K3HvkKabhN0

#test dev -4582491435
#test user 4553197920
#test bot 7495350405:AAF7GoT513arnQgWBSDxyyng77QbRLhheGg

default_config = {
    "web": 'https://cholimexfood.dmsone.vn',
    "username": 'GADMIN',
    "password": 'Ch0oL!iM3ex#2024',
    "token_bot": '7495350405:AAF7GoT513arnQgWBSDxyyng77QbRLhheGg',
    "chat_id_dev": '-4582491435',
    "chat_id_user": '-4553197920',
    "procedure_name" : 'RunSelectTest',
    "server" : '10.211.55.3,1433',
    "database" : 'SQLTEST',
    "usernameSQL" : 'sa',
    "passwordSQL" : '1',
    "time_auto":'09:30',
    "number_repeat_order": '1'
}

def get_setup_file_path():
    if getattr(sys, 'frozen', False):
        # Đường dẫn khi ứng dụng đã được đóng gói
        return os.path.join(os.path.dirname(sys.executable), 'setup.json')
    else:
        # Đường dẫn trong chế độ phát triển
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'setup.json')

# Hàm load cấu hình từ file setup.json
def load_config():
    # Kiểm tra sự tồn tại của file setup.json trong chế độ phát triển
    setup_file_path = get_setup_file_path()

    # Kiểm tra sự tồn tại của file setup.json
    if not os.path.exists(setup_file_path):
        # Tạo file setup.json với giá trị mặc định
        with open(setup_file_path, 'w') as file:
            json.dump(default_config, file, indent=4)
        print(f"File {setup_file_path} đã được tạo với giá trị mặc định.")
    else:
        # Đọc file setup.json
        with open(setup_file_path, 'r') as file:
            config = json.load(file)
        
        # Kiểm tra và thêm giá trị nếu thiếu
        updated = False
        for key in default_config:
            if key not in config:
                config[key] = default_config[key]
                updated = True
        
        if updated:
            # Cập nhật file setup.json với giá trị mới
            with open(setup_file_path, 'w') as file:
                json.dump(config, file, indent=4)
            print(f"File {setup_file_path} đã được cập nhật với giá trị mới.")
    
    # Đọc lại file setup.json để lấy giá trị cuối cùng
    with open(setup_file_path, 'r') as file:
        config = json.load(file)
    
    return config

config = load_config()

procedure_name = config['procedure_name']
server = config['server']
database = config['database']
username = config['usernameSQL']
password = config['passwordSQL']

conn_str = (
    f'DRIVER={{ODBC Driver 17 for SQL Server}};'
    f'SERVER={server};'
    f'DATABASE={database};'
    f'UID={username};'
    f'PWD={password}'
)

alert_mod_dev = 'dev'
alert_mod_user ='user'
info_app_dev =f'[App: Order Status]:[{alert_mod_dev}]'
info_app_user =f'[App: Order Status]:[{alert_mod_user}]'
info_app_error =f'[App: Order Status]:[error]'



class Timer:
    def __init__(self):
        self.start_time = None
        self.stop_event = threading.Event()
        self.thread = None

    def start(self):
        """Bắt đầu đồng hồ thời gian"""
        if self.thread and self.thread.is_alive():
            print("Đồng hồ thời gian đã bắt đầu.")
            return
        self.start_time = time.time()
        self.stop_event.clear()
        self.thread = threading.Thread(target=self._run)
        self.thread.daemon = True
        self.thread.start()

    def _run(self):
        """Chạy đồng hồ thời gian và cập nhật mỗi giây"""
        while not self.stop_event.is_set():
            elapsed_time = time.time() - self.start_time
            minutes, seconds = divmod(int(elapsed_time), 60)
            print(f"\rThời gian: {minutes:02}:{seconds:02}", end='', flush=True)
            time.sleep(1)

    def stop(self):
        """Dừng đồng hồ thời gian"""
        if not self.thread or not self.thread.is_alive():
            print("Đồng hồ thời gian không đang chạy.")
            return
        self.stop_event.set()
        self.thread.join()
        elapsed_time = time.time() - self.start_time
        minutes, seconds = divmod(int(elapsed_time), 60)
        print(f"\r(Kết thúc: {minutes:02}:{seconds:02})", end='', flush=True)

    def get_elapsed_time(self):
        """Lấy thời gian đã trôi qua"""
        if not self.start_time:
            return "00:00"
        elapsed_time = time.time() - self.start_time
        minutes, seconds = divmod(int(elapsed_time), 60)
        return f"{minutes:02}:{seconds:02}"