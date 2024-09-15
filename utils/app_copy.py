import os
import pandas as pd
from termcolor import colored
import warnings
import asyncio
import utils.set_mod
from utils.send_telegram import send_message, send_message_async
from colorama import Fore, Back, Style, init
init(autoreset=True) 
output_directory = os.path.join(utils.set_mod.current_directory, 'data_import')
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

download_directory = os.path.join(utils.set_mod.current_directory, 'downloads')

# Danh sách các file form yêu cầu
required_form_files = ['tbl_tuyen.xlsx', 'tbl_khachhang.xlsx', 'tbl_donhang_import.xlsx']

# Mapping các file tải về với form tương ứng
file_mapping = {
    'Tuyến_bán_hàng.xlsx': {'form_file': 'tbl_tuyen.xlsx', 'skiprows': 5, 'header': [
        'STT', 'MA_NPP', 'MA_TUYEN', 'TEN_TUYEN', 'MA_NVBH', 'MA_KH', 'TEN_KH', 
        'DC', 'TU_NGAY', 'DEN_NGAY', 'TAN_SUAT', 'W1', 'W2', 'W3', 'W4', 
        'T2', 'T3', 'T4', 'T5', 'T6', 'T7', 'CN', 'TT2', 'TT3', 'TT4', 
        'TT5', 'TT6', 'TT7', 'TCN', 'CHAM_SOC', 'NGAY_CAP_NHAT'
    ]},
    'Danh_sách_khách_hàng.xlsx': {'form_file': 'tbl_khachhang.xlsx', 'skiprows': 3, 'header': [
        'STT', 'MA_NPP', 'MA_KH', 'TEN_KH', 'TEN_LOAI_KH', 'TRANG_THAI', 
        'TEN_DAI_DIEN', 'EMAIL', 'DT_BAN', 'DT_DD', 'FAX', 'MA_DIA_BAN', 
        'DUONG_CHO', 'SO_NHA', 'NGAY_SINH', 'TAN_SUAT', 'TUYEN_GH', 'KENH', 
        'GSBH', 'NGAY_THU_NO', 'LAT', 'LNG', 'MA_BRAVO', 'MA_SIEU_THI', 
        'MA_NVGH', 'MA_NV_THU_TIEN', 'MUC_NO', 'HAN_NO', 'VI_TRI_BH', 
        'MST', 'STK', 'TEN_HD', 'TEN_DAI_DIEN_HD', 'MUC', 'CHI_NHANH', 
        'CHU_TK', 'DC_HD', 'NGAY_BAT_DAU', 'NGAY_KET_THUC', 'CO_IN_HD', 
        'THU_TU_XUAT_HIEN', 'LOAI_HINH_KD', 'THUOC_TINH_KH', 'HINH_THUC_TT', 
        'THOI_GIAN_TT', 'DC_GH', 'NONGTHON_THANHTHI', 'TRENDUONG_TRONGCHO', 
        'KHUYEN_MAI', 'CHIET_KHAU', 'CTKM_DB', 'TUYEN', 'NGAY_TAO', 
        'CAU_HINH_THUE_GTGT', 'THUE_TNCN'
    ]},
    'Danh_sách_đơn_hàng.xlsx': {'form_file': 'tbl_donhang_import.xlsx', 'skiprows': 4, 'header': [
        'RSM', 'ASM', 'MA_NPP', 'TEN_NPP', 'MA_TEN_NVBH', 'SO_DH', 'SO_DH_TRA', 
        'NGAY_DAT', 'NGAY_DUYET', 'NGAY_GIAO', 'DO_UU_TIEN', 'MA_KH', 'TEN_KH', 
        'DIA_CHI', 'TRANG_THAI', 'LOAI_DH', 'NHOM_HANG', 'MA_SP', 'TEN_SP', 
        'DONVI_LE', 'QUYCACH', 'GIA', 'TONG_SL', 'SLKM', 'SL_TRA', 'DOANH_SO', 
        'CHIET_KHAU', 'DOANH_SO_TRA', 'DOANH_SO_NET', 'CT_HTTM', 'LOAI_KM', 
        'PT_THUE', 'THUE', 'TT_TT_THUE', 'NHOM', 'MUC', 'MA_VUNG', 'MA_MIEN', 
        'MA_RSM', 'MA_ASM', 'XUAT_HD', 'NVGH', 'NGUOI_DUYET', 'GHI_CHU'
    ]}
}

async def process_files(download_dir, import_dir, file_mapping):
    for source_file, config in file_mapping.items():
        source_path = os.path.join(download_dir, source_file)
        dest_file = config['form_file']
        if os.path.exists(source_path):
            message = (Fore.BLUE +f"Tạo form import {dest_file}...")
            print(message, end="", flush=True)

            df = pd.read_excel(source_path, skiprows=config['skiprows'], engine='openpyxl')
            df.columns = config['header']
            dest_path = os.path.join(import_dir, dest_file)
            if os.path.exists(dest_path):
                try:
                    os.remove(dest_path)
                except OSError:
                    pass  # Nếu tệp không tồn tại, không cần xử lý
            df.to_excel(dest_path, index=False)
            print(Fore.GREEN +f"Done! Đã lưu file {dest_file}.")
            await send_message_async('dev', f'\U0000274E {dest_file} \U00002714',delay=2)
        else:
            print(Fore.RED +f"File nguồn {source_file} không tồn tại trong thư mục {download_dir}.")
            await send_message_async('dev', f'Error: File Excel {source_file} không tồn tại',delay=2)

async def run_main_copy_async():
    warnings.simplefilter(action='ignore', category=UserWarning)
    print(Style.BRIGHT + Fore.YELLOW +'---------- START APP TẠO FILE IMPORT TỰ ĐỘNG ----------')
    await process_files(download_directory, output_directory, file_mapping)
    print(Style.BRIGHT + Fore.YELLOW +'---------- STOP  APP TẠO FILE IMPORT TỰ ĐỘNG ----------')

if __name__ == "__main__":
    asyncio.run(run_main_copy_async())
