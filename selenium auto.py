from datetime import datetime, timedelta, date
import gspread
import time
import os
import webbrowser
import zipfile
import shutil
import math
import warnings
warnings.filterwarnings('ignore')
from python_calamine import CalamineWorkbook
import pandas as pd
import numpy as np

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC



def delete_file(folder):
    file_list = os.listdir(folder)
    for file in file_list:
        file_path = os.path.join(folder, file)
        os.remove(file_path)
        print(f'Removed {file_path}')
    print(f'Complete removed all files in {folder}')

def read_xlsx_files(folder):
    #check xlsx files in the folder
    excel_files = [file for file in os.listdir(folder) if file.endswith('.xlsx')]
    #create blank df
    df = []
    #read each file and append the the blank df
    for file in excel_files:
        file_path = os.path.join(folder, file)
        wb = CalamineWorkbook.from_path(file_path)
        row_list = wb.get_sheet_by_index(0).to_python()
        data = pd.DataFrame(row_list[1:], columns=row_list[0])
        print(data)
        df.append(data)
    print('Completed reading all files')  
    #return a completely concated df  
    return pd.concat(df, ignore_index=True)
today = date.today()
f_time = today - timedelta(days=1)

email = 'tam.hoangthanh'
password = 'pa$$w0rd'

folder_dict = {
    'folder_vns_inv_transaction': r"C:\Users\tam.hoangthanh\data\1_data_source\inv\vns_transaction",
    'folder_vns_inv_sku_map': r"C:\Users\tam.hoangthanh\data\1_data_source\inv\vns_map",
    'download_folder': r"C:\Users\tam.hoangthanh\Downloads"
}

for key, value in folder_dict.items():
    delete_file(value)

service = Service(executable_path=r"C:\Users\tam.hoangthanh\data\2_main\chromedriver-win64\chromedriver.exe")

driver = webdriver.Chrome(service=service)
url = 'https://wms.ssc.shopee.vn/v2/login?redirect_url=https%3A%2F%2Fwms.ssc.shopee.vn%2Fv2%2Freportcenter%2Freportcenter'
driver.get(url)
driver.maximize_window()


user_name = driver.find_element(By.CSS_SELECTOR, '.ssc-input input')
user_name.send_keys(email)

pass_word = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/form/div[2]/div/div[1]/input')
pass_word.send_keys(password)

login_button = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[2]/div[2]/form/div[4]/div/button')
login_button.click()
time.sleep(10)
#vns_on_hand

choose_wh = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[1]/section[1]/span/span[1]/div/div/span[1]/span/input')
choose_wh.click()
time.sleep(1)

vns = driver.find_element(By.XPATH, '/html/body/span/div/div/div/ul/li[3]/span').click()
time.sleep(5)

rc = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/section/div[2]/div[1]/div/button').click()
time.sleep(1)

module_button = driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div/div[2]/div/form/div[1]/div/span/span[1]/div/div/span[1]').click()
time.sleep(1)

report = driver.find_element(By.XPATH, '/html/body/span/div/div/div/ul/li[3]').click()
time.sleep(1)

task_type_button = driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div/div[2]/div/form/div[2]/div/span[2]/span[1]/div/div/span[2]').click()
time.sleep(1)

task = driver.find_element(By.XPATH, '/html/body/span/div/div/div/ul/li[10]').click()
time.sleep(1)

# select detail
include_batch = driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div/div[2]/div/form/div[5]/div/div[5]/div/span/span[1]/div/div/span[1]').click()
time.sleep(0.1)

choose_Y = driver.find_element(By.XPATH, '/html/body/span/div/div/div/ul/li[2]').click()
time.sleep(0.1)

export_angle = driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div/div[2]/div/form/div[5]/div/div[6]/div/span/span[1]/div/div/span[1]').click()
time.sleep(0.1)

sku_angle = driver.find_element(By.XPATH, '/html/body/span/div/div/div/ul/li[1]/span').click()
time.sleep(0.1)


confirm = driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div/div[3]/div[2]/div/button[2]').click()
time.sleep(5)

task_id = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/section/div[2]/div[3]/div[1]/div/div/div[2]/div[1]/div[1]/div[2]/div/div/div/table/tbody[2]/tr[1]/td[2]/div/span').text
time.sleep(3)

search_task = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/section/div[2]/form/div[4]/div/div/input')
search_task.send_keys(task_id)
time.sleep(1)

while True:
    search_button = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/section/div[2]/form/div[7]/button[1]')
    search_button.click()
    time.sleep(1)
    status = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/section/div[2]/div[3]/div[1]/div/div/div[2]/div[1]/div[1]/div[2]/div/div/div/table/tbody[2]/tr/td[7]/div/div').text
    if status == 'Fail':
        break
    elif status != 'Done':
        print(status)
        time.sleep(3)
    else:
        file_downloaded = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/section/div[2]/div[3]/div[1]/div/div/div[2]/div[1]/div[1]/div[2]/div/div/div/table/tbody[2]/tr/td[6]/div').text
        print(file_downloaded)
        download_button = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/section/div[2]/div[3]/div[1]/div/div/div[2]/div[1]/div[1]/div[2]/div/div/div/table/tbody[2]/tr/td[9]/div/div/button').click()
        break

file_path2 = os.path.join(folder_dict['download_folder'], file_downloaded)

while True:
    if os.path.exists(file_path2):
        print("Tải file xong, tiến hành xử lý...")
        time.sleep(3)
        break
    else:
        time.sleep(1)
        print("Đang tải file")

file_list_in_download_folder = os.listdir(folder_dict['download_folder'])
for file_name in file_list_in_download_folder:
    file_path = os.path.join(folder_dict['download_folder'], file_name)
    
    if file_name.endswith('.zip'):
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(folder_dict['folder_vns_inv_sku_map'])
        shutil.move(file_path, folder_dict['folder_vns_inv_sku_map'])
        
    elif file_name.endswith('.xlsx'):
        shutil.move(file_path, folder_dict['folder_vns_inv_sku_map'])


#vns_transaction

choose_wh = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[1]/section[1]/span/span[1]/div/div/span[1]/span/input')
choose_wh.click()
time.sleep(1)

vns = driver.find_element(By.XPATH, '/html/body/span/div/div/div/ul/li[3]/span').click()
time.sleep(5)

rc = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/section/div[2]/div[1]/div/button').click()
time.sleep(1)

module_button = driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div/div[2]/div/form/div[1]/div/span/span[1]/div/div/span[1]').click()
time.sleep(1)

report = driver.find_element(By.XPATH, '/html/body/span/div/div/div/ul/li[3]').click()
time.sleep(1)

task_type_button = driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div/div[2]/div/form/div[2]/div/span[2]/span[1]/div/div/span[2]').click()
time.sleep(1)

task = driver.find_element(By.XPATH, '/html/body/span/div/div/div/ul/li[12]').click()
time.sleep(1)

# select detail
transaction_type = driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div/div[2]/div/form/div[5]/div/div[7]/div/span/span[1]/div/div/span[1]/span/input').click()
time.sleep(1)

picking = driver.find_element(By.XPATH, '/html/body/span/div/div/div/ul/li[13]/span').click()
time.sleep(1)

date_range = driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div/div[2]/div/form/div[15]/div/div[1]/span/span[1]/div/div/span[1]').click()
user_define = driver.find_element(By.XPATH, '/html/body/span/div/div/div/ul/li[1]/span').click()

a = str(f_time) + ' 00:00:00'
b = str(f_time) + ' 23:59:59'
start_date = driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div/div[2]/div/form/div[15]/div/div[2]/span/div/div/div[1]/input').send_keys(a)
end_date = driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div/div[2]/div/form/div[15]/div/div[2]/span/div/div/div[2]/input').send_keys(b)

confirm = driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div/div[3]/div[2]/div/button[2]').click()
time.sleep(5)

task_id = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/section/div[2]/div[3]/div[1]/div/div/div[2]/div[1]/div[1]/div[2]/div/div/div/table/tbody[2]/tr[1]/td[2]/div/span').text
time.sleep(3)

search_task = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/section/div[2]/form/div[4]/div/div/input')
search_task.send_keys(task_id)
time.sleep(1)

while True:
    search_button = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/section/div[2]/form/div[7]/button[1]')
    search_button.click()
    time.sleep(1)
    status = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/section/div[2]/div[3]/div[1]/div/div/div[2]/div[1]/div[1]/div[2]/div/div/div/table/tbody[2]/tr/td[7]/div/div').text
    if status == 'Fail':
        break
    elif status != 'Done':
        print(status)
        time.sleep(1)
    else:
        file_downloaded = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/section/div[2]/div[3]/div[1]/div/div/div[2]/div[1]/div[1]/div[2]/div/div/div/table/tbody[2]/tr/td[6]/div').text
        print(file_downloaded)
        download_button = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/section/div[2]/div[3]/div[1]/div/div/div[2]/div[1]/div[1]/div[2]/div/div/div/table/tbody[2]/tr/td[9]/div/div/button').click()
        break

file_path2 = os.path.join(folder_dict['download_folder'], file_downloaded)

while True:
    if os.path.exists(file_path2):
        print("Tải file xong, tiến hành xử lý...")
        time.sleep(3)
        break
    else:
        time.sleep(1)
        print("Đang tải file")

file_list_in_download_folder = os.listdir(folder_dict['download_folder'])
for file_name in file_list_in_download_folder:
    file_path = os.path.join(folder_dict['download_folder'], file_name)
    
    if file_name.endswith('.zip'):
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(folder_dict['folder_vns_inv_transaction'])
        shutil.move(file_path, folder_dict['folder_vns_inv_transaction'])
        
    elif file_name.endswith('.xlsx'):
        shutil.move(file_path, folder_dict['folder_vns_inv_transaction'])



driver.quit()
df_inv = read_xlsx_files(folder_dict['folder_vns_inv_sku_map'])
df_rack = myfunction.read_xlsx_files(folder_dict['folder_vns_inv_transaction'])
df_rack['to zone'] = df_rack['To Location'].str.split('-').str[1]
df_rack['from zone'] = df_rack['From Location'].str.split('-').str[1]
zone_id = ['DO','AV']
df_rack_zone = df_rack[(df_rack['to zone'].isin(zone_id)) | (df_rack['from zone'].isin(zone_id))]
df_rack_zone
df_rack_zone_use = df_rack_zone[['from zone','to zone','(L1)Category Name','Quantity','Sheet Time']]
df_rack_zone_use['Sheet Time'] = pd.to_datetime(df_rack_zone_use['Sheet Time'], errors='coerce').dt.strftime('%Y-%m-%d')
df_rack_zone_use_gr = df_rack_zone_use.groupby(['Sheet Time','from zone','to zone','(L1)Category Name'])['Quantity'].sum().reset_index()

df_inv_zone = df_inv[(df_inv['Zone id'] == "DO") | (df_inv['Zone id'] == "AV")]
df_inv_zone_gr = df_inv_zone.groupby(['Zone id','(L1)Category Name'])['On-rack Qty'].sum().reset_index()
df_inv_zone_gr['date'] = str(datetime.today().strftime('%Y-%m-%d'))


gc = gspread.service_account(r"C:\Users\tam.hoangthanh\Data\api_gsheet.json")
do_av = gc.open_by_key('1bnc_HuK4Uh7d9Ab_PWrUm-i23v3Dwfl11DI4RCmrDoo')
sh_on_rack = do_av.worksheet('on_rack')
sh_on_rack.append_rows(df_inv_zone_gr.values.tolist())

sh_rack_transfer = do_av.worksheet('rack_transfer')
sh_rack_transfer.append_rows(df_rack_zone_use_gr.values.tolist())


df_inv_zone_upload = df_inv_zone[['SKU ID','SKU Name','On-rack Qty','Location','Zone id','(L1)Category Name']]
sh_sku_do = do_av.worksheet('sku_do')
sh_sku_do.clear()
sh_sku_do.update([df_inv_zone_upload.columns.values.tolist()] +df_inv_zone_upload.values.tolist())