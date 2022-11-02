import requests
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import os
import pandas as pd
import time
from utils import get_info

start_year = 109
end_year = 112

save_path = '\\Share_Holder_Ann\\'
if not os.path.exists(os.getcwd() + save_path):
    os.makedirs(os.getcwd() + save_path)

# server = 'XAVIOR\SQLEXPRESS'
# database = 'cmoney'
# cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER=' + server + ';DATABASE=' + database)
# cursor = cnxn.cursor()
#
# sql_query = "SELECT * FROM firm_info_listed;"
# df_listed = pd.read_sql_query(sql_query, cnxn)
# sql_query = "SELECT * FROM firm_info_delisted;"
# df_delisted = pd.read_sql_query(sql_query, cnxn)
# df = pd.concat([df_listed, df_delisted])
df_info_tse, df_info_otc = get_info()
df = pd.concat([df_info_tse, df_info_otc])


ChromeDriverManager(path='./chromedriver/').install()
chrome_options = Options()
chrome_options.headless = True

ticker_list = df['公司代號']
year_list = [str(x) for x in range(109, 112)]
fail_list = []
for year in year_list:
    print('-' * 100)
    print('-' * 100)
    print('-' * 100)
    print('processing : {}'.format(year))
    for ticker in ticker_list:
        time.sleep(10)
        try:
            print('processing : {}'.format(ticker))
            driver = webdriver.Chrome(ChromeDriverManager(path='./chromedriver/').install(),
                                    chrome_options=chrome_options)
            driver.get('https://www.google.com/?hl=zh_tw')
            url = 'https://doc.twse.com.tw/server-java/t57sb01?id=&key=&step=1&co_id={}&year={}&seamon=&mtype=F&dtype=F04'.format(
            ticker, year)
            driver.get(url)
            print(url)
            # 點選PDF URL
            time.sleep(1)
            driver.find_element(By.TAG_NAME,'a').click()
            handles = driver.window_handles
            driver.switch_to.window(handles[1])
            # 抓取pdf資訊
            time.sleep(1)
            link = driver.find_element(By.TAG_NAME,'a').get_attribute('href')
            name = driver.find_element(By.TAG_NAME,'a').text
            print(link)
            # 存成PDF
            response = requests.get(link, allow_redirects=True)
            file_name = name

            exist_files_list = os.listdir(os.getcwd() + save_path)
            if file_name not in exist_files_list:
                with open(os.getcwd() + save_path + file_name, 'wb') as f:
                    f.write(response.content)
                    f.close()
                print(file_name + '-' * 10 + 'ok')
            else:
                print('existing file : {}'.format(file_name))

        except Exception as e:
            print('processing fail : {}'.format(ticker))
            print(e)
            fail_list.append((year,ticker))