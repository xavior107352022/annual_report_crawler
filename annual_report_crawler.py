import requests
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import os
import pandas as pd
import time
from utils import get_info
from datetime import datetime

start_year = 112
end_year = 113
use_list = True
year_list = [str(x) for x in range(start_year, end_year)]

save_path = "\\Share_Holder_Ann\\"
if not os.path.exists(os.getcwd() + save_path):
    os.makedirs(os.getcwd() + save_path)

if use_list:
    df_list = pd.read_excel("股票代號清單.xlsx")
    ticker_list = df_list["股票代號"]
else:
    df_info_tse, df_info_otc = get_info()
    df = pd.concat([df_info_tse, df_info_otc])
    ticker_list = df["公司代號"]

ChromeDriverManager(path="./chromedriver/").install()
chrome_options = Options()
chrome_options.add_argument("--headless")

fail_list = []
for year in year_list:
    print("-" * 100)
    print("-" * 100)
    print("-" * 100)
    print("processing : {}".format(year))
    for ticker in ticker_list:
        time.sleep(10)
        try:
            print("processing : {}".format(ticker))
            driver = webdriver.Chrome(options=chrome_options)
            driver.get("https://www.google.com/?hl=zh_tw")
            url = "https://doc.twse.com.tw/server-java/t57sb01?id=&key=&step=1&co_id={}&year={}&seamon=&mtype=F&dtype=F04".format(
                ticker, year
            )
            driver.get(url)
            print(url)
            # 點選PDF URL
            time.sleep(1)
            driver.find_element(By.TAG_NAME, "a").click()
            handles = driver.window_handles
            driver.switch_to.window(handles[1])
            # 抓取pdf資訊
            time.sleep(1)
            link = driver.find_element(By.TAG_NAME, "a").get_attribute("href")
            name = driver.find_element(By.TAG_NAME, "a").text
            print(link)
            # 存成PDF
            response = requests.get(link, allow_redirects=True)
            file_name = name

            exist_files_list = os.listdir(os.getcwd() + save_path)
            if file_name not in exist_files_list:
                with open(os.getcwd() + save_path + file_name, "wb") as f:
                    f.write(response.content)
                    f.close()
                print(file_name + "-" * 10 + "ok")
            else:
                print("existing file : {}".format(file_name))

        except Exception as e:
            print("processing fail : {}".format(ticker))
            print(e)
            fail_list.append((year, ticker))
final_fail_list = []
for year, ticker in fail_list:
    try:
        print("-" * 100)
        print("-" * 100)
        print("-" * 100)
        print("processing : {}".format(year))
        time.sleep(10)

        print("processing : {}".format(ticker))
        driver = webdriver.Chrome(options=chrome_options)
        driver.get("https://www.google.com/?hl=zh_tw")
        url = "https://doc.twse.com.tw/server-java/t57sb01?id=&key=&step=1&co_id={}&year={}&seamon=&mtype=F&dtype=F04".format(
            ticker, year
        )
        driver.get(url)
        print(url)
        # 點選PDF URL
        time.sleep(1)
        driver.find_element(By.TAG_NAME, "a").click()
        handles = driver.window_handles
        driver.switch_to.window(handles[1])
        # 抓取pdf資訊
        time.sleep(1)
        link = driver.find_element(By.TAG_NAME, "a").get_attribute("href")
        name = driver.find_element(By.TAG_NAME, "a").text
        print(link)
        # 存成PDF
        response = requests.get(link, allow_redirects=True)
        file_name = name

        exist_files_list = os.listdir(os.getcwd() + save_path)
        if file_name not in exist_files_list:
            with open(os.getcwd() + save_path + file_name, "wb") as f:
                f.write(response.content)
                f.close()
            print(file_name + "-" * 10 + "ok")
        else:
            print("existing file : {}".format(file_name))

    except Exception as e:
        print("processing fail : {}".format(ticker))
        print(e)
        final_fail_list.append((year, ticker))
df_final_fail = pd.DataFrame()
df_final_fail["year"] = [x[0] for x in final_fail_list]
df_final_fail["ticker"] = [x[1] for x in final_fail_list]
df_final_fail["time"] = datetime.now().strftime("%Y%m%d")
