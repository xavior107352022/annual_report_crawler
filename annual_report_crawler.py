import os
import shutil
import time
import zipfile
from datetime import datetime

import pandas as pd
import requests
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

from utils import get_info

start_year = 112
end_year = 114
use_list = False
year_list = [str(x) for x in range(start_year, end_year)]

download_path = "C:/Users/wa-00100690/Downloads/"
save_path = "./Share_Holder_Ann/"
if not os.path.exists(os.getcwd() + save_path):
    os.makedirs(os.getcwd() + save_path)

if use_list:
    df_list = pd.read_excel("股票代號清單.xlsx")
    ticker_list = df_list["股票代號"]
else:
    df_info_tse, df_info_otc = get_info()
    df = pd.concat([df_info_tse, df_info_otc])
    ticker_list = df["公司代號"]
options = Options()
options.add_argument("--headless=new")

fail_list = []
for year in year_list:
    for ticker in ticker_list:
        try:
            print("-" * 100)
            print("-" * 100)
            print("processing : {}".format(year))
            print("processing : {}".format(ticker))
            driver = webdriver.Chrome(options=options)
            driver.get("https://www.google.com/?hl=zh_tw")
            url = "https://doc.twse.com.tw/server-java/t57sb01?id=&key=&step=1&co_id={}&year={}&seamon=&mtype=F&dtype=F04".format(
                ticker, year
            )
            driver.get(url)
            print(url)
            # 點選PDF URL
            time.sleep(5)
            try:
                file_element = driver.find_element(By.TAG_NAME, "a")
            except NoSuchElementException:
                over_download_tag = driver.find_element(By.TAG_NAME, "center").text
                print(over_download_tag, "暫停5分鐘再抓取")
                time.sleep(300)
                driver.get(url)
                print(url)
                time.sleep(5)
                file_element = driver.find_element(By.TAG_NAME, "a")
                
            file_name = file_element.text
            download_file_list = os.listdir(download_path)
            download_file_list = [x for x in download_file_list if "t57sb01" in x]
            file_element.click()
            for download_file in download_file_list:
                os.remove(f"{download_path}{download_file}")

            if ".pdf" in file_name:
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
                    with open(f"{save_path}{file_name}", "wb") as f:
                        f.write(response.content)
                        f.close()
                    print(file_name + "-" * 10 + "ok")
                else:
                    print("existing file : {}".format(file_name))
            if ".doc" in file_name:
                download_file = "t57sb01.doc"
                sub_file_src_path = os.path.join(download_path, download_file)
                sub_file_dst_path = os.path.join(
                    save_path,
                    file_name,
                )
                shutil.copyfile(sub_file_src_path, sub_file_dst_path)
                os.remove(sub_file_src_path)
            if ".zip" in file_name:
                download_file = "t57sb01.zip"
                download_file_path = f"{download_path}{download_file}"
                while not os.path.isfile(download_file_path):
                    pass
                download_file_baename = download_file.replace(".zip", "")
                extract_dir = f"{download_path}{download_file_baename}/"
                with zipfile.ZipFile(download_file_path, "r") as zf:
                    if not os.path.exists(extract_dir):
                        os.makedirs(extract_dir)
                    zf.extractall(extract_dir)
                    zf.close()
                for s, sub_file in enumerate(os.listdir(extract_dir)):
                    print(s, sub_file)
                    sub_file_src_path = os.path.join(
                        download_path, download_file_baename, sub_file
                    )
                    sub_file_dst_path = os.path.join(
                        save_path,
                        f"{file_name.replace('.zip', '')}_{s}.{sub_file.split('.')[-1]}",
                    )
                    shutil.copyfile(
                        sub_file_src_path,
                        sub_file_dst_path,
                    )
                    os.remove(sub_file_src_path)
                os.removedirs(extract_dir)
                os.remove(download_file_path)

        except Exception as e:
            print("processing fail : {}".format(ticker))
            print(e)
            fail_list.append((year, ticker))

df_final_fail = pd.DataFrame()
df_final_fail["year"] = [x[0] for x in fail_list]
df_final_fail["ticker"] = [x[1] for x in fail_list]
df_final_fail["time"] = datetime.now().strftime("%Y%m%d")
