# %% Import Packages
import os
import shutil
import time
import zipfile
from datetime import datetime

import pandas as pd
import requests
from cathaysite import cathay_db as db
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

# 設定起始年和結束年
start_year = 114
end_year = 115
year_list = [str(x) for x in range(start_year, end_year)]

use_list = False

# 暫存路徑
download_path = "C:/Users/wa-00100809/Downloads/"

db_config = db.db_config(db="tdds")
sql_query = "select * from tej_twn_anprcstd where [stype] in ('STOCK','FSTOCK') and [mkt] in ('TSE','OTC')"
df = db.pd_read_mssql_data(sql_query, database="FinanceOther", **db_config)

ticker_list = df["coid"].str.strip()
ticker_list.reset_index(drop=True, inplace=True)
options = Options()
options.add_argument("--headless=new")
service = Service()

# %% Main loop
fail_list = []
for year in year_list:
    # 最終存放路徑
    save_path = f"./output/{year}/"
    if not os.path.exists(os.getcwd() + save_path):
        os.makedirs(os.getcwd() + save_path)

    result = [
        f
        for f in os.listdir(os.getcwd() + save_path)
        if os.path.isfile(os.path.join(os.getcwd() + save_path, f))
    ]
    ok_ticker_list = [t.split("_")[1] for t in result]

    not_yet_list = sorted(list(set(ticker_list.tolist()) - set(ok_ticker_list)))
    print(len(not_yet_list))
    print(not_yet_list)

    for ticker in not_yet_list:  # 執行目前不在資料夾內的
        try:
            print("-" * 100)
            print("-" * 100)
            print("processing : {}".format(year))
            print("processing : {}".format(ticker))
            driver = webdriver.Chrome(service=service, options=options)
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
                print(over_download_tag, "暫停3分鐘再抓取")
                time.sleep(180)
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
                response = requests.get(link, allow_redirects=True, verify=False)
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

# output fail list
df_final_fail = pd.DataFrame()
df_final_fail["year"] = [x[0] for x in fail_list]
df_final_fail["ticker"] = [x[1] for x in fail_list]
df_final_fail["datetime"] = datetime.now().strftime("%Y%m%d%H%M")
df_final_fail.to_csv(
    f"./fail/fail_list_{datetime.now().strftime('%Y%m%d%H%M')}.csv", index=False
)

# %%
