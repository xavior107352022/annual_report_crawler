import shutil
import os
import pandas as pd

src_path = "./Share_Holder_Ann/"
dst_path = "./Share_Holder_Ann_List/"
src_file_list = os.listdir(os.getcwd() + src_path)
df_list = pd.read_excel("股票代號清單.xlsx")
ticker_list = df_list["股票代號"].astype("str").to_list()
src_file_list = [x for x in src_file_list if x.split("_")[1] in ticker_list]

for src_file in src_file_list:
    shutil.copy(src=src_path + src_file, dst=dst_path + src_file)
