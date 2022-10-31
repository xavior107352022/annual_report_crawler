import requests
import pandas as pd


def get_info():
    post_dict = dict(encodeURIComponent=1, step=1, firstin=1, TYPEK='sii')
    r = requests.post("https://mops.twse.com.tw/mops/web/ajax_t51sb01", data=post_dict)
    r.encoding = 'utf-8'
    df_list = pd.read_html(r.text)
    df_info_tse = df_list[0]
    df_info_tse = df_info_tse[df_info_tse['公司代號'] != '公司代號']

    post_dict = dict(encodeURIComponent=1, step=1, firstin=1, TYPEK='otc')
    r = requests.post("https://mops.twse.com.tw/mops/web/ajax_t51sb01", data=post_dict)
    r.encoding = 'utf-8'
    df_list = pd.read_html(r.text)
    df_info_otc = df_list[0]
    df_info_otc = df_info_otc[df_info_otc['公司代號'] != '公司代號']
    return df_info_tse, df_info_otc