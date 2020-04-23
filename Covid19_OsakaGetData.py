#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options


# In[2]:


url = "https://covid19-osaka.info"
table_class_name='v-data-table__wrapper'
output_file_name='Coivd19_Osaka_Table.xlsx'
options = Options()
#options.add_argument('--headless')
options.add_experimental_option('prefs', {'intl.accept_languages': 'ja'})
driver = webdriver.Chrome(options=options)
driver.implicitly_wait(30)


# In[3]:


driver.get(url)
driver.implicitly_wait(15)

tableElem = driver.find_element_by_class_name(table_class_name)
ths = tableElem.find_elements(By.TAG_NAME, 'th')
trs = tableElem.find_elements(By.TAG_NAME, "tr")

l_header = []
# テーブルのヘッダを取得してリストに格納する
for th in ths:
    l_header.append(th.text)
# テーブルのヘッダのリストをDataFrameテーブルのコラムにする
df_Table = pd.DataFrame(columns=l_header)
# 進捗表示用
len_trs = len(trs)
cnt_trs = 0
# テーブルの内容を取得する
for tr in trs:
    # 進捗表示
    cnt_trs += 1
    print("\rProcessing ({0}/{1})".format(cnt_trs, len_trs), end="")
    l_data = []
    tds = tr.find_elements(By.TAG_NAME, "td")
    # テーブルのヘッダ(長さ０)も拾ってしまうのでここでふるいにかける
    if len(tds) > 0:
        # テーブルの内容を１行毎に取得してリストに格納する
        for td in tds:
             l_data.append(td.text)
        # １行毎のテーブルの内容のリストをSeriesに格納する
        s_data = pd.Series(l_data, index = df_Table.columns)
        # １行毎のテーブルの内容のSeriesをDataFrameテーブルに追加する
        df_Table = df_Table.append(s_data, ignore_index=True)
# 取得結果をExcelファイルに出力する
df_Table.to_excel(output_file_name)
driver.close()

print('\nFinish !')

