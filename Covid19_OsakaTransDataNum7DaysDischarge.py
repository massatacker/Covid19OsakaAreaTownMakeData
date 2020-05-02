#!/usr/bin/env python
# coding: utf-8

# In[ ]:


from datetime import datetime
from datetime import timedelta
import pandas as pd
import numpy as np
source_filename = 'Coivd19_Osaka_Table.xlsx'
trans_filename = 'OsakaAreaTownTransList.xlsx'
area_output_filename = 'Covid19_Osaka_AreaNumData.xlsx'
town_output_filename = 'Covid19_Osaka_TownNumData.xlsx'
clmns_area = ['日付', '区域', '日別', '退院・解除', '週平均', '累計', '退院・解除累計']
clmns_town = ['日付', '市町村', '日別', '退院・解除', '週平均', '累計', '退院・解除累計']
mrk_discharge = '○'


# In[ ]:


# _startから_endまで、１日毎にdatetimeの変数を返す
# 注意．通常のPythonのrange関数は_startから(_end-1)までを返すが
# 　　　今回は使用目的の関係上、+1の(_end±0)まで返す
def daterange_p1(_start, _end):
    _endp1 = _end + timedelta(days=1)
    for n in range((_endp1 - _start).days):
        yield _start + timedelta(n)


# In[ ]:


# 抽出された単位で感染者数を集計する関数
def add_up_num_infection( df_area, area, first_day, last_day, df_num ):
    # 感染日の一覧を抽出する
    infection_dates = df_area['日付'].unique()
    # 感染日の一覧を昇順に並べ換える
    infection_dates.sort()
    # 陽性者の累計人数計算用の変数をクリアする
    num_total_infection_people = 0
    # 退院・解除者の累計人数計算用の変数をクリアする
    num_total_discharge_people = 0
    # 感染日初日から最終日までループする
    #first_day = pd.to_datetime(infection_dates[0])
    #last_day = pd.to_datetime(infection_dates[-1])
    for infection_date in daterange_p1(first_day, last_day):
        # 感染日毎のデータを抽出する
        df_area_date = df_area[ df_area['日付']==infection_date ]
        # 感染日毎の陽性者の人数を求める
        num_infection_people = len(df_area_date)
        # 陽性者の累計人数を求める
        num_total_infection_people += num_infection_people
        # 感染日毎の退院・解除者の人数を求める
        num_discharge_people = len(df_area_date[ df_area_date['退院・解除']==mrk_discharge ])
        # 退院・解除者の累計人数を求める
        num_total_discharge_people += num_discharge_people
        # [感染日、居住地、感染日の人数]のSeriesを作成する
        s_infection = pd.Series(
                            [infection_date, 
                             area,
                             num_infection_people,
                             num_discharge_people,
                             0,
                             num_total_infection_people,
                             num_total_discharge_people],
                             index = df_num.columns)
        # 作成したSeriesをdf_numに追加する
        df_num = df_num.append(s_infection, ignore_index=True)
        # 日別感染者数の７日間移動平均を計算して格納する
        df_num['週平均'] = df_num['日別'].rolling(7).mean()
    # 集計結果を返す
    return df_num    


# In[ ]:


#集計結果を格納するためのDataFrameを宣言する
df_areas_num = pd.DataFrame(columns=clmns_area)
df_towns_num = pd.DataFrame(columns=clmns_town)

# 元データを読み込む
df_source = pd.read_excel(source_filename)

# 元データの「居住地」に対応する区域と市町村の一覧表を読み込む
df_trans = pd.read_excel(trans_filename)
df_trans = df_trans.set_index('居住地')


# In[ ]:


# 元データの「日付」の文字列をdatetime型に変換する
# 「日付」に西暦を付加する
df_source['日付']='2020/'+df_source['日付']
# 「日付」の文字列をdatetime型に変換する
df_source['日付']=pd.to_datetime(df_source['日付'], format='%Y/%m/%d')

# データ格納開始日と終了日を決定する
first_day = df_source.iloc[-1]['日付']
last_day = df_source.iloc[0]['日付']


# In[ ]:


# 元データの「居住地」に対応する区域と市町村を割当てる
# 「居住地の」一覧を抽出する
residences = df_source['居住地'].unique()
# 「居住地」毎にループする
for residence in residences:
    l_trans = df_trans.index.tolist()
    if residence in l_trans:
        # 「居住地」が登録されている場合
        df_source.loc[ df_source['居住地']==residence, '区域' ] = df_trans.at[residence,'区域']
        df_source.loc[ df_source['居住地']==residence, '市町村' ] = df_trans.at[residence,'市町村']
    else:
        # 「居住地」が登録されていない場合
        df_source.loc[ df_source['居住地']==residence, '区域' ] = 'その他'
        df_source.loc[ df_source['居住地']==residence, '市町村' ] = residence


# In[ ]:


# 区域単位で集計する
# 区域の一覧を抽出する
areas = df_source['区域'].unique()
# 区域単位毎にループする
for area in areas:
    print('\r{0}        '.format(area), end='')
    # 区域単位のデータを抽出する
    df_area = df_source[ df_source['区域']==area ]
    # 区域単位の集計結果を格納するためのデータフレームを宣言する
    df_area_num = pd.DataFrame(columns=clmns_area)
    # 集計する
    df_area_num = add_up_num_infection( df_area, area, first_day, last_day, df_area_num )
    # 区域全体の集計結果に追加格納する
    df_areas_num = df_areas_num.append(df_area_num, ignore_index=True)
# 全域で集計する(おまけ)
print('\r全域        ', end='')
# 区域単位の集計結果を格納するためのデータフレームを宣言する
df_area_num = pd.DataFrame(columns=clmns_area)
# 集計する
df_area_num = add_up_num_infection( df_source, '全域', first_day, last_day, df_area_num )
# 区域全体の集計結果に追加格納する
df_areas_num = df_areas_num.append(df_area_num, ignore_index=True)
# 結果をExcelファイルに出力する
df_areas_num.to_excel(area_output_filename)

# 市町村単位で集計する
# 市町村の一覧を抽出する
towns = df_source['市町村'].unique()
# 区域単位毎にループする
for town in towns:
    print('\r{0}        '.format(town), end='')
    # 市町村単位のデータを抽出する
    df_town = df_source[ df_source['市町村']==town ]
    # 市町村単位の集計結果を格納するためのデータフレームを宣言する
    df_town_num = pd.DataFrame(columns=clmns_town)
    # 集計する
    df_town_num = add_up_num_infection( df_town, town, first_day, last_day, df_town_num )
    # 区域全体の集計結果に追加格納する
    df_towns_num = df_towns_num.append(df_town_num, ignore_index=True)
# 結果をExcelファイルに出力する
df_towns_num.to_excel(town_output_filename)

print('\rFinish !              ')


# In[ ]:





# In[ ]:




