# -*- coding: utf-8 -*-
"""
Created on Wed Oct 14 15:29:16 2020

@author: zoutianshu
"""

import pandas as pd
from WindPy import w
w.start()

df = pd.read_excel("501001.xls")
df1 = df[['日期','证券代码','证券名称','市值比净值(%)']]
df1 = df1.iloc[:-1]
df1['证券代码'] = [str(int(i)).zfill(6) for i in df1['证券代码']]

stk_list = df1['证券代码'].unique().tolist()
stk_str = ','.join(stk_list)
data = w.htocode(stk_str, "stocka")
stk_list1 = pd.DataFrame(data.Data[0])[0].tolist()
stk_str1 = ','.join(stk_list1)

date_list = df1['日期'].sort_values().unique().tolist()

data = w.wsd(stk_str1, "industry_sw", date_list[-1], date_list[-1], "industryType=1")

df_indus = pd.DataFrame(data.Data[0],index = stk_list,columns = ['行业'])

df1 = pd.merge(df1, df_indus,left_on = ['证券代码'],right_index = True, how = 'left')

df2 = df1.pivot_table(index='日期',columns='行业',values='市值比净值(%)',aggfunc=sum).fillna(0)
df_indus_avg = df2.T.mean(axis = 1).sort_values(ascending = False)
