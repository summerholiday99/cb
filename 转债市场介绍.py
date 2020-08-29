# -*- coding: utf-8 -*-
"""
针对数据库cb_panel做转债介绍
1.转债支数
2.转债余额
3.各等级角度
f_month = lambda x:str(x.year)+'-'+str(x.month).zfill(2)
temp = df.groupby(by=['month']).size()
"""

import datetime as dt
import numpy as np
import pandas as pd
import warnings
warnings.filterwarnings("ignore")

def tool_process_data(df):
    f_date = lambda x:dt.datetime.strptime(str(x),'%Y%m%d')
    df['交易日'] = df['交易日'].apply(f_date)
    df = df[df['交易日']>pd.Timestamp(year=2010,month=1,day=1)]
    return df

def tool_df2group(df_r,str_index_name,str_column_name,str_data):
    '''选定一列，按其值展开成一个矩阵'''
    # 前提条件，index columns 定了，不能有重复项
    if max(df_r.groupby(by=[str_index_name,str_column_name]).size())==1:
        df_base = pd.Series(df_r[str_data].values, index=[df_r[str_index_name], df_r[str_column_name]]) #多级索引
        df_aim = df_base.unstack() # 索引展开，你这编程，有待长进啊!! 太菜了，多看看别人怎么写
        return df_aim
    else:
        print('有重复项！')

def stat_num_amount(df):
    df['数量'] = 1
    gg = df.groupby(by=['交易日'])
    temp = gg[['数量','转债余额']].sum()
    return temp

def stat_num_credit(df):
    df['数量'] = 1
    gg = df.groupby(by=['交易日','信用等级'], as_index=False)
    temp = gg[['数量']].sum() 
    temp = tool_df2group(df_r=temp, str_index_name='交易日', str_column_name='信用等级', str_data='数量')
    return temp

#%%
str_road = r'C:\Users\huangjili\Desktop\cb_panel_2020-08-06.xlsx'
df = pd.read_excel(str_road)
df = tool_process_data(df)

df_1 = stat_num_amount(df)
df_2 = stat_num_credit(df)
df_2.to_clipboard()

str_road = r'D:\东证资管\可转债\#相关课题\转债市场概述\转债市场介绍.xlsx'
df = pd.read_excel(str_road,sheet_name='正股行业')
len(df['一级行业'].unique())
len(df['二级行业'].unique())


str_road = r'C:\Users\huangjili\Desktop\cb_panel_2020-08-06.xlsx'
df = pd.read_excel(str_road)
