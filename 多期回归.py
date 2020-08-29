# -*- coding: utf-8 -*-
"""
Created on %(date)s
f_month = lambda x:str(x.year)+'-'+str(x.month).zfill(2)
f_date =  lambda x:dt.datetime.strptime('','%Y%m%d')
temp = df.groupby(by=['month']).size()
"""

import numpy as np
import pandas as pd
import warnings
warnings.filterwarnings("ignore")
import datetime as dt
import shelve
import statsmodels.api as sm

#%% 数据准备&主函数
#df_data = pd.read_excel(r'D:\东证资管\可转债\#相关课题\转债历史估值\估值偏离-全历史-20191129.xlsx')
#df_data.drop_duplicates(subset=['转债代码','ctime'], inplace=True)              # 之前有一点重复值
#df_data.to_clipboard()

#group = df_data.groupby(by=['转债代码'])
#df_data['半年前定价偏离'] = group['定价偏离'].shift(24)
#group = df_data.groupby(by=['ctime']) 

df_data = pd.read_excel(r'C:\Users\jili\Desktop\估值变动-正股涨跌幅.xlsx')
df_data['估值偏离_delta'] = df_data['定价偏离'] - df_data['定价偏离_半年前'] 
df_data.dropna(inplace=True)
df_data['信用等级_num'] = df_data['信用等级'].apply(f_credict)

list_date = list(df_data['ctime'].unique())
df_aim = jili_hurry_up(df_data,list_date)
df_aim.loc[df_aim['类别']=='系数','近半年正股涨幅'].plot()
df_aim.to_clipboard()

#%%多期回归

def ols(df):                  
    dict_ols = {'Y':['估值偏离_delta'],
                'X':['信用等级_num','正股区间涨跌幅']}  
    #data_ols = df_ols2[df_ols2['业绩预告']!=0]
    #data_ols = df_ols2.loc[df_ols2['int_month']==int_month]  # data_ols.shape
    data_ols = df 
    regy = np.array(data_ols[dict_ols['Y'][0]].astype('float'))
    regx = np.array(data_ols[dict_ols['X']]) 
    #regx = sm.add_constant(regx)
    regr = sm.OLS(regy, regx)                                              
    res = regr.fit()
    a = pd.DataFrame([res.params,res.tvalues],
                     index=[df['ctime'].unique()[0]]*2,
                     columns=['信用等级','近半年正股涨幅'])
    a['类别'] = ['系数','t值']
    return a

def jili_hurry_up(df_data, list_date):
    con = []
    for ii in list_date:
        temp = df_data[df_data['ctime']==ii]
        temp2 = ols(temp)
        con.append(temp2)
    df_aim = pd.concat(con)    
    return df_aim

#%% 单期回归
def f_credict(x):
    list_cre = ['A+','AA-','AA', 'AA+','AAA']
    return list_cre.index(x)+1 

df = pd.read_excel('D:\东证资管\可转债\#相关课题\转债因子模型.xlsx')
df['信用等级'] = df['信用等级'].apply(f_credict)
df.to_clipboard()
                       
def reg(df):
#    dict_ols = {'Y':['区间收益率'],
#                'X':['000832.CSI','对数自由流通市值','BPS/P','信用等级','发行费用占募资规模比例']}  
    dict_ols = {'Y':['区间收益率'],
                'X':['000832.CSI','对数自由流通市值','BPS/P','信用等级','发行费用占募资规模比例']}         
    data_ols = df#[df['信用等级'].isin([4,5])]
    regy = np.array(data_ols[dict_ols['Y'][0]].astype('float'))
    regx = np.array(data_ols[dict_ols['X']]) 
    regx = sm.add_constant(regx)
    regr = sm.OLS(regy, regx)                                              
    res = regr.fit()
    res.summary()