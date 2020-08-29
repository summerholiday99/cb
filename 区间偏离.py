# -*- coding: utf-8 -*-
"""
f_month = lambda x:str(x.year)+'-'+str(x.month).zfill(2)

temp = df.groupby(by=['month']).size()
"""
import datetime as dt
import numpy as np
import pandas as pd
import warnings
warnings.filterwarnings("ignore")
import WindPy as wind
wind.w.start()

def tool_df2group(df_r):
    '''选定一列，按其值展开成一个矩阵'''
    # 前提条件，index columns 定了，不能有重复项
    str_index_name = '信用等级'
    str_column_name = '价格区间'
    str_data = '定价偏离'
    if max(df_r.groupby(by=[str_index_name,str_column_name]).size())==1:
        df_base = pd.Series(df_r[str_data].values, index=[df_r[str_index_name], df_r[str_column_name]]) #多级索引
        df_aim = df_base.unstack() # 索引展开，你这编程，有待长进啊!! 太菜了，多看看别人怎么写
        return df_aim
    else:
        print('有重复项！')
        
def tool_f(x,list_r):
    if x<list_r[0]:
        return '0~{}'.format(str(int(list_r[0])))
    elif x>=list_r[0] and x<list_r[-1]:
        for ii in range(0,len(list_r)-1):
            if x>=list_r[ii] and x<list_r[ii+1]:
                return '{}~{}'.format(str(int(list_r[ii])),str(int(list_r[ii+1])))
            else:
                pass
    elif x>=list_r[-1]:
        return '{}~'.format(str(int(list_r[-1])))


def tool_turn(df,str_date_s, str_date):
    temp = wind.w.wsd(df['转债代码'].tolist(), "pq_amount", str_date_s, str_date, "unit=1;bondPriceType=2")
    temp = pd.DataFrame([ii[-1] for ii in temp.Data], index=df['转债代码'].tolist(),columns=temp.Fields)/100000000
    temp.rename(columns={'PQ_AMOUNT':'周成交额'}, inplace=True)
    temp2 = wind.w.wsd(df['转债代码'].tolist(), "outstandingbalance",str_date,str_date)
    temp2 = pd.DataFrame(temp2.Data[0], index=df['转债代码'].tolist(),columns=temp2.Fields)
    temp2.rename(columns={'OUTSTANDINGBALANCE':'转债余额'}, inplace=True)
    temp['转债余额'] = temp2['转债余额']
    temp['周换手率'] = temp['周成交额']/temp2['转债余额']
    return temp 

def tool_credit(df,str_date_s, str_date):
    temp = wind.w.wsd(df['转债代码'].tolist(),"latestissurercreditrating", str_date_s, str_date, "")
    temp = pd.DataFrame([ii[-1] for ii in temp.Data], index=df['转债代码'].tolist(),columns=['信用等级'])
    return temp 

def process_data(df, str_date):    
    
    list_r = [100, 110, 120, 130, 150]
    df['价格区间'] = df['结算价'].apply(lambda x:tool_f(x, list_r))
    #下载转债余额和周成交额, 计算换手率
    f_date =  lambda x:dt.datetime.strptime(x,'%Y-%m-%d')
    f_date2 =  lambda x:dt.datetime.strftime(x,'%Y-%m-%d')
    str_date_s = f_date2(f_date(str_date) - pd.Timedelta(days=7))
    
    temp = tool_turn(df,str_date_s, str_date)
    df = pd.merge(df,temp[['周换手率']],left_on=['转债代码'],right_index=True)
    temp = tool_credit(df,str_date_s, str_date)
    df = pd.merge(df,temp[['信用等级']],left_on=['转债代码'],right_index=True)
    
    df = df[(df['结算价']-df['纯债价值']>0)&(df['周换手率']<5)]        
    return df

def stat_group(df):
    gg = df.groupby(by=['信用等级','价格区间'],as_index=False)
    a = gg[['定价偏离']].mean()
    #b = gg[['定价偏离']].median()
    df_aim = tool_df2group(df_r=a)
    return df_aim


#%%
str_date = '2020-05-08'
str_road = r'C:\Users\huangjili\Desktop\同步\OneDrive\可转债\量化定价\delta\{}.xlsx'.format(str_date)
df = pd.read_excel(str_road)
df = process_data(df, str_date)
df_aim = stat_group(df)
df_aim.to_clipboard()

