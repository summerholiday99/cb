# -*- coding: utf-8 -*-
"""
通过每周转债名单做每日转债名单
计算转债每日的涨幅，偏离度
(不小心把手动的删掉了，现在来补上）

f_month = lambda x:str(x.year)+'-'+str(x.month).zfill(2)
f_date =  lambda x:dt.datetime.strptime('','%Y%m%d')
temp = df.groupby(by=['month']).size()
import datetime as dt
import numpy as np
"""

import pandas as pd
import warnings
warnings.filterwarnings("ignore")

def holdingmarketvalue_balance_ratio(str_raw,date):
#    date = '20200814'
    str_cb = r'{}\量化定价\转债数据.xlsx'.format(str_raw)
    str_pos = r'{}\相关工作\#持仓统计\input\综合信息查询_组合证券_{}.xls'.format(str_raw, date)
    str_mutual_priviate = r'{}\相关工作\#持仓统计\公募私募名单.xlsx'.format(str_raw)
    str_out = r'{}\相关工作\#持仓统计\output\转债持仓占比统计-{}.xlsx'.format(str_raw, date)

    def f(df_aa):
        # 这边发来的持仓数据是3.20的，计算的是3.20的持仓（XX张）/债券剩余张数(余额乘1亿，除以100元，就是剩余张数)
        gg = df_aa.groupby(by=['证券名称'])
        df_s = pd.DataFrame(gg[['持仓','市值']].sum())
        temp = pd.merge(left=df_s, right=df[['转债简称','债券剩余张数']],left_index=True , right_on = '转债简称')
        temp = temp[['转债简称','市值','持仓','债券剩余张数']]
        temp['市值'] = temp['市值']/10000
        temp.rename(columns={'市值':'市值(万)'},inplace=True)
        temp['持仓占比'] = temp['持仓']/temp['债券剩余张数']
        temp.set_index('转债简称' ,inplace=True)
        return temp

    df = pd.read_excel(str_cb,sheet_name='条款')    
    df_aim = pd.read_excel(str_pos,skipfooter=1) 
    df_ = pd.read_excel(str_mutual_priviate)    
    df['债券剩余张数'] = df['债券余额']*1000000 
    df_aim = df_aim[df_aim['证券名称'].isin(df['转债简称'].values.tolist())] 
    df_aim = pd.merge(left=df_aim, right=df_)         # 区分公募私募
    
    temp_total = f(df_aa=df_aim)
    temp_hedge = f(df_aa=df_aim[df_aim['性质']=='私募'])
    temp_mutual = f(df_aa=df_aim[df_aim['性质']=='公募'])
    temp_total['公募持仓占比'] =  temp_mutual['持仓占比']
    temp_total['私募持仓占比'] =  temp_hedge['持仓占比']
    temp_total.sort_values(by=['公募持仓占比'], ascending=False,inplace=True)
    
    writer = pd.ExcelWriter(str_out)
    temp_total.to_excel(writer,sheet_name='汇总')
    writer.save()
    return

if __name__ =='__main__':
    date='20200814'
    str_raw = r'D:\同步\我的坚果云\可转债'
    holdingmarketvalue_balance_ratio(str_raw,date)
