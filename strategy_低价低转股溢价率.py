# -*- coding: utf-8 -*-
'''
看下标准的框架怎么命名
输出：
作图：加上转债余额
'''

import datetime as dt
import numpy as np
import pandas as pd
import warnings
warnings.filterwarnings("ignore")
import matplotlib.pyplot as plt
import seaborn as sns

def cb_strategy_1(df, list_q):
    # 2*价格分位数倒序+溢价率分位数倒序    #越低越好
    df = df.copy()
    gg = df.groupby(by=['交易日'])
    df[['收盘价_r','转股溢价率_r']] = gg[['收盘价','转股溢价率']].rank(ascending=True)
    df['score'] = 2*df['收盘价_r'] + df['转股溢价率_r']
    gg = df.groupby(by=['交易日'])
    df['score_r'] = gg['score'].rank(pct=True)  
    df_pick = df[(df['score_r']>=list_q[0])&(df['score_r']<list_q[1])]   
    return df_pick

def cb_rev_weight(df_pick, list_q):
    '''转债余额加权'''
    gg = df_pick.groupby(by=['交易日'])
    
    temp = gg['转债余额'].sum().to_frame()
    temp.columns = ['转债余额和']
    df_pick = pd.merge(df_pick, temp, left_on='交易日', right_index=True, how='left')
    df_pick['组合权重'] = df_pick['转债余额']/df_pick['转债余额和']
    df_pick['涨跌幅_下一交易日_weighted'] = df_pick['涨跌幅_下一交易日']*df_pick['组合权重']
    df_pick.sort_values(by=['交易日','组合权重'],inplace=True)
    
    gg = df_pick.groupby(by=['交易日'])
    k0 = gg.size().to_frame()
    k1 = gg['涨跌幅_下一交易日_weighted'].sum()   
    k2 = (gg['涨跌幅_下一交易日_weighted'].sum()+1).cumprod().to_frame()
    df_aim  = pd.concat([k2,k1,k0], axis=1)
    str_add = str(list_q[0])+'~'+str(list_q[1])
    df_aim.columns = [ii+ str_add for ii in ['净值','组合涨跌幅','持仓数量']]
    return df_pick,df_aim

def cb_rev_mean(df_pick, list_q):
    '''注意未来函数：不要当日选券再用当日收益率算'''
    gg = df_pick.groupby(by=['交易日'])
    k0 = gg.size().to_frame()  # 数量
    k1 = gg['涨跌幅_下一交易日'].mean()   
    k2 = (gg['涨跌幅_下一交易日'].mean()+1).cumprod().to_frame()
    df_aim  = pd.concat([k2,k1,k0], axis=1)
    str_add = str(list_q[0])+'~'+str(list_q[1])
    df_aim.columns = [ii+ str_add for ii in ['净值','组合涨跌幅','持仓数量']]
    return df_aim

def cb_data_prepare(df):
    '''日期，收益率'''
    f_date =  lambda x:dt.datetime.strptime(str(x),'%Y%m%d')
    df['交易日'] = df['交易日'].apply(f_date)
    df.sort_values(by=['交易日'],inplace=True)
    df = df[df['交易日']>=dt.datetime.strptime('20171231','%Y%m%d')]
    
    df['涨跌幅'] = df['涨跌幅']/100
    gg = df.groupby(by=['转债代码'])
    df['涨跌幅_下一交易日'] = gg['涨跌幅'].shift(-1)
    return df

def tool_trade_analysis(df, col_name, rate='D'):
    '''输入: df_r_trade
    统计收益率, 年化收益, 年化波动, 年化夏普, 最大回撤;
    交易次数，胜率, 盈亏比'''
    if rate=='D':
        multi = 245
    elif rate=='M':
        multi = 12
    #df.dropna(axis=0,inplace=True)   #如果要全部相同时段，就用这个
    pnlseries = df[col_name]
    pnlseries.dropna(axis=0,inplace=True)    # dropna只针对那一列  
    nav = (pnlseries + 1).cumprod()
    ret0 = (pnlseries + 1).groupby(pnlseries.index.year).prod() - 1
    dd = (nav / nav.groupby(nav.index.year).cummax() - 1).groupby(nav.index.year).min()
    alldd = (nav / nav.cummax() - 1).min()
    allret0 = (1 + pnlseries).prod() - 1
    # rets = pnlseries.groupby(pnlseries.index.year).mean() * multi  #你需要一个准一点的年化 
    rets = ret0*multi/(pnlseries + 1).groupby(pnlseries.index.year).size() # 成功
    std = pnlseries.groupby(pnlseries.index.year).std() * (multi** 0.5)
    sharpe = rets / std
    #allret = pnlseries.mean() * multi   # 之前这个算的太糙
    allret = (allret0+1)**(multi/len(pnlseries))-1
    allstd = pnlseries.std() * (multi ** 0.5)
    if allstd!=0:
        allsharpe = allret / allstd
    else:
        allsharpe = 0

    allsummary = pd.DataFrame([allret0, allret, allstd, alldd, allsharpe]).T
    allsummary.index = ['all']
    allsummary.columns = ['收益率', '年化收益', '年化波动', '最大回撤', '年化夏普']
    df_r = pd.concat([ret0, rets, std, dd, sharpe], axis=1)  # 先横着合并
    df_r.columns = ['收益率', '年化收益', '年化波动', '最大回撤', '年化夏普']
    df_r = pd.concat([df_r, allsummary], axis=0)  # 再竖着合并
    df_r['标的'] = col_name
    df_r = df_r[['标的','收益率','年化收益','年化波动','最大回撤','年化夏普']]
    return df_r


def tool_output(df_pick, df_aim, df_r,str_road_out):
    '''输出：更新时间，净值:最新持仓及权重'''
    df_pick = df_pick[df_pick['交易日']==dt.datetime.strptime(str_date,'%Y-%m-%d')]
    file = pd.ExcelWriter(str_road_out)
    df_aim.to_excel(file,sheet_name='净值')
    df_r.to_excel(file,sheet_name='风险收益')
    df_pick.to_excel(file,sheet_name='持仓')
    file.save()
    return

#%% Main   
str_date = '2020-08-21'
str_raw = r'C:\Users\huangjili\Desktop\同步\OneDrive\可转债'  # 在家在公司换一下
str_road = r'{}\cb_panel_{}.xlsx'.format(str_raw,str_date)       
str_road_out = r'{}\策略回测\策略-{}.xlsx'.format(str_raw,str_date)       

list_q = [0, 0.2] #分位数

df = pd.read_excel(str_road)
df_c = cb_data_prepare(df)
df_pick = cb_strategy_1(df_c, list_q)
df_pick,df_aim = cb_rev_weight(df_pick, list_q)
df_r = tool_trade_analysis(df_aim, col_name='组合涨跌幅0~0.2', rate='D')
tool_output(df_pick, df_aim, df_r, str_road_out)

sns.set_style('whitegrid')
plt.rcParams["font.sans-serif"] = ["SimHei"]
plt.rcParams["axes.unicode_minus"] = False
df_aim[['净值0~0.2']].plot()



#%%
import tool
dd = tool.rev_risk.drawdown(df_aim,'净值0~0.2')
tool.plot.p_netvalue_drawdown(dd,['净值0~0.2','回撤_滚动1年'])

import os
os.chdir(r'C:\Users\huangjili\Desktop\同步\OneDrive')
tool.rev_risk.rev_yqm(df_aim[['净值0~0.2']], rate='M').to_clipboard()

df_pick.to_clipboard()
