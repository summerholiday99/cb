# -*- coding: utf-8 -*-
""" WORK
1.历史数据要有，可以用周频, 但要对
2.写因子，选持仓
3.计算表现，表现评价
"""

import pandas as pd
import numpy as np

#%%  策略信号 写信号-出持仓&权重 (目前不用管停牌啥的)

def order_proc(last_hold,pool,b,s,aas):
    # 先决定本期收盘前标的是谁
    ke = last_hold 
    if aas in ke: ke.remove(aas)        # 先删了后面再加
    ke = set(ke)|set(b) - set(s)  # 不在hold中的,加上去, 在卖单中的，不再取 
    ke = list(ke&set(pool))
    # 然后来决定各自权重 (权重要求不高)
    if len(ke)==0:
        ke = [aas]
        kv = [1]
    elif (len(ke)>0)&(len(ke)<10):
        kv = [0.1]*len(ke)
        kv.append(1-0.1*len(ke))
        ke.append(aas)
    elif (len(ke)>=10):
        kv = [1/len(ke)]*len(ke)
    return dict(zip(ke,kv))

def order_pro(buy, sell, df_price, alter_asset=ass_name):
    '''1.存在 2.满足买卖阈值 两个条件要同时满足'''
    judge = ~df_price.isna()
    arr_prepare = np.array(df_price.columns)
    list_pool = [arr_prepare[np.array(judge.iloc[i,:])] for i in range(df_price.shape[0])]
        
    # 然后开始看每一期持仓
    list_hold = []      
    list_pro = list(zip(buy,sell))  # 不会同时要买又要卖的
    for period in range(len(list_pro)):
        # b,s = list_pro[30]
        b,s = list_pro[period]
        if len(list_hold)==0:
            last_hold = []
            pool = list_pool[0+1] 
            temp = order_proc(last_hold,pool,b,s,alter_asset)
            list_hold.append(temp)        
        elif len(list_hold)>0:    
            last_hold = list(list_hold[-1].keys())
            pool = list_pool[min(period+1,len(list_pool)-1)] 
            temp = order_proc(last_hold,pool,b,s,alter_asset)
            list_hold.append(temp)   

    # 每天是一个字典，最后DataFrame,concat起来 
    con = []
    for i in range(len(list_hold)):
        t1 = pd.DataFrame(list(list_hold[i].values()), index=list(list_hold[i].keys()),columns=['weight'])
        t1['date'] = df_price.index[i]
        con.append(t1)
    df_weight = pd.concat(con)
    df_weight['code'] = list(df_weight.index)
    df_weight.set_index('date',inplace=True)
    return df_weight

def signalPort(df_factor, df_price):
    # 买入信号，卖出信号，持仓数据
    indi_buy = 100
    indi_sell = 130
    df_factor = df_price.copy()  # 转债资产
    buy = [list(df_factor.loc[t1,:][df_factor.loc[t1,:]<indi_buy].index) for t1 in df_factor.index]
    sell = [list(df_factor.loc[t1,:][df_factor.loc[t1,:]>indi_sell].index) for t1 in df_factor.index]
    df_hold = order_pro(buy, sell, df_price, alter_asset=ass_name)
    # df_hold = df_weight.to_clipboard()
    return df_hold
    
def rev(list_hold, df_price):   
    
    t1 = df_price.pct_change()
    t1[ass_name] = df_asset.pct_change()[ass_name]
    t1 = t1.shift(-1)  #这是暂时性的，最后日期要移动，在第一个单位补1
    rev = []
    for i in range(df_price.shape[0]-1):
        se_1 = np.array(t1.loc[df_price.index[i],list(list_hold[i].keys())])  #涨跌幅
        se_2 = np.array(list(list_hold[i].values()))
        rev.append(sum(se_1*se_2))
    df_rev = pd.DataFrame([0] + rev ,index=t1.index,columns=['rev'])
    return df_rev

def trade_analysis(df):
    '''输入: df_trade
    统计收益率, 年化收益, 年化波动, 年化夏普, 最大回撤;
    交易次数，胜率, 盈亏比'''
    pnlseries = df['rev']
    nav = (pnlseries + 1).cumprod()
    ret0 = (pnlseries + 1).groupby(pnlseries.index.year).prod() - 1
    dd = (nav / nav.groupby(nav.index.year).cummax() - 1).groupby(nav.index.year).min()
    alldd = (nav / nav.cummax() - 1).min()
    allret0 = (1 + pnlseries).prod() - 1

    rets = pnlseries.groupby(pnlseries.index.year).mean() * 245
    std = pnlseries.groupby(pnlseries.index.year).std() * (245 ** 0.5)
    sharpe = rets / std
    allret = pnlseries.mean() * 245
    allstd = pnlseries.std() * (245 ** 0.5)
    if allstd!=0:
        allsharpe = allret / allstd
    else:
        allsharpe = 0
        
    allsummary = pd.DataFrame([allret0, allret, allstd, allsharpe, alldd]).T
    allsummary.index = ['all']
    allsummary.columns = ['收益率', '年化收益', '年化波动', '年化夏普', '最大回撤']
    df = pd.concat([ret0, rets, std, sharpe, dd], axis=1)  # 先横着合并
    df.columns = ['收益率', '年化收益', '年化波动', '年化夏普', '最大回撤']
    df = pd.concat([df, allsummary], axis=0)  # 再竖着合并
#    torseries = df['trade']
#    tt = abs(torseries).groupby(torseries.index.year).sum()/2
#    alltt = sum(tt)
#
#    allsummary = pd.DataFrame([allret0, allret, allstd, allsharpe, alldd, alltt]).T
#    allsummary.index = ['all']
#    allsummary.columns = ['收益率', '年化收益', '年化波动', '年化夏普', '最大回撤', '交易次数']
#    df = pd.concat([ret0, rets, std, sharpe, dd, tt], axis=1)  # 先横着合并
#    df.columns = ['收益率', '年化收益', '年化波动', '年化夏普', '最大回撤', '交易次数']
#    df = pd.concat([df, allsummary], axis=0)  # 再竖着合并
    return df

def output(df_rev, df_asset):
    nav = (df_rev['rev'] + 1).cumprod()
    list_pick = ['CBA00301', '000001.SH']
    t1 = df_asset[list_pick]
    t1 = t1/t1.iloc[0,:]
    t1['策略'] = nav
    t1.to_clipboard()
    return t1

    
df_price = pd.read_excel('D:\东证资管\可转债\转债策略\data\转债价格.xlsx')
df_price.set_index('date',inplace=True)
df_asset = pd.read_excel('D:\东证资管\可转债\转债策略\data\资产价格.xlsx')
df_asset.set_index('date',inplace=True)
ass_name = 'H11025'

df_ana = trade_analysis(df_rev)
df_ana.to_clipboard()

df_weight['time'] = df_weight.index
pd.DataFrame(df_weight.groupby(by='time').size()).to_clipboard()





