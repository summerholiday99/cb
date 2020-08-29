# -*- coding: utf-8 -*-
"""
现金流判定可以用这边的，
然后解个方程
"""
import pandas as pd 
import numpy as np
import math
import scipy.optimize as so 
import datetime as dt
 
#%% 算历史YTM

df_price = pd.read_excel('D:\东证资管\可转债\量化定价\data\history\价格.xlsx')
df_price.set_index('date',inplace=True)
df_info = pd.read_excel('D:\东证资管\可转债\量化定价\data\history\转债数据FULL.xlsx', sheet_name='条款')
df_info = df_info[df_info['转债上市日期'].isna()==False]
df_info['转债上市日期'] = df_info['转债上市日期'].apply(lambda x: dt.datetime.strptime(x,'%Y-%m-%d'))

def compute_ytm(ttm,cashflow,price):
    '''未来现金流，相应折算时间'''
    time = [ttm - i for i in range(math.ceil(ttm))][::-1]
    cash = cashflow[-math.ceil(ttm):] # 对了
    def f(x):
        eq = '+'.join([str(m*x[0]**(-n)) for m,n in zip(cash,time)])+'-'+str(price)
        return np.array(eval(eq))
    try:
        init_guess = np.array([1.00])   # 初始猜一个,离正确值太远会没有解
        fsolve = so.fsolve(f, init_guess)
    except:
        init_guess = np.array([(cash[0]/price)**(1/ttm)])   # 聪明！早就该这么估
        fsolve = so.fsolve(f, init_guess, maxfev=1000) 
    ytm = (max(fsolve) - 1)*100
    return ytm

con = []
for cb in [df_info[df_info['转债代码']==j] for j in df_info['转债代码'].values]: # 循环啦
    # cb = df_info[df_info['转债代码']=='110016.SH']
    code = cb['转债代码'].values[0]
    t_duration = cb['发行期限'].values[0]
    t_start  = cb['转债上市日期'].values[0]
    t_arr = np.array(df_price[cb['转债代码']].dropna().index)
    t_arr = t_arr[t_arr>t_start]
    cashflow = [float(i) for i in cb['利率条款_结构化'].values[0].split(',')]
    cashflow[-1] = cb['利率补偿_结构化_全包含'].values[0]
    print(code)    
    for t_node in t_arr:   # 日期
        # t_node = t_arr[0]
        ttm =  t_duration - ((t_node-t_start)/np.timedelta64(1, 'D'))/365
        price = df_price.loc[t_node, code]
        ytm = compute_ytm(ttm, cashflow, price)
        con.append([code, t_node, price, ytm])
 
df_ytm = pd.DataFrame(con, columns = ['代码','时间','转债价格','YTM'])
df_ytm.to_clipboard()
group0 = df_ytm.groupby(by = '时间')
group0['YTM'].mean().to_clipboard()
group0.size().to_clipboard()
group0.size().plot()                                                           # 太少
df_price.count(axis=1).to_clipboard()


#%% 适用《现有转债》表 适合一个一个来

df_info = pd.read_excel('D:\东证资管\可转债\量化定价\data\转债数据.xlsx', sheet_name='条款')
df_info.replace(np.nan, 0, inplace=True)                            # 这个放前面防退市的干扰
df_info = df_info[(df_info['结算价']!=100)&(df_info['结算价']!=0)]   # 还没上市的

class bond():
    
    def __init__(self, cb):
        self.ttm = cb['转债剩余期限'].values[0]
        self.price = cb['结算价'].values[0]
        self.cashflow = [float(i) for i in cb['利率条款_结构化'].values[0].split(',')]
        if cb['是否有利率补偿'].values[0]=='是':
            self.cashflow[-1] = cb['利率补偿_全包含_结构化'].values[0]
        elif cb['是否有利率补偿'].values[0]=='否':
            self.cashflow[-1] += 100

def ytm(t1):
    '''未来现金流，相应折算时间'''
    time = [t1.ttm - i for i in range(math.ceil(t1.ttm))][::-1]
    cash = t1.cashflow[-math.ceil(t1.ttm):] # 对了
    def f(x):
        eq = '+'.join([str(m*x[0]**(-n)) for m,n in zip(cash,time)])+'-'+str(t1.price)
        return np.array(eval(eq))
    init_guess = np.array([1.03])   # 初始猜一个
    # root = so.root(f, init_guess) # 
    fsolve = so.fsolve(f, init_guess)
    ytm = (max(fsolve) - 1)*100
    return ytm


para = [df_info[df_info['转债简称']==j] for j in df_info['转债简称'].values]
con = []
for cb in para:
    t1 = bond(cb)
    con.append(ytm(t1))
df_info['ytm'] = con
df_info = df_info[['转债代码','结算价','ytm']] 
df_info.to_clipboard(index=False)

df_info['ytm'].mean()
