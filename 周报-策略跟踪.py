# -*- coding: utf-8 -*-
"""
# 先在数据库更新数据

实现功能
1 基础角度策略净值及excel
2 策略净值graph
4.债底线性graph(虽然不是策略)
df = df[df['交易日'] < pd.Timestamp(year=2020, month=6, day=26)]  #有时放假数据不对，部分剔除
"""

import datetime as dt
import numpy as np
import pandas as pd
import seaborn as sns
import warnings
import matplotlib.pyplot as plt
warnings.filterwarnings("ignore")

def tool_preproc(df):    
    f_date = lambda x:dt.datetime.strptime(str(x),'%Y%m%d')
    df['交易日'] = df['交易日'].apply(lambda x:f_date(x))
    tt = pd.Timestamp(year=2016, month=12, day=31)
    df = df[df['交易日']>tt]                     # 这个就砍掉一半
    df['涨跌幅'] = df['涨跌幅']/100
    #df['交易额(万)'] = df['交易额(万)']/10  #这个以后就不用了
    df = df[df['转债余额']>1]
    df = df[~df['信用等级'].isin(['A-','A'])]
    return df

def tool_df2group(df_r,str_index_name,str_column_name,str_data):
    '''选定一列，按其值展开成一个矩阵'''
    # 前提条件，index columns 定了，不能有重复项
#    str_index_name = '重仓股简称'
#    str_column_name = '代码'
#    str_data = '2020Q1'
    if max(df_r.groupby(by=[str_index_name,str_column_name]).size())==1:
        df_base = pd.Series(df_r[str_data].values, index=[df_r[str_index_name], df_r[str_column_name]]) #多级索引
        df_aim = df_base.unstack() # 索引展开，你这编程，有待长进啊!! 太菜了，多看看别人怎么写
        return df_aim
    else:
        print('有重复项！')

def tool_plot_netvalue(di_dict, str_road3):
    '''绘制多图，直接保存'''
    sns.set_style('whitegrid')
    plt.rcParams['font.sans-serif']=['SimHei']
    plt.rcParams['axes.unicode_minus'] = False
    
    for ii in list(di_dict.keys()):
        #ii='规模'
        ax = di_dict[ii]['净值'].plot(title=ii,figsize=(6,4))
        fig = ax.get_figure()
        fig.savefig(str_road3 +'\\'+'{}.png'.format(ii))
    return

def tool_plot_purebond(df, str_road3):
    '''债底'''    
    # 这个函数画分类图特别合适  lineplot
    sns.set_style('whitegrid')
    plt.rcParams['font.sans-serif'] = ['SimHei']
    plt.rcParams['axes.unicode_minus'] = False
    ax = sns.lineplot(x='交易日', y='纯债价值', data=df, ci='sd')
    fig = ax.get_figure()
    fig.savefig(str_road3+'//'+'纯债价值.png',figsize=(6,4) )
    return

def tool_output(di_dict,df_con,str_road):
    ''''''
    a = pd.ExcelWriter(str_road)
    for ii in list(di_dict.keys()):
        di_dict[ii]['净值'].to_excel(a, ii+'-净值')
    df_con.to_excel(a, '统计')
    a.save()
    return 

def stat_num(df):
    df['数量'] = 1
    gg = df.groupby(by=['交易日','价格区间'], as_index=False)    
    temp = gg[['数量']].sum()
    df_t = tool_df2group(temp, str_index_name='交易日',str_column_name='价格区间',str_data='数量')
    #df_t.plot()
    df_t.to_clipboard()
    return k

def stat_num_amount(df,field):
    '''统计数量、规模'''
    dd = df.copy()
    dd['数量'] = 1     # 多一列方便求和
    gg = dd.groupby(by=field)
    k = gg[['数量','转债余额']].sum()
    return k

def stat_thisweek(di_dict, str_date, num=7):
    '''
    已有净值，统计当周表现
    然后和数量表并起来
    '''    
    dt_period = [pd.Timestamp(str_date)-pd.Timedelta(days=num),
                 pd.Timestamp(str_date)]
    
    con = []
    for ii in list(di_dict.keys()):
        temp = di_dict[ii]['净值']
        rev = pd.DataFrame(temp.loc[dt_period[1],:]/temp.loc[dt_period[0],:], columns=['周涨跌幅'])-1
        stat2 = pd.concat([di_dict[ii]['统计'],rev], axis=1)
        con.append(stat2)
    df_con = pd.concat(con)
    return df_con

def stat_netvalue(dd2, field, shift):
    '''重要函数
    有了filed判断后, 累乘求净值'''
    df = dd2
    if shift==1: 
        df.sort_values(by=['转债代码','交易日'],inplace=True)
        gg = df.groupby(by=['转债代码'],as_index=False)    
        df[field] = gg[field].shift(1)       #这个一定要有 要不然就未来信息了
    else:
        pass
    gg = df.groupby(by=['交易日',field], as_index=False)    
    temp = gg[['涨跌幅']].mean()
    df_t = tool_df2group(temp, str_index_name='交易日',str_column_name=field,str_data='涨跌幅')
    df_t.fillna(0, inplace=True)
    df_t2 = (df_t+1).cumprod()
    return df_t2
   
def compute_credit(dd):
    df = dd.copy()
    df.sort_values(by=['交易日'],inplace=True)
    df_t2 = stat_netvalue(df, field='信用等级', shift=0)
    df_s = stat_num_amount(df[df['交易日']==df_t2.index[-1]],'信用等级')
    df_t2 = df_t2[['A+','AA-','AA','AA+','AAA']]
    df_s = df_s.T[['A+','AA-','AA','AA+','AAA']].T
    di = {'净值':df_t2,'统计':df_s}
    return di

def compute_scale(dd):
    df = dd.copy()
    df.sort_values(by=['交易日'],inplace=True)
    def f(x):
        if x>40:
            return '40~'
        elif x>15 and x<=40:
            return '15~40'
        elif x<=15 and x>5:
            return '5~15'
        elif x<=5:
            return '0~5'
    df['规模'] = df['发行额度'].apply(f)     
    df_t2 = stat_netvalue(df, field='规模', shift=0)
    df_s = stat_num_amount(df[df['交易日']==df_t2.index[-1]],'规模')
    df_t2 = df_t2[['0~5','5~15','15~40','40~']]
    df_s = df_s.T[['0~5','5~15','15~40','40~']].T
    di = {'净值':df_t2, '统计':df_s}
    return di

def compute_type(dd):
    '''这个比较麻烦点'''
    df = dd.copy()
    df.sort_values(by=['交易日'],inplace=True)
    def f(x):
        # 纯债到期收益率 收盘价 转股溢价率
        kk = [float(ii) for ii in x.split('~')]
        tt = '-'
        if kk[1]>115 or kk[2]<10:
            tt = '偏股型'
        elif kk[0]>3:
            tt = '偏债型'
        elif kk[2]>=10 and kk[2]<35:
            tt = '平衡型'
        if tt =='-':
            tt = '其他型'
        return tt    
    
    df['类型'] = df['纯债到期收益率'].astype('str')\
                +"~"+df['收盘价'].astype('str')\
                +"~"+df['转股溢价率'].astype('str')
    df['类型'] = df['类型'].apply(f) 

    df_t2 = stat_netvalue(df, field='类型', shift=1)
    df_s = stat_num_amount(df[df['交易日']==df_t2.index[-1]],'类型')
    df_t2 = df_t2[['偏股型','偏债型','平衡型','其他型']]
    df_s = df_s.T[['偏股型','偏债型','平衡型','其他型']].T
    di = {'净值':df_t2,'统计':df_s}
    return di

def compute_price(dd):
    '''统计不同价格区间的转债表现，确实有个券数量的问题
    18年末的时候120以上的转债基本没有，20年4月100元以下的转债也基本没有
    你不方便去管制  可以试着数量 那部分的就不投了？那最近的你怎么跟踪呢？ 
    '''
    df = dd.copy()
    df.sort_values(by=['交易日'],inplace=True)
#    def f2(x):
#        if x<70:
#            return '0~70'
#        elif x>=70 and x<90:
#            return '70~90'
#        elif x>=90 and x<110:
#            return '90~110'
#        elif x>=110 and x<130:
#            return '110~130'
#        elif x>=130:
#            return '130~'
    #  df['价格区间'] = df['转股价值'].apply(f2) 
    
    def f(x):
        if x<100:
            return '0~100'
        elif x>=100 and x<110:
            return '100~110'
        elif x>=110 and x<120:
            return '110~120'
        elif x>=120 and x<130:
            return '120~130'
        elif x>=130:
            return '130~'
  
    df['价格区间'] = df['收盘价'].apply(f) 
    df_t2 = stat_netvalue(df, field='价格区间', shift=1)
    df_s = stat_num_amount(df[df['交易日']==df_t2.index[-1]],'价格区间')
    df_t2 = df_t2[['0~100','100~110','110~120','120~130','130~']]
    df_s = df_s.T[['0~100','100~110','110~120','120~130','130~']].T
    di = {'净值':df_t2, '统计':df_s}
    return di


#%% Main
str_date = '2020-08-21'
str_raw = r'C:\Users\huangjili\Desktop\同步\OneDrive\可转债'  # 在家在公司换一下
str_road = r'{}\cb_panel_{}.xlsx'.format(str_raw, str_date)
str_road2 = r'{}\市场观察&周报\周度统计\统计&跟踪\周报策略跟踪_{}.xlsx'.format(str_raw,str_date)
str_road3 = r'{}\市场观察&周报\周度统计\graph'.format(str_raw)

# str_road = r'C:\Users\jili\Desktop\周报20200718\cb_panel_{}.xlsx'.format(str_date)
# str_road2 = r'C:\Users\jili\Desktop\周报20200718\周度策略跟踪_{}.xlsx'.format(str_date)
# str_road3 = r'C:\Users\jili\Desktop\周报20200718\gragh'

df = pd.read_excel(str_road)
df = tool_preproc(df)
di_credit = compute_credit(dd=df)
di_scale = compute_scale(dd=df)    
di_type = compute_type(dd=df)
di_price = compute_price(dd=df)

di_dict = {'信用等级':di_credit, 
           '规模':di_scale, 
           '类型':di_type,
           '价格区间':di_price}
df_con = stat_thisweek(di_dict, str_date, num=7)

tool_output(di_dict, df_con, str_road2)   
tool_plot_netvalue(di_dict, str_road3)    # 作图，合适长宽，直接保存
tool_plot_purebond(df, str_road3)         # 作图，纯债价值
