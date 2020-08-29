# -*- coding: utf-8 -*-
import datetime as dt
import numpy as np
import pandas as pd
import warnings
warnings.filterwarnings("ignore")
import matplotlib.pyplot as plt
import seaborn as sns

def compare_week(df_new, df_old, list_date):
    '''读取周报两期对比'''
    t1 = df_new[['转债代码', '转债简称', '信用等级','定价偏离(%)']]
    t1.rename(columns={'定价偏离(%)':list_date[0]}, inplace=True) 
    t2 = df_old[['转债代码','定价偏离(%)']]
    t2.rename(columns={'定价偏离(%)':list_date[1]}, inplace=True) 
    df_aim = pd.merge(t1,t2,on='转债代码',how='right')   #生成对比表 
    df_aim.dropna(inplace=True)
    df_aim['变化'] = df_aim[list_date[0]] - df_aim[list_date[1]]
    df_aim['是否转债'] = df_aim['转债简称'].apply(lambda x:'EB' if 'EB' in x else '转债')
    return df_aim

def change_test():
    return

def stat_week(df, list_date):
    '''几个统计，做成一张表'''
    def f(s,ll):
        gg = s.groupby(by=['信用等级'])
        temp = gg[ll].mean().T    
        temp = temp[['A+','AA-','AA','AA+','AAA']]
        return temp
    
    s1 = df[df['是否转债']=='转债']
    t1 = f(s1,list_date+['变化'])
    s2 = df[(df['是否转债']=='转债')&(df[list_date[0]].abs()<10)]
    t2 = f(s2,list_date)
    return {'去EB':t1,'去EB+10%以内':t2}

def tool_output(dict_di,str_road2):
    writer = pd.ExcelWriter(str_road2)
    for ii in dict_di.keys():
        dict_di[ii].to_excel(writer, sheet_name=ii)
    writer.save()
    return

def tool_plot(df,str_road2):
    sns.set_style('whitegrid')
    plt.rcParams["font.sans-serif"] = ["SimHei"]
    plt.rcParams["axes.unicode_minus"] = False

    x = np.arange(5)
    y = df.loc[list_date[0],:].tolist()
    y1 = df.loc['变化',:].tolist()
    bar_width = 0.35
    tick_label = df.columns.tolist()
    
    plt.bar(x, y, bar_width, align="center", label=list_date[0], alpha=0.8)
    plt.bar(x+bar_width, y1, bar_width, align="center", label='变化', alpha=0.8)
    plt.xticks(x+bar_width/2, tick_label)
    plt.legend()
    plt.show()
    plt.savefig(str_road2)
    return

if __name__ == '__main__':
    list_date = ['2020-08-28','2020-08-21']  #第一个新的，第二个旧的
    str_raw = r'D:\同步\我的坚果云\可转债'  # 在家/在公司换一下
    str_road = r'{}\市场观察&周报\周报偏离\{}-{}.xlsx'.format(str_raw,list_date[0],list_date[1])
    str_road2 = r'{}\市场观察&周报\周报偏离\{}-{}.png'.format(str_raw,list_date[0],list_date[1])
    str_di = {'最新':'{}\量化定价\output\{}.xlsx'.format(str_raw,list_date[0]),
              '前一周':'{}\量化定价\output\{}.xlsx'.format(str_raw,list_date[1])}
    df_new = pd.read_excel(str_di['最新'])
    df_old = pd.read_excel(str_di['前一周'])

    df_aim = compare_week(df_new, df_old, list_date)
    dict_ = stat_week(df_aim, list_date)
    dict_.update({'周对比':df_aim})
    tool_output(dict_, str_road)
    tool_plot(dict_['去EB'],str_road2)
