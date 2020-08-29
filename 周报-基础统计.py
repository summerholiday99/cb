# -*- coding: utf-8 -*-
"""
实现功能
1.基本统计excel
2.换手率统计、高关注度转债统计excel
3.平价区间溢价率graph
TODO: 价格区间~90以下的, 有可能没转债,程序上调一下
"""

import datetime as dt
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns  
import warnings
warnings.filterwarnings("ignore")
import WindPy as wind
wind.w.start()

def tool_download_cb1(df, str_date):
    '''单日下载'''
    # str_date = '20181231'
    print('download'+str_date)
    ns_date = dt.datetime.strptime(str_date,'%Y%m%d')
    list_col = list(set(df['转债代码'].values.tolist()))  # 下面下载怕有重复代码    
    str_fields = "sec_name,creditrating,amt,close,pct_chg,ytm_cb,"+\
                 "convvalue,convpremiumratio,ipo_date,outstandingbalance"
    temp = wind.w.wss(list_col, str_fields, tradeDate=str_date, cycle="D", priceAdj="U")
    df_dd = pd.DataFrame(temp.Data,index=temp.Fields).T
    df_dd['code'] = temp.Codes
    df_dd['date'] = ns_date 
    df_dd = df_dd[(df_dd['IPO_DATE']<ns_date)&(df_dd['OUTSTANDINGBALANCE']>0)] # 清除部分转债
    return df_dd

def tool_download_cb(df, list_date):
    '''单日下载整合'''
    con = [tool_download_cb1(df, str_date) for str_date in list_date]
    df_aim = pd.concat(con)
    return df_aim

def tool_f(x,list_r):
    if x<list_r[0]:
        return '~{}'.format(str(int(list_r[0])))
    elif x>=list_r[0] and x<list_r[-1]:
        for ii in range(0,len(list_r)-1):
            if x>=list_r[ii] and x<list_r[ii+1]:
                return '{}~{}'.format(str(int(list_r[ii])),str(int(list_r[ii+1])))
            else:
                pass
    elif x>=list_r[-1]:
        return '{}~'.format(str(int(list_r[-1])))

def tool_process_data(df):
    '''数值修正;辅助列'''
    df['date'] = df['date'].apply(lambda x:x.date())
    df['CONVPREMIUMRATIO'] = df['CONVPREMIUMRATIO'].astype('float')
    df['OUTSTANDINGBALANCE'] = df['OUTSTANDINGBALANCE'].astype('float')
    df['num'] = 1   #到时候数个数用
    df['AMT'] = df['AMT']/100000000
    list_r = [90,100,105,110,115,120,125,130]
    df['price_p'] = df['CLOSE'].apply(lambda x:tool_f(x,list_r))
    list_r2 = [60,70,80,90,100,110,120,130]
    df['convv_p'] = df['CONVVALUE'].apply(lambda x:tool_f(x,list_r2))
    return df

def tool_df2group(df_r, di):
    '''选定一列，按其值展开成一个矩阵'''
    # 前提条件，index columns 定了，不能有重复项
#    str_index_name = '重仓股简称'
#    str_column_name = '代码'
#    str_data = '2020Q1'
    str_index_name = 'date'
    str_column_name = di['x']
    str_data = di['y']   
    if max(df_r.groupby(by=[str_index_name,str_column_name]).size())==1:
        df_base = pd.Series(df_r[str_data].values, index=[df_r[str_index_name], df_r[str_column_name]]) #多级索引
        df_aim = df_base.unstack() # 索引展开，你这编程，有待长进啊!! 太菜了，多看看别人怎么写
        return df_aim
    else:
        print('有重复项！')


def tool_df2group_general(df_r, str_index_name, str_column_name, str_data):
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
        print('有重复项!')

def tool_process_col(s):
    if s.columns.name=='CREDITRATING':
        return s[['A+', 'AA-','AA', 'AA+',  'AAA']]
    elif s.columns.name=='convv_p':
        return s[['~60', '60~70', '70~80', '80~90','90~100',
                  '100~110', '110~120', '120~130', '130~']]
    elif s.columns.name=='price_p':
        return s[['~90','90~100','100~105', '105~110', '110~115', 
                  '115~120', '120~125', '125~130', '130~']]  #这边不好写 先这样

def tool_output(dict_di,str_road2):
    writer = pd.ExcelWriter(str_road2)
    for ii in dict_di.keys():
        dict_di[ii].to_excel(writer, sheet_name=ii)
    writer.save()
    return

def tool_linesplot_conv(df, list_dd):    
    '''输入: 选定的几天'''
    # 再阴影
    # df = dict_di['转股溢价率_平价价格区间']
    f_date = lambda x: dt.datetime.strptime(x,'%Y%m%d')
    list_d = [f_date(ii).date() for ii in list_dd]
    df = df[df.index.isin(list_d)]    #偏债的就不看了，这样图的关键部分点
    df.drop(columns={df.columns[0]}, inplace=True)  
    temp = df.T
    sns.set(style="white", palette="muted", color_codes=True)
    plt.figure()
    plt.rcParams['font.sans-serif'] = ['SimHei']
    plt.rcParams['axes.unicode_minus'] = False
    plt.xticks(np.arange(len(temp.index)), temp.index, rotation=30)
    for ii in range(temp.shape[1]):    
        plt.plot(temp.iloc[:,ii].tolist(), marker='o', label=str(temp.columns[ii]))              
    plt.xlabel('转换价值')
    plt.ylabel('转股溢价率')
    plt.legend()
    plt.grid()
    plt.tight_layout()
    plt.show()
    return

def tool_violinplot_conv(df_con, dd):
    '''输入原始数据：
    y转股溢价率 x转股价值
    '''
    df = df_con[df_con['date']==dt.datetime.strptime(dd,'%Y%m%d').date()]
    df.rename(columns={'convv_p':'转股价值','CONVPREMIUMRATIO':'转股溢价率'},
              inplace=True)
    df['转股溢价率'] = df['转股溢价率'].astype('float')
    ll = ['60~70', '70~80', '80~90','90~100','100~110', '110~120', '120~130', '130~']
    df = df[(df['转股溢价率']<200)&(df['转股价值'].isin(ll))]
        
    plt.style.use('seaborn')
    plt.rcParams['font.sans-serif'] = ['SimHei']
    plt.rcParams['axes.unicode_minus'] = False
    plt.xticks(list(range(len(ll))),ll)
    sns.violinplot(x='转股价值', y='转股溢价率', data=df, 
                   order=ll, inner='point', cut=0.1,
                   showmeans=False, showmedians=True, showextrema=False,
                   linewidth = 1, width = 1.5,  color ='skyblue')
    plt.tight_layout() 
    return 

def stat_y_x(df,di):
    # df = df_con
    # di = {'y':'num','x':'CREDITRATING','m':'sum'}
    df = df.copy()
    gg = df.groupby(by=['date', di['x']], as_index=False)
    s1 = eval("gg[di['y']].{}()".format(di['m']))    
    s2 = tool_df2group(s1, di)
    s3 = tool_process_col(s2)
    return s3

def stat_con(df):
    df = df_con
    # 数量: 信用等级/价格区间
    s1_di_num_credit = stat_y_x(df, {'y':'num','x':'CREDITRATING','m':'sum'})
    s1_di_num_price_p = stat_y_x(df, {'y':'num','x':'price_p','m':'sum'})
    # 成交量: 信用等级/价格价格区间
    s1_di_amt_credit = stat_y_x(df, {'y':'AMT','x':'CREDITRATING','m':'sum'})
    s1_di_amt_price_p = stat_y_x(df, {'y':'AMT','x':'price_p','m':'sum'})
    # 转股溢价率: 信用等级/平价区间/价格区间   
    s1_di_stopre_credit = stat_y_x(df, {'y':'CONVPREMIUMRATIO','x':'CREDITRATING','m':'median'})
    s1_di_stopre_convv_p = stat_y_x(df, {'y':'CONVPREMIUMRATIO','x':'convv_p','m':'median'})
    
    dict_di={'数量_信用等级':s1_di_num_credit,
             '数量_价格区间':s1_di_num_price_p,
             '成交额_信用等级':s1_di_amt_credit,
             '成交额_价格区间':s1_di_amt_price_p,
             '转股溢价率_信用等级':s1_di_stopre_credit,
             '转股溢价率_平价价格区间':s1_di_stopre_convv_p}
    return dict_di

def stat_turn(df_con):
    '''
    这里计算周换手率
    周成交额加总/'''
    gg = df_con.groupby(by=['code'])
    k1 = gg['AMT'].sum()
    k2 = gg['OUTSTANDINGBALANCE'].mean()
    df_aim = pd.concat([k1,k2], axis=1)
    df_aim['turn'] = df_aim['AMT']/df_aim['OUTSTANDINGBALANCE']
    return df_aim

def stat_highturn(df_con, df_turn):
    '''取最近1天的价格区间'''
    df_ = df_con[df_con['date']==df_con['date'].unique()[-1]]
    df_.set_index('code', inplace=True)
    df_turn['price_p'] = df_['price_p']
    df_turn['转债简称'] = df_['SEC_NAME']
    
    # 把转债选出来
    gg = df_turn.groupby(by=['price_p'])
    k = gg[['AMT','turn']].rank(pct=True)
    df_turn[['AMT_t','turn_t']] = k
    df_turn['score'] = df_turn['AMT_t']*0.8 + df_turn['turn_t']*0.2
    gg = df_turn.groupby(by=['price_p'])
    df_turn['score_r'] = gg['score'].rank(ascending=False, method='first')
    df_turn = df_turn[df_turn['score_r']<=5]
    df_turn['code'] = df_turn.index

    # 输出    
    temp = tool_df2group_general(df_turn, 'score_r', 'price_p','转债简称')
    temp= temp[['~90','90~100','100~105', '105~110', '110~115',
                '115~120', '120~125', '125~130', '130~']]
    return temp

def tool_turn_appendix(dict_di,df_con):
    df_turn = stat_turn(df_con)       #成交量方面增加的统计
    df_highturn_cb = stat_highturn(df_con, df_turn)
    
    df_turn_o = df_turn[['AMT','转债简称','OUTSTANDINGBALANCE','turn']]
    df_turn_o.sort_values(by=['OUTSTANDINGBALANCE'],inplace=True)
    dict_di.update({'df_turn':df_turn_o,
                    'df_highturn_cb':df_highturn_cb})
    return dict_di

#%% 数据统计
#str_road = r'D:\Tools\转债\转债代码-更新.xlsx' #自己更新下
#str_road2 = 'D:\东证资管\可转债\市场观察&周报\周度统计\{}.xlsx'.format(d_date) 
    
str_raw = r'C:\Users\huangjili\Desktop\同步\OneDrive\可转债'  # 在家在公司换一下
list_date = "20190802,20200710,20200821".split(',')   # 必须是交易日  #2019之前没数据
str_road = r'{}\量化定价\转债数据.xlsx'.format(str_raw)   
str_road2 = r'{}\市场观察&周报\周度统计\统计&跟踪\周度基础统计_{}.xlsx'.format(str_raw,list_date[-1]) 

df = pd.read_excel(str_road,sheet_name='条款')        # 提供转债名单
df_con = tool_download_cb(df, list_date)              # 原始数据 df_con.to_clipboard()
df_con = tool_process_data(df_con)
dict_di = stat_con(df_con)
dict_di = tool_turn_appendix(dict_di,df_con)
tool_output(dict_di, str_road2)                       # excel文件输出

#%% 作图
tool_linesplot_conv(dict_di['转股溢价率_平价价格区间'], list_date)
#tool_violinplot_conv(df, dd='20200605') 