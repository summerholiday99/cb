# -*- coding: utf-8 -*-
"""
2020/8/7：维护自动化(主要不出错,省心)
2020/8/18：粘表自动化(省时省力)
"""
import datetime as dt
import numpy as np
import pandas as pd
import warnings
warnings.filterwarnings("ignore")
import xlwt

def tool_his_grade(dfs):
    '''从老文件，先做一个历史评分记录文件'''
    def f(a,name):
        con = []
        for ii in a.columns:    
            temp = a[[ii]]
            temp.dropna(inplace=True)
            temp['时间'] = ii
            temp.rename(columns={ii:'评分'}, inplace=True) 
            temp.reset_index(inplace=True)
            con.append(temp)
        df_con = pd.concat(con)
        df_con['类型'] = name
        return df_con
    list_a = ['偏股历史打分', '偏债历史打分', '平衡历史打分', '其他历史打分']
    df_aim = pd.concat([f(dfs[kk],kk) for kk in list_a])
    return df_aim
    
def tool_output(dict_di,str_road2):
    writer = pd.ExcelWriter(str_road2)
    for ii in dict_di.keys():
        dict_di[ii].to_excel(writer, sheet_name=ii, index=False)
    writer.save()
    return

def tool_output_2(tt,rr):
    '''这个专门为输出粘贴表写的,实现五个style:
    整体列宽行宽
    首行自动换行，加粗
    下面四类四种颜色
    '''
    font_b = xlwt.Font() # 为样式创建字体
    font_b.name = '微软雅黑' 
    font_b.bold = True  # 加粗
    font_b.height = 20*10
    
    font = xlwt.Font() # 为样式创建字体
    font.name = '微软雅黑' 
    font.height = 20*10
    
    border = xlwt.Borders()        
    border.top=xlwt.Borders.THIN   # THIN的意思是细边框
    border.bottom=xlwt.Borders.THIN
    border.left=xlwt.Borders.THIN
    border.right=xlwt.Borders.THIN
    border.left_colour = 0x40
    border.right_colour = 0x40
    border.top_colour = 0x40
    border.bottom_colour = 0x40
    
    alignment=xlwt.Alignment()  #初始化一个对齐方式
    alignment.horz = 0x02       # 设置水平居中
    alignment.vert = 0x01       # 设置垂直居中
    alignment.wrap = 1
    
    def f(ii):
        pat = xlwt.Pattern()
        pat.pattern = xlwt.Pattern.SOLID_PATTERN  # 设置背景颜色
        pat.pattern_fore_colour = xlwt.Style.colour_map[ii]
        return pat
   
    def f_main(f,b,a,p):
        style = xlwt.XFStyle() # 初始化样式
        style.font = f # 设定样式
        style.borders = b
        style.alignment = a
        if p:
            style.pattern = p
        return style
    
    # 不同的style，四个不同的区间, 不同颜色
    list_color = ['sea_green','sky_blue','ivory','silver_ega']
    dict_color = dict(zip(list_color,[f(name) for name in list_color]))
    style_line = f_main(font_b,border,alignment,None)
    list_style = [f_main(font,border,alignment,dict_color[c_]) for c_ in list_color]
    list_type = tt['类型'].unique()
    list_p = [tt[tt['类型']==k_].index[0] for k_ in list_type]   # 第一次出现的几个index
    list_p = list_p+[tt.shape[0]]                                
    list_p2 = [range(list_p[ii-1]+1,list_p[ii]+1) for ii in range(1,len(list_p))] 
    list_block = list(zip(list_p2,list_style))                                   
                             
    #按理说都开始画表了，可以写成四张表，但这里先简化一下了
    workbook = xlwt.Workbook(encoding = 'ascii')
    worksheet = workbook.add_sheet('本周表格')    
    for w_ in range(tt.shape[0]):      #设置下宽度，基本OK了
        col_ = worksheet.col(w_)       
        col_.width = 256*13   

    for j_ in range(tt.shape[1]):   
        worksheet.write(0, j_, tt.columns[j_], style_line) 
    for bb in list_block:
        for i_ in bb[0]:  #迭代器
            for j_ in range(tt.shape[1]):
                worksheet.write(i_, j_,tt.iloc[i_-1,j_], bb[1])  # bb[1] style list
    workbook.save(rr)                
    return

def cb_update_grade(rr,rr2):
    '''每次跑之前就更新\检查一下
    老的版本是在那四张表上读的，后面直接从新的一张表来做
    '''
    df_his = pd.read_excel(rr)
    df_g =  pd.read_excel(rr2, sheet_name=None)
    t1 = max(df_his['时间']).date()
    t2 = df_g['时间信息']['评分时间'][0].date()
    t3 = df_g['时间信息']['评分已完成？'][0]
    print('历史打分文件最新数据:{}'.format(str(t1)))
    print('打分文件最新打分时间:{},状态已完成?:{}'.format(str(t2),t3))
    str_choice = input('是否更新历史打分数据(y/n)?')

    if str_choice=='y':
        df_m = df_g['打分表'][['转债代码','打分','类型']]
        df_m['时间'] = pd.Timestamp(t2)
        df_m = df_m[['转债代码','打分','时间','类型']]
        df_his = pd.concat([df_m,df_his])
        df_his.to_excel(rr)
        print('打分数据已更新到最新')
        return df_his
    elif str_choice=='n':
        return df_his
    else:
        print('输入错误')

def cb_class_sheet(df_dev, df_his, df_s, rr):
    '''
    读一定是从可以手动刷新的大表读，所以这里生成是打分文件
    输入:现有转债大表（数据已刷新）取相应字段，带估值偏离，带历史评分
    输出:新的偏股偏债四个类型表(带历史评分/估值偏离数据，带估值偏离)以供评分
    '''
    # df_dev '估值偏离'  df_his '历史评分' 加到大表上去
    for ii in ['偏股','偏债','平衡','其他']:
        df_s[ii] = df_s[ii].apply(lambda x: ii if x==1 else '')
    df_s['类型'] = df_s['偏股']+df_s['偏债']+df_s['平衡']+df_s['其他']
    df_s['类型'] = df_s['类型'].apply(lambda x:x[:2])       # 有一些新债会有两个类型
    df_his = df_his[df_his['时间'] == max(df_his['时间'])]  # 这里就取最新评分
    temp = pd.merge(df_s, df_dev[['转债代码','定价偏离(%)']], on='转债代码', how='left')
    temp = pd.merge(temp, df_his[['转债代码','打分']], on='转债代码', how='left')
    
    # 取数据，格式控制及输出
    list_di = ['转债代码', '转债名称','本周收盘价', '周涨跌幅(%)', '正股周涨跌幅(%)',
               '本周转股溢价率(%)','转股溢价率周变动(%)', '本周纯债到期收益率(%)', 
               '本周纯债溢价率(%)', '定价偏离(%)', '打分','类型','信用等级','债券余额(亿元)']
    df_table = temp[list_di]
    df_table = df_table.round(2)

    di_kk = {'偏股':1,'偏债':2,'平衡':3,'其他':4}
    df_table['类型2'] = df_table['类型'].apply(lambda x:di_kk[x])
    df_table.sort_values(by=['类型2'], inplace=True)
    df_table.drop(columns=['类型2'], inplace=True)
    
    df_dev['本周老破小'] = None
    df_time = pd.DataFrame({'评分时间':[dt.datetime.today().date()], '评分已完成？':['否']}) 
    
    dict_di = {'时间信息':df_time,'打分表':df_table,'老破小': df_dev}
    tool_output(dict_di,rr)
    return 

def cb_recommend_sheet(df_his, path):
    '''
    输入：打分文件
    输出：四张表供粘贴&评分有变动的标注出来（这个暂时没做）
    '''
    df_g = pd.read_excel(path.str_grade, sheet_name=None)
    print('评分已完成？{}'.format(df_g['时间信息']['评分已完成？'][0]))
    
    def f(tt,ii):
        tt['粘贴类型'] = ii
        return tt
    def f2(tt):
        dict_c = {'偏股':1,'偏债':2,'平衡':3, '其他':4}
        tt['类型2'] = tt['类型'].apply(lambda x: dict_c[x])
        tt.sort_values(by=['类型2'], inplace=True)
        tt.drop(columns=['类型2'], inplace=True)
        return tt
    
    if df_g['时间信息']['评分已完成？'][0]=='是':
        temp = df_g['打分表']
        list_lpx = df_g['老破小'].loc[df_g['老破小']['本周老破小']==1,'转债代码'].tolist()
        list_aim = ['转债代码', '转债名称', '本周收盘价', '周涨跌幅(%)', '正股周涨跌幅(%)',
                    '本周转股溢价率(%)', '转股溢价率周变动(%)', '本周纯债到期收益率(%)', 
                    '本周纯债溢价率(%)', '定价偏离(%)', '类型']
        df_1 = temp.loc[temp['打分']==1, list_aim]
        df_2 = temp.loc[temp['打分']==2, list_aim]
        df_25 = temp.loc[temp['打分']==2.5, list_aim]        
        df_lpx = temp.loc[temp['转债代码'].isin(list_lpx), list_aim]        
        
        dict_di = {'买入':df_1, '可持有':df_2, '正股佳':df_25, '老破小':df_lpx}
        df_aim = pd.concat([f(dict_di[ii],ii) for ii in dict_di.keys()])
        df_aim = f2(df_aim)
        df_aim.index = range(df_aim.shape[0])
        tool_output_2(df_aim, path.str_list)
                
    elif df_g['时间信息']['评分已完成？'][0]=='否':
        print('打分去')

#%% Main    
class road():
    str_raw = r'C:\Users\huangjili\Desktop\同步\OneDrive\可转债'
    str_road = r'{}\市场观察&周报\周报打分\{}.xlsx'
    str_road2 = r'{}\量化定价\{}.xlsx'
    str_road3 = r'{}\市场观察&周报\周报打分\历史表格\{}.xls' #用的xls
    
    str_f1 = '转债分类维护'    
    str_f2 = '转债历史评分' 
    str_f3 = '2020-08-28'  # 最新的估值偏离
    str_f4 = '打分文件'
    str_f5 = '粘表文件-'+str_f3
    
    str_class = str_road.format(str_raw, str_f1)  # 手动大表
    str_his = str_road.format(str_raw, str_f2)    # 转债历史评分
    str_dev = str_road2.format(str_raw, str_f3)   # 估值偏离
    str_grade = str_road.format(str_raw, str_f4)  # 打分文件
    str_list = str_road3.format(str_raw, str_f5)  # 名单文件

path = road()
dfs = pd.read_excel(path.str_class, sheet_name=None)
df_dev = pd.read_excel(path.str_dev)      # 输出打分文件前先算估值
df_s = pd.read_excel(path.str_class, sheet_name='现有转债', skiprows=1)

# 0）地址/日期更新
# 1) update_grade()  最好每次跑完就再运行一次
df_his = cb_update_grade(path.str_his, path.str_grade)
# 2) 手动刷新那个大表（新券加上）
# 3) 输出偏股偏债四张表+估值偏离表以打分
cb_class_sheet(df_dev, df_his, df_s, path.str_grade)
# 4)评分结束后，生成买入/可持有/正股佳/老破小四张表
cb_recommend_sheet(df_his, path)
# 作图（这样更好, 但难度有点高）
