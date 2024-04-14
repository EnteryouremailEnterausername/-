
#%%
import os
import pandas as pd 
import numpy as np
from math import floor
import docx
from docx.shared import Cm, RGBColor, Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re
from WindPy import w
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.styles import numbers
from openpyxl.utils.dataframe import dataframe_to_rows
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.ticker as ticker
from tqdm import tqdm
import warnings
warnings.filterwarnings('ignore')
w.start()







# %%
### data to excel
#所有参数
path_data=r"./输入/"
path_template=r'./输入/精准模板/模板布局/'
path_index=r'./输入/精准模板/模板参数/完整参数.pkl'
path_result=r"./输出/"
today='2024-03-29'
rptDate_all='2023-12-31'
rptDate_top='2023-12-31'
input_date = '2024Q1'
language_type='zh'



#运行更新excel数据
#读取数据
print('读取数据')
df_link=pd.read_pickle(path_data+'关联基金.pkl')
df_link.set_index('FundCode',inplace=True)
df_FundType=pd.read_excel(path_data+f'上海证券基金类型变更表（{today}）.xlsx',sheet_name='基金类型变更表',dtype='str', header=0, skiprows=1)
df_FundType['基金代码']=df_FundType['基金代码'].apply(lambda x:x+'.OF')
df_FundType.set_index('基金代码',inplace=True)
df_rating=pd.read_excel(path_data+'基金评级查询结果.xlsx',sheet_name='基金评价',dtype='str')
df_rating.set_index('基金代码',inplace=True)
#df_fund_special_date基本不变，故请手动到wind下载
#打开基金数据浏览器
#待选范围为：全部基金(含未成立、已到期)（注：是否可只选择全部基金，尚未测试）
#待选指标为：净值披露首日、基金到期日
df_fund_special_date=pd.read_excel(path_data+f'基金特殊日期（{today}）.xlsx',dtype=str)
df_fund_special_date=df_fund_special_date.iloc[:-2]#删除最后两行的wind水印
df_fund_special_date.rename(columns={'证券代码':'FundCode','证券简称':'FundName'},inplace=True)
df_fund_special_date['基金到期日'].fillna(pd.Timestamp(today)+pd.Timedelta(days=1),inplace=True)
df_fund_special_date['基金到期日']=df_fund_special_date['基金到期日'].apply(pd.Timestamp)
df_fund_special_date=df_fund_special_date[(df_fund_special_date['基金到期日']>=pd.Timestamp(today))]#筛选出尚未到期的基金
df_fund_special_date['净值披露首日']=df_fund_special_date['净值披露首日'].apply(pd.Timestamp)

def get_rank(df_FundType,FundCode,today):
    ###筛选出同类型基金列表
    FundInvestType=df_FundType.loc[FundCode,'三级分类']
    df_FundType_part=df_FundType[df_FundType['三级分类']==FundInvestType]
    ###同类型基金列表中，筛选出尚未到期的基金
    df_FundType_part=df_FundType_part.loc[df_FundType_part.index.isin(df_fund_special_date['FundCode'].unique())]
    str_funds=','.join(df_FundType_part.index.tolist())


    ###获取各时间点
    #上年年末时间点
    time_ytd=datetime(datetime.strptime(today, '%Y-%m-%d').year-1 , 12, 31).strftime('%Y-%m-%d')
    #前3月、6月、1年、2年时间点
    time_3m=w.tdaysoffset(-3, today, "Days=Alldays;Period=M").Data[0][0].strftime('%Y-%m-%d')
    time_6m=w.tdaysoffset(-6, today, "Days=Alldays;Period=M").Data[0][0].strftime('%Y-%m-%d')
    time_1y=w.tdaysoffset(-1, today, "Days=Alldays;Period=Y").Data[0][0].strftime('%Y-%m-%d')
    time_2y=w.tdaysoffset(-2, today, "Days=Alldays;Period=Y").Data[0][0].strftime('%Y-%m-%d')
    #成立的时间点
    #由于成立日的净值可能为nan值，故需要寻找净值披露首日
    se_fund_special_date_part=df_fund_special_date[df_fund_special_date['FundCode']==FundCode].iloc[0]
    time_setup=se_fund_special_date_part['净值披露首日'].strftime('%Y-%m-%d')


    ###检测是否要创建文件或更新数据
    flag1=True#是否存有截面净值文件
    flag2=True#截面净值文件是否有指定类型的数据
    if os.path.exists(path_data+f'截面净值_{today}.pkl'):
        df=pd.read_pickle(path_data+f'截面净值_{today}.pkl')
        if FundInvestType not in df['三级分类'].unique():
            flag2=False
    else:
        flag1=False

    if flag1==True and flag2==True:
        print('已存在数据')
        if time_setup in df.columns:
            print('数据已存在成立日净值，不再更新')
        else:
            print('数据无存在成立日净值，进行更新')
            tmp=w.wss(str_funds, "NAV_adjusted_transform",f"tradeDate={time_setup}")
            tmp=pd.DataFrame({'FundCode': tmp.Codes,time_setup:tmp.Data[0]})
            tmp.set_index('FundCode',inplace=True)
            df=df.merge(tmp,left_index=True,right_index=True,how='outer')
            df.to_pickle(path_data+f'截面净值_{today}.pkl')
    else:
        print('不存在数据，开始更新')
        ###获取各时间点的净值
        dic={'今日':today,time_setup:time_setup,
             '年初':time_ytd,'三月前':time_3m,
             '六月前':time_6m,'一年前':time_1y,
             '两年前':time_2y}
        lst_df=[]
        for key in dic.keys():
            time=dic[key]
            tmp=w.wss(str_funds, "NAV_adjusted_transform",f"tradeDate={time}")
            tmp=pd.DataFrame({'FundCode': tmp.Codes,key:tmp.Data[0]})
            tmp.set_index('FundCode',inplace=True)
            lst_df.append(tmp)
        df_new=pd.concat(lst_df,axis=1,join='outer')
        df_new['三级分类']=FundInvestType

        ###合并到已有数据，并保存
        if flag1==False:
            df=df_new.copy()
        elif flag2==False:
            df=pd.concat([df,df_new])
        df.to_pickle(path_data+f'截面净值_{today}.pkl')
        
    ####在更新好数据的基础上，筛选出同类基金的截面净值数据
    df=df[df['三级分类']==FundInvestType]
    ###计算各收益率
    df['收益_成立以来']=df['今日']/df[time_setup]-1
    df['收益_ytd']=df['今日']/df['年初']-1
    df['收益_3m']=df['今日']/df['三月前']-1
    df['收益_6m']=df['今日']/df['六月前']-1
    df['收益_1y']=df['今日']/df['一年前']-1
    df['收益_2y']=df['今日']/df['两年前']-1

    #如果成立日期比今年年初/前三月...大，则不计算收益率和排名
    ret_setup_self=f"{df.loc[FundCode,'收益_成立以来']:.2%}"
    if datetime.strptime(time_setup,'%Y-%m-%d')>datetime.strptime(time_ytd,'%Y-%m-%d'):
        ret_ytd_self='--'
    else:
        ret_ytd_self=f"{df.loc[FundCode,'收益_ytd']:.2%}"
    if datetime.strptime(time_setup,'%Y-%m-%d')>datetime.strptime(time_3m,'%Y-%m-%d'):
        ret_3m_self='--'
    else:
        ret_3m_self=f"{df.loc[FundCode,'收益_3m']:.2%}"
    if datetime.strptime(time_setup,'%Y-%m-%d')>datetime.strptime(time_6m,'%Y-%m-%d'):
        ret_6m_self='--'
    else:
        ret_6m_self=f"{df.loc[FundCode,'收益_6m']:.2%}"
    if datetime.strptime(time_setup,'%Y-%m-%d')>datetime.strptime(time_1y,'%Y-%m-%d'):
        ret_1y_self='--'
    else:
        ret_1y_self=f"{df.loc[FundCode,'收益_1y']:.2%}"
    if datetime.strptime(time_setup,'%Y-%m-%d')>datetime.strptime(time_2y,'%Y-%m-%d'):
        ret_2y_self='--'
    else:
        ret_2y_self=f"{df.loc[FundCode,'收益_2y']:.2%}"

    ###计算排名
    ret_setup_rank=f"{int(df['收益_成立以来'].dropna().rank(ascending=False,na_option='bottom').loc[FundCode])}/{int(len(df['收益_成立以来'].dropna()))}"
    if datetime.strptime(time_setup,'%Y-%m-%d')>datetime.strptime(time_ytd,'%Y-%m-%d'):
        ret_ytd_rank='--'
    else:
        ret_ytd_rank=f"{int(df['收益_ytd'].dropna().rank(ascending=False,na_option='bottom').loc[FundCode])}/{int(len(df['收益_ytd'].dropna()))}"
    if datetime.strptime(time_setup,'%Y-%m-%d')>datetime.strptime(time_3m,'%Y-%m-%d'):
        ret_3m_rank='--'
    else:
        ret_3m_rank=f"{int(df['收益_3m'].dropna().rank(ascending=False,na_option='bottom').loc[FundCode])}/{int(len(df['收益_3m'].dropna()))}"
    if datetime.strptime(time_setup,'%Y-%m-%d')>datetime.strptime(time_6m,'%Y-%m-%d'):
        ret_6m_rank='--'
    else:
        ret_6m_rank=f"{int(df['收益_6m'].dropna().rank(ascending=False,na_option='bottom').loc[FundCode])}/{int(len(df['收益_6m'].dropna()))}"
    if datetime.strptime(time_setup,'%Y-%m-%d')>datetime.strptime(time_1y,'%Y-%m-%d'):
        ret_1y_rank='--'
    else:
        ret_1y_rank=f"{int(df['收益_1y'].dropna().rank(ascending=False,na_option='bottom').loc[FundCode])}/{int(len(df['收益_1y'].dropna()))}"
    if datetime.strptime(time_setup,'%Y-%m-%d')>datetime.strptime(time_2y,'%Y-%m-%d'):
        ret_2y_rank='--'
    else:
        ret_2y_rank=f"{int(df['收益_2y'].dropna().rank(ascending=False,na_option='bottom').loc[FundCode])}/{int(len(df['收益_2y'].dropna()))}"


    ###计算同类平均收益率
    ret_setup_avg=f"{df['收益_成立以来'].dropna().mean():.2%}"
    ret_ytd_avg=f"{df['收益_ytd'].dropna().mean():.2%}"
    ret_3m_avg=f"{df['收益_3m'].dropna().mean():.2%}"
    ret_6m_avg=f"{df['收益_6m'].dropna().mean():.2%}"
    ret_1y_avg=f"{df['收益_1y'].dropna().mean():.2%}"
    ret_2y_avg=f"{df['收益_2y'].dropna().mean():.2%}"


    ###计算业绩基准收益率
    if FundCode!='000259.OF':
        BchCode=w.wss(FundCode, 'fund_benchindexcode').Data[0][0]
        ret_setup_bch=f'{(((w.wss(BchCode, "close",f"tradeDate={today}").Data[0][0]/w.wss(BchCode, "close",f"tradeDate={time_setup}").Data[0][0]))-1):.2%}'
        if datetime.strptime(time_setup,'%Y-%m-%d')>datetime.strptime(time_ytd,'%Y-%m-%d'):
            ret_ytd_bch='--'
        else:
            ret_ytd_bch=f'{(((w.wss(BchCode, "close",f"tradeDate={today}").Data[0][0]/w.wss(BchCode, "close",f"tradeDate={time_ytd}").Data[0][0]))-1):.2%}'
        if datetime.strptime(time_setup,'%Y-%m-%d')>datetime.strptime(time_3m,'%Y-%m-%d'):
            ret_3m_bch='--'
        else:
            ret_3m_bch=f'{(((w.wss(BchCode, "close",f"tradeDate={today}").Data[0][0]/w.wss(BchCode, "close",f"tradeDate={time_3m}").Data[0][0]))-1):.2%}'
        if datetime.strptime(time_setup,'%Y-%m-%d')>datetime.strptime(time_6m,'%Y-%m-%d'):
            ret_6m_bch='--'
        else:
            ret_6m_bch=f'{(((w.wss(BchCode, "close",f"tradeDate={today}").Data[0][0]/w.wss(BchCode, "close",f"tradeDate={time_6m}").Data[0][0]))-1):.2%}'
        if datetime.strptime(time_setup,'%Y-%m-%d')>datetime.strptime(time_1y,'%Y-%m-%d'):
            ret_1y_bch='--'
        else:
            ret_1y_bch=f'{(((w.wss(BchCode, "close",f"tradeDate={today}").Data[0][0]/w.wss(BchCode, "close",f"tradeDate={time_1y}").Data[0][0]))-1):.2%}'
        if datetime.strptime(time_setup,'%Y-%m-%d')>datetime.strptime(time_2y,'%Y-%m-%d'):
            ret_2y_bch='--'
        else:
            ret_2y_bch=f'{(((w.wss(BchCode, "close",f"tradeDate={today}").Data[0][0]/w.wss(BchCode, "close",f"tradeDate={time_2y}").Data[0][0]))-1):.2%}'
    else:
        ret_setup_bch='--'
        ret_ytd_bch='--'
        ret_3m_bch='--'
        ret_6m_bch='--'
        ret_1y_bch='--'
        ret_2y_bch='--'

    ###存入dataframe
    df_rank=pd.DataFrame(index=['今年以来','近三个月','近六个月',
                                '近一年','近两年','成立以来'],
                         columns=['本基金收益','同类排名','同类平均','比较基准'])
    df_rank.iloc[0]=[ret_ytd_self,ret_ytd_rank,ret_ytd_avg,ret_ytd_bch]
    df_rank.iloc[1]=[ret_3m_self,ret_3m_rank,ret_3m_avg,ret_3m_bch]
    df_rank.iloc[2]=[ret_6m_self,ret_6m_rank,ret_6m_avg,ret_6m_bch]
    df_rank.iloc[3]=[ret_1y_self,ret_1y_rank,ret_1y_avg,ret_1y_bch]
    df_rank.iloc[4]=[ret_2y_self,ret_2y_rank,ret_2y_avg,ret_2y_bch]
    df_rank.iloc[5]=[ret_setup_self,ret_setup_rank,ret_setup_avg,ret_setup_bch]

    df_rank=df_rank.T
    df_rank.loc['']=['','','','','','']
    df_rank.loc['截止日期']=[today,'','','','','']
    df_rank.loc['年初日期']=[time_ytd,'','','','','']
    df_rank.loc['前三个月日期']=[time_3m,'','','','','']
    df_rank.loc['前六个月日期']=[time_6m,'','','','','']
    df_rank.loc['前一年日期']=[time_1y,'','','','','']
    df_rank.loc['前两年日期']=[time_2y,'','','','','']
    df_rank.loc[f'{FundCode}成立日期']=[time_setup,'','','','','']
    if language_type=='en':
        df_rank.rename(index={'本基金收益':'Total Return',
                              '同类排名':'Rank',
                              '同类平均':'Category Mean',
                              '比较基准':'benchmark'},inplace=True)
    print('业绩排名获取完毕')
    return df_rank

tmp=pd.read_excel('target.xlsx',sheet_name='en',dtype=str)
tmp['FundCode']=tmp['FundCode'].apply(lambda x:x.zfill(6)+'.OF')
for FundCode in tmp['FundCode'].unique():
    df_rank=get_rank(df_FundType,FundCode,today)
    df_rank.to_excel(FundCode+'.xlsx')










