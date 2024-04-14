'''
Date: 2023-10-25 15:24:45
LastEditors: Lei_T
LastEditTime: 2024-03-28 16:37:20
'''

#%%
import os
import pandas as pd 
from WindPy import w
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.ticker as ticker
import warnings
warnings.filterwarnings('ignore')
w.start()

def get_nav(FundCode,today,language_type):
    '''
    获取全历史净值走势
    '''
    ###指定IndexCode_1
    FundType=w.wss(FundCode, "fund_investtype").Data[0][0]
    IndexCode_2='nan'#设定一个初始值，如果是指定类型的基金，则IndexCode_2自动设置为空值，否则再自动匹配合适的指数
    if FundType=='普通股票型基金' or FundType=='Normal Equity Funds':
        IndexCode_1='885000.WI'
    elif FundType=='偏股混合型基金' or FundType=='Aggressive Allocation Funds':
        IndexCode_1='885001.WI'
    elif FundType=='平衡混合型基金' or FundType=='Balanced Funds':
        IndexCode_1='885002.WI'
    elif FundType=='偏债混合型基金' or FundType=='Moderate Allocation Funds':
        IndexCode_1='885003.WI'
        IndexCode_2=''
    elif FundType=='被动指数型基金' or FundType=='Passive Equity Index Funds':
        IndexCode_1='885004.WI'
    elif FundType=='混合债券型二级基金' or FundType=='Enhanced Bond Funds (Secondary Market)':
        IndexCode_1='885007.WI'
        IndexCode_2=''
    elif FundType=='中长期纯债型基金' or FundType=='Mid / Long-term Bond Funds':
        IndexCode_1='885008.WI'
        IndexCode_2=''
    elif FundType=='货币市场型基金' or FundType=='Money Market Funds':
        IndexCode_1='885009.WI'
        IndexCode_2=''
    elif FundType=='增强指数型基金' or FundType=='Enhanced Equity Index Funds':
        IndexCode_1='885044.WI'
    elif FundType=='灵活配置型基金' or FundType=='Flexible Allocation Funds':
        IndexCode_1='885061.WI'
    elif FundType=='短期纯债型基金' or FundType=='Short-term Bond Funds':
        IndexCode_1='885062.WI'
        IndexCode_2=''
    elif FundType=='被动指数型债券基金' or FundType=='Passive Bond Index Funds':
        IndexCode_1='885063.WI'
        IndexCode_2=''
    elif FundType=='混合型FOF基金' or FundType=='Hybrid fof fund':
        IndexCode_1='885072.WI'


    ###指定IndexCode_2
    if IndexCode_2=='nan':
        df_fund_bench=pd.read_excel(path_data+'农银基金基准.xlsx',index_col='FundCode')#之所以不从wind直接获取，是因为英文模式下不好判断IndexCode_2
        bench=df_fund_bench.loc[FundCode,'Benchmark']
        if '沪深300' in bench:
            IndexCode_2='000300.SH'
            if language_type=='zh':
                IndexCode_2_name='沪深300'
            else:
                IndexCode_2_name='CSI 300 INDEX'
        elif '中证700' in bench:
            IndexCode_2='000907.CSI'
            if language_type=='zh':
                IndexCode_2_name='中证700'
            else:
                IndexCode_2_name='CSI Small & MidCap 700 index'
        elif '中证800' in bench:
            IndexCode_2='000906.SH'
            if language_type=='zh':
                IndexCode_2_name='中证800'
            else:
                IndexCode_2_name='CSI 800 index'
        elif '中证1000' in bench:
            IndexCode_2='000852.SH'
            if language_type=='zh':
                IndexCode_2_name='中证1000'
            else:
                IndexCode_2_name='CSI 1000 Index'
        elif '国有企业综合' in bench:
            IndexCode_2='000955.CSI'
            if language_type=='zh':
                IndexCode_2_name='中证国有企业综合指数'
            else:
                IndexCode_2_name='CSI State-owned Enterprises Composite Index'
        elif '国有企业改革' in bench:
            IndexCode_2='399974.SZ'
            if language_type=='zh':
                IndexCode_2_name='中证国有企业改革指数'
            else:
                IndexCode_2_name='CSI State-Owned Enterprises Reform Index'
        elif '新能源' in bench:
            IndexCode_2='399808.SZ'
            if language_type=='zh':
                IndexCode_2_name='中证新能源指数'
            else:
                IndexCode_2_name='CSI New Energy Index'
        elif '大农业' in bench:
            IndexCode_2='399814.SZ'
            if language_type=='zh':
                IndexCode_2_name='中证大农业指数'
            else:
                IndexCode_2_name='CSI Grand Agriculture Index'
        elif '医药' in bench:
            IndexCode_2='000933.SH'
            if language_type=='zh':
                IndexCode_2_name='中证医药卫生指数'
            else:
                IndexCode_2_name='CSI Health Care Index'
        elif 'TMT' in bench:
            IndexCode_2='000998.CSI'
            if language_type=='zh':
                IndexCode_2_name='中证TMT产业主题指数'
            else:
                IndexCode_2_name='CSI TMT Industries Index'
        elif '内地消费' in bench:
            IndexCode_2='000942.CSI'
            if language_type=='zh':
                IndexCode_2_name='中证内地消费主题指数'
            else:
                IndexCode_2_name='CSI China Mainland Consumer Index'
        elif '新华社民族品牌工程' in bench:
            IndexCode_2='931403.CSI'
            if language_type=='zh':
                IndexCode_2_name='中证新华社民族品牌工程指数'
            else:
                IndexCode_2_name='CSI National Brands Project of Xinhua Index'
        elif '新兴产业' in bench:
            IndexCode_2='000964.CSI'
            if language_type=='zh':
                IndexCode_2_name='中证新兴产业指数'
            else:
                IndexCode_2_name='CSI Emerging Industries index'
        else:
            IndexCode_2=''

    ###基金净值
    df_nav_1=pd.read_excel(path_data+'农银基金净值.xlsx',sheet_name='净值')
    df_nav_1.rename(columns={df_nav_1.columns[0]:'TradingDay'},inplace=True)
    df_nav_1['TradingDay']=df_nav_1['TradingDay'].apply(pd.Timestamp)
    df_nav_1=df_nav_1.loc[:,['TradingDay',FundCode]]

    valid_index=(df_nav_1[df_nav_1[FundCode]>0]).index
    df_nav_1=df_nav_1.loc[valid_index]
    
    FundName=w.wss(FundCode,"fund_info_name").Data[0][0]
    if language_type=='zh':#如果是中文，则图例末尾加上“净值”字样
        FundName+='净值'
    elif language_type=='en':#如果是英文，则不加
        pass
    df_nav_1.rename(columns={FundCode:FundName},inplace=True)    

    ###基准净值
    df_nav_2=pd.read_excel(path_data+'农银基金净值.xlsx',sheet_name='基准指数价格')
    df_nav_2.rename(columns={df_nav_2.columns[0]:'TradingDay'},inplace=True)
    df_nav_2['TradingDay']=df_nav_2['TradingDay'].apply(pd.Timestamp)
    BchCode=FundCode.split('.')[0]+'BI.WI'
    df_nav_2=df_nav_2.loc[valid_index,['TradingDay',BchCode]]

    if language_type=='zh':
        BchName='业绩比较基准'
    elif language_type=='en':
        BchName='Benchmark'
    df_nav_2.rename(columns={BchCode:BchName},inplace=True)

    ###指数1净值
    df_nav_3=pd.read_excel(path_data+'农银基金净值.xlsx',sheet_name='基金指数价格')
    df_nav_3.rename(columns={df_nav_3.columns[0]:'TradingDay'},inplace=True)
    df_nav_3['TradingDay']=df_nav_3['TradingDay'].apply(pd.Timestamp)
    df_nav_3=df_nav_3.loc[valid_index,['TradingDay',IndexCode_1]]

    IndexName_1=w.wss(IndexCode_1, "sec_name").Data[0][0]
    if language_type=='zh':#如果是中文，则图例末尾加上“净值”字样
        if IndexCode_1=='885008.WI':
            IndexName_1='中长期纯债型基金指数净值'
        else:
            IndexName_1+='净值'
    elif language_type=='en':#如果是英文，则不加
        if IndexName_1.split(' ')[0]=='wind':
            IndexName_1=' '.join(IndexName_1.split(' ')[1:])
    df_nav_3.rename(columns={IndexCode_1:IndexName_1},inplace=True)

    ###指数2净值
    if IndexCode_2!='':
        df_nav_4=pd.read_excel(path_data+'农银基金净值.xlsx',sheet_name='宽基指数价格')
        df_nav_4.rename(columns={df_nav_4.columns[0]:'TradingDay'},inplace=True)
        df_nav_4['TradingDay']=df_nav_4['TradingDay'].apply(pd.Timestamp)
        df_nav_4=df_nav_4.loc[valid_index,['TradingDay',IndexCode_2]]
        df_nav_4.rename(columns={IndexCode_2:IndexCode_2_name},inplace=True)

    ###合并
    df_price=df_nav_1.merge(df_nav_2,on='TradingDay',how='left')
    df_price=df_price.merge(df_nav_3,on='TradingDay',how='left')
    if IndexCode_2!='':
        df_price=df_price.merge(df_nav_4,on='TradingDay',how='left')
    df_price.fillna(method='ffill',inplace=True)
    df_price.dropna(axis=1,how='all',inplace=True)
    df_price.dropna(axis=0,inplace=True)#防止开头基金净值为一直为1、基准没有数据的情况

    ###归一化
    factor=df_price[FundName].iloc[0]
    if BchName in df_price.columns:
        df_price[BchName]/=df_price[BchName].iloc[0]
        df_price[BchName]*=factor
    df_price[IndexName_1]/=df_price[IndexName_1].iloc[0]
    df_price[IndexName_1]*=factor
    if IndexCode_2!='':
        df_price[IndexCode_2_name]/=df_price[IndexCode_2_name].iloc[0]
        df_price[IndexCode_2_name]*=factor   
    df_price.set_index('TradingDay',inplace=True)
    df_price.dropna(axis=1,how='all',inplace=True)#000259.OF没有基准数据，需要drop

    print('净值走势获取完毕')
    return df_price




def line_plt(path,data,width=13, height=8,x_interval=6,dpi=300,linewidth=1.5,save=False,fontsize=6):
    """
        Description
        ----------
        画线段图
        
        Parameters
        ----------
        path: str. 保存图片的路径
        data: pd.DataFrame. 每列为需要绘画的数据
        width: int. 图表输出宽度
        height: int. 图表输出高度
        x_interval: int. 横坐标日期间隔N月
        dpi: int. 图表清晰度
        linewidth: int. 图表线段粗细
        save: bool. 是否保存

        Return
        ----------
        None
        """
    fig,ax = plt.subplots(figsize=(width,height),dpi=dpi)
    plt.rcParams['font.sans-serif'] =['Microsoft YaHei']
    plt.rcParams['font.size'] = 5
    line_color = ['#488FD0','#FF9F9F','#BFBFBF','#FFE389']
    if len(data.columns) <= 4:
        for i in range(len(data.columns)):
            ax.plot(data.iloc[:,i],color=line_color[i],label=data.iloc[:,i].name,linewidth=linewidth)
    else:
        for i in range(len(data.columns)):
            ax.plot(data.iloc[:,i],label=data.iloc[:,i].name,linewidth=linewidth)
    figheight = ax.get_window_extent().height
    figwidth = ax.get_window_extent().width
    ax.legend(loc='upper center',
              frameon=False,
              ncol=(int(len(data.columns)/2) if len(data.columns) == 4 else len(data.columns)),
              fontsize=fontsize,
              bbox_to_anchor=(0.5, 1.15))
    ax.spines['right'].set_visible(False)
    ax.spines['top'].set_visible(False)
    ax.spines['left'].set_color('#D9D9D9')
    ax.spines['bottom'].set_color('#D9D9D9')

    ###设置x轴
    #fmt_half_year = mdates.MonthLocator(interval = x_interval)
    #ax.xaxis.set_major_locator(fmt_half_year)
    last_month_start_time=data.index[data.index.map(lambda x:x.strftime('%Y%m'))==data.index[-1].strftime('%Y%m')][0]
    last_month_start_index=data.index.tolist().index(last_month_start_time)
    x_ticks = [data.index[0], data.index[last_month_start_index],
               data.index[int(len(data[:last_month_start_index].index)//4)],
               data.index[int(len(data[:last_month_start_index].index)//2)],
               data.index[int(len(data[:last_month_start_index].index)//4*3)]]
    x_labels = [data.index[0].strftime('%Y-%m'), data.index[last_month_start_index].strftime('%Y-%m'),
                data.index[int(len(data[:last_month_start_index].index)//4)].strftime('%Y-%m'),
                data.index[int(len(data[:last_month_start_index].index)//2)].strftime('%Y-%m'),
                data.index[int(len(data[:last_month_start_index].index)//4*3)].strftime('%Y-%m')]
    ax.set_xticks(x_ticks)
    ax.set_xticklabels(x_labels)

    ###设置y轴
    #_max=data.values.max()+(data.values.max()-data.values.min())/10
    #_min=data.values.min()-(data.values.max()-data.values.min())/10
    #ax.set_ylim(_min,_max)
    #ax.set_ylim(data.values.min()*0.99,data.values.max()*1.01)
    ax.set_ylim(data.values.min()-0.15,data.values.max()+0.6)
    ax.set_xlim(data.index[0],data.index[-1])
    plt.xticks(fontsize=fontsize)
    plt.yticks(fontsize=fontsize)
    plt.tight_layout()
    if save:
        plt.savefig(path+f'\{data.columns[0][:-2]}.png')
    else:
        plt.show()








# %%
### data to excel
#所有参数
path_data=r"./输入/"
path_template=r'./输入/精准模板/模板布局/'
path_index=r'./输入/精准模板/模板参数/完整参数.pkl'
path_result=r"./输出/"
FundCode='002190.OF'
today='2024-03-29'
rptDate_all='2023-12-31'
rptDate_top='2023-12-31'
input_date = '2024Q1'
language_type='zh'


width=11.71
height=4.15

df_price=get_nav(FundCode,today,language_type)
line_plt(r"输出",df_price,width/2.54,height/2.54,x_interval=15,dpi=300,linewidth=1.5,save=True)









