# encoding: utf-8
import pandas as pd
import numpy as np
import xlrd
from WindPy import *
w.start()

# 数据源的数据格式为xlsx，包含四列：Date、Id、Name、Weight，其中Id为万德代码，日期从上到下顺序排列

end_date = '2017-09-30'
file = 'C:\\Users\\yz283\\Desktop\\holding.xlsx'
write_file = 'C:\\Users\\yz283\\Desktop\\Brinson_result.xlsx'

# 主要参考资料如下：
# http://www.docin.com/p-1514243338.html
# 辅助参考材料如下：
# https://www.ricequant.com/community/topic/4204/%E7%BB%A9%E6%95%88%E5%88%86%E6%9E%90%E4%B9%8Bbrinson%E6%A8%A1%E5%9E%8B
# http://myfof.org/thread-191-1-1.html
# http://myfof.org/forum.php?mod=viewthread&tid=320&highlight=%BB%F9%BD%F0%D2%B5%BC%A8%B9%E9%D2%F2
# https://cran.r-project.org/web/packages/pa/vignettes/pa.pdf

# 请注意Brinson模型一个无法避免的问题就是截面持仓数据时间间隔越大
# 通过Brinson方法按持仓数据计算的收益率与实际收益率的偏差就越大
# 基金持仓数据的密度越大越可规避此问题，如在主要参考资料中，采用的实证数据为基金每日交易数据

# Single Period Brinson Model
def brinson_single(end_date, holding_add, write_file):
    # 由于不少管理人涉及港股头寸，行业分类统一采用GICS分类标准
    Q = pd.DataFrame(['能源', '材料', '工业', '可选消费', '日常消费', '房地产', '信息技术', '公用事业',
                      '电信服务', '金融', '医疗保健'], columns=['GICS'])

    # 读取持仓数据，并计算对应的行业权重和区间收益率
    excel = xlrd.open_workbook(holding_add)
    sheet = excel.sheets()[0]
    holding = pd.DataFrame(sheet._cell_values[1:], columns=['Date', 'Id', 'Name', 'Weight_hld'])
    start_date = holding.Date[0]
    holding = holding[holding.Date == start_date]
    holding['Industry'] = \
    w.wss(holding['Id'].tolist(), "industry_gics", "industryType=1;tradeDate=" + start_date).Data[0]
    holding['Rtn_hld'] = \
    w.wss(holding['Id'].tolist(), "pct_chg_per", "startDate=" + start_date + ";endDate=" + end_date).Data[0]
    holding['Rtn_Prod'] = holding['Rtn_hld'] * holding['Weight_hld']
    holding_return = holding['Rtn_Prod'].sum()
    hld = holding.groupby(['Industry']).sum().reset_index()
    hld['Rtn_hld'] = hld['Rtn_Prod'] / hld['Weight_hld']
    del hld['Rtn_Prod']

    # 基准设定为沪深300，下面计算沪深300的行业权重和区间收益率
    hs300 = w.wset('IndexConstituent', date=start_date, windcode='000300.SH').Data
    rtn = w.wss(hs300[1], "pct_chg_per", "startDate=" + start_date + ";endDate=" + end_date).Data[0]
    hs300 = pd.DataFrame(hs300 + w.wss(hs300[1], "industry_gics", "industryType=1;tradeDate=" + start_date).Data \
                         + [rtn] + [[a * b for a, b in zip(hs300[3], rtn)]],
                         index=['Date', 'Id', 'Name', 'Weight_index', 'Industry', 'Rtn', 'Rtn_Prod']).T
    index = hs300[['Industry', 'Weight_index', 'Rtn_Prod']].groupby(['Industry']).sum().reset_index()
    index['Weight_index'] = index['Weight_index'] / 100
    # 通过期初start_date的权重对个股区间收益进行加权，一定会与沪深300的区间涨跌幅有误差，通过沪深300真实的涨跌幅按比例调整行业涨跌幅
    index['Rtn_Prod'] = index['Rtn_Prod'] / index['Weight_index']
    index_return = w.wss('000300.SH', "pct_chg_per", "startDate=" + start_date + ";endDate=" + end_date).Data[0][0]
    index['Rtn_Prod'] = index['Rtn_Prod'] * (index_return / sum(index['Rtn_Prod'] * index['Weight_index']))
    index.rename(columns={'Rtn_Prod': 'Rtn_index'}, inplace=True)

    Q = Q.merge(hld, left_on='GICS', right_on='Industry', how='left').merge(index, left_on='GICS', right_on='Industry',
                                                                            how='left')
    Q.drop(['Industry_x', 'Industry_y'], axis=1, inplace=True)
    Q.fillna(0, inplace=True)

    Q['Q1'] = Q['Weight_index'] * Q['Rtn_index']
    Q['Q2'] = Q['Weight_hld'] * Q['Rtn_index']
    Q['Q3'] = Q['Weight_index'] * Q['Rtn_hld']
    Q['Q4'] = Q['Weight_hld'] * Q['Rtn_hld']

    Q['行业配置'] = Q['Q2'] - Q['Q1']
    Q['个股选择'] = Q['Q3'] - Q['Q1']
    Q['交互收益'] = Q['Q4'] - Q['Q3'] - Q['Q2'] + Q['Q1']
    Q['总超额收益'] = Q['Q4'] - Q['Q1']
    Q.ix['合计'] = Q.apply(lambda x: x.sum())
    Q.GICS.iat[-1] = '合计'
    Q.Rtn_hld.iat[-1] = holding_return
    Q.Rtn_index.iat[-1] = index_return

    writer = pd.ExcelWriter(write_file)
    Q.to_excel(writer, 'Brinson_Single')
    writer.save()

    return Q

# Multi Period Brinson Model
def brinson_multi(end_date_, holding_add, write_file):
    # 由于不少管理人涉及港股头寸，行业分类统一采用GICS分类标准
    Q = pd.DataFrame(['能源', '材料', '工业', '可选消费', '日常消费', '房地产', '信息技术', '公用事业',
                      '电信服务', '金融', '医疗保健'], columns=['GICS'])

    # 读取持仓数据，并计算对应的行业权重和区间收益率
    excel = xlrd.open_workbook(holding_add)
    sheet = excel.sheets()[0]
    holding_total = pd.DataFrame(sheet._cell_values[1:], columns=['Date', 'Id', 'Name', 'Weight_hld'])
    date = holding_total['Date'].unique()
    Rp, Raa, Rss, Rb = [0], [0], [0], [0]

    for i in range(date.shape[0]):
        start_date = date[i]
        end_date = date[i + 1] if (i < date.shape[0] - 1) else end_date_
        holding = holding_total.loc[holding_total.Date == start_date]
        holding['Industry'] = \
            w.wss(holding['Id'].tolist(), "industry_gics", "industryType=1;tradeDate=" + start_date).Data[0]
        holding['Rtn_hld'] = \
            w.wss(holding['Id'].tolist(), "pct_chg_per", "startDate=" + start_date + ";endDate=" + end_date).Data[0]
        holding['Rtn_Prod'] = holding['Rtn_hld'] * holding['Weight_hld']
        holding = holding.groupby(['Industry']).sum().reset_index()
        holding['Rtn_hld'] = holding['Rtn_Prod'] / holding['Weight_hld']
        del holding['Rtn_Prod']
        holding.rename(columns={'Weight_hld': 'Weight_hld_' + start_date, 'Rtn_hld': 'Rtn_hld_' + start_date},
                       inplace=True)

        # 基准设定为沪深300，下面计算沪深300的行业权重和区间收益率
        hs300 = w.wset('IndexConstituent', date=start_date, windcode='000300.SH').Data
        rtn = w.wss(hs300[1], "pct_chg_per", "startDate=" + start_date + ";endDate=" + end_date).Data[0]
        hs300 = pd.DataFrame(hs300 + w.wss(hs300[1], "industry_gics", "industryType=1;tradeDate=" + start_date).Data \
                             + [rtn] + [[a * b for a, b in zip(hs300[3], rtn)]],
                             index=['Date', 'Id', 'Name', 'Weight_index', 'Industry', 'Rtn', 'Rtn_Prod']).T
        index = hs300[['Industry', 'Weight_index', 'Rtn_Prod']].groupby(['Industry']).sum().reset_index()
        index['Weight_index'] = index['Weight_index'] / 100
        # 通过期初start_date的权重对个股区间收益进行加权，一定会与沪深300的区间涨跌幅有误差，通过沪深300真实的涨跌幅按比例调整行业涨跌幅
        index['Rtn_Prod'] = index['Rtn_Prod'] / index['Weight_index']
        index_return = w.wss('000300.SH', "pct_chg_per", "startDate=" + start_date + ";endDate=" + end_date).Data[0][0]
        index['Rtn_Prod'] = index['Rtn_Prod'] * (index_return / sum(index['Rtn_Prod'] * index['Weight_index']))
        index.rename(columns={'Rtn_Prod': 'Rtn_index_' + start_date, 'Weight_index': 'Weight_index_' + start_date},
                     inplace=True)

        Q = Q.merge(holding, left_on='GICS', right_on='Industry', how='left').merge(index, left_on='GICS',
                                                                                    right_on='Industry',
                                                                                    how='left')
        Q.drop(['Industry_x', 'Industry_y'], axis=1, inplace=True)
        Q.fillna(0, inplace=True)

        # 计算当期收益率：实际组合收益率Rp，积极资产配置组合收益率Raa，积极股票选择组合收益率Rss，基准组合收益率Rb
        if i == date.shape[0] - 1:
            continue
        Rp.append((Q.iloc[:, -3] * Q.iloc[:, -4]).sum())
        Rb.append((Q.iloc[:, -1] * Q.iloc[:, -2]).sum())
        Raa.append((Q.iloc[:, -1] * Q.iloc[:, -4]).sum())
        Rss.append((Q.iloc[:, -3] * Q.iloc[:, -2]).sum())

    # 计算四个组合的累积收益率
    Rp = (np.array(Rp) / 100 + 1).cumprod()
    Rb = (np.array(Rb) / 100 + 1).cumprod()
    Raa = (np.array(Raa) / 100 + 1).cumprod()
    Rss = (np.array(Rss) / 100 + 1).cumprod()

    industry_attribution_p = pd.DataFrame()
    industry_attribution_b = pd.DataFrame()
    industry_attribution_aa = pd.DataFrame()
    industry_attribution_ss = pd.DataFrame()
    for i in range(date.shape[0]):

        # 计算各个行业的收益贡献度
        industry_attribution_p['Ip_' + date[i]] = Q.iloc[:, 1 + 4 * i + 0] * Q.iloc[:, 1 + 4 * i + 1] * Rp[i]
        industry_attribution_b['Ib_' + date[i]] = Q.iloc[:, 1 + 4 * i + 2] * Q.iloc[:, 1 + 4 * i + 3] * Rb[i]
        industry_attribution_aa['Iaa_' + date[i]] = Q.iloc[:, 1 + 4 * i + 0] * Q.iloc[:, 1 + 4 * i + 3] * Raa[i]
        industry_attribution_ss['Iss_' + date[i]] = Q.iloc[:, 1 + 4 * i + 1] * Q.iloc[:, 1 + 4 * i + 2] * Rss[i]

    industry_attribution_p = industry_attribution_p.cumsum(axis=1)
    industry_attribution_b = industry_attribution_b.cumsum(axis=1)
    industry_attribution_aa = industry_attribution_aa.cumsum(axis=1)
    industry_attribution_ss = industry_attribution_ss.cumsum(axis=1)


    industry_attribution_p.index, industry_attribution_b.index = Q['GICS'].tolist(), Q['GICS'].tolist()

    Qp = industry_attribution_p.cumsum().iloc[-1,:].tolist()
    Qb = industry_attribution_b.cumsum().iloc[-1,:].tolist()
    Qaa = industry_attribution_aa.cumsum().iloc[-1,:].tolist()
    Qss = industry_attribution_ss.cumsum().iloc[-1,:].tolist()

    result = pd.DataFrame([Qp, Qb, Qaa, Qss], index=['Qp', 'Qb', 'Qaa', 'Qss'],
                          columns=date.tolist()[1:] + [end_date]).T
    result['行业配置'] = result['Qaa'] - result['Qb']
    result['个股选择'] = result['Qss'] - result['Qb']
    result['交互收益'] = result['Qp'] - result['Qss'] - result['Qaa'] + result['Qb']
    result['总超额收益'] = result['Qp'] - result['Qb']

    industry_excess = pd.DataFrame(industry_attribution_p.as_matrix() - industry_attribution_b.as_matrix(),
                 columns=industry_attribution_p.columns, index=Q['GICS'].tolist())
    industry_aa = pd.DataFrame(industry_attribution_aa.as_matrix() - industry_attribution_b.as_matrix(),
                 columns=industry_attribution_p.columns, index=Q['GICS'].tolist())
    industry_ss = pd.DataFrame(industry_attribution_ss.as_matrix() - industry_attribution_b.as_matrix(),
                 columns=industry_attribution_p.columns, index=Q['GICS'].tolist())
    industry_ = pd.DataFrame(industry_attribution_p.as_matrix() - industry_attribution_ss.as_matrix() \
                 - industry_attribution_ss.as_matrix() + industry_attribution_b.as_matrix(),
                 columns=industry_attribution_p.columns, index=Q['GICS'].tolist())

    # 将结果写到对应路径的xlsx文件中
    writer = pd.ExcelWriter(write_file)
    result.to_excel(writer, 'Brinson_Multi')
    industry_attribution_p.to_excel(writer, '组合行业收益')
    industry_attribution_b.to_excel(writer, '基准行业收益')
    industry_excess.to_excel(writer, '行业超额收益贡献')
    industry_aa.to_excel(writer, '行业配置收益贡献')
    industry_ss.to_excel(writer, '行业选股收益贡献')
    industry_.to_excel(writer, '行业交互收益贡献')
    writer.save()

    return result, industry_attribution_p, industry_attribution_b, industry_excess,industry_aa,industry_ss,industry_,Q

brinson_multi(end_date, file, write_file)