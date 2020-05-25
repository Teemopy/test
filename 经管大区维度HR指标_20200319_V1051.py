import numpy as np
import pandas as pd

folderPath = r'E:\2020年提供经管数据\大区维度HR指标\文件读取\0525'
start = '2020-05-01'
end = '2020-05-24'
period = start+'_'+end
sp = folderPath+'\\'+'花名册完整版_{start}.xlsx'.format(start=start)
ep = folderPath+'\\'+'花名册完整版_{end}.xlsx'.format(end=end)
resignp = folderPath+'\\'+'离职减员表（新）_{period}.xlsx'.format(period=period)
entryp = folderPath+'\\'+'入职增员表（新）_{period}.xlsx'.format(period=period)
to_path = r'E:\2020年提供经管数据\大区维度HR指标\生成报表\2020年5月截至第3周大区维度人员指标test.xlsx'

stardata = pd.read_excel(sp)
enddata = pd.read_excel(ep)
resigndata = pd.read_excel(resignp)
entrydata = pd.read_excel(entryp)

position = ['买卖经纪人', '综合经纪人', '租赁经纪人', '买卖店经理', '综合店经理', '租赁店经理']
include = ['京北运营', '京东环运营', '京东北运营', '京东运营', '京东南运营', '京西北运营', '京西南运营', '京西运营', '京中运营', '京南运营']
star1 = stardata[(stardata['运营/职能'] == '运营') & (stardata['营销大区/部门'].isin(include)) & (stardata['职务'].isin(position))]
end1 = enddata[(enddata['运营/职能'] == '运营') & (enddata['营销大区/部门'].isin(include)) & (enddata['职务'].isin(position))]
resign1 = resigndata[(resigndata['运营/职能'] == '运营') & (resigndata['营销大区/部门'].isin(include)) & (resigndata['职务'].isin(position))]
entrydata['职务'] = entrydata['主要职位'].apply(lambda x: x[4:-2])
entry1 = entrydata[(entrydata['运营/职能'] == '运营') & (entrydata['营销大区/部门'].isin(include)) & (entrydata['职务'].isin(['经纪人', '店经理']))]

# 经纪人数据
star_agent = star1.groupby(['门店']).count()
star_agent = star_agent.rename(columns={'员工编号': '期初经纪人数量'})
end_agent = end1.groupby(['门店']).count()
end_agent = end_agent.rename(columns={'员工编号': '期末经纪人数量'})
resign_agent = resign1.groupby(['门店']).count()
resign_agent = resign_agent.rename(columns={'员工编号': '流失经纪人数量'})
entry_agent = entry1.groupby(['门店']).count()
entry_agent = entry_agent.rename(columns={'系统号': '入职经纪人数量'})

# 将数据进行合并,准备接入组织架构信息
store_data = pd.concat([star_agent, end_agent, resign_agent, entry_agent], axis=1, sort=False)
store_data = store_data.fillna(0)
store_data['门店1'] = store_data.index
store_data.rename(columns={'营销大区/部门': '事业部1', '业务区域/组': '大区1'}, inplace=True)

# 获取组织架构信息的数据框
org_str = pd.concat([end1, entry1, resign1, star1], axis=0, sort=False)
org_str = org_str[['营销大区/部门', '业务区域/组', '门店']]
org_str = org_str.drop_duplicates(['门店'])
org_str.rename(columns={'营销大区/部门': '事业部', '业务区域/组': '大区'}, inplace=True)

# 数据接入组织架构信息
m_data = pd.merge(org_str, store_data, how='left', left_on='门店', right_on='门店1')
m_data = m_data[['事业部', '大区', '门店', '期初经纪人数量', '期末经纪人数量', '入职经纪人数量', '流失经纪人数量']]


def turn(a, b, c):
    t = round(a * 2 / (b + c), 7)
    return t


# 完成门店维度后进行大区汇总
d_data = m_data.groupby(['大区']).sum()
d_data['大区1'] = d_data.index
org_dstr = org_str[['大区', '事业部']]
org_dstr = org_dstr.drop_duplicates(['大区'])
d_data = pd.merge(org_dstr, d_data, how='left', left_on='大区', right_on='大区1')
d_data['经纪人流失率'] = turn(d_data['流失经纪人数量'], d_data['期初经纪人数量'], d_data['期末经纪人数量'])
d_data['人员净增长'] = d_data['入职经纪人数量']-d_data['流失经纪人数量']
d_data['经纪人流失率排名'] = d_data['经纪人流失率'].rank(method='min', ascending=True)
d_data['人员净增长排名'] = d_data['人员净增长'].rank(method='min', ascending=False)
d_data = d_data[['事业部', '大区', '期初经纪人数量', '期末经纪人数量', '经纪人流失率', '经纪人流失率排名', '入职经纪人数量', '流失经纪人数量', '人员净增长', '人员净增长排名']]
d_data['事业部'] = d_data['事业部'].apply(lambda x: x[:-2])

d_data.to_excel(to_path, index=False)
