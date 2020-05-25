# -*- coding: utf-8 -*-

import pandas as pd
import copy

folderPath = r'E:\2020年2月人力指标报表\文件读取\0522'
start = '2020-05-01'
end = '2020-05-21'
sp = folderPath + '\\' + '花名册完整版_{start}.xlsx'.format(start=start)
ep = folderPath + '\\' + '花名册完整版_{end}.xlsx'.format(end=end)
tp = folderPath + '\\' + '离职减员表（新）_{start}_{end}.xlsx'.format(start=start, end=end)
nep = folderPath + '\\' + '入职增员表（新）_{start}_{end}.xlsx'.format(start=start, end=end)
path = r'E:\2020年2月人力指标报表\生成报表\2020年5月第3周人力指标test.xlsx'
Sheet = ['事业部情况', '大区情况', '门店情况']

Sd = pd.read_excel(sp)
Ed = pd.read_excel(ep)
Td = pd.read_excel(tp)
Ned = pd.read_excel(nep)

# 筛选出运营经纪人数据
part = ['京东北事业部', '京东南事业部', '京东事业部', '京中事业部', '京北事业部', '京南事业部', '京西南事业部', '京西北事业部', '京西事业部']
Sd1 = Sd[(Sd['运营/职能'] == '运营') & (Sd['运营管理大区/中心'].isin(part)) & (Sd['PS职等'].isin(['A', 'M']))]
Ed1 = Ed[(Ed['运营/职能'] == '运营') & (Ed['运营管理大区/中心'].isin(part)) & (Ed['PS职等'].isin(['A', 'M']))]
Td1 = Td[(Td['运营/职能'] == '运营') & (Td['运营管理大区/中心'].isin(part)) & (Td['PS职等'].isin(['A', 'M']))]
Ned['主要职位1'] = Ned['主要职位'].apply(lambda x: x[4:-2])
Ned1 = Ned[(Ned['运营/职能'] == '运营') & (Ned['主要职位1'].isin(['经纪人', '店经理']))]

# 经纪人数据
Sdm_agent = Sd1.groupby(['门店']).count()
Sdm_agent = Sdm_agent.rename(columns={'员工编号': '期初经纪人数量'})
Edm_agent = Ed1.groupby(['门店']).count()
Edm_agent = Edm_agent.rename(columns={'员工编号': '期末经纪人数量'})
Tdm_agent = Td1.groupby(['门店']).count()
Tdm_agent = Tdm_agent.rename(columns={'员工编号': '流失经纪人数量'})
Ned_agent = Ned1.groupby(['门店']).count()
Ned_agent = Ned_agent.rename(columns={'系统号': '入职经纪人数量'})

# 统招本经纪人数据
Ed_tb = Ed1[Ed1['最高统招学历'].isin(['本科（统招）', '本科（实习生）', '本科', '硕士研究生', '硕士研究生（统招）', '硕士研究生（实习生）', '博士研究生', '博士研究生（统招）', '国外学历（本科）', '国外学历（硕士）'])].groupby(['门店']).count()
Ed_tb = Ed_tb.rename(columns={'员工编号': '期末统招本经纪人数量'})

# 将数据进行合并,准备接入组织架构信息
Mdata = pd.concat([Sdm_agent, Edm_agent, Tdm_agent, Ned_agent, Ed_tb], axis=1, sort=True)
Mdata = Mdata.fillna(0)

Mdata = Mdata[['期初经纪人数量', '期末经纪人数量', '流失经纪人数量', '入职经纪人数量', '期末统招本经纪人数量']]

# 获取组织架构信息的数据框
Zzjg = pd.concat([Ed1, Td1, Ned1, Sd1], axis=0, sort=True)
Zzjg = Zzjg[['业务区域/组', '营销大区/部门', '门店']]
Zzjg = Zzjg.drop_duplicates(['门店'])
Zzjg.index = Zzjg['门店']

# 深拷贝3份数据方便计算各维度流失情况
Data = Zzjg.join(Mdata, how='inner')

Datam = copy.deepcopy(Data)
Dataq = copy.deepcopy(Data)
Datab = copy.deepcopy(Data)


def turn(a, b, c):
    t = round(a*2/(b+c), 7)
    return t


def zb(a, b):
    t = round(a/b, 7)
    return t


# 门店维度流失数据输出
Datam['经纪人流失率'] = turn(Datam['流失经纪人数量'], Datam['期初经纪人数量'], Datam['期末经纪人数量'])
Datam['期末统招本占比'] = zb(Datam['期末统招本经纪人数量'], Datam['期末经纪人数量'])
Datam = Datam.rename(columns={'业务区域/组': '所属大区', '营销大区/部门': '所属事业部'})
mcolumns = ['门店', '所属大区', '所属事业部', '期初经纪人数量', '期末经纪人数量', '流失经纪人数量', '入职经纪人数量', '期末统招本占比', '经纪人流失率']
Datam['所属事业部'] = Datam['所属事业部'].apply(lambda x: x[:-2])
Datam = Datam[mcolumns]


# 大区维度流失数据输出
Dataq = Dataq.drop('门店', axis=1)
Dataq = Dataq.groupby(['业务区域/组']).sum()
Dataq['大区1'] = Dataq.index
Dataq = pd.merge(Dataq, Zzjg, how='left', left_on='大区1', right_on='业务区域/组')
Dataq = Dataq.rename(columns={'业务区域/组': '大区', '营销大区/部门': '所属事业部'})
Dataq['经纪人流失率'] = turn(Dataq['流失经纪人数量'], Dataq['期初经纪人数量'], Dataq['期末经纪人数量'])
Dataq['期末统招本占比'] = zb(Dataq['期末统招本经纪人数量'], Dataq['期末经纪人数量'])
Dataq = Dataq.drop_duplicates('大区')
qcolumns = ['大区', '所属事业部', '期初经纪人数量', '期末经纪人数量', '流失经纪人数量', '入职经纪人数量', '期末统招本占比', '经纪人流失率']
Dataq['所属事业部'] = Dataq['所属事业部'].apply(lambda x: x[:-2])
Dataq = Dataq[qcolumns]


# 事业部部维度数据输出
Datab = Datab.groupby(['营销大区/部门']).sum()
Datab['事业部'] = Datab.index
Datab = pd.merge(Datab, Zzjg, how='left', left_on='事业部', right_on='营销大区/部门')
Datab = Datab.drop_duplicates('事业部')
Datab.loc['整体情况'] = Datab.apply(lambda x: x.sum())
Datab['经纪人流失率'] = turn(Datab['流失经纪人数量'], Datab['期初经纪人数量'], Datab['期末经纪人数量'])
Datab['期末统招本占比'] = zb(Datab['期末统招本经纪人数量'], Datab['期末经纪人数量'])
Datab.loc['整体情况', '事业部'] = '整体情况'
qcolumns = ['事业部', '期初经纪人数量', '期末经纪人数量', '流失经纪人数量', '入职经纪人数量', '期末统招本占比', '经纪人流失率']
Datab['事业部'] = Datab['事业部'].apply(lambda x: x[:-2])
Datab = Datab[qcolumns]


# 输出至Excel
writer = pd.ExcelWriter(path)
Datab.to_excel(writer, sheet_name=Sheet[0], index=False)
Dataq.to_excel(writer, sheet_name=Sheet[1], index=False)
Datam.to_excel(writer, sheet_name=Sheet[2], index=False)

writer.save()
writer.close()
