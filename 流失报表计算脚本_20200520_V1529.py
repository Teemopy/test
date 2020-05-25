import pandas as pd

folderPath = r'E:\2019经纪人流失报表\文件读取\test'
start = '2020-04-30'
end = '2020-05-01'
start_path = folderPath + '\\' + '员工花名册_{start}.xlsx'.format(start=start)
end_path = folderPath + '\\' + '员工花名册_{end}.xlsx'.format(end=end)
quit_path = folderPath + '\\' + '离职减员表_{start}_{end}.xlsx'.format(start=start, end=end)
Sheet = ['整体情况', '事业部情况', '大区情况', '门店情况']
path = r'E:\2019经纪人流失报表\生成报表\新版流失报表test.xlsx'

start_data = pd.read_excel(start_path)
end_data = pd.read_excel(end_path)
quit_data = pd.read_excel(quit_path)

part = ['京东北事业部', '京东南事业部', '京东事业部', '京中事业部', '京北事业部', '京南事业部', '京西南事业部', '京西北事业部', '京西事业部']
startagent_data = start_data[(start_data['三级组织'] == '运营') & (start_data['四级组织'].isin(part)) & (start_data['PS职等'].isin(['A', 'M']))]
endagent_data = end_data[(end_data['三级组织'] == '运营') & (end_data['四级组织'].isin(part)) & (end_data['PS职等'].isin(['A', 'M']))]
quitagent_data = quit_data[(quit_data['三级组织'] == '运营') & (quit_data['四级组织'].isin(part)) & (quit_data['PS职等'].isin(['A', 'M']))]


# 增加职级信息
def rank(data):
    data.loc[data['PS职等'] == 'A', '职级'] = data.loc[data['PS职等'] == 'A', '岗位_linkhr原职位'].apply(lambda x: x[x.index('A'):])
    data.loc[data['PS职等'] == 'M', '职级'] = data.loc[data['PS职等'] == 'M', '岗位_linkhr原职位'].apply(lambda x: x[x.index('M'):])
    return data


startagent_data = rank(startagent_data)
endagent_data = rank(endagent_data)
quitagent_data = rank(quitagent_data)

# 经纪人数据
startagent = startagent_data.groupby(['七级组织']).count()
startagent = startagent.rename(columns={'员工系统号': '期初经纪人数量'})
endagent = endagent_data.groupby(['七级组织']).count()
endagent = endagent.rename(columns={'员工系统号': '期末经纪人数量'})
quitagent = quitagent_data.groupby(['七级组织']).count()
quitagent = quitagent.rename(columns={'员工系统号': '流失经纪人数量'})

# A0-A2经纪人数据
lowrank_startagent = startagent_data[startagent_data['职级'].isin(['A0', 'A1', 'A2', 'A0/1', 'A0/2'])].groupby(['七级组织']).count()
lowrank_startagent = lowrank_startagent.rename(columns={'员工系统号': '期初A0-A2经纪人数量'})
lowrank_endagent = endagent_data[endagent_data['职级'].isin(['A0', 'A1', 'A2', 'A0/1', 'A0/2'])].groupby(['七级组织']).count()
lowrank_endagent = lowrank_endagent.rename(columns={'员工系统号': '期末A0-A2经纪人数量'})
lowrank_quitagent = quitagent_data[quitagent_data['职级'].isin(['A0', 'A1', 'A2', 'A0/1', 'A0/2'])].groupby(['七级组织']).count()
lowrank_quitagent = lowrank_quitagent.rename(columns={'员工系统号': '流失A0-A2经纪人数量'})

# 储备人才数据
reserve_startagent = startagent_data[startagent_data['职级'].isin(['A0', 'A1', 'A2', 'M1', 'M2', 'A0/1', 'A0/2'])].groupby(['七级组织']).count()
reserve_startagent = reserve_startagent.rename(columns={'员工系统号': '期初储备人才数量'})
reserve_endagent = endagent_data[endagent_data['职级'].isin(['A0', 'A1', 'A2', 'M1', 'M2', 'A0/1', 'A0/2'])].groupby(['七级组织']).count()
reserve_endagent = reserve_endagent.rename(columns={'员工系统号': '期末储备人才数量'})
reserve_quitagent = quitagent_data[quitagent_data['职级'].isin(['A0', 'A1', 'A2', 'M1', 'M2', 'A0/1', 'A0/2'])].groupby(['七级组织']).count()
reserve_quitagent = reserve_quitagent.rename(columns={'员工系统号': '流失储备人才数量'})

# 统招本经纪人数据
bachelor_startagent = startagent_data[startagent_data['统招最高教育程度'].isin(
    ['本科（统招）', '本科（实习生）', '本科', '硕士研究生', '硕士研究生（统招）', '硕士研究生（实习生）', '博士研究生', '博士研究生（统招）', '国外学历（本科）',
     '国外学历（硕士）'])].groupby(['七级组织']).count()
bachelor_startagent = bachelor_startagent.rename(columns={'员工系统号': '期初统招本经纪人数量'})
bachelor_endagent = endagent_data[endagent_data['统招最高教育程度'].isin(
    ['本科（统招）', '本科（实习生）', '本科', '硕士研究生', '硕士研究生（统招）', '硕士研究生（实习生）', '博士研究生', '博士研究生（统招）', '国外学历（本科）',
     '国外学历（硕士）'])].groupby(['七级组织']).count()
bachelor_endagent = bachelor_endagent.rename(columns={'员工系统号': '期末统招本经纪人数量'})
bachelor_quitagent = quitagent_data[quitagent_data['统招最高教育程度'].isin(
    ['本科（统招）', '本科（实习生）', '本科', '硕士研究生', '硕士研究生（统招）', '硕士研究生（实习生）', '博士研究生', '博士研究生（统招）', '国外学历（本科）',
     '国外学历（硕士）'])].groupby(['七级组织']).count()
bachelor_quitagent = bachelor_quitagent.rename(columns={'员工系统号': '流失统招本经纪人数量'})

# 租赁经纪人数据
lease_startagent = startagent_data[startagent_data['LinkHR职务'].isin(['租赁经纪人', '租赁店经理'])].groupby(['七级组织']).count()
lease_startagent = lease_startagent.rename(columns={'员工系统号': '期初租赁经纪人数量'})
lease_endagent = endagent_data[endagent_data['LinkHR职务'].isin(['租赁经纪人', '租赁店经理'])].groupby(['七级组织']).count()
lease_endagent = lease_endagent.rename(columns={'员工系统号': '期末租赁经纪人数量'})
lease_quitagent = quitagent_data[quitagent_data['LinkHR职务'].isin(['租赁经纪人', '租赁店经理'])].groupby(['七级组织']).count()
lease_quitagent = lease_quitagent.rename(columns={'员工系统号': '流失租赁经纪人数量'})

# 将数据进行合并,准备接入组织架构信息
Mdata = pd.concat(
    [startagent, endagent, quitagent, lowrank_startagent, lowrank_endagent, lowrank_quitagent, reserve_startagent,
     reserve_endagent, reserve_quitagent, bachelor_startagent, bachelor_endagent, bachelor_quitagent,
     lease_startagent, lease_endagent, lease_quitagent], axis=1, sort=False)
Mdata = Mdata.fillna(0)
Mdata['期初成熟经纪人数量'] = Mdata['期初经纪人数量'] - Mdata['期初储备人才数量']
Mdata['期末成熟经纪人数量'] = Mdata['期末经纪人数量'] - Mdata['期末储备人才数量']
Mdata['流失成熟经纪人数量'] = Mdata['流失经纪人数量'] - Mdata['流失储备人才数量']
Mdata['期初综合经纪人数量'] = Mdata['期初经纪人数量'] - Mdata['期初租赁经纪人数量']
Mdata['期末综合经纪人数量'] = Mdata['期末经纪人数量'] - Mdata['期末租赁经纪人数量']
Mdata['流失综合经纪人数量'] = Mdata['流失经纪人数量'] - Mdata['流失租赁经纪人数量']

Mdata = Mdata[['期初经纪人数量', '期末经纪人数量', '流失经纪人数量', '期初A0-A2经纪人数量', '期末A0-A2经纪人数量', '流失A0-A2经纪人数量', '期初储备人才数量', '期末储备人才数量', '流失储备人才数量',
               '期初成熟经纪人数量', '期末成熟经纪人数量', '流失成熟经纪人数量', '期初统招本经纪人数量', '期末统招本经纪人数量', '流失统招本经纪人数量',
               '期初综合经纪人数量', '期末综合经纪人数量', '流失综合经纪人数量', '期初租赁经纪人数量', '期末租赁经纪人数量', '流失租赁经纪人数量']]
Mdata['门店'] = Mdata.index

# 获取组织架构信息的数据框
org_end = endagent_data[['六级组织', '五级组织', '七级组织']]
org_quit = quitagent_data[['六级组织', '五级组织', '七级组织']]
org_start = startagent_data[['六级组织', '五级组织', '七级组织']]
organization = pd.concat([org_end, org_quit, org_start], axis=0, sort=False)
organization = organization.drop_duplicates(['七级组织'])
organization.reset_index()


def turn(a, b, c):
    t = round(a * 2 / (b + c), 7)
    return t


# 门店维度流失数据输出
Mdata['经纪人流失率'] = turn(Mdata['流失经纪人数量'], Mdata['期初经纪人数量'], Mdata['期末经纪人数量'])
Mdata['储备人才流失率'] = turn(Mdata['流失储备人才数量'], Mdata['期初储备人才数量'], Mdata['期末储备人才数量'])
Mdata['成熟经纪人流失率'] = turn(Mdata['流失成熟经纪人数量'], Mdata['期初成熟经纪人数量'], Mdata['期末成熟经纪人数量'])
Mdata['统招本经纪人流失率'] = turn(Mdata['流失统招本经纪人数量'], Mdata['期初统招本经纪人数量'], Mdata['期末统招本经纪人数量'])
Mdata['租赁经纪人流失率'] = turn(Mdata['流失租赁经纪人数量'], Mdata['期初租赁经纪人数量'], Mdata['期末租赁经纪人数量'])
Mdata['综合经纪人流失率'] = turn(Mdata['流失综合经纪人数量'], Mdata['期初综合经纪人数量'], Mdata['期末综合经纪人数量'])
Mdata['A0-A2经纪人流失率'] = turn(Mdata['流失A0-A2经纪人数量'], Mdata['期初A0-A2经纪人数量'], Mdata['期末A0-A2经纪人数量'])
Mdata = pd.merge(Mdata, organization, how='left', left_on='门店', right_on='七级组织')
Mdata = Mdata.rename(columns={'六级组织': '所属大区', '五级组织': '所属事业部'})
mcolumns = ['所属事业部', '所属大区', '门店',
            '期初经纪人数量', '期末经纪人数量', '流失经纪人数量', '经纪人流失率', '期初A0-A2经纪人数量', '期末A0-A2经纪人数量', '流失A0-A2经纪人数量', 'A0-A2经纪人流失率',
            '期初成熟经纪人数量', '期末成熟经纪人数量', '流失成熟经纪人数量', '成熟经纪人流失率', '期初统招本经纪人数量', '期末统招本经纪人数量', '流失统招本经纪人数量', '统招本经纪人流失率',
            '期初综合经纪人数量', '期末综合经纪人数量', '流失综合经纪人数量', '综合经纪人流失率', '期初租赁经纪人数量', '期末租赁经纪人数量', '流失租赁经纪人数量', '租赁经纪人流失率',
            '期初储备人才数量', '期末储备人才数量', '流失储备人才数量', '储备人才流失率']
Mdata['所属事业部'] = Mdata['所属事业部'].apply(lambda x: x[:-2])
Mdata = Mdata[mcolumns]

# 大区维度流失数据输出
organization = organization[['六级组织', '五级组织']]
organization = organization.drop_duplicates(['六级组织'])
Dataq = Mdata.groupby(['所属大区']).sum()
Dataq['大区'] = Dataq.index
Dataq = pd.merge(Dataq, organization, how='left', left_on='大区', right_on='六级组织')
Dataq = Dataq.rename(columns={'五级组织': '所属事业部'})
Dataq['经纪人流失率'] = turn(Dataq['流失经纪人数量'], Dataq['期初经纪人数量'], Dataq['期末经纪人数量'])
Dataq['储备人才流失率'] = turn(Dataq['流失储备人才数量'], Dataq['期初储备人才数量'], Dataq['期末储备人才数量'])
Dataq['成熟经纪人流失率'] = turn(Dataq['流失成熟经纪人数量'], Dataq['期初成熟经纪人数量'], Dataq['期末成熟经纪人数量'])
Dataq['统招本经纪人流失率'] = turn(Dataq['流失统招本经纪人数量'], Dataq['期初统招本经纪人数量'], Dataq['期末统招本经纪人数量'])
Dataq['租赁经纪人流失率'] = turn(Dataq['流失租赁经纪人数量'], Dataq['期初租赁经纪人数量'], Dataq['期末租赁经纪人数量'])
Dataq['综合经纪人流失率'] = turn(Dataq['流失综合经纪人数量'], Dataq['期初综合经纪人数量'], Dataq['期末综合经纪人数量'])
Dataq['A0-A2经纪人流失率'] = turn(Dataq['流失A0-A2经纪人数量'], Dataq['期初A0-A2经纪人数量'], Dataq['期末A0-A2经纪人数量'])
qcolumns =  ['所属事业部', '大区',
            '期初经纪人数量', '期末经纪人数量', '流失经纪人数量', '经纪人流失率', '期初A0-A2经纪人数量', '期末A0-A2经纪人数量', '流失A0-A2经纪人数量', 'A0-A2经纪人流失率',
            '期初成熟经纪人数量', '期末成熟经纪人数量', '流失成熟经纪人数量', '成熟经纪人流失率', '期初统招本经纪人数量', '期末统招本经纪人数量', '流失统招本经纪人数量', '统招本经纪人流失率',
            '期初综合经纪人数量', '期末综合经纪人数量', '流失综合经纪人数量', '综合经纪人流失率', '期初租赁经纪人数量', '期末租赁经纪人数量', '流失租赁经纪人数量', '租赁经纪人流失率',
            '期初储备人才数量', '期末储备人才数量', '流失储备人才数量', '储备人才流失率']
Dataq['所属事业部'] = Dataq['所属事业部'].apply(lambda x: x[:-2])
Dataq = Dataq[qcolumns]

# 事业部部维度数据输出
Datab = Dataq.groupby(['所属事业部']).sum()
Datab['事业部'] = Datab.index
Datab.loc['整体情况'] = Datab.apply(lambda x: x.sum())
Datab['经纪人流失率'] = turn(Datab['流失经纪人数量'], Datab['期初经纪人数量'], Datab['期末经纪人数量'])
Datab['储备人才流失率'] = turn(Datab['流失储备人才数量'], Datab['期初储备人才数量'], Datab['期末储备人才数量'])
Datab['成熟经纪人流失率'] = turn(Datab['流失成熟经纪人数量'], Datab['期初成熟经纪人数量'], Datab['期末成熟经纪人数量'])
Datab['统招本经纪人流失率'] = turn(Datab['流失统招本经纪人数量'], Datab['期初统招本经纪人数量'], Datab['期末统招本经纪人数量'])
Datab['租赁经纪人流失率'] = turn(Datab['流失租赁经纪人数量'], Datab['期初租赁经纪人数量'], Datab['期末租赁经纪人数量'])
Datab['综合经纪人流失率'] = turn(Datab['流失综合经纪人数量'], Datab['期初综合经纪人数量'], Datab['期末综合经纪人数量'])
Datab['A0-A2经纪人流失率'] = turn(Datab['流失A0-A2经纪人数量'], Datab['期初A0-A2经纪人数量'], Datab['期末A0-A2经纪人数量'])
Datab.loc['整体情况', '事业部'] = '整体情况'
qcolumns = ['事业部',
            '期初经纪人数量', '期末经纪人数量', '流失经纪人数量', '经纪人流失率', '期初A0-A2经纪人数量', '期末A0-A2经纪人数量', '流失A0-A2经纪人数量', 'A0-A2经纪人流失率',
            '期初成熟经纪人数量', '期末成熟经纪人数量', '流失成熟经纪人数量', '成熟经纪人流失率', '期初统招本经纪人数量', '期末统招本经纪人数量', '流失统招本经纪人数量', '统招本经纪人流失率',
            '期初综合经纪人数量', '期末综合经纪人数量', '流失综合经纪人数量', '综合经纪人流失率', '期初租赁经纪人数量', '期末租赁经纪人数量', '流失租赁经纪人数量', '租赁经纪人流失率',
            '期初储备人才数量', '期末储备人才数量', '流失储备人才数量', '储备人才流失率']
Datab = Datab[qcolumns]

# 整体情况数据输出
Dataz = Datab.rename(columns={'期末经纪人数量': '经纪人规模（期末）'})
Dataz = Dataz.loc[['整体情况'], ['经纪人规模（期末）', '流失经纪人数量', '经纪人流失率', 'A0-A2经纪人流失率',
                                 '成熟经纪人流失率', '综合经纪人流失率', '租赁经纪人流失率', '统招本经纪人流失率', '储备人才流失率']]
Dataz.loc[:, '时间'] = start+'_'+end
Dataz = Dataz.T

# 输出至Excel
writer = pd.ExcelWriter(path)
Dataz.to_excel(writer, sheet_name=Sheet[0])
Datab.to_excel(writer, sheet_name=Sheet[1], index=False)
Dataq.to_excel(writer, sheet_name=Sheet[2], index=False)
Mdata.to_excel(writer, sheet_name=Sheet[3], index=False)

writer.save()
writer.close()
