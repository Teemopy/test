import pandas as pd
import numpy as np

read_path = r'E:\2020小石墩\基尼系数\202002.csv'
data = pd.read_csv(read_path, engine='python', encoding='utf-8-sig')


def gini_coef(wealths):
    cum_wealths = np.cumsum(sorted(np.append(wealths, 0)))
    sum_wealths = cum_wealths[-1]
    xarray = np.array(range(0, len(cum_wealths))) / np.float(len(cum_wealths)-1)
    yarray = cum_wealths / sum_wealths
    B = np.trapz(yarray, x=xarray)
    A = 0.5 - B
    return A / (A+B)


data.loc[data['业绩'] < 0, '业绩'] = 0
data.loc[data['收入'] < 0, '收入'] = 0
print('业绩基尼系数为：', gini_coef(data['业绩']), '\n收入基尼系数为：', gini_coef(data['收入']))





