#!/usr/bin/python3
# encoding='utf-8'
# author:weibk
# @time:2021/10/9 10:34

import xlrd

wd = xlrd.open_workbook(filename=r'2020年每个月的销售情况.xlsx', encoding_override=True)
# 所有工作表的名字列表
# ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月']
names = wd.sheet_names()
# 1.全年销售总额
total = 0
# 循环每个工作表
for i in names:
    table = wd.sheet_by_name(i)
    print(table.nrows)
    # 每个工作表的销售量列
    count = table.col_values(-1)
    # 每个工作表的单价列
    price = table.col_values(-3)
    sum1 = 0
    for j in range(1, table.nrows):
        sum1 += count[j] * price[j]
    total += sum1
print(f'全年销售总额：{total}')

# 2.每种衣服儿销售（件数）占比 每件衣服12个月的总销售量/所有衣服12个月总销售量
# 所有衣服12个月总销售量
print('-' * 60)
sales = 0
for i in names:
    table = wd.sheet_by_name(i)
    count = table.col_values(-1)
    sum1 = 0
    for j in range(1, len(count)):
        sum1 += count[j]
    sales += sum1
print(f'全年销售总量：{sales}')

# 计算每种衣服的年销售量
# 一共有多少种衣服
allname = set()
for i in names:
    table = wd.sheet_by_name(i)
    monthname = table.col_values(1)
    monthname = set(monthname[1:])
    for j in range(len(names)):
        allname = allname | monthname
allname = list(allname)
print('所有衣服种类', allname)

# 计算每种衣服的总销售量
for i in allname:
    sum1 = 0
    for j in names:
        table = wd.sheet_by_name(j)
        for k in range(1, table.nrows):
            row = table.row_values(k)
            if row[1] == i:
                sum1 += row[-1]
    result = sum1 * 100 / sales
    print(f'{i}的销售占比为：{round(result, 2)}%')

# 3.每件衣服的月销售占比
print('-' * 60)
for i in names:
    print(f'{i}份的每种衣服销售占比：')
    table = wd.sheet_by_name(i)
    # 每个月内的衣服种类
    clusname = list(set(table.col_values(1)))[1:]
    monthprice = table.col_values(-3)[1:]
    monthcount = table.col_values(-1)[1:]
    # 月销售总额
    monthtotal = 0
    for x in range(len(monthprice)):
        monthtotal += monthcount[x] * monthprice[x]
    # 当月每种衣服的销售额
    for j in clusname:
        sum1 = 0
        for k in range(1, table.nrows):
            if j == table.row_values(k)[1]:
                sum1 += table.row_values(k)[-1] * table.row_values(k)[-3]
        result = sum1 * 100 / monthtotal
        print(f'\t{j}的销售额占比为：{round(result, 2)}%')

# 4.每种衣服的销售额占比
# 每种衣服的销售额
for i in allname:
    sum2 = 0
    for j in names:
        table = wd.sheet_by_name(j)
        for k in range(1, table.nrows):
            if i == table.row_values(k)[1]:
                sum2 += table.row_values(k)[-1] * table.row_values(k)[-3]
    result = round(sum2 * 100 / total, 2)
    print(f'{i}的销售额占比为：{result}%')

# 5.最畅销的衣服
print('-' * 60)
changxiao = {}
for i in allname:
    counts = 0
    for j in names:
        table = wd.sheet_by_name(j)
        for k in range(1, table.nrows):
            if table.row_values(k)[1] == i:
                counts += table.row_values(k)[-1]
    changxiao[i] = counts
result = max(changxiao, key=changxiao.get)
print('最畅销的衣服是：', result)
result = min(changxiao, key=changxiao.get)
print('全年销量最低的衣服是：', result)

# 6.每个季度最畅销的衣服
print('-'*60)
name1 = [['3月', '4月', '5月'], ['6月', '7月', '8月'], ['9月', '10月', '11月'],
         ['12月', '1月', '2月']]
for m in name1:
    for i in allname:
        counts = 0
        for j in m:
            table = wd.sheet_by_name(j)
            for k in range(1, table.nrows):
                if table.row_values(k)[1] == i:
                    counts += table.row_values(k)[-1]
        changxiao[i] = counts
    result = max(changxiao, key=changxiao.get)
    print(f'第{name1.index(m)+1}季度最畅销的衣服是：', result)