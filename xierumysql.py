#!/usr/bin/python3
# encoding='utf-8'
# author:weibk
# @time:2021/10/9 14:44
import xlrd
from DBUtils import update

wd = xlrd.open_workbook(filename=r'2020年每个月的销售情况.xlsx', encoding_override=True)
names = wd.sheet_names()

# 读取数据
for i in names:
    table = wd.sheet_by_name(i)
    for j in range(1, table.nrows):
        data = table.row_values(j)
        print(data)
        # 写入数据库
        update(f'insert into market{names.index(i)+1} '
               f'values (%s, %s, %s, %s, %s)', data)
