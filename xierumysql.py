#!/usr/bin/python3
# encoding='utf-8'
# author:weibk
# @time:2021/10/9 14:44

from DBUtils import update
import xlrd

wd = xlrd.open_workbook(filename=r'2020年每个月的销售情况.xlsx', encoding_override=True)
names = wd.sheet_names()

# 读取数据
for i in names:
    table = wd.sheet_by_name(i)
    for j in range(1, table.nrows):
        data = table.row_values(j)
        print(data)
        # 写入数据库
        update(f'insert into market '
               f'values (%s, %s, %s, %s, %s)', (i + data[0], data[1],
                                                data[2], data[3], data[4]))
