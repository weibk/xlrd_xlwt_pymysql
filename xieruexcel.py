#!/usr/bin/python3
# encoding='utf-8'
# author:weibk
# @time:2021/10/9 17:35
import xlwt
import pymysql

connect = pymysql.connect(host="localhost", user="root", password="123456",
                          database="company")
cursor = connect.cursor()

workbook = xlwt.Workbook(encoding="utf-8")


def write_e(sheetname="", table=""):
    worksheet = workbook.add_sheet(sheetname)

    # 获取数据
    cursor.execute(f'select * from {table}')
    rows = list(cursor.fetchall())
    column = list(cursor.description)
    # 写入表头
    for i in range(len(column)):
        worksheet.write(0, i, column[i][0])

    # 写入数据
    for i in rows:
        for j in range(len(i)):
            worksheet.write(rows.index(i) + 1, j, i[j])

write_e('t_employees', 't_employees')
write_e('t_dept', 't_dept')
workbook.save('sanguo.xls')