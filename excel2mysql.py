#!/usr/bin/env python
# encoding: utf-8
'''
@Author: Joven Chu
@Email: jovenchu@gmail.com
@File: excel2mysql.py
@Time: 2019-12-27 14:19
@Project: Audit_report
@About: 将表格转换为mysql数据库存储
'''

import pymysql
import xlrd

def newconnectToMysql(sql):
    """
    连接数据库并新建数据库和表
    :param sql: 新建数据表语句
    :return: db,cur游标
    """
    # 与数据库建立连接
    db = pymysql.connect(host='127.0.0.1', user='root', password='123456', database='audit_report',
                         port=3306, charset='utf8')
    # 创建游标链接
    cur = db.cursor()

    # 新建一个database
    # cur.execute("drop database if exists audit_report")
    # cur.execute("create database audit_report")
    # cur.execute("use audit_report")

    # 如果存在audit_problem这个表则删除
    cur.execute("drop table if exists audit_problem")
    # 创建表
    cur.execute(sql)
    # 返回游标
    return db,cur




def importExcelToMysql(cur, path, sql):
    # 读取excel中内容到数据库
    num = 1
    # 读取excel文件
    workbook = xlrd.open_workbook(path)
    # 获得第一张表
    worksheet = workbook.sheet_by_index(0)
    # 获取表的行数
    row_num = worksheet.nrows
    # 第一行是标题名，对应表中的字段名所以应该从第二行开始，计算机以0开始计数，所以值是1
    for i in range(1, row_num):
        sqlstr = worksheet.row_values(i)
        valuestr = [str(sqlstr[0]), int(sqlstr[1]), str(sqlstr[2]), str(sqlstr[3]), str(sqlstr[4]), str(sqlstr[5]),
                    str(sqlstr[6])]
        # valuestr = tuple(valuestr)
        # 将每行数据存到数据库中
        cur.execute(sql, valuestr)


def readTable(cursor,sql):
    """
    输出数据库中的内容
    :param cursor:
    :return:
    """
    # 选择全部
    cursor.execute(sql)
    # 获得返回值， 返回多条记录， 若没有结果则返回
    results = cursor.fetchall()
    for i in range(0, results.__len__()):
        for j in range(0, 7):
            print(results[i][j], end='\t')
    print('\n')


def closeMysql(db,cur):
    """
    关闭数据库
    :param db:
    :param cur:
    :return:
    """
    # 关闭游标链接
    cur.close()
    db.commit()
    # 关闭数据库服务连接, 释放内存
    db.close()


if __name__ == '__main__':
    # 与数据库建立连接，新建游标、数据库、数据表
    sql = "CREATE TABLE  audit_problem(num INT,time INT,problem_title VARCHAR (100),problem_description VARCHAR (1000),sanctions VARCHAR (100),influence_level VARCHAR (10),card VARCHAR (10))"
    db,cur = newconnectToMysql(sql)


    # 将excel中的数据导入数据库中
    path = "audit_report.xlsx"
    sql2 = "insert into audit_problem(num, time, problem_title, problem_description, sanctions, influence_level, card) VALUES(%s, %s, %s, %s, %s, %s, %s)"
    importExcelToMysql(cur, path, sql2)
    sql3 = "select * from audit_problem"
    readTable(cur, sql3)

    # 关闭数据库
    closeMysql(db,cur)