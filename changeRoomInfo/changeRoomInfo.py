# -*- coding:utf-8 -*-
__author__ = 'TanXing'

import xlrd
import pymssql
import time
import logging
import ConfigParser
import decimal
'''
Time:2017.12.20
  执行该脚本前，需先确认配置表字段是否正确，ServerName是否唯一，目前暂不支持大师赛和活动赛房间的配置.
Time:2017.12.21
  新增了日志模块，修复一些已知问题.
Time:2017.12.25
  优化了查询方式，执行sql前对gameroominfo进行备份.
'''

'''日志配置'''
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)   #总日志级别
logfile = 'log.txt'
fh = logging.FileHandler(logfile,mode='a')  #日志记录方式，w为覆盖记录，a为追加记录。
# fh.setLevel(logging.loglevel)
formatter = logging.Formatter('[%(asctime)s] - %(levelname)s: %(message)s ')
fh.setFormatter(formatter)
logger.addHandler(fh)

'''读取数据库信息'''
try:
    conf = ConfigParser.ConfigParser()
    conf.read("config.ini")
    server = conf.get("DatabaseInfo","HOST").strip("'")
    user = conf.get("DatabaseInfo","USER").strip("'")
    passwd = conf.get("DatabaseInfo","PASSWORD").strip("'")
    table_name = conf.get("DatabaseInfo","TABLE_NAME").strip("'")
    database = "Databaseby"
except Exception as e:
    logger.error(e)
    exit(-1)

'''读取配置表内容，以字典形式返回servername和对应字段（缺serverID）的sql'''
def exls2sql():
    BEGIN = 'UPDATE %s SET ' %(table_name)
    END = ' WHERE ServerID ='
    result_dict={}
    for i in range(2,rowscount):    #循环顺序可以优化，目前影响不大，Mark，暂不改。
        servername = sheet1.cell_value(i,0)
        tmp_list = []
        for j in range(2,colscount):
            if sheet1.cell_value(1,j):
                try:
                    usql = sheet1.cell_value(1,j) +'='+ str(int(sheet1.cell_value(i,j)))
                    tmp_list.append(usql)
                except ValueError as e:
                    logger.warn(sheet1.cell_value(i,0)+sheet1.cell_value(1,j)+u"值非法！请确认！")
                    pass
            else:pass
        tmp = BEGIN + ','.join(tmp_list) + END
        result_dict[servername] = tmp
    return result_dict

'''查询函数,根据serverName和version查询对应的serverID'''
def selectSQL():
    id_dict = {}
    logger.debug("查询sql,数据库连接中...")
    try:
        conn = pymssql.connect(server,user,passwd,database,login_timeout=5)
        cursor = conn.cursor()
        if not cursor:
            raise Exception('数据库连接数失败!')
            logger.error("数据库连接数失败!")
    except Exception as e:
        logger.error(e)
        exit(-1)
    logger.debug("查询sql,数据库已连接...")
    for i in range(2,rowscount):
        id_list = []
        servername = sheet1.cell_value(i,0)
        servername_code = servername.encode('utf8')
        version = str(sheet1.cell_value(i,1))
        sSQL = u"SELECT ServerID FROM %s WHERE ServerName LIKE '%%%s%%' and  ServerName LIKE '%%%s%%'" %(table_name,servername,version)
        try:
            cursor.execute(sSQL.encode('utf8'))
        except Exception as e:
            logger.error(e)
            exit(-1)
        serverid = cursor.fetchall()
        if serverid:    #将servername和以此查询到的所有serverID以列表形式存入字典
            for x in serverid:
                id_list.append(x[0])
                id_dict[servername]= id_list
        else:
            messg = "未查询到 %s 请手动添加！" %(servername_code)
            logger.warn(messg)
    logger.debug("查询完成，断开数据库连接！")
    return id_dict
    cursor.close()
    conn.close()

'''执行修改房间配置的sql'''
def execsql(usql):
    conn = pymssql.connect(server, user, password=passwd, database=database,login_timeout=5)
    logger.debug("执行sql,数据库连接中...")
    cursor = conn.cursor()
    if not cursor:
        raise Exception('数据库连接数失败!')
    logger.debug("执行sql,数据库已连接...")
    # timestamp = time.strftime("%Y%m%d%H%M%S", time.localtime()) #复制sql前先备份表,只能复制数据,不能复制索引,主键等信息
    # copySQL = "select * into %s_%s from gameroominfo_copy" %(table_name,timestamp)
    # cursor.execute(copySQL)
    # logger.debug("GameroomInfo_copy已备份.")
    try:
        for i in usql:
            cursor.execute(i)
    except Exception as e:
        logger.error(e)
        exit(-1)
    conn.commit()   #提交事务
    logger.debug("执行完成，断开数据库连接！")
    cursor.close()
    conn.close()

def main():
    sql_list = []
    name_sql = exls2sql()   #获取servername和sql的字典
    name_id = selectSQL()   #获取servername和ID的字典
    logger.debug("准备生成UPDATE语句...")
    for k in name_sql:
        id_list = name_id.get(k)
        if not id_list:
            pass
        else:
            for i in range(len(id_list)):
                d = name_sql[k] + str(id_list[i])
                sql_list.append(d)
                logger.info(k+":"+d)
    logger.debug("UPDATE语句已生成完毕，准备执行sql.")
    execsql(sql_list)

if __name__ == '__main__':
    '''读取房间配置表'''
    start = time.time()
    try:
        workbook = xlrd.open_workbook(r'room.xlsx')
    except IOError as e:
        # print "房间配置表是否存在？名称是否正确？"
        logger.error(e)
        exit(-1)
    sheet1 = workbook.sheet_by_index(0)
    rowscount = sheet1.nrows
    colscount = sheet1.ncols
    main()
    end = time.time()
    cost_time = "执行完成，共耗时%.2f秒！" %(end - start)
    logger.info(cost_time)