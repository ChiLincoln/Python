import sqlite3
import xlrd
import configparser
import os
from datetime import datetime
from xlrd import xldate_as_tuple

def CreateIni(path):  # 创建INI
    config = configparser.ConfigParser()
    if os.path.exists(path):
        return True
    else:
        with open(path, 'w') as configfile:
            config.write(configfile)

        return True


def readstr(path, unitstr, fieldstr):  # 读取INI文件

    if CreateIni(path):
        config = configparser.ConfigParser()
        config.sections()
        config.read(path,encoding="utf-8-sig")
        # config.read(path, encoding="gbk")
        if config[unitstr][fieldstr]:
            return config[unitstr][fieldstr]
        else:
            return None


def bulidconnection(path):
    '''建立连接'''
    print(path)
    try:
        return sqlite3.connect(path)
    except Exception as e:
        print(e)
        return None


def UpdateData(con: sqlite3.Connection, sqlstr):
    # print(sqlstr)
    if con:
        try:
            con.execute(sqlstr)
            con.commit()
            return True
        except Exception as e:
            print(e)
            return False


def DirTraval(dirpath):  # 遍历文件
    filenamelist = []
    for filename in os.listdir(dirpath):
        # 只遍历XLS和XLSX
        if filename[-4:] == 'xlsx' or filename[-4:] == '.xls' or filename[-4:] == '.csv':
            path = os.path.join(dirpath, filename)
            filenamelist.append(path)
    print(filenamelist)
    return filenamelist


def InsertDataMany(con: sqlite3.Connection, sqlstr, params):
    # sqlstr = "insert into ProviceData_AllConfirm  (BeiJing) values (?)"
    #
    # params = [("2"),("3")]
    if con:
        try:
            con.executemany(sqlstr, params)
            con.commit()
            return True
        except Exception as e:
            print(e)
            return False


def writeIntosqlietthread(file):
    # 加载EXCEL
    book = xlrd.open_workbook(file)
    # 激活第一个sheet
    sheet = book.sheet_by_index(0)

    params = []

    for i in range(3, sheet.nrows - 5):  # 遍历Excel
        list1 = []
        # 遍历Excel的列
        # print(sheet.cell(i, num).value)
        for num in range(30):
            ctype = sheet.cell(i, num).ctype
            cell = sheet.cell_value(i, num)
            if ctype == 3:
                date = datetime(*xldate_as_tuple(cell, 0))
                cell = date.strftime('%Y-%m-%d')
                list1.append(cell)
            else:
                list1.append(str(sheet.cell(i, num).value))

        tuple1 = tuple(list1)
        # print(tuple1)
        params.append(tuple1)

    ## 插入数据
    sqlstr = "insert into ImportData(商户号, 交易日期 ,交易时间 ,交易类型 ,卡号 ,交易金额 , 终端号 , 清算金额 , 手续费 , 参考号 ,流水号 ,商户名称 ,卡类型 ,商户订单号 ,支付类型 ,银商订单号 ,退货订单号 ,实际支付金额 , 备注 ,付款附言 ,钱包优惠金额 ,商户优惠金额 ,发卡行 ,分店简称 ,其他优惠金额,分期期数 ,分期手续费 ,分期服务方 ,分期付息方 ,子订单号) values (" \
             "?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
    conn = bulidconnection(sqlitepath)
    if InsertDataMany(conn, sqlstr, params):
        print("程序运行success!")


def WriteIntoSqlite(dirpath, sqlitepath):
    # 删除旧数据
    conn = bulidconnection(sqlitepath)
    sqlstr = "delete from  ImportData"
    UpdateData(conn, sqlstr)

    # 遍历下载文件夹
    filenamelist = DirTraval(dirpath)

    if len(filenamelist) != 0:
        for file in filenamelist:
            # print(file)
            # tfile=threading.Thread(target=writeIntosqlietthread,args=(file,), name="run")
            writeIntosqlietthread(file)
            # tfile.start()
            # tfile.join()


# 加载INI文件
path = "setting.ini"
# 读取EXCEL文件路径
srcpath = readstr(path, "YWX", "srcdir")
# 读取数据库路径
sqlitepath = readstr(path, "YWX", "sqlitepath")

WriteIntoSqlite(srcpath, sqlitepath)