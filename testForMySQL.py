import datetime

import pandas as pd
import pymysql

# 打开数据库连接
import pyodbc
def uploaddata(insert_name, insert_type,insert_data,insert_shift,insert_yield):
    # 打开数据库连接
    _db = pymysql.connect(host='localhost',
                         user='user1',
                         password='ruanjianjishu',
                         database='test')
    # 使用 cursor() 方法创建一个游标对象 cursor
    _cursor = _db.cursor()
    _insert_name = "'" + insert_name + "'"
    _insert_type = "'" + insert_type + "'"
    _insert_data = "'" + insert_data + "'"
    _insert_shift = "'" + insert_shift + "'"
    _insert_yield = insert_yield

    _cursor.execute(
        "INSERT INTO yield_all(name,type,date,shift,yield)VALUES(" + _insert_name + "," + _insert_type + "," + _insert_data + "," + _insert_shift + "," + _insert_yield + ")")
    _db.commit()
    # 关闭数据库连接
    _db.close()
# 数据库连接
cnxn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=./DataSend03.accdb')
crsr = cnxn.cursor()
crsr = cnxn.cursor()
timenow = (datetime.datetime.now() + datetime.timedelta(days=-0)).strftime("%Y-%m-%d %H:%M:%S")
timelast = (datetime.datetime.now() + datetime.timedelta(days=-1)).strftime("%Y-%m-%d %H:%M:%S")
crsr.execute(
    "SELECT * FROM 产量2 WHERE 机种 = '" +
                            "7302-1-FML" + "'")
list0 = crsr.fetchall()
if list0 != []:
    list_Data = []
    list_Data_x = []
    list_Data_x1 = []

    for i in range(len(list0)):
        list_Data.append(list0[i][3])
        list_Data_x.append(list0[i][2])
        list_Data_x1.append(i)
        uploaddata("7302-2-FML","7302","2021-12-14 00:00:00","白班","10000")
    d = {'a': list_Data_x,
         'b': list_Data}
    df = pd.DataFrame(d)
    print(df)
    df.drop_duplicates(subset=['a'], keep='first', inplace=True)
    ts = pd.Series(df['b'].tolist(), index=df['a'])
    ts_10T = ts.resample('1T').bfill()

# uploaddata("7302-1-FML","7302","2021-12-14 00:00:00","白班","10000")
# 写入单次数据
# 打开数据库连接
_db = pymysql.connect(host='localhost',
                     user='user1',
                     password='ruanjianjishu',
                     database='test')
# 使用 cursor() 方法创建一个游标对象 cursor
_cursor = _db.cursor()

_cursor.execute(
   "SELECT * FROM yield_all")
list0 = _cursor.fetchall()
print(list0)
# 关闭数据库连接
_db.close()
