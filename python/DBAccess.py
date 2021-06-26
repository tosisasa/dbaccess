# 2020/03/11 新規作成

import cx_Oracle

def connectDB(schema):
    if schema == "ユーザ名@ホスト名":
        host = "ホスト名"
        port = 1521
        sid = "orcl"
        user = "ユーザ名"
        password = "パスワード"
 

    tns = cx_Oracle.makedsn(host, port, sid)
    conn = cx_Oracle.connect(user, password, tns)

    return conn

def closeDB(conn, isCommit):
    if isCommit == True:
        conn.commit()
    else:
        conn.rollback()

    conn.close()


def selectDB(sql, schema):
    conn = connectDB(schema)

    cur = conn.cursor()
    cur.execute(sql)
    print(sql) #sql文を標準出力
    rows = cur.fetchall()

    cur.close()

    closeDB(conn, True)

    return rows
