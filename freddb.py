#!/usr/bin/python2

import MySQLdb

class FredDB(object):
    conn = False
    cursor = False
    connected = False

    def __init__(self, user, passwd, dbname):
        self.user = user
        self.passwd = passwd
        self.dbname = dbname
        self.host = '127.0.0.1'

    def connect(self):
        self.conn = MySQLdb.connect(host=self.host, port=3306,
                                    user=self.user, passwd=self.passwd,
                                    db=self.dbname)
        self.connected = True
        self.cursor = self.conn.cursor()

    def close(self):
        if self.cursor:
            self.cursor.close()
            self.cursor = False
        if self.conn:
            self.conn.close()
            self.conn = False

    def query(self, sql, data=[]):
        if not self.conn:
            self.connect()
        self.conn.query(sql, data)
        return self.conn.use_result()

    def execute(self, sql, data=[]):
        if not self.conn:
            self.connect()
        self.cursor.execute(sql, data)
        return self.cursor.fetchall()

    def executemany(self, sql, data=[]):
        if not self.conn:
            self.connect()
        self.cursor.executemany(sql, data)
        self.conn.commit()
