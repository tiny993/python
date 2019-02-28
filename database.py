#!/usr/bin/env python
# -*- coding: utf-8 -*-
""" Operate database """

from enum import Enum, unique


@unique
class DatabaseMethod(Enum):
    Oracle = 1
    Mysql = 2
    PostgreSql = 3


class Database(object):
    def __init__(self, database):
        self._conn = None
        self._database = database

    def close(self):
        if self._conn is not None:
            try:
                self._conn.close()
            except Exception as e:
                print("Close database Error: ", e)

    def connect(self, user, pwd, sid, ip, port='1521'):
        # Oracle
        if self._database is DatabaseMethod.Oracle:
            import cx_Oracle
            try:
                conn_str = r'%s/%s@%s:%s/%s' % (user, pwd, ip, port, sid)
                self._conn = cx_Oracle.connect(conn_str)
            except Exception as e:
                print("Oracle Error: ", e)

        # MySql
        if self._database is DatabaseMethod.Mysql:
            import mysql.connector
            try:
                self._conn = mysql.connector.connect(host=ip, port=port, user=user, password=pwd, database=sid)
            except Exception as e:
                print("MySql Error: ", e)

        # PostgreSql
        if self._database is DatabaseMethod.PostgreSql:
            import psycopg2
            try:
                self._conn = psycopg2.connect(host=ip, port=port, user=user, password=pwd, database=sid)
            except Exception as e:
                print('PostgreSql Error: ', e)

    def query(self, sql_str):
        if self._conn is not None:
            try:
                cur = self._conn.cursor()
                res = cur.execute(sql_str)
                display = res.fetchall()
                cur.close()
                return display

            except Exception as e:
                print('Error: ', e)
        else:
            print('Call connect() to connect database first.')

    def insert(self, sql_str):
        cur = self._conn.cursor()
        if self._conn is not None:
            try:
                if isinstance(sql_str, list):
                    for x in sql_str:
                        cur.execute(x)
                else:
                    cur.execute(sql_str)

                cur.close()
                self._conn.commit()

            except Exception as e:
                print('Error: ', e)
        else:
            print('Call connect() to connect database first.')
