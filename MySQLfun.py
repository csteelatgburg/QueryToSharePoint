import http.client
import json
import platform
import getpass
import urllib.request
import urllib.parse
import sys
import pymysql.cursors
from datetime import date, datetime, timedelta
import os
from dateutil import tz

from config import *

dateTimeTypes = [7, 10, 11, 12, 13, 14]
numberTypes = [0, 1, 2, 4, 5, 16, 246, 247]
integerTypes = [3, 8, 9]

def MySQLQuery(query):
    mysqlconn = pymysql.connect(user=config['MySQL_user'],
                                        password=config['MySQL_pass'],
                                        host=config['MySQL_server'],
                                        database=config['MySQL_db'],
                                        cursorclass=pymysql.cursors.DictCursor)
    cursor = mysqlconn.cursor()
    cursor.execute(query)
    result = cursor.fetchall()
    cursor.close()
    mysqlconn.close()
    return result

def MySQLColumnTypes(queryfile):
    fh = open("queries/" + queryfile, "r")
    sqlQuery = fh.read()
    mysqlconn = pymysql.connect(user=config['MySQL_user'],
                                        password=config['MySQL_pass'],
                                        host=config['MySQL_server'],
                                        database=config['MySQL_db'],
                                        cursorclass=pymysql.cursors.DictCursor)
    cursor = mysqlconn.cursor()
    cursor.execute(sqlQuery)
    result = cursor.fetchone()
    types = cursor.description
    cursor.close()
    mysqlconn.close()
    return types
