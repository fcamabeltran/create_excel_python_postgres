#!/usr/bin/python
import psycopg2
import sys
import pprint
#LIBRARIES FOR WORBOOK
from openpyxl import Workbook
import os

"""rows = (
            ('cabecera','data','asunto'),
            (88, 46, 57),
            (89, 38, 12),
            (23, 59, 78),
            (56, 21, 98),
            (24, 18, 43),
            (34, 15, 67)
)"""
def create_excel(LIST_SQL,HEAD_EXCEL,NAME):
    #-----------SET THE PARAMETERS-----------
    #QUERY_RESULT=LIST_SQL.insert(0,HEAD_EXCEL)
    QUERY_RESULT=LIST_SQL
    HEAD_XLS=HEAD_EXCEL
    NAME_FILE=NAME
    #-----------PREPARE THE DATA-----------
    #if the data is None  as change NONE for EMPTY
    for item in QUERY_RESULT:
        [item_x==' 'for item_x in item if (item_x is None)]

    QUERY_RESULT.insert(0,HEAD_EXCEL)
    ROWS=tuple(list(QUERY_RESULT))
    BOOK = Workbook()
    SHEET = BOOK.active
    for row in ROWS:
        SHEET.append(row)
    #-----------CREATE THE FILE(XLS)-----------
    FILE=NAME_FILE +'.xlsx'
    BOOK.save(FILE)
    RUTA=os.getcwd()
    RUTA_FILE=RUTA+'/'+FILE
    print("The file was generated in this---> %s" % RUTA_FILE)
    return RUTA_FILE;

def main_postgresql():
    conn_string = "host='127.0.0.1' dbname='BD_TIWS' user='postgres' password='postgres'"
    # print the connection string we will use to connect
    print ("Connecting to database\n	->%s" % (conn_string))
    # get a connection, if a connect cannot be made an exception will be raised here
    conn = psycopg2.connect(conn_string)
    # conn.cursor will return a cursor object, you can use this cursor to perform queries
    cursor = conn.cursor()
    # execute our Query
    cursor.execute('SELECT * FROM  tiws_pais limit 50')
    # retrieve the renamecords from the database
    records = cursor.fetchall()

    cabecera=('numero01','numero02','numero03','numero04','numero05','numero06','numero07','numero09','numero10')
    #records.insert(0,cabecera)
    name='ejemplo'
    # Send the data for the def with 3 arguments
    create_excel(records, cabecera,name)

if __name__ == "__main__":
    main_postgresql()
    #create_excel(0,'demostracion')
