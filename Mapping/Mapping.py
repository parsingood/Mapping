#import pandas as pd
import pyodbc 
#import json
#import imaplib
#import base64
#import os
#import email
#import datetime
#from email.header import decode_header
#from email.message import EmailMessage
#from win32com import client as wc
#import re
#import base64
#import quopri
#from  os.path import splitext
#import io
#import csv
#import sys
#import docx2txt
#import codecs
#import win32com.client
#import docx
#import PyPDF2
#import pikepdf
#import chardet    
#import subprocess



connection_string='Driver={SQL Server Native Client 11.0};Server=PARSING01\SQLEXPRESS;Database=Parsing;Trusted_Connection=yes;'
#f = open("c:\parsing\connection_string.txt")
#connection_string =  f.read()  
#f.close

conn = pyodbc.connect(connection_string)
cursor = conn.cursor()

cursor.execute('select * from Mapping.dbo.Params')


xxx  =  [(x.ParamID, x.ParamName) for x in cursor]

ParamIDs = [x[0] for x in  xxx]
ParamNames = [x[1] for x in  xxx]
col = dict(zip(ParamNames, ParamIDs))

cursor.execute('select * from Parsing.dbo.PMS_Mapping')

#col_names = [x[0] for x in cursor.description]
#for col_name in col_names:
#    conn1 = pyodbc.connect(connection_string)
#    cursor1 = conn1.cursor()
#    cursor1.execute('insert into Mapping.dbo.Params([ParamName]) select ?',col_name )
#    cursor1.commit()
#row = cursor.fetchone()
#row_as_dict = dict(zip(col_names, row))

col_names = [x[0] for x in cursor.description]
for r in cursor:
    conn1 = pyodbc.connect(connection_string)
    cursor1 = conn1.cursor()
    cursor1.execute('insert into Mapping.dbo.PropositionSet (UpdateUserID,UpdateTime) select ?,  CURRENT_TIMESTAMP', 1)
    cursor1.execute("  select @@IDENTITY SetID ")
    recs=cursor1.fetchall()
    SetID = 0
    if len(recs) > 0:
        SetID=int(recs[0].SetID)
    cursor1.commit()

    
    row_as_dict = dict(zip(col_names, r))
    for col_name in col_names:
        conn2 = pyodbc.connect(connection_string)
        cursor2 = conn2.cursor()
        cursor2.execute('insert into Mapping.dbo.Premise(SetID,ParamID,ParamValue) select ?,?,?'
                        , SetID, col[col_name],row_as_dict[col_name] )
        cursor2.commit()
    
    #print(row_as_dict['Agent'])  # no error  row_as_dict['FolderName'] === row.FolderName












