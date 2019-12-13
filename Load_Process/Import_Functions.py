import openpyxl
import os 
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from email import encoders
from email.mime.base import MIMEBase
import math
import sqlite3
import datetime
from copy import copy,deepcopy
from openpyxl.utils import range_boundaries
from openpyxl import Workbook
import inspect
import pandas as pd
import pyodbc

timeReference = datetime.datetime.now()


#*****************************************************
#NAME: savexlsxFile
#
#DESCRIPTION: Save xlsx files
#
#PARAMETERS:   	-wb                     excel workbook
#               -path                   leave '' if no path  
#               -filename               Name of the document 
#               -Time                   True to write the creation time in the file name
#               -extension              '.xlsx'
#*****************************************************
def savexlsxFile(wb, path,filename ,Time =False, extension = '.xlsx'):
    
    if Time ==True:
        tN=pd.datetime.now()
        tempSave='_'+str(tN.year)+'_'+str(tN.month)+'_'+str(tN.day)+';'+str(tN.hour)+'.'+str(tN.minute)+'.'+str(tN.second)
    else:
        tempSave=''
    ErrorMessage=''
    
    saved=False
    nameExtension=''
    while(not saved):  
        try:
            saved = True
            wb.save(path+filename+nameExtension+tempSave+extension)
        except:
            msg = "Someone already has the file " +filename+nameExtension+extension+" open" 
            ErrorMessage+=msg
            saved = False
            nameExtension+=" Copy"
    return path+filename+nameExtension+tempSave+extension


#*****************************************************
#NAME: send_email
#
#DESCRIPTION: Sends emails
#
#PARAMETERS:   	-recipient
#               -subject
#               -text
#               -listfile           Directory of file to send with email
#               -CC
#*****************************************************
def send_email(recipient, subject, text,listfile = [], CC =[] ):
#Example: send_email(["felix.theroux-joncas@brp.com","mireille.larouche-mailhot@brp.com"], "Subject", "text", ["C:/Users/test.txt", "C:/Users/test2.xlsx"])
    email = "SVC_CAVL_SMTP_OTD"
    password = "!983IIdj111" 
    msg = MIMEMultipart()
    msg['From'] = "OrdertoDelivery@brp.com"
    msg['To'] = ", ".join(recipient)
    msg['CC'] = ", ".join(CC)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject
    msg.attach(MIMEText(text))

    # Attach all file in listfile
    for pathfile in listfile:
        part = MIMEBase('application',"octet-stream")
        part.set_payload(open(pathfile, "rb").read())
        encoders.encode_base64(part)
        if('/' in pathfile):
            file = pathfile.split("/")[-1]
        else:
            file = pathfile.split('\\')[-1]
        part.add_header('Content-Disposition','attachment', filename=file)
        msg.attach(part)  
    try: 
        smtpObj = smtplib.SMTP_SSL('smtp-cavl.brp.local',465)
        smtpObj.ehlo()
        smtpObj.login(email, password)
        smtpObj.sendmail("OrdertoDelivery@brp.com", recipient+CC, msg.as_string())
        smtpObj.close()
    except:
        print("Sending email failed")




#*****************************************************
#NAME: timeSinceLastCall
#
#DESCRIPTION: Print the time since the last call, gives an indication of the time required by each function
#
#PARAMETERS:   	-functionName           name of the last executed function
#               -ToPrint                True to print the time required 
#*****************************************************
def timeSinceLastCall(functionName = "",ToPrint=True): 
    """Print the time since the last call, gives an indication of the time required by each function
    """
    global timeReference
    if ToPrint:
        print(functionName +' : ' +str((datetime.datetime.now()-timeReference).total_seconds())+ " sec")
    timeReference = datetime.datetime.now()




#*****************************************************
#NAME: weekdays
#
#DESCRIPTION: Returns the date in (dayAfter) week days from today's date
#
#PARAMETERS:   	-dayAfter               #Number of days from today
#               -startDay               #First day to start from now (0 for today, -1 for yesterday, 1 for tomorrow,...)

#*****************************************************
def weekdays(dayAfter,startDay=0,officialDay=''):#returns the date in X week days
    if officialDay=='':
        today = datetime.date.today() + datetime.timedelta(startDay)
    else:
        today = datetime.datetime.strptime(officialDay, '%Y-%m-%d').date() + datetime.timedelta(startDay)
    wkday = today.weekday()
    #if today's date is a weekend day
    if wkday>4:
        today = datetime.date.today()+datetime.timedelta(startDay+7-wkday)
    dayConversion = dayAfter%5
    week = (dayAfter - dayConversion) /5
    date2= today+datetime.timedelta(7*week + dayConversion)
    if date2.weekday() >4 or date2.weekday()<dayConversion:
        date2= date2+datetime.timedelta(2)
    deltmonth='' #else we have a problem with SQL date format
    delday=''
    if date2.month < 10:
        deltmonth='0'
    if date2.day < 10:
        delday='0'
    return str(date2.year)+'-' +deltmonth+str(date2.month)+'-' +delday+str(date2.day)


#*****************************************************
#NAME: xlsxToTxtFile
#
#DESCRIPTION: writes excel data in a text document
#
#PARAMETERS:   	-readPath
#               -filexlsx
#               -writePath
#               -txtfileName
#*****************************************************
def xlsxToTxtFile(readPath, filexlsx,writePath, txtfileName):
    global ErrorMessage
    wb = openpyxl.load_workbook(readPath+filexlsx)
    ws = wb.active
    try:
        txt_file = open(writePath+txtfileName,'w')
        data =""
        for rows in ws:
            line =""
            for cells in rows: 
                line+=xstr(cells.value)
                line+="\t"
            data+=(line[:-1]+"\n")
        txt_file.write(data[:-1])
        txt_file.close()
    except:
        ErrorMessage+="Error in xlsxToTxtFile function"


#*****************************************************
#NAME: SQLConnection
#
#DESCRIPTION: Connection to SQL servers
#
#PARAMETERS:   	-server
#               -DataBase
#               -table
#*****************************************************
class SQLConnection:
    "Connection to SQL servers"

    server = ''
    DataBase = ''
    table = ''
    conn = ''
    cursor = ''
    headers=''


    def __init__(self,server, DataBase, table,headers=''):
        self.server = server
        self.DataBase = DataBase
        self.table = table
        self.conn = pyodbc.connect('Driver={SQL Server};'
                      'Server='+server+ ';'
                      'Database='+DataBase+';'
                      'Trusted_Connection=yes;')
        self.cursor = cursor = self.conn.cursor()
        self.headers=headers

    def deleteFromSQL(self,conditions):  #Envoit du data dans  une table SQL
        "Delete DATA from SQL"
    
        if conditions=='':
            application=''
        else:
            application=''' WHERE ''' + conditions
        query=''' DELETE FROM '''+ self.table +application

        self.cursor.execute(query)
        self.conn.commit()

    def GetSQLData(self,SQLquery):
        "Get data from a SQL DataBase"   
        
        if SQLquery != """ """:
            self.cursor.execute(SQLquery)
            result = self.cursor.fetchall()
            resultList = [list(i) for i in result] #Else returns a tuple
            return resultList
        else:
            return[]

    #*****************************************************
    #               -RowsToChange       Rows to change ( ex: LAST_CHANGED_DATE = GETDATE())
    #               -conditions         Conditions to update DATA (ex: "country = 'CA' AND shipping_POINT <>'4110'")
    #*****************************************************
    def UpdateSQL(self,RowsToChange,conditions): 
        "Modify data FROM a SQL table"
    
        query='''
                    UPDATE '''+ table +''' SET '''+ RowsToChange +''' WHERE ''' + conditions

        self.cursor.execute(query)
        self.conn.commit()

    #*****************************************************
    #               -headersCount           number of columns in the table, it counts the number of comma in headers +1
    #*****************************************************
    def sendToSQL(self, values, headers='', headersCount = -1 ): 
        if headers =='':
            headers = self.headers
        if headersCount == -1:  
            headersCount = headers.count(',')
        valuesToInsert = '?'+headersCount*',?'        #expected number of values to insert (ex: 4 columns = '(?,?,?,?)'    )
    
        query='''
                    INSERT INTO '''+self.table+ ''' ('''+ headers+''')
                    VALUES ('''+valuesToInsert+ ''')
                    '''


        self.cursor.executemany(query, values)
        self.conn.commit()
