"""

Author : Olivier Lefebre

This file creates loads for P2P

Last update : 2020-01-07
By : Nicolas Raymond

"""



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


def savexlsxFile(wb, path, filename, Time =False, extension='.xlsx'):

    """
    Saves xlsx files

    :param wb: workbook
    :param path: path where the file is going to be save (leave '' if no path)
    :param filename: Name of the document
    :param Time: (bool) True to write the creation time in the file name
    :param extension: extension of the file
    :return: complete path of the file saved
    """
    
    if Time is True:
        tN = pd.datetime.now()
        tempSave = '_'+str(tN.year)+'_'+str(tN.month)+'_'+str(tN.day)+';'+str(tN.hour)+\
                   '.'+str(tN.minute)+'.'+str(tN.second)
    else:
        tempSave = ''
    ErrorMessage = ''
    
    saved = False
    nameExtension = ''
    while not saved:
        try:
            saved = True
            wb.save(path+filename+nameExtension+tempSave+extension)
        except:
            msg = "Someone already has the file "+filename+nameExtension+extension+" open"
            ErrorMessage += msg
            saved = False
            nameExtension += " Copy"
    return path+filename+nameExtension+tempSave+extension


def send_email(recipient, subject, text, listfile=[], CC=[]):
    """
    Sends email

    :param recipient: list of email addresses that will receive the email
    :param subject: subject of the email
    :param text: text contained in the email
    :param listfile: list with directories of the files to send with email
    :param CC: list of email adresses in CC

    Example :
        send_email(["felix.theroux-joncas@brp.com"], "Subject", "text", ["C:/Users/test.txt", "C:/Users/test2.xlsx"])
    """
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
        part = MIMEBase('application', "octet-stream")
        part.set_payload(open(pathfile, "rb").read())
        encoders.encode_base64(part)
        if '/' in pathfile:
            file = pathfile.split("/")[-1]
        else:
            file = pathfile.split('\\')[-1]
        part.add_header('Content-Disposition', 'attachment', filename=file)
        msg.attach(part)  
    try: 
        smtpObj = smtplib.SMTP_SSL('smtp-cavl.brp.local', 465)
        smtpObj.ehlo()
        smtpObj.login(email, password)
        smtpObj.sendmail("OrdertoDelivery@brp.com", recipient+CC, msg.as_string())
        smtpObj.close()
    except:
        print("Sending email failed")


def timeSinceLastCall(functionName="", ToPrint=True):
    """
    Print the time since the last call, gives an indication of the time required by each function

    :param functionName: name of the last executed function
    :param ToPrint: (bool) True to print the time required
    """

    global timeReference
    if ToPrint:
        print(functionName+' : '+str((datetime.datetime.now()-timeReference).total_seconds())+" sec")
    timeReference = datetime.datetime.now()


def weekdays(dayAfter, startDay=0, officialDay=''):
    """
    Returns the date in (dayAfter) week days from today's date

    :param dayAfter: Number of days from today
    :param startDay: First day to start from now (0 for today, -1 for yesterday, 1 for tomorrow,...)
    :param officialDay: date to consider as today in format '%Y-%m-%d'
    """

    if officialDay == '':
        today = datetime.date.today() + datetime.timedelta(startDay)
    else:
        today = datetime.datetime.strptime(officialDay, '%Y-%m-%d').date() + datetime.timedelta(startDay)
    wkday = today.weekday()

    # if today's date is a weekend day
    if wkday > 4:
        today = datetime.date.today()+datetime.timedelta(startDay+7-wkday)
    dayConversion = dayAfter % 5
    week = (dayAfter - dayConversion)/5
    date2 = today+datetime.timedelta(7*week + dayConversion)
    if date2.weekday() > 4 or date2.weekday() < dayConversion:
        date2 = date2+datetime.timedelta(2)
    deltmonth = ''  # else we have a problem with SQL date format
    delday=''
    if date2.month < 10:
        deltmonth = '0'
    if date2.day < 10:
        delday = '0'
    return str(date2.year)+'-'+deltmonth+str(date2.month)+'-'+delday+str(date2.day)


def xlsxToTxtFile(readPath, filexlsx, writePath, txtfileName):
    """
    Writes excel data in a text document

    :param readPath: path where is the file we want to read
    :param filexlsx: name of the file with .xlsx extension included
    :param writePath: path where the new text file will be saved
    :param txtfileName: name of the new text file
    """

    global ErrorMessage
    wb = openpyxl.load_workbook(readPath+filexlsx)
    ws = wb.active
    try:
        txt_file = open(writePath+txtfileName, 'w')
        data = ""
        for rows in ws:
            line = ""
            for cells in rows: 
                line += str(cells.value)
                line += "\t"
            data += (line[:-1]+"\n")
        txt_file.write(data[:-1])
        txt_file.close()
    except:
        ErrorMessage += "Error in xlsxToTxtFile function"


class SQLConnection:

    """Connection to SQL servers"""

    def __init__(self, server, DataBase, table, headers=''):
        self.server = server
        self.DataBase = DataBase
        self.table = table
        self.conn = pyodbc.connect('Driver={SQL Server};'
                                   'Server='+server+';'
                                   'Database='+DataBase+';'
                                   'Trusted_Connection=yes;')
        self.cursor = cursor = self.conn.cursor()
        self.headers = headers

    def deleteFromSQL(self, conditions):
        """
        Deletes data from the sql table
        :param conditions: sql condition that indicates data to delete  (Example : NAMEOFCOLUMN < 0)
        """
    
        if conditions == '':
            application = ''
        else:
            application = ''' WHERE ''' + conditions
        query = ''' DELETE FROM ''' + self.table + application

        self.cursor.execute(query)
        self.conn.commit()

    def GetSQLData(self, SQLquery):
        """
        Gets data from the sql table
        :param SQLquery: sql query to get the data
        :return: list of lists containing the data
        """
        
        if SQLquery != """ """:
            self.cursor.execute(SQLquery)
            result = self.cursor.fetchall()
            resultList = [list(i) for i in result]
            return resultList
        else:  # Else returns an empty list
            return[]

    def UpdateSQL(self, RowsToChange, conditions):
        """
        Modifies data FROM a SQL table

        :param RowsToChange: Rows to change ( ex: LAST_CHANGED_DATE = GETDATE())
        :param conditions: Conditions to update DATA (ex: "country = 'CA' AND shipping_POINT <>'4110'")
        """
    
        query = ''' UPDATE '''+self.table+''' SET '''+RowsToChange+''' WHERE ''' + conditions

        self.cursor.execute(query)
        self.conn.commit()

    def sendToSQL(self, values, headers='', headersCount=-1):
        """

        :param values: list of lists of values
        :param headers:
        :param headersCount: number of columns in the table, it counts the number of comma in headers +1
        :return:
        """
        if headers == '':
            headers = self.headers
        if headersCount == -1:  
            headersCount = headers.count(',')
        valuesToInsert = '?'+headersCount*',?'   # expected number of values to insert (ex: 4 columns = '(?,?,?,?)')
    
        query = ''' INSERT INTO '''+self.table+''' ('''+headers+''')VALUES ('''+valuesToInsert+''')'''

        self.cursor.executemany(query, values)
        self.conn.commit()
