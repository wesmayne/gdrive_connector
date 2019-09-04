'''Program to read an SQL view into pandas, convert it to an excel file, then upload to 
Google Drive for use in the TMS Track n Trace app
'''
import os
import configparser

import smtplib
import pandas as pd
import sqlalchemy
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive


# global variables read from config file
config = configparser.ConfigParser()
config.read('config.ini')
# read google upload path from config
download_path = config['DOWNLOAD']['download_path']
folder_id = config['DOWNLOAD']['folder_id']
# read database details from config
user = config['DATABASE']['user']
password = config['DATABASE']['password']
servername = config['DATABASE']['servername']
database = config['DATABASE']['database']
query_orders = config['DATABASE']['query_orders']
query_products = config['DATABASE']['query_products']
# read email details from config
sender_email = config['EMAIL']['sender_email']
receiver_email = config['EMAIL']['receiver_email']
email_password = config['EMAIL']['email_password']


def sql_import():
    connection_string = f'mssql+pymssql://{user}:{password}@{servername}/{database}'
    engine = sqlalchemy.create_engine(connection_string)
    conn = engine.connect()
    # create a pandas dataframe from each query
    orders = pd.read_sql(query_orders, conn)
    products = pd.read_sql(query_products, conn)
    # combine the dataframes into a single excel worbook
    with pd.ExcelWriter(download_path) as writer:  # doctest: +SKIP
        orders.to_excel(writer, sheet_name='Shipments', index = None)
        products.to_excel(writer, sheet_name='Products', index = None)

def google_upload():
    g_login = GoogleAuth()
    g_login.LocalWebserverAuth()
    # Create GoogleDrive instance with authenticated GoogleAuth instance.
    drive = GoogleDrive(g_login)

    # create google drive file with specified file name and matching doc id and save to specified folder id
    file1 = drive.CreateFile({"id":"1-_C8UttiVf46wF3krBHWTL1Q3eENnLei" , "title": "trackntrace", 'parents': [{'id': folder_id}]})
    # set the file path to retrieve file from
    file1.SetContentFile(download_path)
    # upload file to google drive
    file1.Upload()

def error_email(ex):
    message = ('''Subject: TracknTrace App - Data Import Error

    ***AUTOMATED EMAIL***\n\nData import to Track n' Trace app has failed.\n\nError Code:\n\n'''+str(ex))       

    mailserver = smtplib.SMTP('smtp.office365.com',587)
    mailserver.ehlo()
    mailserver.starttls()
    mailserver.login(sender_email, email_password)
    mailserver.sendmail(sender_email,receiver_email, message)
    mailserver.quit()
    
    
try:
    sql_import()
    google_upload()
except Exception as ex:
    error_email(ex)
