'''Program to read an SQL view into pandas, convert it to an excel file, then upload to Google Drive for use in the TMS Track n Trace app
'''
import os
import configparser

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
    # scan through the gdrive subfolder and delete any files with matching titles
    file_list = drive.ListFile({'q':"'1AhKTdDTYrGILm50sjlROQ09Be3TOfXN8' in parents and trashed=False"}).GetList()
    try:
        for f in file_list:
            if f['title'] == 'trackntracetest':
                f.Delete()                
    except:
        pass
    # create google drive file with specified file name and save to specified folder id
    file1 = drive.CreateFile({"title": "trackntracetest", 'parents': [{'id': folder_id}]})
    # set the file path to retrieve file from
    file1.SetContentFile(download_path)
    # upload file to google drive
    file1.Upload()

#sql_import()
#google_upload()

print(os.getcwd())