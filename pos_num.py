# -*- coding: utf-8 -*-
"""
Created on Wed Jan 22 08:57:38 2020

@author: sayers
"""
#from emailautosend import mailthis
#from emailautosend import getemail
import os
import pandas as pd
import re
import win32com.client as win32
from cleansheet import *

def newest(path,fname):
    files = os.listdir(path)
    paths = [os.path.join(path, basename) for basename in files if basename.startswith(fname)]
    return max(paths, key=os.path.getmtime)

path = "S:\\Downloads\\"     # Give the location of the files
fname = "CU_R_POSTN_STATUS"         # Give filename prefix
df = pd.read_excel(newest(path,fname))  #getting the newest of these files in the directory and converting to df
#stripping out the 2 metadata columns in CJR files
if re.match("Position Status Report ",df.columns.values[0]).group() == "Position Status Report ":
    new_header = df.iloc[1] #grab the first row for the header
    df = df[2:] #take the data less the header row
    df.columns = new_header #set the header row as the df header
#standardizing the column names
df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_').str.replace('(', '').str.replace(')', '')
list(df.columns.values.tolist()) 



status_report = df[(df['vacant_/_filled']=='Vacant')&(df['position_status']=='A')][['deptid',
   'deptid_description','position_#','position_#_description',
   'position_effective_date','position_active_/_inactive','position_status',
   'position_full/part','reports_to_name','reports_to','pay_serv_position','budget_line_#']]
nfname = 'S:\\Downloads\\vacant_positions.xlsx'


cleansheet(status_report,nfname)

hrisgroup = "'lolsson@york.cuny.edu';'pcaceres901@york.cuny.edu';'ajackson1@york.cuny.edu';'mwilliams@york.cuny.edu';'lwilkinson901@york.cuny.edu'"

def reportmail():
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = hrisgroup
    mail.Subject = "Successful Vacant Position Update"
    
       
    text = """
    Good Day,
    
    The Vacant Position Report has been refreshed. You can find the most recent copy on the shared HR drive.    


    Best Regards,
    Shane Ayers
    Acting Human Resources Information Systems Manager
    Office of Human Resources
    York College
    The City University of New York"""
    
    html = """
    <html>
    <head>
    <style>     
     table, th, td {{ border: 1px solid black; border-collapse: collapse; }}
      th, td {{ padding: 10px; }}
    </style>
    </head>
    <body><p>Good Day,</p>
    <p>The Vacant Position Report has been refreshed. You can find the most recent copy on the shared HR drive </p>
    <p></p>
    <p>Best Regards,</p>
    <p>Shane Ayers</p>
    <p>Acting Human Resources Information Systems Manager</p>
    <p>Office of Human Resources</p>
    <p>York College</p>
    <p>The City University of New York</p>
    </body></html>
    """
    
    # above line took every col inside csv as list
    mail.Body = text
    mail.HTMLBody = html
    mail.Send()
#reportmail()
