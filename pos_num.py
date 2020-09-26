# -*- coding: utf-8 -*-
"""
Created on Wed Jan 22 08:57:38 2020

@author: sayers
"""

import pandas as pd
import re
from src.admin import newest,colclean
from src.cleansheet import dl_clean as dl
from src.subset import subsetlist as sl
from src.emailautosend import mailer

path = "S:\\Downloads\\"     # Give the location of the files
fname = "CU_R_POSTN_STATUS"         # Give filename prefix
df = pd.read_excel(newest(path,fname))  #getting the newest of these files in the directory and converting to df
#stripping out the 2 metadata columns in CJR files
if re.match("Position Status Report ",df.columns.values[0]).group() == "Position Status Report ":
    new_header = df.iloc[1] #grab the first row for the header
    df = df[2:] #take the data less the header row
    df.columns = new_header #set the header row as the df header
#standardizing the column names
df=colclean(df)
#filtering for only active positions and removing unnecessary columns
collist=[['deptid','deptid_description','position_#','position_#_description',
   'position_effective_date','position_active_/_inactive','position_status',
   'position_full/part','reports_to_name','reports_to','pay_serv_position','budget_line_#']]
status_report = sl(df,[['vacant_/_filled','Vacant'],['position_status','A']],str1=collist)
nfname = 'S:\\Downloads\\vacant_positions.xlsx'
#saving to a file and cleaning presentation
dl(nfname,status_report)
#sending confirmation to end-users
hrisgroup = "'lolsson@york.cuny.edu';'pcaceres901@york.cuny.edu';'ajackson1@york.cuny.edu';'mwilliams@york.cuny.edu';'lwilkinson901@york.cuny.edu'"
body="The Vacant Position Report has been refreshed. You can find the most recent copy on the shared HR drive."
subj="Successful Vacant Position Update"
mailer(hrisgroup,'',subj,body)