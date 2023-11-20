import itertools
import pandas as pd
import numpy as np
from openpyxl import load_workbook
import datetime
from functools import reduce
from numpy import inf
import glob
import os.path
import xlsxwriter
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
pd.options.mode.chained_assignment = None
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)
import win32com.client as win32
import pywintypes as py
from pywintypes import com_error

pd.set_option('display.float_format','{:.2f}'.format)
pd.set_option('precision', 2)


w1= int(input("Week no.:"))

year_data=2021

path=r'C:/Users/nkcho/Desktop/POE Dashboard'

path1=r'C:\Users\nkcho\Documents\Python\usn'

#master_data import

login_data= pd.read_excel(path+'/Master_data.xlsx', sheet_name='Login_data')
login_data=pd.DataFrame(login_data)

target_data=pd.read_excel(path+'/Master_data.xlsx', sheet_name='Target')
target_data=pd.DataFrame(target_data)

#importing data

#importing data from tracker folder
filenames = glob.glob(path +'/Tracker/'+ "\*.xlsx")
print('File names:', filenames)

# Initializing empty data frame
finalexcelsheet = pd.DataFrame()

# to iterate excel file one by one 
# inside the folder
for file in filenames:
    df = pd.concat(pd.read_excel(file, sheet_name=None),
                   ignore_index=True, sort=False)
    finalexcelsheet = finalexcelsheet.append(
    df, ignore_index=True)

finalexcelsheet.to_csv(path+'/Tracker/Merged_tracker.csv')

filename1=path+'/Tracker/Merged_tracker.csv'

filename2=path+'/Connections.xlsx'


tracker_final=pd.read_csv(filename1)
tracker_final= pd.DataFrame(tracker_final)

tracker_final.loc[:, "Leave"].fillna("NO", inplace=True)

tracker_final.loc[:,"Year"]= pd.to_datetime(tracker_final.loc[:, "Date"]).dt.year

tracker_final= tracker_final.drop(tracker_final[(tracker_final['Year']!= year_data)].index)

tracker_final.loc[:,'Date']=pd.to_datetime(tracker_final.loc[:,'Date']).dt.date

tracker_final=tracker_final.rename(columns={'User Login': 'Login'})

tracker_final= tracker_final.drop(columns=['Unnamed: 0'])

tracker_final.loc[:, "End Date"]= pd.to_datetime(tracker_final.loc[:,"End Date"]).dt.date

tracker_final.loc[:, "Login"].fillna(0, inplace=True)

tracker_final= tracker_final.drop(tracker_final[(tracker_final['Login']==0)].index)


#leave split

tracker_final.loc[:, "Date"]= pd.to_datetime(tracker_final.loc[:,"Date"]).dt.date

tracker_final.loc[:, "End Date"]= pd.to_datetime(tracker_final.loc[:,"End Date"]).dt.date

tracker_final.loc[:, "Date"]= tracker_final.loc[:, "Date"].astype('str')

tracker_final.loc[:, "ID"]= tracker_final.loc[:, "Date"]+tracker_final.loc[:, "Login"]



tracker_final.fillna(0, inplace=True)

tracker_final.loc[:, "End Date"]=[x if y==0 and z==0 else w for x,y,z,w in zip(
tracker_final.loc[:, "Date"], tracker_final.loc[:, "End Date"], tracker_final.loc[:, "Leave Hours"], tracker_final.loc[:, "End Date"])]

tracker_final.loc[:, "End Date"]= [x if y==0 else z for x,y,z in zip(
tracker_final.loc[:, "Date"], tracker_final.loc[:, "End Date"], tracker_final.loc[:, "End Date"])]

 

def datesplit(df):
    tracker_final = df.rename(columns={'Date':'sdate','End Date':'edate', 'ID':'ID'})
    return  (pd.concat([pd.Series(r.ID,pd.date_range(r.sdate, r.edate)) 
                        for r in tracker_final.itertuples()])
                .rename_axis('Date_1')
                .reset_index())

df = datesplit(tracker_final)
print (df)

df['WEEKDAY'] = pd.to_datetime(df['Date_1']).dt.dayofweek
df.loc[:, "Date_1"]= pd.to_datetime(df.loc[:,"Date_1"]).dt.date

index_names = df[ (df['WEEKDAY'] >=5)].index
df.drop(index_names, inplace = True)
df= df.rename(columns={0:'ID'})

df= df.drop(columns=['WEEKDAY'])

df= df.rename(columns={'Date_1':'Start Date'})

df.loc[:, "End Date"]= df.loc[:, "Start Date"]


tracker_final= pd.merge(df,tracker_final, how='left', on='ID')

tracker_final= tracker_final.rename(columns={'End Date_x': 'End Date'})

tracker_final= tracker_final.drop(columns=['End Date_y', 'Date', 'ID'])

tracker_final= tracker_final.rename(columns={'Start Date':'Date'})

tracker_final.loc[:,"Week"]= pd.to_datetime(tracker_final.loc[:, "Date"]).dt.week

#droppping rows where login is not present
tracker_final= tracker_final.drop(tracker_final[(tracker_final['Login']==0)].index)

tracker_final[['Reporting Manager', 'Program']].fillna("NA", inplace=True)

tracker_final.fillna(0, inplace=True)

convert_dict = {'Login':'category','Reporting Manager':'category','Program':'category','Leave Type': 'category','Leave Hours': 'int8','Instances':'category',
                'Send Email': 'category','Bincheck EPCR Hours': 'float16','Bincheck USN Hours':'float16','Bincheck IXD Hours':'float16',
                'Bincheck CCR Hours':'float16','CCR EN':'float16','CCR Non EN':'float16','Brand Owner Audits':'float16',
                'Non AVP Audits EN':'float16','Non AVP Audits Non EN':'float16',
                'EIM FP Count':'int8','EIM FP Hours':'float16','EIM FP QA Count':'float16','EIM FP QA Hours':'float16',
                'EIM Esc Count':'float16','EIM Esc QA Count':'float16','EIM Esc Hours':'float16',
                'EIM Esc QA Hours':'float16','EPCR EN':'float16',
                'EPCR Non EN':'float16','EACP Hours':'float16','EACP Appeals Hours':'float16',
                'MCM EN':'float16','MCM Non EN':'float16',
                'MCM Live EN QA':'float16','MCM Live Non EN QA':'float16','MCM Practice Files Count':'float16','MCM Practice File hours':'float16',
                'Safety Sensitivity Hours':'float16','Safety Sensitivity Count':'int16',
                'Test Buy Audits counts':'int16','MCM Others NE':'float16','MCM DE by Non DE resources':'float16',
                'Test Buy Hrs':'float16','USN EN':'float16',
                'USN Non EN':'float16','DCR KAI EN(Live)':'float16','DCR KAI Non EN (Live)':'float16',
                'DCR Live EN QA':'float16','DCR Live Non EN QA':'float16',
                'DCR Practice File Count':'float16','DCR Practice File hours':'float16',
                'WI KAI EN (Live)':'float16','WI KAI Non EN (Live)':'float16',
                'WI Live EN QA':'float16','WI Live Non EN QA':'float16','WI Practice File Count':'float16',
                'WI Practice File Hours':'float16','SLIM Hours':'float16','PAD Hours':'float16',
                'Coverage Hours':'float16','Retails Hours':'float16','Backlog & Ops Update':'float16','HTMS Data Reporting':'float16','Flash':'float16','Keyword Analysis':'float16',
                'MOM':'float16','Ops Tracker':'float16','Other Reportings':'float16',
                'PR Doc':'float16','Productivity Report':'float16','WBR Report':'float16',
                'Work Allocation':'float16','EN QA Hours':'float16','QA Reporting':'float16',
                'Non EN QA Hrs':'float16','QA Calibration':'float16',
                'Project Hours':'float16','Other Tasks Hours':'float16',
                'Meeting':'float16','No Work':'float16','Other NPTS':'float16','System Issues':'float16',
                'Training':'float16','RAMP Hours':'float16','Total Hours':'float16',
                'Productive Hours':'float16','NPT SUM':'float16','Leave Sum':'float16',
                'Quality Hours':'float16','CSM Hours':'float16','NPT Hours':'float16', 
                'Deep Dive Hours':'float16','Deep Dive Name':'category',
                'Adhoc-Business Name':'category','Adhoc-Business Hours':'float16',
                'Adhoc-Internal Name':'category','Adhoc-Internal Hours':'float16',
                'DCR Total':'float16','EACP Appeals Count':'float16',
                'WI Total':'float16','Week':'int8','Year':'int16','Date':'category','End Date':'category'}
  
tracker_final = tracker_final.astype(convert_dict)

#new dataframe to merge data from both tracker and connections one by one

tracker_final=tracker_final[['Login','Reporting Manager', 'Program', 'Year','Week','Date',
                             'End Date','Leave','Leave Type','Leave Hours',
                             'Instances','Bincheck EPCR Hours','Bincheck USN Hours',
                             'Bincheck CCR Hours','Bincheck IXD Hours',
                             'CCR EN','CCR Non EN','Brand Owner Audits','Non AVP Audits EN','Non AVP Audits Non EN',
                             'EIM FP Hours','EIM FP QA Count','EIM FP QA Hours','EIM Esc QA Count',
                             'EIM Esc Hours','EIM Esc QA Hours','EPCR EN','EPCR Non EN','EACP Hours','EACP Appeals Hours',
                             'MCM EN','MCM Non EN','MCM Live EN QA','MCM Live Non EN QA',
                             'MCM DE by Non DE resources','MCM Others NE','MCM Practice File hours',
                             'Safety Sensitivity Hours','Test Buy Hrs','USN EN','USN Non EN','DCR KAI EN(Live)',
                             'DCR KAI Non EN (Live)','DCR Live EN QA','DCR Live Non EN QA','DCR Practice File hours',
                             'WI KAI EN (Live)','WI KAI Non EN (Live)','WI Live EN QA',
                             'WI Live Non EN QA','WI Practice File Hours','SLIM Hours','PAD Hours','Coverage Hours',
                             'Retails Hours','Backlog & Ops Update','HTMS Data Reporting','Flash','Keyword Analysis',
                             'MOM','Ops Tracker','Other Reportings','PR Doc','Productivity Report',
                             'WBR Report','Work Allocation','EN QA Hours','QA Reporting',
                             'Non EN QA Hrs','QA Calibration','Deep Dive Hours','Deep Dive Name',
                             'Project Name','Project Hours',
                             'Other Task Name','Other Tasks Hours','Meeting',
                             'No Work','Other NPTS','System Issues','Training','RAMP Hours','Quality Hours',
                             'Audits', 'QA Meeting','Adhoc-Business Name', 'Adhoc-Business Hours','Adhoc-Internal Name', 'Adhoc-Internal Hours',
                             'Comments','Productive Hours','NPT SUM','NPT Hours','Leave Sum','Total Hours']]

options = ['akanssum','anspriya','birudp','boppudi','chandaky','haritann','hkoyiri',
           'jarshiy','khliqmo','kishya','kotkows','maithily','mmuner','otteeste',
           'pbbedi','pirans','portabp','saibezaw','sjjml','sreehb','sunishka','vshaldv',
           'yavanab']


# selecting rows based on condition
tracker_final = tracker_final[tracker_final['Login'].isin(options)]

tracker_final.drop_duplicates(subset=['Date', 'Login'], inplace=True)

#columns not present in initial sharepoint site
#deep dive hours, deep dive name
#adhoc based columns

#project

#en qa hours was present in old SP

#column names to be changed either in old SP or new SP
#ixd count and ixd hours

#changed name of below columns in new Sharepoint
#eacp hours
#dcr practice file hours
#WI practice file hours and count
#other reporting
#system issues
#productive hours
#ramp hours
#Test Buy audit hours
#training
#bincheck ccr and usn hours
#eim esc hours and count
#eim fp hours and count
#non avp audits non en



tracker_final.loc[:,'Date']=tracker_final.loc[:,'Date'].dt.date

tracker_final.loc[:,'End Date']=tracker_final.loc[:,'End Date'].dt.date

#collecting data from sharepoint connections in the "connections sheet"



#Bincheck

bincheck= pd.read_excel(filename2, sheet_name='Bincheck_Data')
bincheck= pd.DataFrame(bincheck)

bincheck=bincheck[['Date', 'Login', 'Bincheck EPCR Count', 'Bincheck USN Count','Bincheck CCR Count', 'Bincheck IXD Count']]

bincheck.loc[:,'Date']=bincheck.loc[:,'Date'].dt.date

bincheck[ 'Bincheck EPCR Count'].fillna(0, inplace=True)
bincheck['Bincheck USN Count'].fillna(0, inplace=True)

bincheck.drop_duplicates(subset=['Date', 'Login'], inplace=True)

convert_dict = {'Login':'category','Bincheck EPCR Count':'float16','Bincheck USN Count': 'float16', 
                'Bincheck IXD Count':'float16','Bincheck CCR Count':'float16'}

bincheck = bincheck.astype(convert_dict)

tracker_final= pd.merge(tracker_final, bincheck, how='left', on=['Login', 'Date'])

tracker_final=pd.merge(tracker_final, target_data[['Week', 'Bincheck Target/hr']], on=['Week'], how='left')

tracker_final[ 'Bincheck EPCR Count'].fillna(0, inplace=True)
tracker_final['Bincheck USN Count'].fillna(0, inplace=True)
tracker_final['Bincheck IXD Count'].fillna(0, inplace=True)
tracker_final['Bincheck CCR Count'].fillna(0, inplace=True)
        
convert_dict = {'Bincheck Target/hr':'float16'}
  
tracker_final = tracker_final.astype(convert_dict)

tracker_final.loc[:, "temp1"]= [0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:,'Bincheck EPCR Hours'],tracker_final.loc[:,'Bincheck Target/hr'],
                                                                                      (tracker_final.loc[:, "Bincheck EPCR Count"]/tracker_final.loc[:, "Bincheck Target/hr"]))]

tracker_final.loc[:, "temp2"]= [0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:,'Bincheck CCR Hours'],tracker_final.loc[:,'Bincheck Target/hr'],
                                                                                      (tracker_final.loc[:, "Bincheck CCR Count"]/tracker_final.loc[:, "Bincheck Target/hr"]))]
tracker_final.loc[:, "temp3"]= [0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:,'Bincheck USN Hours'],tracker_final.loc[:,'Bincheck Target/hr'],
                                                                                      (tracker_final.loc[:, "Bincheck USN Count"]/tracker_final.loc[:, "Bincheck Target/hr"]))]
tracker_final.loc[:, "temp4"]= [0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:,'Bincheck IXD Hours'],tracker_final.loc[:,'Bincheck Target/hr'],
                                                                                      (tracker_final.loc[:, "Bincheck IXD Count"]/tracker_final.loc[:, "Bincheck Target/hr"]))]

tracker_final.loc[:, "Achieved Bincheck Weighted"]=tracker_final.loc[:, "temp1"]+tracker_final.loc[:, "temp2"]+tracker_final.loc[:, "temp3"]+tracker_final.loc[:, "temp4"]

tracker_final= tracker_final.drop(columns=['temp1', 'temp2', 'temp3','temp4'])

tracker_final.loc[:, "temp1"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "Bincheck EPCR Hours"], tracker_final.loc[:, "Bincheck EPCR Count"])]

tracker_final.loc[:, "temp2"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "Bincheck USN Hours"], tracker_final.loc[:, "Bincheck USN Count"])]

tracker_final.loc[:, "temp3"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "Bincheck IXD Hours"], tracker_final.loc[:, "Bincheck IXD Count"])]

tracker_final.loc[:, "temp4"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "Bincheck CCR Hours"], tracker_final.loc[:, "Bincheck CCR Count"])]

tracker_final.loc[:, "Bincheck Total Achieved"]= tracker_final.loc[:, "temp1"]+tracker_final.loc[:, "temp2"]+tracker_final.loc[:, "temp3"]+tracker_final.loc[:, "temp4"]

tracker_final= tracker_final.drop(columns=['temp1', 'temp2', 'temp3','temp4'])

tracker_final.loc[:, "temp1"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "Bincheck EPCR Hours"], (tracker_final.loc[:,"Bincheck EPCR Hours"]*tracker_final.loc[:, "Bincheck Target/hr"]))]

tracker_final.loc[:, "temp2"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "Bincheck USN Hours"], (tracker_final.loc[:,"Bincheck USN Hours"]*tracker_final.loc[:, "Bincheck Target/hr"]))]

tracker_final.loc[:, "temp3"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "Bincheck IXD Hours"], (tracker_final.loc[:,"Bincheck IXD Hours"]*tracker_final.loc[:, "Bincheck Target/hr"]))]

tracker_final.loc[:, "temp4"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "Bincheck CCR Hours"], (tracker_final.loc[:,"Bincheck CCR Hours"]*tracker_final.loc[:, "Bincheck Target/hr"]))]


tracker_final.loc[:, "Bincheck Total Target"]= tracker_final.loc[:, "temp1"]+tracker_final.loc[:, "temp2"]+tracker_final.loc[:, "temp3"]+tracker_final.loc[:, "temp4"]

tracker_final= tracker_final.drop(columns=['temp1', 'temp2', 'temp3','temp4'])

tracker_final.loc[:, "Bincheck %"]=[0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "Bincheck Total Achieved"],tracker_final.loc[:, "Bincheck Total Target"],
                                                                              (tracker_final.loc[:, "Bincheck Total Achieved"]/tracker_final.loc[:, "Bincheck Total Target"]))]
tracker_final['Bincheck %'].fillna(0, inplace=True)

tracker_final['Bincheck %']=tracker_final['Bincheck %']*100


tracker_final=tracker_final[['Login','Reporting Manager', 'Program', 'Year','Week','Date', 'End Date','Leave','Leave Type','Leave Hours',
                         'Instances','Bincheck EPCR Hours','Bincheck EPCR Count','Bincheck USN Hours','Bincheck USN Count',
                         'Bincheck IXD Hours','Bincheck IXD Count',
                         'Bincheck CCR Hours','Bincheck CCR Count','Bincheck Target/hr',
                         'Achieved Bincheck Weighted','Bincheck Total Achieved',
                         'Bincheck Total Target','Bincheck %','CCR EN','CCR Non EN','Brand Owner Audits','Non AVP Audits EN','Non AVP Audits Non EN',
                         'EIM FP Hours','EIM FP QA Count','EIM FP QA Hours','EIM Esc QA Count',
                         'EIM Esc Hours','EIM Esc QA Hours','EPCR EN','EPCR Non EN','EACP Hours','EACP Appeals Hours',
                         'MCM EN','MCM Non EN','MCM Live EN QA','MCM Live Non EN QA',
                         'MCM DE by Non DE resources','MCM Others NE','MCM Practice File hours',
                         'Safety Sensitivity Hours','Test Buy Hrs','USN EN','USN Non EN','DCR KAI EN(Live)',
                         'DCR KAI Non EN (Live)','DCR Live EN QA','DCR Live Non EN QA','DCR Practice File hours','WI KAI EN (Live)','WI KAI Non EN (Live)','WI Live EN QA',
                         'WI Live Non EN QA','WI Practice File Hours','SLIM Hours','PAD Hours','Coverage Hours','Retails Hours',
                         'Backlog & Ops Update','HTMS Data Reporting','Flash','Keyword Analysis',
                         'MOM','Ops Tracker','Other Reportings','PR Doc','Productivity Report',
                         'WBR Report','Work Allocation','EN QA Hours','QA Reporting',
                         'Non EN QA Hrs','QA Calibration','Deep Dive Hours','Deep Dive Name',
                         'Project Name','Project Hours',
                         'Other Task Name','Other Tasks Hours','Meeting',
                         'No Work','Other NPTS','System Issues','Training','RAMP Hours','Quality Hours',
                         'Audits', 'QA Meeting','Adhoc-Business Name', 'Adhoc-Business Hours','Adhoc-Internal Name', 'Adhoc-Internal Hours',
                         'Comments','Productive Hours','NPT SUM','NPT Hours','Leave Sum','Total Hours']]


tracker_final= tracker_final.round({"Achieved Bincheck Weighted":2, "Bincheck %":2})

convert_dict = {'Achieved Bincheck Weighted':'float32',
                'Bincheck Total Achieved':'float32',
                'Bincheck Total Target':'float32',
                'Bincheck %':'float32'}
  
tracker_final = tracker_final.astype(convert_dict)

tracker_final['Achieved Bincheck Weighted'] = tracker_final['Achieved Bincheck Weighted'].round(2)
tracker_final['Bincheck Total Achieved'] = tracker_final['Bincheck Total Achieved'].round(2)
tracker_final['Bincheck Total Target'] = tracker_final['Bincheck Total Target'].round(2)
tracker_final['Bincheck %'] = tracker_final['Bincheck %'].round(2)

bincheck=''
    

#CCR

ccr= pd.read_excel(filename2, sheet_name='Tagged_Data_Workdocs')
ccr=pd.DataFrame(ccr)

ccr=ccr[['Date', 'Login', 'CCR EN Count', 'CCR Non EN Count']]

ccr.loc[:,'Date']=pd.to_datetime(ccr.loc[:,'Date']).dt.date

ccr['CCR EN Count'].fillna(0, inplace=True)
ccr['CCR Non EN Count'].fillna(0, inplace=True)

ccr.drop_duplicates(subset=['Date', 'Login'], inplace=True)

convert_dict = {'Login':'category','CCR EN Count':'float16',
            'CCR Non EN Count': 'float16'}
  
ccr = ccr.astype(convert_dict)

#merging ccr count from connections sheets to new_data

tracker_final=pd.merge(tracker_final, ccr, how='left', on=['Login', 'Date'])

tracker_final['CCR EN Count'].fillna(0, inplace=True)
tracker_final['CCR Non EN Count'].fillna(0, inplace=True)

#getting BO Audits count from connections sheet
bo= pd.read_excel(filename2, sheet_name='BOA_Data')
bo=pd.DataFrame(bo)

bo=bo[['Login', 'Date', 'Brand Owner Audits Count']]

bo.loc[:,'Date']=pd.to_datetime(bo.loc[:,'Date']).dt.date

bo['Brand Owner Audits Count'].fillna(0, inplace=True)

bo.drop_duplicates(subset=['Date', 'Login'], inplace=True)

convert_dict = {'Login':'category','Brand Owner Audits Count':'float16'}
  
bo = bo.astype(convert_dict)

#merging BO Audits data to new_data

tracker_final=pd.merge(tracker_final, bo, how='left', on=['Login', 'Date'])

tracker_final['Brand Owner Audits Count'].fillna(0, inplace=True)

#getting non avp data from connections
non_avp=pd.read_excel(filename2, sheet_name='NON_AVP_Data')
non_avp=pd.DataFrame(non_avp)

non_avp=non_avp[['Login', 'Date', 'Non AVP Audits EN Count', 'Non AVP Audits Non EN Count']]

non_avp.loc[:,'Date']=pd.to_datetime(non_avp.loc[:,'Date']).dt.date

non_avp['Non AVP Audits EN Count'].fillna(0, inplace=True)
non_avp['Non AVP Audits Non EN Count'].fillna(0, inplace=True)

non_avp.drop_duplicates(subset=['Date', 'Login'], inplace=True)

convert_dict = {'Login':'category','Non AVP Audits EN Count':'float16',
            'Non AVP Audits Non EN Count': 'float16'}
  
non_avp = non_avp.astype(convert_dict)


#merge non_avp data with new_data

tracker_final= pd.merge(tracker_final, non_avp, how='left', on=['Login', 'Date'])


tracker_final['Non AVP Audits EN Count'].fillna(0, inplace=True)
tracker_final['Non AVP Audits Non EN Count'].fillna(0, inplace=True)

#getting target data for ccr, bo and non_avp

tracker_final= pd.merge(tracker_final, target_data[['Week', 'CCR Target/hr', 'BOA NA IN Target/hr',
                                      'NON AVP Target/hr']], how='left', on=['Week'])

convert_dict = {'Week':'int8','CCR Target/hr':'float16',
            'NON AVP Target/hr': 'float16','BOA NA IN Target/hr':'float16'}
  
tracker_final = tracker_final.astype(convert_dict)


tracker_final.loc[:, "temp1"]= [0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "CCR EN"],
                                                                  tracker_final.loc[:, "CCR Target/hr"], (tracker_final.loc[:, "CCR EN Count"]/tracker_final.loc[:, "CCR Target/hr"]))]
                       
tracker_final.loc[:, "temp2"]=[0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "CCR Non EN"],
                                                                  tracker_final.loc[:, "CCR Target/hr"], (tracker_final.loc[:, "CCR Non EN Count"]/tracker_final.loc[:, "CCR Target/hr"]))]

tracker_final.loc[:, "temp3"]=[0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "Brand Owner Audits"],
                                                                  tracker_final.loc[:, "BOA NA IN Target/hr"], (tracker_final.loc[:, "Brand Owner Audits Count"]/tracker_final.loc[:, "BOA NA IN Target/hr"]))]

tracker_final.loc[:, "temp4"]=[0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "Non AVP Audits EN"],
                                                                  tracker_final.loc[:, "NON AVP Target/hr"], (tracker_final.loc[:, "Non AVP Audits EN Count"]/tracker_final.loc[:, "NON AVP Target/hr"]))]

tracker_final.loc[:, "temp5"]=[0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "Non AVP Audits Non EN"],
                                                                  tracker_final.loc[:, "NON AVP Target/hr"], (tracker_final.loc[:, "Non AVP Audits Non EN Count"]/tracker_final.loc[:, "NON AVP Target/hr"]))]

tracker_final.loc[:, "Achieved CCR Weighted"]= tracker_final.loc[:, "temp1"]+tracker_final.loc[:, "temp2"]+tracker_final.loc[:, "temp3"]+tracker_final.loc[:, "temp4"]+tracker_final.loc[:, "temp5"]

tracker_final= tracker_final.drop(columns=['temp1', 'temp2', 'temp3','temp4', 'temp5'])

tracker_final.loc[:, "temp1"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "CCR EN"], tracker_final.loc[:, "CCR EN Count"])]

tracker_final.loc[:, "temp2"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "CCR Non EN"], tracker_final.loc[:, "CCR Non EN Count"])]

tracker_final.loc[:, "temp3"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "Brand Owner Audits"], tracker_final.loc[:, "Brand Owner Audits Count"])]

tracker_final.loc[:, "temp4"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "Non AVP Audits Non EN"], tracker_final.loc[:, "Non AVP Audits Non EN Count"])]

tracker_final.loc[:, "temp5"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "Non AVP Audits EN"], tracker_final.loc[:, "Non AVP Audits EN Count"])]

tracker_final.loc[:, "CCR Total Achieved"]= tracker_final.loc[:, "temp1"]+tracker_final.loc[:, "temp2"]+tracker_final.loc[:, "temp3"]+tracker_final.loc[:, "temp4"]+tracker_final.loc[:, "temp5"]

tracker_final= tracker_final.drop(columns=['temp1', 'temp2', 'temp3','temp4', 'temp5'])

tracker_final.loc[:, "temp1"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "CCR EN"],(tracker_final.loc[:,"CCR EN"]*tracker_final.loc[:, "CCR Target/hr"]))]

tracker_final.loc[:, "temp2"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "CCR Non EN"], (tracker_final.loc[:,"CCR Non EN"]*tracker_final.loc[:, "CCR Target/hr"]))]

tracker_final.loc[:, "temp3"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "Brand Owner Audits"], (tracker_final.loc[:,"Brand Owner Audits"]*tracker_final.loc[:, "BOA NA IN Target/hr"]))]

tracker_final.loc[:, "temp4"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "Non AVP Audits Non EN"], (tracker_final.loc[:,"Non AVP Audits EN"]*tracker_final.loc[:, "NON AVP Target/hr"]))]

tracker_final.loc[:, "temp5"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "Non AVP Audits EN"], (tracker_final.loc[:,"Non AVP Audits Non EN"]*tracker_final.loc[:, "NON AVP Target/hr"]))]


tracker_final.loc[:, "CCR Total Target"]= tracker_final.loc[:, "temp1"]+tracker_final.loc[:, "temp2"]+tracker_final.loc[:, "temp3"]+tracker_final.loc[:, "temp4"]+tracker_final.loc[:, "temp5"]

tracker_final= tracker_final.drop(columns=['temp1', 'temp2', 'temp3','temp4', 'temp5'])


tracker_final.loc[:, "CCR %"]=[0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "CCR Total Achieved"],tracker_final.loc[:, "CCR Total Target"],
                                                                         (tracker_final.loc[:, "CCR Total Achieved"]/tracker_final.loc[:, "CCR Total Target"]))]

tracker_final.loc[:, "CCR %"]=tracker_final.loc[:, "CCR %"]*100

tracker_final['CCR %'].fillna(0, inplace=True)

tracker_final=tracker_final[['Login','Reporting Manager', 'Program', 'Year','Week','Date', 'End Date','Leave','Leave Type','Leave Hours',
                         'Instances','Bincheck EPCR Hours','Bincheck EPCR Count','Bincheck USN Hours','Bincheck USN Count',
                         'Bincheck CCR Hours','Bincheck CCR Count',
                         'Bincheck IXD Hours', 'Bincheck IXD Count','Bincheck Target/hr',
                         'Achieved Bincheck Weighted','Bincheck Total Achieved', 'Bincheck Total Target','Bincheck %','CCR EN','CCR EN Count','CCR Non EN','CCR Non EN Count','Brand Owner Audits',
                         'Brand Owner Audits Count','Non AVP Audits EN','Non AVP Audits EN Count','Non AVP Audits Non EN','Non AVP Audits Non EN Count',
                         'CCR Target/hr','BOA NA IN Target/hr', 'NON AVP Target/hr', 'Achieved CCR Weighted', 'CCR Total Achieved', 'CCR Total Target','CCR %',
                         'EIM FP Hours','EIM FP QA Count','EIM FP QA Hours','EIM Esc QA Count',
                         'EIM Esc Hours','EIM Esc QA Hours','EPCR EN','EPCR Non EN','EACP Hours','EACP Appeals Hours',
                         'MCM EN','MCM Non EN','MCM Live EN QA','MCM Live Non EN QA',
                         'MCM DE by Non DE resources','MCM Others NE','MCM Practice File hours',
                         'Safety Sensitivity Hours','Test Buy Hrs','USN EN','USN Non EN','DCR KAI EN(Live)',
                         'DCR KAI Non EN (Live)','DCR Live EN QA','DCR Live Non EN QA','DCR Practice File hours','WI KAI EN (Live)','WI KAI Non EN (Live)','WI Live EN QA',
                         'WI Live Non EN QA','WI Practice File Hours','SLIM Hours','PAD Hours','Coverage Hours','Retails Hours',
                         'Backlog & Ops Update','HTMS Data Reporting','Flash','Keyword Analysis',
                         'MOM','Ops Tracker','Other Reportings','PR Doc','Productivity Report',
                         'WBR Report','Work Allocation','EN QA Hours','QA Reporting',
                         'Non EN QA Hrs','QA Calibration','Deep Dive Hours','Deep Dive Name',
                         'Project Name','Project Hours',
                         'Other Task Name','Other Tasks Hours','Meeting',
                         'No Work','Other NPTS','System Issues','Training','RAMP Hours','Quality Hours',
                         'Audits', 'QA Meeting','Adhoc-Business Name', 'Adhoc-Business Hours','Adhoc-Internal Name', 'Adhoc-Internal Hours',
                         'Comments','Productive Hours','NPT SUM','NPT Hours','Leave Sum','Total Hours']]


tracker_final= tracker_final.round({"Achieved CCR Weighted":2, "CCR %":2,'CCR Total Achieved': 2, 'CCR Total Target':2})

convert_dict = {'Achieved CCR Weighted':'float32','CCR Total Achieved': 'float32',
            'CCR Total Target':'float32','CCR %':'float32'}
  
tracker_final = tracker_final.astype(convert_dict)

tracker_final['Achieved CCR Weighted'] = tracker_final['Achieved CCR Weighted'].round(2)
tracker_final['CCR Total Achieved'] = tracker_final['CCR Total Achieved'].round(2)
tracker_final['CCR Total Target'] = tracker_final['CCR Total Target'].round(2)
tracker_final['CCR %'] = tracker_final['CCR %'].round(2)



ccr=""
bo=""
non_avp=""

#EIM

eim= pd.read_excel(filename2, sheet_name='EIM_FP')
eim=pd.DataFrame(eim)

eim=eim[['Login', 'Date', 'EIM FP Count']]

eim.loc[:,'Date']=pd.to_datetime(eim.loc[:,'Date']).dt.date

eim['EIM FP Count'].fillna(0, inplace=True)

eim.drop_duplicates(subset=['Date', 'Login'], inplace=True)

convert_dict = {'Login':'category','EIM FP Count':'float16'}
  
eim = eim.astype(convert_dict)


tracker_final=pd.merge(tracker_final, eim, how='left', on=['Date', 'Login'])

tracker_final['EIM FP Count'].fillna(0, inplace=True)


tracker_final=pd.merge(tracker_final, target_data[['Week', 'EIM Target/hr']], how='left', on=['Week'])

convert_dict = {'Week':'int8','EIM Target/hr':'float16'}
  
tracker_final = tracker_final.astype(convert_dict)

tracker_final.loc[:, "temp1"]= [0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "EIM FP Hours"],
                                                                  tracker_final.loc[:, "EIM Target/hr"], (tracker_final.loc[:, "EIM FP Count"]/tracker_final.loc[:, "EIM Target/hr"]))]
                       
tracker_final.loc[:, "temp2"]=[0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "EIM FP QA Hours"],
                                                                  tracker_final.loc[:, "EIM Target/hr"], (tracker_final.loc[:, "EIM FP QA Count"]/tracker_final.loc[:, "EIM Target/hr"]))]

tracker_final.loc[:, "Achieved EIM FP Weighted"]= tracker_final.loc[:, "temp1"]+tracker_final.loc[:, "temp2"]

tracker_final= tracker_final.drop(columns=['temp1', 'temp2'])

tracker_final.loc[:, "temp1"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "EIM FP Hours"], tracker_final.loc[:, "EIM FP Count"])]

tracker_final.loc[:, "temp2"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "EIM FP QA Hours"], tracker_final.loc[:, "EIM FP QA Count"])]

tracker_final.loc[:, "EIM FP Total Achieved"]= tracker_final.loc[:, "temp1"]+tracker_final.loc[:, "temp2"]

tracker_final= tracker_final.drop(columns=['temp1', 'temp2'])

tracker_final.loc[:, "temp1"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "EIM FP Hours"], (tracker_final.loc[:,"EIM FP Hours"]*tracker_final.loc[:, "EIM Target/hr"]))]

tracker_final.loc[:, "temp2"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "EIM FP QA Hours"], (tracker_final.loc[:,"EIM FP QA Hours"]*tracker_final.loc[:, "EIM Target/hr"]))]


tracker_final.loc[:, "EIM FP Total Target"]= tracker_final.loc[:, "temp1"]+tracker_final.loc[:, "temp2"]

tracker_final= tracker_final.drop(columns=['temp1', 'temp2'])

tracker_final.loc[:, "EIM FP %"]=[0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "EIM FP Total Achieved"],tracker_final.loc[:, "EIM FP Total Target"],
                                                                            (tracker_final.loc[:, "EIM FP Total Achieved"]/tracker_final.loc[:, "EIM FP Total Target"]))]
tracker_final['EIM FP %']=tracker_final['EIM FP %']*100

tracker_final['EIM FP %'].fillna(0, inplace=True)

tracker_final=tracker_final[['Login','Reporting Manager', 'Program', 'Year','Week','Date', 'End Date','Leave','Leave Type','Leave Hours',
                         'Instances','Bincheck EPCR Hours','Bincheck EPCR Count','Bincheck USN Hours','Bincheck USN Count',
                         'Bincheck CCR Hours','Bincheck CCR Count',
                         'Bincheck IXD Hours', 'Bincheck IXD Count','Bincheck Target/hr',
                         'Achieved Bincheck Weighted','Bincheck Total Achieved', 'Bincheck Total Target','Bincheck %','CCR EN','CCR EN Count','CCR Non EN','CCR Non EN Count','Brand Owner Audits',
                         'Brand Owner Audits Count','Non AVP Audits EN','Non AVP Audits EN Count','Non AVP Audits Non EN','Non AVP Audits Non EN Count',
                         'CCR Target/hr','BOA NA IN Target/hr', 'NON AVP Target/hr', 'Achieved CCR Weighted', 'CCR Total Achieved', 'CCR Total Target','CCR %',
                         'EIM FP Hours','EIM FP Count','EIM FP QA Hours','EIM FP QA Count',
                         'EIM Target/hr', 'Achieved EIM FP Weighted','EIM FP Total Achieved', 'EIM FP Total Target', 'EIM FP %','EIM Esc Hours','EIM Esc QA Hours',
                         'EIM Esc QA Count','EPCR EN','EPCR Non EN','EACP Hours','EACP Appeals Hours',
                         'MCM EN','MCM Non EN','MCM Live EN QA','MCM Live Non EN QA',
                         'MCM DE by Non DE resources','MCM Others NE','MCM Practice File hours',
                         'Safety Sensitivity Hours','Test Buy Hrs','USN EN','USN Non EN','DCR KAI EN(Live)',
                         'DCR KAI Non EN (Live)','DCR Live EN QA','DCR Live Non EN QA','DCR Practice File hours','WI KAI EN (Live)','WI KAI Non EN (Live)','WI Live EN QA',
                         'WI Live Non EN QA','WI Practice File Hours','SLIM Hours','PAD Hours','Coverage Hours','Retails Hours',
                         'Backlog & Ops Update','HTMS Data Reporting','Flash','Keyword Analysis',
                         'MOM','Ops Tracker','Other Reportings','PR Doc','Productivity Report',
                         'WBR Report','Work Allocation','EN QA Hours','QA Reporting',
                         'Non EN QA Hrs','QA Calibration','Deep Dive Hours','Deep Dive Name',
                         'Project Name','Project Hours',
                         'Other Task Name','Other Tasks Hours','Meeting',
                         'No Work','Other NPTS','System Issues','Training','RAMP Hours','Quality Hours',
                         'Audits', 'QA Meeting','Adhoc-Business Name', 'Adhoc-Business Hours','Adhoc-Internal Name', 'Adhoc-Internal Hours',
                         'Comments','Productive Hours','NPT SUM','NPT Hours','Leave Sum','Total Hours']]


tracker_final= tracker_final.round({"Achieved EIM FP Weighted":2, "EIM FP %":2,'EIM FP Total Achieved': 2, 'EIM FP Total Target':2})

convert_dict = {'Achieved EIM FP Weighted':'float32','EIM FP Total Achieved':'float32',
            'EIM FP Total Target': 'float32','EIM FP %':'float32'}
  
tracker_final = tracker_final.astype(convert_dict)

tracker_final['Achieved EIM FP Weighted'] = tracker_final['Achieved EIM FP Weighted'].round(2)
tracker_final['EIM FP Total Achieved'] = tracker_final['EIM FP Total Achieved'].round(2)
tracker_final['EIM FP Total Target'] = tracker_final['EIM FP Total Target'].round(2)
tracker_final['EIM FP %'] = tracker_final['EIM FP %'].round(2)



eim=""

#EIM Escalations


eim_esc= pd.read_excel(filename2, sheet_name='EIM_Esc')
eim_esc=pd.DataFrame(eim_esc)

eim_esc=eim_esc[['Login', 'Date', 'EIM Esc Count']]

eim_esc.loc[:,'Date']=pd.to_datetime(eim_esc.loc[:,'Date']).dt.date

eim_esc['EIM Esc Count'].fillna(0, inplace=True)

eim_esc.drop_duplicates(subset=['Date', 'Login'], inplace=True)

convert_dict = {'Login':'category','EIM Esc Count':'float16',}
  
eim_esc = eim_esc.astype(convert_dict)

tracker_final=pd.merge(tracker_final, eim_esc, how='left', on=['Date', 'Login'])

tracker_final['EIM Esc Count'].fillna(0, inplace=True)

tracker_final=pd.merge(tracker_final, target_data[['Week', 'EIM ESC Target/hr']], how='left', on=['Week'])

convert_dict = {'Week':'int16','EIM ESC Target/hr':'float16'}
  
tracker_final = tracker_final.astype(convert_dict)



tracker_final.loc[:, "temp1"]= [0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "EIM Esc Hours"],
                                                                  tracker_final.loc[:, "EIM ESC Target/hr"], (tracker_final.loc[:, "EIM Esc Count"]/tracker_final.loc[:, "EIM ESC Target/hr"]))]
                       
tracker_final.loc[:, "temp2"]=[0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "EIM Esc QA Hours"],
                                                                  tracker_final.loc[:, "EIM ESC Target/hr"], (tracker_final.loc[:, "EIM Esc QA Count"]/tracker_final.loc[:, "EIM ESC Target/hr"]))]

tracker_final.loc[:, "Achieved EIM Esc Weighted"]= tracker_final.loc[:, "temp1"]+tracker_final.loc[:, "temp2"]

tracker_final= tracker_final.drop(columns=['temp1', 'temp2'])

tracker_final.loc[:, "temp1"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "EIM Esc Hours"], tracker_final.loc[:, "EIM Esc Count"])]

tracker_final.loc[:, "temp2"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "EIM Esc QA Hours"], tracker_final.loc[:, "EIM Esc QA Count"])]

tracker_final.loc[:, "EIM Esc Total Achieved"]= tracker_final.loc[:, "temp1"]+tracker_final.loc[:, "temp2"]

tracker_final= tracker_final.drop(columns=['temp1', 'temp2'])

tracker_final.loc[:, "temp1"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "EIM Esc Hours"],(tracker_final.loc[:,"EIM Esc Hours"]*tracker_final.loc[:, "EIM ESC Target/hr"]))]

tracker_final.loc[:, "temp2"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "EIM Esc QA Hours"], (tracker_final.loc[:,"EIM Esc QA Hours"]*tracker_final.loc[:, "EIM ESC Target/hr"]))]

tracker_final.loc[:, "EIM Esc Total Target"]= tracker_final.loc[:, "temp1"]+tracker_final.loc[:, "temp2"]

tracker_final= tracker_final.drop(columns=['temp1', 'temp2'])

tracker_final.loc[:, "EIM Escalations %"]=[0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "EIM Esc Total Achieved"],tracker_final.loc[:, "EIM Esc Total Target"],
                                                                                     (tracker_final.loc[:, "EIM Esc Total Achieved"]/tracker_final.loc[:, "EIM Esc Total Target"]))]
tracker_final['EIM Escalations %']=tracker_final['EIM Escalations %']*100

tracker_final['EIM Escalations %'].fillna(0, inplace=True)

tracker_final=tracker_final[['Login','Reporting Manager', 'Program', 'Year','Week','Date', 'End Date','Leave','Leave Type','Leave Hours',
                         'Instances','Bincheck EPCR Hours','Bincheck EPCR Count','Bincheck USN Hours','Bincheck USN Count',
                         'Bincheck IXD Hours', 'Bincheck IXD Count','Bincheck CCR Hours','Bincheck CCR Count','Bincheck Target/hr',
                         'Achieved Bincheck Weighted','Bincheck Total Achieved', 'Bincheck Total Target','Bincheck %','CCR EN','CCR EN Count','CCR Non EN','CCR Non EN Count','Brand Owner Audits',
                         'Brand Owner Audits Count','Non AVP Audits EN','Non AVP Audits EN Count','Non AVP Audits Non EN','Non AVP Audits Non EN Count',
                         'CCR Target/hr','BOA NA IN Target/hr', 'NON AVP Target/hr', 'Achieved CCR Weighted', 'CCR Total Achieved', 'CCR Total Target','CCR %',
                         'EIM FP Hours','EIM FP Count','EIM FP QA Hours','EIM FP QA Count',
                         'EIM Target/hr', 'Achieved EIM FP Weighted','EIM FP Total Achieved', 'EIM FP Total Target', 'EIM FP %','EIM Esc Hours','EIM Esc Count','EIM Esc QA Hours',
                         'EIM Esc QA Count','EIM ESC Target/hr', 'Achieved EIM Esc Weighted', 'EIM Esc Total Achieved', 'EIM Esc Total Target', 'EIM Escalations %',
                         'EPCR EN','EPCR Non EN','EACP Hours','EACP Appeals Hours',
                         'MCM EN','MCM Non EN','MCM Live EN QA','MCM Live Non EN QA',
                         'MCM DE by Non DE resources','MCM Others NE','MCM Practice File hours',
                         'Safety Sensitivity Hours','Test Buy Hrs','USN EN','USN Non EN','DCR KAI EN(Live)',
                         'DCR KAI Non EN (Live)','DCR Live EN QA','DCR Live Non EN QA','DCR Practice File hours','WI KAI EN (Live)','WI KAI Non EN (Live)','WI Live EN QA',
                         'WI Live Non EN QA','WI Practice File Hours','SLIM Hours','PAD Hours','Coverage Hours','Retails Hours',
                         'Backlog & Ops Update','HTMS Data Reporting','Flash','Keyword Analysis',
                         'MOM','Ops Tracker','Other Reportings','PR Doc','Productivity Report',
                         'WBR Report','Work Allocation','EN QA Hours','QA Reporting',
                         'Non EN QA Hrs','QA Calibration','Deep Dive Hours','Deep Dive Name',
                         'Project Name','Project Hours',
                         'Other Task Name','Other Tasks Hours','Meeting',
                         'No Work','Other NPTS','System Issues','Training','RAMP Hours','Quality Hours',
                         'Audits', 'QA Meeting','Adhoc-Business Name', 'Adhoc-Business Hours','Adhoc-Internal Name', 'Adhoc-Internal Hours',
                         'Comments','Productive Hours','NPT SUM','NPT Hours','Leave Sum','Total Hours']]


tracker_final= tracker_final.round({"Achieved EIM Esc Weighted":2, "EIM Escalations %":2,'EIM Esc Total Achieved': 2, 'EIM Esc Total Target':2})

convert_dict = {'Achieved EIM Esc Weighted':'float32','EIM Esc Total Achieved':'float32',
            'EIM Esc Total Target': 'float32','EIM Escalations %':'float32'}
  
tracker_final = tracker_final.astype(convert_dict)

tracker_final['Achieved EIM Esc Weighted'] = tracker_final['Achieved EIM Esc Weighted'].round(2)
tracker_final['EIM Esc Total Achieved'] = tracker_final['EIM Esc Total Achieved'].round(2)
tracker_final['EIM Esc Total Target'] = tracker_final['EIM Esc Total Target'].round(2)
tracker_final['EIM Escalations %'] = tracker_final['EIM Escalations %'].round(2)



eim_esc=""

#EPCR


epcr= pd.read_excel(filename2, sheet_name='Tagged_Data_Workdocs')
epcr=pd.DataFrame(epcr)

epcr=epcr[['Login', 'Date', 'EPCR EN Count', 'EPCR Non EN Count']]

epcr.loc[:,'Date']=pd.to_datetime(epcr.loc[:,'Date']).dt.date

epcr['EPCR EN Count'].fillna(0, inplace=True)

epcr['EPCR Non EN Count'].fillna(0, inplace=True)

epcr.drop_duplicates(subset=['Date', 'Login'], inplace=True)

convert_dict = {'Login':'category','EPCR EN Count':'float16',
            'EPCR Non EN Count': 'float16'}
  
epcr = epcr.astype(convert_dict)


tracker_final=pd.merge(tracker_final, epcr, how='left', on=['Date', 'Login'])

tracker_final['EPCR EN Count'].fillna(0, inplace=True)

tracker_final['EPCR Non EN Count'].fillna(0, inplace=True)

tracker_final=pd.merge(tracker_final, target_data[['Week', 'EPCR ENG Target/hr', 'EPCR NON ENG Target/hr']], how='left', on=['Week'])


convert_dict = {'Week':'int8','EPCR ENG Target/hr':'float16',
            'EPCR NON ENG Target/hr': 'float16'}
  
tracker_final = tracker_final.astype(convert_dict)


#Eacp data

eacp= pd.read_excel(filename2, sheet_name='EACP_Data')
eacp=pd.DataFrame(eacp)

eacp=eacp[['Login', 'Date', 'EACP Count']]

eacp.loc[:,'Date']=pd.to_datetime(eacp.loc[:,'Date']).dt.date

eacp['EACP Count'].fillna(0, inplace=True)

eacp.drop_duplicates(subset=['Date', 'Login'], inplace=True)

convert_dict = {'Login':'category','EACP Count':'float16'}
  
eacp = eacp.astype(convert_dict)

tracker_final=pd.merge(tracker_final, eacp, how='left', on=['Date', 'Login'])

tracker_final['EACP Count'].fillna(0, inplace=True)

tracker_final=pd.merge(tracker_final, target_data[['Week', 'EACP Target/hr']], how='left', on=['Week'])

convert_dict = {'Week':'int8','EACP Target/hr':'float16'}
  
tracker_final = tracker_final.astype(convert_dict)

#Eacp appeals

eacp_appeals= pd.read_excel(filename2, sheet_name='EACP_Appeals')
eacp_appeals=pd.DataFrame(eacp_appeals)

eacp_appeals=eacp_appeals[['Login', 'Date', 'EACP Appeals Count']]

eacp_appeals.loc[:,'Date']=pd.to_datetime(eacp_appeals.loc[:,'Date']).dt.date

eacp_appeals['EACP Appeals Count'].fillna(0, inplace=True)

eacp_appeals.drop_duplicates(subset=['Date', 'Login'], inplace=True)

convert_dict = {'Login':'category','EACP Appeals Count':'float16'}
  
eacp_appeals = eacp_appeals.astype(convert_dict)


tracker_final=pd.merge(tracker_final, eacp_appeals, how='left', on=['Date', 'Login'])

tracker_final['EACP Appeals Count'].fillna(0, inplace=True)

tracker_final=pd.merge(tracker_final, target_data[['Week', 'EACP Appeals Target/hr']], how='left', on=['Week'])


convert_dict = {'Week':'int8','EACP Appeals Target/hr':'float16'}
  
tracker_final = tracker_final.astype(convert_dict)

tracker_final.loc[:, "temp1"]= [0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "EPCR EN"],
                                                                  tracker_final.loc[:, "EPCR ENG Target/hr"], (tracker_final.loc[:, "EPCR EN Count"]/tracker_final.loc[:, "EPCR ENG Target/hr"]))]
                       
tracker_final.loc[:, "temp2"]=[0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "EPCR Non EN"],
                                                                  tracker_final.loc[:, "EPCR NON ENG Target/hr"], (tracker_final.loc[:, "EPCR Non EN Count"]/tracker_final.loc[:, "EPCR NON ENG Target/hr"]))]

tracker_final.loc[:, "temp3"]=[0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "EACP Hours"],
                                                                  tracker_final.loc[:, "EACP Target/hr"], (tracker_final.loc[:, "EACP Count"]/tracker_final.loc[:, "EACP Target/hr"]))]

tracker_final.loc[:, "temp4"]=[0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "EACP Appeals Hours"],
                                                                  tracker_final.loc[:, "EACP Appeals Target/hr"], (tracker_final.loc[:, "EACP Appeals Count"]/tracker_final.loc[:, "EACP Appeals Target/hr"]))]

tracker_final.loc[:, "Achieved EPCR Weighted"]= tracker_final.loc[:, "temp1"]+tracker_final.loc[:, "temp2"]+tracker_final.loc[:, "temp3"]+tracker_final.loc[:, "temp4"]

tracker_final= tracker_final.drop(columns=['temp1', 'temp2','temp3', 'temp4'])

tracker_final.loc[:, "temp1"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "EPCR EN"], tracker_final.loc[:, "EPCR EN Count"])]

tracker_final.loc[:, "temp2"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "EPCR Non EN"], tracker_final.loc[:, "EPCR Non EN Count"])]

tracker_final.loc[:, "temp3"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "EACP Hours"], tracker_final.loc[:, "EACP Count"])]

tracker_final.loc[:, "temp4"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "EACP Appeals Hours"], tracker_final.loc[:, "EACP Appeals Count"])]

tracker_final.loc[:, "EPCR Total Achieved"]= tracker_final.loc[:, "temp1"]+tracker_final.loc[:, "temp2"]+tracker_final.loc[:, "temp3"]+tracker_final.loc[:, "temp4"]

tracker_final= tracker_final.drop(columns=['temp1', 'temp2','temp3', 'temp4'])


tracker_final.loc[:, "temp1"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "EPCR EN"], (tracker_final.loc[:,"EPCR EN"]*tracker_final.loc[:, "EPCR ENG Target/hr"]))]

tracker_final.loc[:, "temp2"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "EPCR Non EN"],(tracker_final.loc[:,"EPCR Non EN"]*tracker_final.loc[:, "EPCR NON ENG Target/hr"]))]

tracker_final.loc[:, "temp3"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "EACP Hours"], (tracker_final.loc[:,"EACP Hours"]*tracker_final.loc[:, "EACP Target/hr"]))]

tracker_final.loc[:, "temp4"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "EACP Appeals Hours"], (tracker_final.loc[:,"EACP Appeals Hours"]*tracker_final.loc[:, "EACP Appeals Target/hr"]))]


tracker_final.loc[:, "EPCR Total Target"]=tracker_final.loc[:, "temp1"]+tracker_final.loc[:, "temp2"]+tracker_final.loc[:, "temp3"]+tracker_final.loc[:, "temp4"]

tracker_final= tracker_final.drop(columns=['temp1', 'temp2','temp3', 'temp4'])


tracker_final.loc[:, "EPCR %"]=[0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "EPCR Total Achieved"],tracker_final.loc[:, "EPCR Total Target"],
                                                                          (tracker_final.loc[:, "EPCR Total Achieved"]/tracker_final.loc[:, "EPCR Total Target"]))]
tracker_final['EPCR %']=tracker_final['EPCR %']*100

tracker_final['EPCR %'].fillna(0, inplace=True)

tracker_final=tracker_final[['Login','Reporting Manager', 'Program', 'Year','Week','Date', 'End Date','Leave','Leave Type','Leave Hours',
                             'Instances','Bincheck EPCR Hours','Bincheck EPCR Count','Bincheck USN Hours','Bincheck USN Count',
                             'Bincheck IXD Hours', 'Bincheck IXD Count','Bincheck CCR Hours','Bincheck CCR Count','Bincheck Target/hr',
                             'Achieved Bincheck Weighted','Bincheck Total Achieved', 'Bincheck Total Target','Bincheck %','CCR EN','CCR EN Count','CCR Non EN','CCR Non EN Count','Brand Owner Audits',
                             'Brand Owner Audits Count','Non AVP Audits EN','Non AVP Audits EN Count','Non AVP Audits Non EN','Non AVP Audits Non EN Count',
                             'CCR Target/hr','BOA NA IN Target/hr', 'NON AVP Target/hr', 'Achieved CCR Weighted', 'CCR Total Achieved', 'CCR Total Target','CCR %',
                             'EIM FP Hours','EIM FP Count','EIM FP QA Hours','EIM FP QA Count',
                             'EIM Target/hr', 'Achieved EIM FP Weighted','EIM FP Total Achieved', 'EIM FP Total Target', 'EIM FP %','EIM Esc Hours','EIM Esc Count','EIM Esc QA Hours',
                             'EIM Esc QA Count','EIM ESC Target/hr', 'Achieved EIM Esc Weighted', 'EIM Esc Total Achieved', 'EIM Esc Total Target', 'EIM Escalations %',
                             'EPCR EN','EPCR EN Count','EPCR Non EN','EPCR Non EN Count','EACP Hours','EACP Count','EACP Appeals Hours','EACP Appeals Count',
                             'EACP Appeals Target/hr', 'EPCR ENG Target/hr', 'EPCR NON ENG Target/hr','EACP Target/hr','Achieved EPCR Weighted','EPCR Total Achieved',
                             'EPCR Total Target','EPCR %',
                             'MCM EN','MCM Non EN','MCM Live EN QA','MCM Live Non EN QA',
                             'MCM DE by Non DE resources','MCM Others NE','MCM Practice File hours',
                             'Safety Sensitivity Hours','Test Buy Hrs','USN EN','USN Non EN','DCR KAI EN(Live)',
                             'DCR KAI Non EN (Live)','DCR Live EN QA','DCR Live Non EN QA','DCR Practice File hours','WI KAI EN (Live)','WI KAI Non EN (Live)','WI Live EN QA',
                             'WI Live Non EN QA','WI Practice File Hours','SLIM Hours','PAD Hours','Coverage Hours','Retails Hours',
                             'Backlog & Ops Update','HTMS Data Reporting','Flash','Keyword Analysis',
                             'MOM','Ops Tracker','Other Reportings','PR Doc','Productivity Report',
                             'WBR Report','Work Allocation','EN QA Hours','QA Reporting',
                             'Non EN QA Hrs','QA Calibration','Deep Dive Hours','Deep Dive Name',
                             'Project Name','Project Hours',
                             'Other Task Name','Other Tasks Hours','Meeting',
                             'No Work','Other NPTS','System Issues','Training','RAMP Hours','Quality Hours',
                             'Audits', 'QA Meeting','Adhoc-Business Name', 'Adhoc-Business Hours','Adhoc-Internal Name', 'Adhoc-Internal Hours',
                             'Comments','Productive Hours','NPT SUM','NPT Hours','Leave Sum','Total Hours']]



tracker_final= tracker_final.round({"Achieved EPCR Weighted":2, "EPCR %":2,'EPCR Total Achieved': 2, 'EPCR Total Target':2})

convert_dict = {'Achieved EPCR Weighted':'float32','EPCR Total Achieved':'float32',
            'EPCR Total Target': 'float32','EPCR %':'float32'}
  
# tracker_final = tracker_final.astype(convert_dict)

tracker_final['Achieved EPCR Weighted'] = tracker_final['Achieved EPCR Weighted'].round(2)
tracker_final['EPCR Total Achieved'] = tracker_final['EPCR Total Achieved'].round(2)
tracker_final['EPCR Total Target'] = tracker_final['EPCR Total Target'].round(2)
tracker_final['EPCR %'] = tracker_final['EPCR %'].round(2)



epcr=""
eacp=""
eacp_appeals=""


#MCM



#getting mcm data from different sheets in connections file
mcm1=pd.read_excel(filename2, sheet_name='Tagged_Data_Workdocs')
mcm1=pd.DataFrame(mcm1)

mcm1=mcm1[['Login', 'Date', 'MCM EN Count', 'MCM Non EN Count','MCM Live EN QA Count',
           'MCM Live Non EN QA Count']]

mcm1.loc[:,'Date']=pd.to_datetime(mcm1.loc[:,'Date']).dt.date

mcm1['MCM EN Count'].fillna(0, inplace=True)

mcm1['MCM Non EN Count'].fillna(0, inplace=True)

mcm1['MCM Live EN QA Count'].fillna(0, inplace=True)

mcm1['MCM Live Non EN QA Count'].fillna(0, inplace=True)


mcm1.drop_duplicates(subset=['Date', 'Login'], inplace=True)

convert_dict={'Login':'category','MCM EN Count':'float16', 'MCM Non EN Count':'float16',
          'MCM Live EN QA Count':'float16','MCM Live Non EN QA Count':'float16'}

mcm1 = mcm1.astype(convert_dict)

tracker_final=pd.merge(tracker_final, mcm1, how='left', on=['Date', 'Login'])

mcm2= pd.read_excel(filename2, sheet_name="MCM_NE")
mcm2=pd.DataFrame(mcm2)

mcm2=mcm2[['Login', 'Date', 'MCM Others NE Count']]

mcm2.loc[:,'Date']=pd.to_datetime(mcm2.loc[:,'Date']).dt.date

mcm2['MCM Others NE Count'].fillna(0, inplace=True)

convert_dict={'Login':'category', 'MCM Others NE Count':'float16'}

mcm2 = mcm2.astype(convert_dict)

tracker_final=pd.merge(tracker_final, mcm2, how='left', on=['Date', 'Login'])

mcm3=pd.read_excel(filename2, sheet_name='MCM_Practice_Files')
mcm3=pd.DataFrame(mcm3)

mcm3=mcm3[['Login', 'Date', 'MCM Practice Files Count']]

mcm3.loc[:,'Date']=pd.to_datetime(mcm3.loc[:,'Date']).dt.date

mcm3['MCM Practice Files Count'].fillna(0, inplace=True)

convert_dict={'Login':'category', 'MCM Practice Files Count':'float16'}

mcm3 = mcm3.astype(convert_dict)

tracker_final=pd.merge(tracker_final, mcm3, how='left', on=['Date', 'Login'])

tracker_final['MCM EN Count'].fillna(0, inplace=True)

tracker_final['MCM Non EN Count'].fillna(0, inplace=True)

tracker_final['MCM Live EN QA Count'].fillna(0, inplace=True)

tracker_final['MCM Practice Files Count'].fillna(0, inplace=True)

tracker_final['MCM Others NE Count'].fillna(0, inplace=True)

mcm3.drop_duplicates(subset=['Date', 'Login'], inplace=True)


tracker_final=pd.merge(tracker_final, target_data[['Week','MCM ENG Target/hr',
                                               'MCM NON ENG Target/hr',
                                               'MCM Others NE Target/hr']], how='left', on=['Week'])

convert_dict = {'Week':'int8','MCM ENG Target/hr':'float16',
            'MCM NON ENG Target/hr':'float16', "MCM Others NE Target/hr":'float16',
            }
  
tracker_final = tracker_final.astype(convert_dict)

tracker_final.loc[:, "temp1"]= [0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "MCM EN Count"],
                                                                          tracker_final.loc[:, "MCM ENG Target/hr"], (tracker_final.loc[:, "MCM EN Count"]/tracker_final.loc[:, "MCM ENG Target/hr"]))]
                       
tracker_final.loc[:, "temp2"]=[0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "MCM Non EN Count"],
                                                                         tracker_final.loc[:, "MCM NON ENG Target/hr"], (tracker_final.loc[:, "MCM Non EN Count"]/tracker_final.loc[:, "MCM NON ENG Target/hr"]))]

tracker_final.loc[:, "temp3"]=[0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "MCM Live EN QA Count"],
                                                                         tracker_final.loc[:, "MCM ENG Target/hr"], (tracker_final.loc[:, "MCM Live EN QA Count"]/tracker_final.loc[:, "MCM ENG Target/hr"]))]

tracker_final.loc[:, "temp4"]=[0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "MCM Live Non EN QA Count"],
                                                                         tracker_final.loc[:, "MCM NON ENG Target/hr"], (tracker_final.loc[:, "MCM Live Non EN QA Count"]/tracker_final.loc[:, "MCM NON ENG Target/hr"]))]

tracker_final.loc[:, "temp5"]=[0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "MCM Others NE Count"],
                                                                         tracker_final.loc[:, "MCM Others NE Target/hr"], (tracker_final.loc[:, "MCM Others NE Count"]/tracker_final.loc[:, "MCM Others NE Target/hr"]))]


tracker_final.loc[:, "Achieved MCM Weighted"]= (tracker_final.loc[:, "temp1"]+tracker_final.loc[:, "temp2"]+tracker_final.loc[:, "temp3"]+
                                                tracker_final.loc[:, "temp4"]+tracker_final.loc[:, "temp5"])

tracker_final= tracker_final.drop(columns=['temp1', 'temp2','temp3', 'temp4', 'temp5'])

tracker_final.loc[:, "temp1"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "MCM EN Count"], tracker_final.loc[:, "MCM EN Count"])]

tracker_final.loc[:, "temp2"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "MCM Non EN Count"], tracker_final.loc[:, "MCM Non EN Count"])]

tracker_final.loc[:, "temp3"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "MCM Live EN QA Count"], tracker_final.loc[:, "MCM Live EN QA Count"])]

tracker_final.loc[:, "temp4"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "MCM Live Non EN QA Count"], tracker_final.loc[:, "MCM Live Non EN QA Count"])]

tracker_final.loc[:, "temp5"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "MCM Others NE Count"], tracker_final.loc[:, "MCM Others NE Count"])]

tracker_final.loc[:, "MCM Total Achieved"]= (tracker_final.loc[:, "temp1"]+tracker_final.loc[:, "temp2"]+
                                             tracker_final.loc[:, "temp3"]+tracker_final.loc[:, "temp4"]+
                                             tracker_final.loc[:, "temp5"])

tracker_final= tracker_final.drop(columns=['temp1', 'temp2','temp3', 'temp4', 'temp5'])

tracker_final.loc[:, "temp1"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "MCM EN"], (tracker_final.loc[:,"MCM EN"]*tracker_final.loc[:, "MCM ENG Target/hr"]))]

tracker_final.loc[:, "temp2"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "MCM Non EN"], (tracker_final.loc[:,"MCM Non EN"]*tracker_final.loc[:, "MCM NON ENG Target/hr"]))]

tracker_final.loc[:, "temp3"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "MCM Live EN QA"], (tracker_final.loc[:,"MCM Live EN QA"]*tracker_final.loc[:, "MCM ENG Target/hr"]))]

tracker_final.loc[:, "temp4"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "MCM Live Non EN QA"], (tracker_final.loc[:,"MCM Live Non EN QA"]*tracker_final.loc[:, "MCM NON ENG Target/hr"]))]

tracker_final.loc[:, "temp5"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "MCM Others NE"], (tracker_final.loc[:,"MCM Others NE"]*tracker_final.loc[:, "MCM Others NE Target/hr"]))]

tracker_final.loc[:, "MCM Total Target"]= (tracker_final.loc[:, "temp1"]+tracker_final.loc[:, "temp2"]+
                                           tracker_final.loc[:, "temp3"]+tracker_final.loc[:, "temp4"]+
                                           tracker_final.loc[:, "temp5"])

tracker_final= tracker_final.drop(columns=['temp1', 'temp2','temp3', 'temp4', 'temp5'])

tracker_final.loc[:, "MCM %"]=[0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "MCM Total Achieved"],tracker_final.loc[:, "MCM Total Target"],
                                                                         (tracker_final.loc[:, "MCM Total Achieved"]/tracker_final.loc[:, "MCM Total Target"]))]
tracker_final['MCM %']=tracker_final['MCM %']*100

tracker_final['MCM %'].fillna(0, inplace=True)

tracker_final=tracker_final[['Login','Reporting Manager', 'Program', 'Year','Week','Date', 'End Date','Leave','Leave Type','Leave Hours',
                             'Instances','Bincheck EPCR Hours','Bincheck EPCR Count','Bincheck USN Hours','Bincheck USN Count',
                             'Bincheck IXD Hours', 'Bincheck IXD Count','Bincheck CCR Hours','Bincheck CCR Count','Bincheck Target/hr',
                             'Achieved Bincheck Weighted','Bincheck Total Achieved', 'Bincheck Total Target','Bincheck %','CCR EN','CCR EN Count','CCR Non EN','CCR Non EN Count','Brand Owner Audits',
                             'Brand Owner Audits Count','Non AVP Audits EN','Non AVP Audits EN Count','Non AVP Audits Non EN','Non AVP Audits Non EN Count',
                             'CCR Target/hr','BOA NA IN Target/hr', 'NON AVP Target/hr', 'Achieved CCR Weighted', 'CCR Total Achieved', 'CCR Total Target','CCR %',
                             'EIM FP Hours','EIM FP Count','EIM FP QA Hours','EIM FP QA Count',
                             'EIM Target/hr', 'Achieved EIM FP Weighted','EIM FP Total Achieved', 'EIM FP Total Target', 'EIM FP %','EIM Esc Hours','EIM Esc Count','EIM Esc QA Hours',
                             'EIM Esc QA Count','EIM ESC Target/hr', 'Achieved EIM Esc Weighted', 'EIM Esc Total Achieved', 'EIM Esc Total Target', 'EIM Escalations %',
                             'EPCR EN','EPCR EN Count','EPCR Non EN','EPCR Non EN Count','EACP Hours','EACP Count','EACP Appeals Hours','EACP Appeals Count',
                             'EACP Appeals Target/hr', 'EPCR ENG Target/hr', 'EPCR NON ENG Target/hr','EACP Target/hr','Achieved EPCR Weighted','EPCR Total Achieved',
                             'EPCR Total Target','EPCR %',
                             'MCM EN','MCM EN Count','MCM Non EN','MCM Non EN Count','MCM Live EN QA',
                             'MCM Live EN QA Count','MCM Live Non EN QA','MCM Live Non EN QA Count',
                             'MCM Others NE','MCM Others NE Count','MCM Practice File hours','MCM Practice Files Count',
                             'MCM ENG Target/hr', 'MCM NON ENG Target/hr','MCM Others NE Target/hr',
                             'Achieved MCM Weighted','MCM Total Achieved', 'MCM Total Target', 'MCM %',
                             'Safety Sensitivity Hours','Test Buy Hrs','USN EN','USN Non EN','DCR KAI EN(Live)',
                             'DCR KAI Non EN (Live)','DCR Live EN QA','DCR Live Non EN QA','DCR Practice File hours','WI KAI EN (Live)','WI KAI Non EN (Live)','WI Live EN QA',
                             'WI Live Non EN QA','WI Practice File Hours','SLIM Hours','PAD Hours','Coverage Hours','Retails Hours',
                             'Backlog & Ops Update','HTMS Data Reporting','Flash','Keyword Analysis',
                             'MOM','Ops Tracker','Other Reportings','PR Doc','Productivity Report',
                             'WBR Report','Work Allocation','EN QA Hours','QA Reporting',
                             'Non EN QA Hrs','QA Calibration','Deep Dive Hours','Deep Dive Name',
                             'Project Name','Project Hours',
                             'Other Task Name','Other Tasks Hours','Meeting',
                             'No Work','Other NPTS','System Issues','Training','RAMP Hours','Quality Hours',
                             'Audits', 'QA Meeting','Adhoc-Business Name', 'Adhoc-Business Hours','Adhoc-Internal Name', 'Adhoc-Internal Hours',
                             'Comments','Productive Hours','NPT SUM','NPT Hours','Leave Sum','Total Hours']]


tracker_final= tracker_final.round({"Achieved MCM Weighted":2, "MCM %":2,'MCM Total Achieved': 2, 'MCM Total Target':2})

convert_dict = {'Achieved MCM Weighted':'float32','MCM Total Achieved':'float32',
            'MCM Total Target': 'float32','MCM %':'float32'}
  
tracker_final = tracker_final.astype(convert_dict)

tracker_final['Achieved MCM Weighted'] = tracker_final['Achieved MCM Weighted'].round(2)
tracker_final['MCM Total Achieved'] = tracker_final['MCM Total Achieved'].round(2)
tracker_final['MCM Total Target'] = tracker_final['MCM Total Target'].round(2)
tracker_final['MCM %'] = tracker_final['MCM %'].round(2)



mcm1=""
mcm3=""

#safety sensitivity

ss= pd.read_excel(filename2, sheet_name='Safety_Sensitive')
ss=pd.DataFrame(ss)

ss=ss[['Login', 'Date', 'Safety Sensitivity Count']]

ss.loc[:,'Date']=pd.to_datetime(ss.loc[:,'Date']).dt.date

ss['Safety Sensitivity Count'].fillna(0, inplace=True)

ss.drop_duplicates(subset=['Date', 'Login'], inplace=True)

tracker_final=pd.merge(tracker_final, ss, how='left', on=['Date', 'Login'])

tracker_final['Safety Sensitivity Count'].fillna(0, inplace=True)

tracker_final=pd.merge(tracker_final, target_data[['Week', '​​Safety Sensitivity Target/hr']], how='left', on=['Week'])

convert_dict = {'Login':'category','Week':'int8','​​Safety Sensitivity Target/hr':'float16','Safety Sensitivity Count':'float16'}
  
tracker_final = tracker_final.astype(convert_dict)


tracker_final.loc[:, "Achieved Safety Sensitivity Weighted"]= [0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "Safety Sensitivity Hours"],
                                                                                                     tracker_final.loc[:, "​​Safety Sensitivity Target/hr"], (tracker_final.loc[:, "Safety Sensitivity Count"]/tracker_final.loc[:, "​​Safety Sensitivity Target/hr"]))]


tracker_final.loc[:, "temp1"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "Safety Sensitivity Hours"], tracker_final.loc[:, "Safety Sensitivity Count"])]

tracker_final.loc[:, "Safety Sensitivity Total Achieved"]= tracker_final.loc[:, "temp1"]

tracker_final= tracker_final.drop(columns=['temp1'])

tracker_final.loc[:, "Safety Sensitivity Total Target"]= (tracker_final.loc[:,"Safety Sensitivity Hours"]*tracker_final.loc[:, "​​Safety Sensitivity Target/hr"])

tracker_final.loc[:, "Safety Sensitivity %"]=[0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "Safety Sensitivity Total Achieved"],tracker_final.loc[:, "Safety Sensitivity Total Target"],
                                                                                        (tracker_final.loc[:, "Safety Sensitivity Total Achieved"]/tracker_final.loc[:, "Safety Sensitivity Total Target"]))]
tracker_final['Safety Sensitivity %']=tracker_final['Safety Sensitivity %']*100

tracker_final['Safety Sensitivity %'].fillna(0, inplace=True)

tracker_final=tracker_final[['Login','Reporting Manager', 'Program', 'Year','Week','Date', 'End Date','Leave','Leave Type','Leave Hours',
                             'Instances','Bincheck EPCR Hours','Bincheck EPCR Count','Bincheck USN Hours','Bincheck USN Count',
                             'Bincheck IXD Hours', 'Bincheck IXD Count','Bincheck CCR Hours','Bincheck CCR Count','Bincheck Target/hr',
                             'Achieved Bincheck Weighted','Bincheck Total Achieved', 'Bincheck Total Target','Bincheck %','CCR EN','CCR EN Count','CCR Non EN','CCR Non EN Count','Brand Owner Audits',
                             'Brand Owner Audits Count','Non AVP Audits EN','Non AVP Audits EN Count','Non AVP Audits Non EN','Non AVP Audits Non EN Count',
                             'CCR Target/hr','BOA NA IN Target/hr', 'NON AVP Target/hr', 'Achieved CCR Weighted', 'CCR Total Achieved', 'CCR Total Target','CCR %',
                             'EIM FP Hours','EIM FP Count','EIM FP QA Hours','EIM FP QA Count',
                             'EIM Target/hr', 'Achieved EIM FP Weighted','EIM FP Total Achieved', 'EIM FP Total Target', 'EIM FP %','EIM Esc Hours','EIM Esc Count','EIM Esc QA Hours',
                             'EIM Esc QA Count','EIM ESC Target/hr', 'Achieved EIM Esc Weighted', 'EIM Esc Total Achieved', 'EIM Esc Total Target', 'EIM Escalations %',
                             'EPCR EN','EPCR EN Count','EPCR Non EN','EPCR Non EN Count','EACP Hours','EACP Count','EACP Appeals Hours','EACP Appeals Count',
                             'EACP Appeals Target/hr', 'EPCR ENG Target/hr', 'EPCR NON ENG Target/hr','EACP Target/hr','Achieved EPCR Weighted','EPCR Total Achieved',
                             'EPCR Total Target','EPCR %',
                             'MCM EN','MCM EN Count','MCM Non EN','MCM Non EN Count','MCM Live EN QA',
                             'MCM Live EN QA Count','MCM Live Non EN QA','MCM Live Non EN QA Count',
                             'MCM Others NE','MCM Others NE Count','MCM Practice File hours','MCM Practice Files Count',
                             'MCM ENG Target/hr', 'MCM NON ENG Target/hr','MCM Others NE Target/hr',
                             'Achieved MCM Weighted','MCM Total Achieved', 'MCM Total Target', 'MCM %',
                             'Safety Sensitivity Hours','Safety Sensitivity Count','​​Safety Sensitivity Target/hr','Achieved Safety Sensitivity Weighted',
                             'Safety Sensitivity Total Achieved', 'Safety Sensitivity Total Target','Safety Sensitivity %',
                             'Test Buy Hrs','USN EN','USN Non EN','DCR KAI EN(Live)',
                             'DCR KAI Non EN (Live)','DCR Live EN QA','DCR Live Non EN QA','DCR Practice File hours','WI KAI EN (Live)','WI KAI Non EN (Live)','WI Live EN QA',
                             'WI Live Non EN QA','WI Practice File Hours','SLIM Hours','PAD Hours','Coverage Hours','Retails Hours',
                             'Backlog & Ops Update','HTMS Data Reporting','Flash','Keyword Analysis',
                             'MOM','Ops Tracker','Other Reportings','PR Doc','Productivity Report',
                             'WBR Report','Work Allocation','EN QA Hours','QA Reporting',
                             'Non EN QA Hrs','QA Calibration','Deep Dive Hours','Deep Dive Name',
                             'Project Name','Project Hours',
                             'Other Task Name','Other Tasks Hours','Meeting',
                             'No Work','Other NPTS','System Issues','Training','RAMP Hours','Quality Hours',
                             'Audits', 'QA Meeting','Adhoc-Business Name', 'Adhoc-Business Hours','Adhoc-Internal Name', 'Adhoc-Internal Hours',
                             'Comments','Productive Hours','NPT SUM','NPT Hours','Leave Sum','Total Hours']]


tracker_final= tracker_final.round({"Achieved Safety Sensitivity Weighted":2, "Safety Sensitivity %":2,
                                    "Safety Sensitivity Total Achieved":2, "Safety Sensitivity Total Target":2})

convert_dict = {'Achieved Safety Sensitivity Weighted':'float32',
                'Safety Sensitivity Total Achieved':'float32',
                'Safety Sensitivity Total Target':'float32',
                'Safety Sensitivity %':'float32'}
  
tracker_final = tracker_final.astype(convert_dict)

tracker_final['Achieved Safety Sensitivity Weighted'] = tracker_final['Achieved Safety Sensitivity Weighted'].round(2)
tracker_final['Safety Sensitivity Total Achieved'] = tracker_final['Safety Sensitivity Total Achieved'].round(2)
tracker_final['Safety Sensitivity Total Target'] = tracker_final['Safety Sensitivity Total Target'].round(2)
tracker_final['Safety Sensitivity %'] = tracker_final['Safety Sensitivity %'].round(2)



ss=""



#Test buy

tb=pd.read_excel(filename2, sheet_name='Test_buy')

tb=pd.DataFrame(tb)

tb=tb[['Login', 'Date', 'Test Buy Count']]

tb.loc[:,'Date']=pd.to_datetime(tb.loc[:,'Date']).dt.date

tb['Test Buy Count'].fillna(0, inplace=True)

tb.drop_duplicates(subset=['Date', 'Login'], inplace=True)

convert_dict = {'Login':'category', 'Test Buy Count':'float16'}
  
tb = tb.astype(convert_dict)

tracker_final=pd.merge(tracker_final, tb, how='left', on=['Date', 'Login'])

tracker_final['Test Buy Count'].fillna(0, inplace=True)

tracker_final=pd.merge(tracker_final, target_data[['Week', 'GPV -Test Buy Target/hr']], how='left', on=['Week'])

convert_dict = {'Week':'int8','GPV -Test Buy Target/hr':'float16'}
  
tracker_final = tracker_final.astype(convert_dict)

tracker_final=tracker_final[['Login','Reporting Manager', 'Program', 'Year','Week','Date', 'End Date','Leave','Leave Type','Leave Hours',
                             'Instances','Bincheck EPCR Hours','Bincheck EPCR Count','Bincheck USN Hours','Bincheck USN Count',
                             'Bincheck IXD Hours', 'Bincheck IXD Count','Bincheck CCR Hours','Bincheck CCR Count','Bincheck Target/hr',
                             'Achieved Bincheck Weighted','Bincheck Total Achieved', 'Bincheck Total Target','Bincheck %','CCR EN','CCR EN Count','CCR Non EN','CCR Non EN Count','Brand Owner Audits',
                             'Brand Owner Audits Count','Non AVP Audits EN','Non AVP Audits EN Count','Non AVP Audits Non EN','Non AVP Audits Non EN Count',
                             'CCR Target/hr','BOA NA IN Target/hr', 'NON AVP Target/hr', 'Achieved CCR Weighted', 'CCR Total Achieved', 'CCR Total Target','CCR %',
                             'EIM FP Hours','EIM FP Count','EIM FP QA Hours','EIM FP QA Count',
                             'EIM Target/hr', 'Achieved EIM FP Weighted','EIM FP Total Achieved', 'EIM FP Total Target', 'EIM FP %','EIM Esc Hours','EIM Esc Count','EIM Esc QA Hours',
                             'EIM Esc QA Count','EIM ESC Target/hr', 'Achieved EIM Esc Weighted', 'EIM Esc Total Achieved', 'EIM Esc Total Target', 'EIM Escalations %',
                             'EPCR EN','EPCR EN Count','EPCR Non EN','EPCR Non EN Count','EACP Hours','EACP Count','EACP Appeals Hours','EACP Appeals Count',
                             'EACP Appeals Target/hr', 'EPCR ENG Target/hr', 'EPCR NON ENG Target/hr','EACP Target/hr','Achieved EPCR Weighted','EPCR Total Achieved',
                             'EPCR Total Target','EPCR %',
                             'MCM EN','MCM EN Count','MCM Non EN','MCM Non EN Count','MCM Live EN QA',
                             'MCM Live EN QA Count','MCM Live Non EN QA','MCM Live Non EN QA Count',
                             'MCM Others NE','MCM Others NE Count','MCM Practice File hours','MCM Practice Files Count',
                             'MCM ENG Target/hr', 'MCM NON ENG Target/hr','MCM Others NE Target/hr',
                             'Achieved MCM Weighted','MCM Total Achieved', 'MCM Total Target', 'MCM %',
                             'Safety Sensitivity Hours','Safety Sensitivity Count','​​Safety Sensitivity Target/hr',
                             'Safety Sensitivity Total Achieved', 'Safety Sensitivity Total Target','Achieved Safety Sensitivity Weighted','Safety Sensitivity %','Test Buy Hrs',
                             'Test Buy Count','GPV -Test Buy Target/hr','USN EN','USN Non EN','DCR KAI EN(Live)',
                             'DCR KAI Non EN (Live)','DCR Live EN QA','DCR Live Non EN QA','DCR Practice File hours','WI KAI EN (Live)','WI KAI Non EN (Live)','WI Live EN QA',
                             'WI Live Non EN QA','WI Practice File Hours','SLIM Hours','PAD Hours','Coverage Hours','Retails Hours',
                             'Backlog & Ops Update','HTMS Data Reporting','Flash','Keyword Analysis',
                             'MOM','Ops Tracker','Other Reportings','PR Doc','Productivity Report',
                             'WBR Report','Work Allocation','EN QA Hours','QA Reporting',
                             'Non EN QA Hrs','QA Calibration','Deep Dive Hours','Deep Dive Name',
                             'Project Name','Project Hours',
                             'Other Task Name','Other Tasks Hours','Meeting',
                             'No Work','Other NPTS','System Issues','Training','RAMP Hours','Quality Hours',
                             'Audits', 'QA Meeting','Adhoc-Business Name', 'Adhoc-Business Hours','Adhoc-Internal Name', 'Adhoc-Internal Hours',
                             'Comments','Productive Hours','NPT SUM','NPT Hours','Leave Sum','Total Hours']]


tb=""

#USN

usn=pd.read_excel(filename2, sheet_name='Tagged_Data_Workdocs')

usn=pd.DataFrame(usn)

usn=usn[['Login', 'Date', 'USN EN Count','USN Non EN Count']]

usn.loc[:,'Date']=pd.to_datetime(usn.loc[:,'Date']).dt.date

usn['USN EN Count'].fillna(0, inplace=True)

usn['USN Non EN Count'].fillna(0, inplace=True)

usn.drop_duplicates(subset=['Date', 'Login'], inplace=True)

convert_dict = {'Login':'category', 'USN EN Count':'float16','USN Non EN Count':'float16'}
  
usn = usn.astype(convert_dict)

tracker_final=pd.merge(tracker_final, usn, how='left', on=['Date', 'Login'])

tracker_final['USN EN Count'].fillna(0, inplace=True)

tracker_final['USN Non EN Count'].fillna(0, inplace=True)

tracker_final=pd.merge(tracker_final, target_data[['Week', 'USN ENG Target/hr','USN NON ENG Target/hr']], how='left', on=['Week'])

convert_dict = {'Week':'int8','USN ENG Target/hr':'float16',
            'USN NON ENG Target/hr': 'float16'}
  
tracker_final = tracker_final.astype(convert_dict)

tracker_final.loc[:, "temp1"]= [0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "USN EN"],
                                                                  tracker_final.loc[:, "USN ENG Target/hr"], (tracker_final.loc[:, "USN EN Count"]/tracker_final.loc[:, "USN ENG Target/hr"]))]
                       
tracker_final.loc[:, "temp2"]=[0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "USN Non EN"],
                                                                  tracker_final.loc[:, "USN NON ENG Target/hr"], (tracker_final.loc[:, "USN Non EN Count"]/tracker_final.loc[:, "USN NON ENG Target/hr"]))]


tracker_final.loc[:, "Achieved USN Weighted"]= tracker_final.loc[:, "temp1"]+tracker_final.loc[:, "temp2"]

tracker_final= tracker_final.drop(columns=['temp1', 'temp2'])

tracker_final.loc[:, "temp1"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "USN EN"], tracker_final.loc[:, "USN EN Count"])]

tracker_final.loc[:, "temp2"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "USN Non EN"], tracker_final.loc[:, "USN Non EN Count"])]

tracker_final.loc[:, "USN Total Achieved"]= tracker_final.loc[:, "temp1"]+tracker_final.loc[:, "temp2"]

tracker_final= tracker_final.drop(columns=['temp1', 'temp2'])

tracker_final.loc[:, "temp1"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "USN EN"], (tracker_final.loc[:,"USN EN"]*tracker_final.loc[:, "USN ENG Target/hr"]))]

tracker_final.loc[:, "temp2"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "USN Non EN"], (tracker_final.loc[:,"USN Non EN"]*tracker_final.loc[:, "USN NON ENG Target/hr"]))]

tracker_final.loc[:, "USN Total Target"]= tracker_final.loc[:, "temp1"]+tracker_final.loc[:, "temp2"]

tracker_final= tracker_final.drop(columns=['temp1', 'temp2'])

tracker_final.loc[:, "USN %"]=[0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "USN Total Achieved"],tracker_final.loc[:, "USN Total Target"],
                                                                         (tracker_final.loc[:, "USN Total Achieved"]/tracker_final.loc[:, "USN Total Target"]))]
tracker_final['USN %']=tracker_final['USN %']*100

tracker_final['USN %'].fillna(0, inplace=True)



tracker_final=tracker_final[['Login','Reporting Manager', 'Program', 'Year','Week','Date', 'End Date','Leave','Leave Type','Leave Hours',
                             'Instances','Bincheck EPCR Hours','Bincheck EPCR Count','Bincheck USN Hours','Bincheck USN Count',
                             'Bincheck IXD Hours', 'Bincheck IXD Count','Bincheck CCR Hours','Bincheck CCR Count','Bincheck Target/hr',
                             'Achieved Bincheck Weighted','Bincheck Total Achieved', 'Bincheck Total Target','Bincheck %','CCR EN','CCR EN Count','CCR Non EN','CCR Non EN Count','Brand Owner Audits',
                             'Brand Owner Audits Count','Non AVP Audits EN','Non AVP Audits EN Count','Non AVP Audits Non EN','Non AVP Audits Non EN Count',
                             'CCR Target/hr','BOA NA IN Target/hr', 'NON AVP Target/hr', 'Achieved CCR Weighted', 'CCR Total Achieved', 'CCR Total Target','CCR %',
                             'EIM FP Hours','EIM FP Count','EIM FP QA Hours','EIM FP QA Count',
                             'EIM Target/hr', 'Achieved EIM FP Weighted','EIM FP Total Achieved', 'EIM FP Total Target', 'EIM FP %','EIM Esc Hours','EIM Esc Count','EIM Esc QA Hours',
                             'EIM Esc QA Count','EIM ESC Target/hr', 'Achieved EIM Esc Weighted', 'EIM Esc Total Achieved', 'EIM Esc Total Target', 'EIM Escalations %',
                             'EPCR EN','EPCR EN Count','EPCR Non EN','EPCR Non EN Count','EACP Hours','EACP Count','EACP Appeals Hours','EACP Appeals Count',
                             'EACP Appeals Target/hr', 'EPCR ENG Target/hr', 'EPCR NON ENG Target/hr','EACP Target/hr','Achieved EPCR Weighted','EPCR Total Achieved',
                             'EPCR Total Target','EPCR %',
                             'MCM EN','MCM EN Count','MCM Non EN','MCM Non EN Count','MCM Live EN QA',
                             'MCM Live EN QA Count','MCM Live Non EN QA','MCM Live Non EN QA Count',
                             'MCM Others NE','MCM Others NE Count','MCM Practice File hours','MCM Practice Files Count',
                             'MCM ENG Target/hr', 'MCM NON ENG Target/hr','MCM Others NE Target/hr',
                             'Achieved MCM Weighted','MCM Total Achieved', 'MCM Total Target', 'MCM %',
                             'Safety Sensitivity Hours','Safety Sensitivity Count','​​Safety Sensitivity Target/hr',
                             'Safety Sensitivity Total Achieved', 'Safety Sensitivity Total Target','Achieved Safety Sensitivity Weighted','Safety Sensitivity %',
                             'Test Buy Hrs',
                             'Test Buy Count','GPV -Test Buy Target/hr','USN EN','USN EN Count','USN Non EN','USN Non EN Count','USN ENG Target/hr',
                             'USN NON ENG Target/hr','Achieved USN Weighted','USN Total Achieved','USN Total Target','USN %','DCR KAI EN(Live)',
                             'DCR KAI Non EN (Live)','DCR Live EN QA','DCR Live Non EN QA','DCR Practice File hours','WI KAI EN (Live)','WI KAI Non EN (Live)','WI Live EN QA',
                             'WI Live Non EN QA','WI Practice File Hours','SLIM Hours','PAD Hours','Coverage Hours','Retails Hours',
                             'Backlog & Ops Update','HTMS Data Reporting','Flash','Keyword Analysis',
                             'MOM','Ops Tracker','Other Reportings','PR Doc','Productivity Report',
                             'WBR Report','Work Allocation','EN QA Hours','QA Reporting',
                             'Non EN QA Hrs','QA Calibration','Deep Dive Hours','Deep Dive Name',
                             'Project Name','Project Hours',
                             'Other Task Name','Other Tasks Hours','Meeting',
                             'No Work','Other NPTS','System Issues','Training','RAMP Hours','Quality Hours',
                             'Audits', 'QA Meeting','Adhoc-Business Name', 'Adhoc-Business Hours','Adhoc-Internal Name', 'Adhoc-Internal Hours',
                             'Comments','Productive Hours','NPT SUM','NPT Hours','Leave Sum','Total Hours']]

tracker_final= tracker_final.round({"Achieved USN Weighted":2, "USN %":2,'USN Total Achieved': 2, 'USN Total Target':2})

convert_dict = {'Achieved USN Weighted':'float32','USN Total Achieved':'float32',
            'USN Total Target': 'float32','USN %':'float32'}
  
tracker_final = tracker_final.astype(convert_dict)

tracker_final['Achieved USN Weighted'] = tracker_final['Achieved USN Weighted'].round(2)
tracker_final['USN Total Achieved'] = tracker_final['USN Total Achieved'].round(2)
tracker_final['USN Total Target'] = tracker_final['USN Total Target'].round(2)
tracker_final['USN %'] = tracker_final['USN %'].round(2)


usn=""

#wrong item

wi=pd.read_excel(filename2, sheet_name='WI_Data')

wi=pd.DataFrame(wi)

wi=wi[['Login', 'Date', 'WI Practice File Count']]

wi.loc[:,'Date']=pd.to_datetime(wi.loc[:,'Date']).dt.date

wi['WI Practice File Count'].fillna(0, inplace=True)

wi.drop_duplicates(subset=['Date', 'Login'], inplace=True)

convert_dict = {'Login':'category', 'WI Practice File Count':'float16'}
  
wi = wi.astype(convert_dict)

tracker_final=pd.merge(tracker_final, wi, how='left', on=['Date', 'Login'])

tracker_final['WI Practice File Count'].fillna(0, inplace=True)

tracker_final=pd.merge(tracker_final, target_data[['Week', 'WI Target/hr']], how='left', on=['Week'])

convert_dict = {'Week':'int8','WI Target/hr':'float16'}
  
tracker_final = tracker_final.astype(convert_dict)

tracker_final.loc[:, "temp1"]= [0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "WI Practice File Hours"],
                                                                  tracker_final.loc[:, "WI Target/hr"], (tracker_final.loc[:, "WI Practice File Count"]/tracker_final.loc[:, "WI Target/hr"]))]
                       
tracker_final.loc[:, "Achieved WI Weighted"]= tracker_final.loc[:, "temp1"]

tracker_final= tracker_final.drop(columns=['temp1'])

tracker_final.loc[:, "temp1"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "WI Practice File Hours"], tracker_final.loc[:, "WI Practice File Count"])]

tracker_final.loc[:, "WI Total Achieved"]= tracker_final.loc[:, "temp1"]

tracker_final= tracker_final.drop(columns=['temp1'])

tracker_final.loc[:, "WI Total Target"]= (tracker_final.loc[:,"WI Practice File Hours"]*tracker_final.loc[:, "WI Target/hr"])

tracker_final.loc[:, "WI %"]=[0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "WI Total Achieved"],tracker_final.loc[:, "WI Total Target"],
                                                                        (tracker_final.loc[:, "WI Total Achieved"]/tracker_final.loc[:, "WI Total Target"]))]
tracker_final['WI %']=tracker_final['WI %']*100

tracker_final['WI %'].fillna(0, inplace=True)

tracker_final=tracker_final[['Login','Reporting Manager', 'Program', 'Year','Week','Date', 'End Date','Leave','Leave Type','Leave Hours',
                             'Instances','Bincheck EPCR Hours','Bincheck EPCR Count','Bincheck USN Hours','Bincheck USN Count',
                             'Bincheck IXD Hours', 'Bincheck IXD Count','Bincheck CCR Hours','Bincheck CCR Count','Bincheck Target/hr',
                             'Achieved Bincheck Weighted','Bincheck Total Achieved', 'Bincheck Total Target','Bincheck %','CCR EN','CCR EN Count','CCR Non EN','CCR Non EN Count','Brand Owner Audits',
                             'Brand Owner Audits Count','Non AVP Audits EN','Non AVP Audits EN Count','Non AVP Audits Non EN','Non AVP Audits Non EN Count',
                             'CCR Target/hr','BOA NA IN Target/hr', 'NON AVP Target/hr', 'Achieved CCR Weighted', 'CCR Total Achieved', 'CCR Total Target','CCR %',
                             'EIM FP Hours','EIM FP Count','EIM FP QA Hours','EIM FP QA Count',
                             'EIM Target/hr', 'Achieved EIM FP Weighted','EIM FP Total Achieved', 'EIM FP Total Target', 'EIM FP %','EIM Esc Hours','EIM Esc Count','EIM Esc QA Hours',
                             'EIM Esc QA Count','EIM ESC Target/hr', 'Achieved EIM Esc Weighted', 'EIM Esc Total Achieved', 'EIM Esc Total Target', 'EIM Escalations %',
                             'EPCR EN','EPCR EN Count','EPCR Non EN','EPCR Non EN Count','EACP Hours','EACP Count','EACP Appeals Hours','EACP Appeals Count',
                             'EACP Appeals Target/hr', 'EPCR ENG Target/hr', 'EPCR NON ENG Target/hr','EACP Target/hr','Achieved EPCR Weighted','EPCR Total Achieved',
                             'EPCR Total Target','EPCR %',
                             'MCM EN','MCM EN Count','MCM Non EN','MCM Non EN Count','MCM Live EN QA',
                             'MCM Live EN QA Count','MCM Live Non EN QA','MCM Live Non EN QA Count',
                             'MCM Others NE','MCM Others NE Count','MCM Practice File hours','MCM Practice Files Count',
                             'MCM ENG Target/hr', 'MCM NON ENG Target/hr','MCM Others NE Target/hr',
                             'Achieved MCM Weighted','MCM Total Achieved', 'MCM Total Target', 'MCM %',
                             'Safety Sensitivity Hours','Safety Sensitivity Count','​​Safety Sensitivity Target/hr',
                             'Safety Sensitivity Total Achieved', 'Safety Sensitivity Total Target','Achieved Safety Sensitivity Weighted','Safety Sensitivity %',
                             'Test Buy Hrs',
                             'Test Buy Count','GPV -Test Buy Target/hr','USN EN','USN EN Count','USN Non EN','USN Non EN Count','USN ENG Target/hr',
                             'USN NON ENG Target/hr','Achieved USN Weighted','USN Total Achieved','USN Total Target','USN %','DCR KAI EN(Live)',
                             'DCR KAI Non EN (Live)','DCR Live EN QA','DCR Live Non EN QA','DCR Practice File hours','WI KAI EN (Live)','WI KAI Non EN (Live)','WI Live EN QA',
                             'WI Live Non EN QA','WI Practice File Hours','WI Practice File Count','WI Target/hr','Achieved WI Weighted','WI Total Achieved','WI Total Target','WI %',
                             'SLIM Hours','PAD Hours','Coverage Hours','Retails Hours',
                             'Backlog & Ops Update','HTMS Data Reporting','Flash','Keyword Analysis',
                             'MOM','Ops Tracker','Other Reportings','PR Doc','Productivity Report',
                             'WBR Report','Work Allocation','EN QA Hours','QA Reporting',
                             'Non EN QA Hrs','QA Calibration','Deep Dive Hours','Deep Dive Name',
                             'Project Name','Project Hours',
                             'Other Task Name','Other Tasks Hours','Meeting',
                             'No Work','Other NPTS','System Issues','Training','RAMP Hours','Quality Hours',
                             'Audits', 'QA Meeting','Adhoc-Business Name', 'Adhoc-Business Hours','Adhoc-Internal Name', 'Adhoc-Internal Hours',
                             'Comments','Productive Hours','NPT SUM','NPT Hours','Leave Sum','Total Hours']]

tracker_final= tracker_final.round({"Achieved WI Weighted":2, "WI %":2,'WI Total Achieved': 2, 'WI Total Target':2})

convert_dict = {'Achieved WI Weighted':'float32','WI Total Achieved':'float32',
            'WI Total Target': 'float32','WI %':'float32'}
  
tracker_final = tracker_final.astype(convert_dict)

tracker_final['Achieved WI Weighted'] = tracker_final['Achieved WI Weighted'].round(2)
tracker_final['WI Total Achieved'] = tracker_final['WI Total Achieved'].round(2)
tracker_final['WI Total Target'] = tracker_final['WI Total Target'].round(2)
tracker_final['WI %'] = tracker_final['WI %'].round(2)


wi=""

#slim

slim=pd.read_excel(filename2, sheet_name='Slim_data')

slim=pd.DataFrame(slim)

slim=slim[['Login', 'Date', 'SLIM Count']]

slim.loc[:,'Date']=pd.to_datetime(slim.loc[:,'Date']).dt.date

slim['SLIM Count'].fillna(0, inplace=True)

slim.drop_duplicates(subset=['Date', 'Login'], inplace=True)

convert_dict = {'Login':'category', 'SLIM Count':'float16'}
  
slim = slim.astype(convert_dict)

tracker_final=pd.merge(tracker_final, slim, how='left', on=['Date', 'Login'])

tracker_final['SLIM Count'].fillna(0, inplace=True)

tracker_final=pd.merge(tracker_final, target_data[['Week', 'SLIM Target/hr']], how='left', on=['Week'])

convert_dict = {'Week':'int8','SLIM Target/hr':'float16'}
  
tracker_final = tracker_final.astype(convert_dict)

tracker_final.loc[:, "temp1"]= [0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "SLIM Hours"],
                                                                  tracker_final.loc[:, "SLIM Target/hr"], (tracker_final.loc[:, "SLIM Count"]/tracker_final.loc[:, "SLIM Target/hr"]))]
                       
tracker_final.loc[:, "Achieved SLIM Weighted"]= tracker_final.loc[:, "temp1"]

tracker_final= tracker_final.drop(columns=['temp1'])

tracker_final.loc[:, "temp1"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "SLIM Hours"], tracker_final.loc[:, "SLIM Count"])]

tracker_final.loc[:, "SLIM Total Achieved"]= tracker_final.loc[:, "temp1"]

tracker_final= tracker_final.drop(columns=['temp1'])

tracker_final.loc[:, "SLIM Total Target"]= (tracker_final.loc[:,"SLIM Hours"]*tracker_final.loc[:, "SLIM Target/hr"])

tracker_final.loc[:, "SLIM %"]=[0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "SLIM Total Achieved"],tracker_final.loc[:, "SLIM Total Target"],
                                                                         (tracker_final.loc[:, "SLIM Total Achieved"]/tracker_final.loc[:, "SLIM Total Target"]))]
tracker_final['SLIM %']=tracker_final['SLIM %']*100

tracker_final['SLIM %'].fillna(0, inplace=True)

tracker_final=tracker_final[['Login','Reporting Manager', 'Program', 'Year','Week','Date', 'End Date','Leave','Leave Type','Leave Hours',
                             'Instances','Bincheck EPCR Hours','Bincheck EPCR Count','Bincheck USN Hours','Bincheck USN Count',
                             'Bincheck IXD Hours', 'Bincheck IXD Count','Bincheck CCR Hours','Bincheck CCR Count','Bincheck Target/hr',
                             'Achieved Bincheck Weighted','Bincheck Total Achieved', 'Bincheck Total Target','Bincheck %','CCR EN','CCR EN Count','CCR Non EN','CCR Non EN Count','Brand Owner Audits',
                             'Brand Owner Audits Count','Non AVP Audits EN','Non AVP Audits EN Count','Non AVP Audits Non EN','Non AVP Audits Non EN Count',
                             'CCR Target/hr','BOA NA IN Target/hr', 'NON AVP Target/hr', 'Achieved CCR Weighted', 'CCR Total Achieved', 'CCR Total Target','CCR %',
                             'EIM FP Hours','EIM FP Count','EIM FP QA Hours','EIM FP QA Count',
                             'EIM Target/hr', 'Achieved EIM FP Weighted','EIM FP Total Achieved', 'EIM FP Total Target', 'EIM FP %','EIM Esc Hours','EIM Esc Count','EIM Esc QA Hours',
                             'EIM Esc QA Count','EIM ESC Target/hr', 'Achieved EIM Esc Weighted', 'EIM Esc Total Achieved', 'EIM Esc Total Target', 'EIM Escalations %',
                             'EPCR EN','EPCR EN Count','EPCR Non EN','EPCR Non EN Count','EACP Hours','EACP Count','EACP Appeals Hours','EACP Appeals Count',
                             'EACP Appeals Target/hr', 'EPCR ENG Target/hr', 'EPCR NON ENG Target/hr','EACP Target/hr','Achieved EPCR Weighted','EPCR Total Achieved',
                             'EPCR Total Target','EPCR %',
                             'MCM EN','MCM EN Count','MCM Non EN','MCM Non EN Count','MCM Live EN QA',
                             'MCM Live EN QA Count','MCM Live Non EN QA','MCM Live Non EN QA Count',
                             'MCM Others NE','MCM Others NE Count','MCM Practice File hours','MCM Practice Files Count',
                             'MCM ENG Target/hr', 'MCM NON ENG Target/hr','MCM Others NE Target/hr',
                             'Achieved MCM Weighted','MCM Total Achieved', 'MCM Total Target', 'MCM %',
                             'Safety Sensitivity Hours','Safety Sensitivity Count','​​Safety Sensitivity Target/hr',
                             'Safety Sensitivity Total Achieved', 'Safety Sensitivity Total Target','Achieved Safety Sensitivity Weighted','Safety Sensitivity %',
                             'Test Buy Hrs',
                             'Test Buy Count','GPV -Test Buy Target/hr','USN EN','USN EN Count','USN Non EN','USN Non EN Count','USN ENG Target/hr',
                             'USN NON ENG Target/hr','Achieved USN Weighted','USN Total Achieved','USN Total Target','USN %','DCR KAI EN(Live)',
                             'DCR KAI Non EN (Live)','DCR Live EN QA','DCR Live Non EN QA','DCR Practice File hours','WI KAI EN (Live)','WI KAI Non EN (Live)','WI Live EN QA',
                             'WI Live Non EN QA','WI Practice File Hours','WI Practice File Count','WI Target/hr','Achieved WI Weighted','WI Total Achieved','WI Total Target','WI %',
                             'SLIM Hours', 'SLIM Count', 'SLIM Target/hr', 'Achieved SLIM Weighted', 'SLIM Total Achieved', 'SLIM Total Target', 'SLIM %',
                             'PAD Hours','Coverage Hours','Retails Hours',
                             'Backlog & Ops Update','HTMS Data Reporting','Flash','Keyword Analysis',
                             'MOM','Ops Tracker','Other Reportings','PR Doc','Productivity Report',
                             'WBR Report','Work Allocation','EN QA Hours','QA Reporting',
                             'Non EN QA Hrs','QA Calibration','Deep Dive Hours','Deep Dive Name',
                             'Project Name','Project Hours',
                             'Other Task Name','Other Tasks Hours','Meeting',
                             'No Work','Other NPTS','System Issues','Training','RAMP Hours','Quality Hours',
                             'Audits', 'QA Meeting','Adhoc-Business Name', 'Adhoc-Business Hours','Adhoc-Internal Name', 'Adhoc-Internal Hours',
                             'Comments','Productive Hours','NPT SUM','NPT Hours','Leave Sum','Total Hours']]

tracker_final= tracker_final.round({"Achieved SLIM Weighted":2, "SLIM %":2,'SLIM Total Achieved': 2, 'SLIM Total Target':2})

convert_dict = {'Achieved SLIM Weighted':'float32','SLIM Total Achieved':'float32',
            'SLIM Total Target': 'float32','SLIM %':'float32'}
  
tracker_final = tracker_final.astype(convert_dict)

tracker_final['Achieved SLIM Weighted'] = tracker_final['Achieved SLIM Weighted'].round(2)
tracker_final['SLIM Total Achieved'] = tracker_final['SLIM Total Achieved'].round(2)
tracker_final['SLIM Total Target'] = tracker_final['SLIM Total Target'].round(2)
tracker_final['SLIM %'] = tracker_final['SLIM %'].round(2)


slim=""
#coverage

coverage=pd.read_excel(filename2, sheet_name='Coverage_data')

coverage=pd.DataFrame(coverage)

coverage=coverage[['Login', 'Date', 'Coverage Count']]

coverage.loc[:,'Date']=pd.to_datetime(coverage.loc[:,'Date']).dt.date

coverage['Coverage Count'].fillna(0, inplace=True)

coverage.drop_duplicates(subset=['Date', 'Login'], inplace=True)

convert_dict = {'Login':'category', 'Coverage Count':'float16'}
  
coverage = coverage.astype(convert_dict)

tracker_final=pd.merge(tracker_final, coverage, how='left', on=['Date', 'Login'])

tracker_final['Coverage Count'].fillna(0, inplace=True)

tracker_final=pd.merge(tracker_final, target_data[['Week', 'Coverage Target/hr']], how='left', on=['Week'])

convert_dict = {'Week':'int8','Coverage Target/hr':'float16'}
  
tracker_final = tracker_final.astype(convert_dict)

tracker_final.loc[:, "temp1"]= [0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "Coverage Hours"],
                                                                  tracker_final.loc[:, "Coverage Target/hr"], (tracker_final.loc[:, "Coverage Count"]/tracker_final.loc[:, "Coverage Target/hr"]))]
                       
tracker_final.loc[:, "Achieved Coverage Weighted"]= tracker_final.loc[:, "temp1"]

tracker_final= tracker_final.drop(columns=['temp1'])

tracker_final.loc[:, "temp1"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "Coverage Hours"], tracker_final.loc[:, "Coverage Count"])]

tracker_final.loc[:, "Coverage Total Achieved"]= tracker_final.loc[:, "temp1"]

tracker_final= tracker_final.drop(columns=['temp1'])

tracker_final.loc[:, "Coverage Total Target"]= (tracker_final.loc[:,"Coverage Hours"]*tracker_final.loc[:, "Coverage Target/hr"])

tracker_final.loc[:, "Coverage %"]=[0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "Coverage Total Achieved"],tracker_final.loc[:, "Coverage Total Target"],
                                                                         (tracker_final.loc[:, "Coverage Total Achieved"]/tracker_final.loc[:, "Coverage Total Target"]))]
tracker_final['Coverage %']=tracker_final['Coverage %']*100

tracker_final['Coverage %'].fillna(0, inplace=True)

tracker_final=tracker_final[['Login','Reporting Manager', 'Program', 'Year','Week','Date', 'End Date','Leave','Leave Type','Leave Hours',
                             'Instances','Bincheck EPCR Hours','Bincheck EPCR Count','Bincheck USN Hours','Bincheck USN Count',
                             'Bincheck IXD Hours', 'Bincheck IXD Count','Bincheck CCR Hours','Bincheck CCR Count','Bincheck Target/hr',
                             'Achieved Bincheck Weighted','Bincheck Total Achieved', 'Bincheck Total Target','Bincheck %','CCR EN','CCR EN Count','CCR Non EN','CCR Non EN Count','Brand Owner Audits',
                             'Brand Owner Audits Count','Non AVP Audits EN','Non AVP Audits EN Count','Non AVP Audits Non EN','Non AVP Audits Non EN Count',
                             'CCR Target/hr','BOA NA IN Target/hr', 'NON AVP Target/hr', 'Achieved CCR Weighted', 'CCR Total Achieved', 'CCR Total Target','CCR %',
                             'EIM FP Hours','EIM FP Count','EIM FP QA Hours','EIM FP QA Count',
                             'EIM Target/hr', 'Achieved EIM FP Weighted','EIM FP Total Achieved', 'EIM FP Total Target', 'EIM FP %','EIM Esc Hours','EIM Esc Count','EIM Esc QA Hours',
                             'EIM Esc QA Count','EIM ESC Target/hr', 'Achieved EIM Esc Weighted', 'EIM Esc Total Achieved', 'EIM Esc Total Target', 'EIM Escalations %',
                             'EPCR EN','EPCR EN Count','EPCR Non EN','EPCR Non EN Count','EACP Hours','EACP Count','EACP Appeals Hours','EACP Appeals Count',
                             'EACP Appeals Target/hr', 'EPCR ENG Target/hr', 'EPCR NON ENG Target/hr','EACP Target/hr','Achieved EPCR Weighted','EPCR Total Achieved',
                             'EPCR Total Target','EPCR %',
                             'MCM EN','MCM EN Count','MCM Non EN','MCM Non EN Count','MCM Live EN QA',
                             'MCM Live EN QA Count','MCM Live Non EN QA','MCM Live Non EN QA Count',
                             'MCM Others NE','MCM Others NE Count','MCM Practice File hours','MCM Practice Files Count',
                             'MCM ENG Target/hr', 'MCM NON ENG Target/hr','MCM Others NE Target/hr',
                             'Achieved MCM Weighted','MCM Total Achieved', 'MCM Total Target', 'MCM %',
                             'Safety Sensitivity Hours','Safety Sensitivity Count','​​Safety Sensitivity Target/hr',
                             'Safety Sensitivity Total Achieved', 'Safety Sensitivity Total Target','Achieved Safety Sensitivity Weighted','Safety Sensitivity %',
                             'Test Buy Hrs',
                             'Test Buy Count','GPV -Test Buy Target/hr','USN EN','USN EN Count','USN Non EN','USN Non EN Count','USN ENG Target/hr',
                             'USN NON ENG Target/hr','Achieved USN Weighted','USN Total Achieved','USN Total Target','USN %','DCR KAI EN(Live)',
                             'DCR KAI Non EN (Live)','DCR Live EN QA','DCR Live Non EN QA','DCR Practice File hours','WI KAI EN (Live)','WI KAI Non EN (Live)','WI Live EN QA',
                             'WI Live Non EN QA','WI Practice File Hours','WI Practice File Count','WI Target/hr','Achieved WI Weighted','WI Total Achieved','WI Total Target','WI %',
                             'SLIM Hours','SLIM Count','SLIM Target/hr', 'Achieved SLIM Weighted', 'SLIM Total Achieved', 'SLIM Total Target', 'SLIM %',
                             'PAD Hours','Coverage Hours','Coverage Count','Coverage Target/hr','Achieved Coverage Weighted','Coverage Total Achieved',
                             'Coverage Total Target','Coverage %','Retails Hours',
                             'Backlog & Ops Update','HTMS Data Reporting','Flash','Keyword Analysis',
                             'MOM','Ops Tracker','Other Reportings','PR Doc','Productivity Report',
                             'WBR Report','Work Allocation','EN QA Hours','QA Reporting',
                             'Non EN QA Hrs','QA Calibration','Deep Dive Hours','Deep Dive Name',
                             'Project Name','Project Hours',
                             'Other Task Name','Other Tasks Hours','Meeting',
                             'No Work','Other NPTS','System Issues','Training','RAMP Hours','Quality Hours',
                             'Audits', 'QA Meeting','Adhoc-Business Name', 'Adhoc-Business Hours','Adhoc-Internal Name', 'Adhoc-Internal Hours',
                             'Comments','Productive Hours','NPT SUM','NPT Hours','Leave Sum','Total Hours']]

tracker_final= tracker_final.round({"Achieved Coverage Weighted":2, "Coverage %":2,'Coverage Total Achieved': 2, 'Coverage Total Target':2})

convert_dict = {'Achieved Coverage Weighted':'float32','Coverage Total Achieved':'float32',
            'Coverage Total Target': 'float32','Coverage %':'float32'}
  
tracker_final = tracker_final.astype(convert_dict)

tracker_final['Achieved Coverage Weighted'] = tracker_final['Achieved Coverage Weighted'].round(2)
tracker_final['Coverage Total Achieved'] = tracker_final['Coverage Total Achieved'].round(2)
tracker_final['Coverage Total Target'] = tracker_final['Coverage Total Target'].round(2)
tracker_final['Coverage %'] = tracker_final['Coverage %'].round(2)


coverage=""

#retail

retail=pd.read_excel(filename2, sheet_name='Retail_data')

retail=pd.DataFrame(retail)

retail=retail[['Login', 'Date', 'Retails Count']]

retail.loc[:,'Date']=pd.to_datetime(retail.loc[:,'Date']).dt.date

retail['Retails Count'].fillna(0, inplace=True)

retail.drop_duplicates(subset=['Date', 'Login'], inplace=True)

convert_dict = {'Login':'category', 'Retails Count':'float16'}
  
retail = retail.astype(convert_dict)

tracker_final=pd.merge(tracker_final, retail, how='left', on=['Date', 'Login'])

tracker_final['Retails Count'].fillna(0, inplace=True)

tracker_final=pd.merge(tracker_final, target_data[['Week', 'Retails Target/hr']], how='left', on=['Week'])

convert_dict = {'Week':'int8','Retails Target/hr':'float16'}
  
tracker_final = tracker_final.astype(convert_dict)

tracker_final.loc[:, "temp1"]= [0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "Retails Hours"],
                                                                  tracker_final.loc[:, "Retails Target/hr"], (tracker_final.loc[:, "Retails Count"]/tracker_final.loc[:, "Retails Target/hr"]))]
                       
tracker_final.loc[:, "Achieved Retails Weighted"]= tracker_final.loc[:, "temp1"]

tracker_final= tracker_final.drop(columns=['temp1'])

tracker_final.loc[:, "temp1"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "Retails Hours"], tracker_final.loc[:, "Retails Count"])]

tracker_final.loc[:, "Retails Total Achieved"]= tracker_final.loc[:, "temp1"]

tracker_final= tracker_final.drop(columns=['temp1'])

tracker_final.loc[:, "Retails Total Target"]= (tracker_final.loc[:,"Retails Hours"]*tracker_final.loc[:, "Retails Target/hr"])

tracker_final.loc[:, "Retails %"]=[0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "Retails Total Achieved"],tracker_final.loc[:, "Retails Total Target"],
                                                                         (tracker_final.loc[:, "Retails Total Achieved"]/tracker_final.loc[:, "Retails Total Target"]))]
tracker_final['Retails %']=tracker_final['Retails %']*100

tracker_final['Retails %'].fillna(0, inplace=True)

tracker_final=tracker_final[['Login','Reporting Manager', 'Program', 'Year','Week','Date', 'End Date','Leave','Leave Type','Leave Hours',
                             'Instances','Bincheck EPCR Hours','Bincheck EPCR Count','Bincheck USN Hours','Bincheck USN Count',
                             'Bincheck IXD Hours', 'Bincheck IXD Count','Bincheck CCR Hours','Bincheck CCR Count','Bincheck Target/hr',
                             'Achieved Bincheck Weighted','Bincheck Total Achieved', 'Bincheck Total Target','Bincheck %','CCR EN','CCR EN Count','CCR Non EN','CCR Non EN Count','Brand Owner Audits',
                             'Brand Owner Audits Count','Non AVP Audits EN','Non AVP Audits EN Count','Non AVP Audits Non EN','Non AVP Audits Non EN Count',
                             'CCR Target/hr','BOA NA IN Target/hr', 'NON AVP Target/hr', 'Achieved CCR Weighted', 'CCR Total Achieved', 'CCR Total Target','CCR %',
                             'EIM FP Hours','EIM FP Count','EIM FP QA Hours','EIM FP QA Count',
                             'EIM Target/hr', 'Achieved EIM FP Weighted','EIM FP Total Achieved', 'EIM FP Total Target', 'EIM FP %','EIM Esc Hours','EIM Esc Count','EIM Esc QA Hours',
                             'EIM Esc QA Count','EIM ESC Target/hr', 'Achieved EIM Esc Weighted', 'EIM Esc Total Achieved', 'EIM Esc Total Target', 'EIM Escalations %',
                             'EPCR EN','EPCR EN Count','EPCR Non EN','EPCR Non EN Count','EACP Hours','EACP Count','EACP Appeals Hours','EACP Appeals Count',
                             'EACP Appeals Target/hr', 'EPCR ENG Target/hr', 'EPCR NON ENG Target/hr','EACP Target/hr','Achieved EPCR Weighted','EPCR Total Achieved',
                             'EPCR Total Target','EPCR %',
                             'MCM EN','MCM EN Count','MCM Non EN','MCM Non EN Count','MCM Live EN QA',
                             'MCM Live EN QA Count','MCM Live Non EN QA','MCM Live Non EN QA Count',
                             'MCM Others NE','MCM Others NE Count','MCM Practice File hours','MCM Practice Files Count',
                             'MCM ENG Target/hr', 'MCM NON ENG Target/hr','MCM Others NE Target/hr',
                             'Achieved MCM Weighted','MCM Total Achieved', 'MCM Total Target', 'MCM %',
                             'Safety Sensitivity Hours','Safety Sensitivity Count','​​Safety Sensitivity Target/hr',
                             'Safety Sensitivity Total Achieved', 'Safety Sensitivity Total Target','Achieved Safety Sensitivity Weighted','Safety Sensitivity %',
                             'Test Buy Hrs',
                             'Test Buy Count','GPV -Test Buy Target/hr','USN EN','USN EN Count','USN Non EN','USN Non EN Count','USN ENG Target/hr',
                             'USN NON ENG Target/hr','Achieved USN Weighted','USN Total Achieved','USN Total Target','USN %','DCR KAI EN(Live)',
                             'DCR KAI Non EN (Live)','DCR Live EN QA','DCR Live Non EN QA','DCR Practice File hours','WI KAI EN (Live)','WI KAI Non EN (Live)','WI Live EN QA',
                             'WI Live Non EN QA','WI Practice File Hours','WI Practice File Count','WI Target/hr','Achieved WI Weighted','WI Total Achieved','WI Total Target','WI %',
                             'SLIM Hours','SLIM Count','SLIM Target/hr', 'Achieved SLIM Weighted', 'SLIM Total Achieved', 'SLIM Total Target', 'SLIM %',
                             'PAD Hours','Coverage Hours','Coverage Count','Coverage Target/hr','Achieved Coverage Weighted','Coverage Total Achieved',
                             'Coverage Total Target','Coverage %','Retails Hours', 'Retails Count', 'Retails Target/hr','Achieved Retails Weighted','Retails Total Achieved','Retails Total Target','Retails %',
                             'Backlog & Ops Update','HTMS Data Reporting','Flash','Keyword Analysis',
                             'MOM','Ops Tracker','Other Reportings','PR Doc','Productivity Report',
                             'WBR Report','Work Allocation','EN QA Hours','QA Reporting',
                             'Non EN QA Hrs','QA Calibration','Deep Dive Hours','Deep Dive Name',
                             'Project Name','Project Hours',
                             'Other Task Name','Other Tasks Hours','Meeting',
                             'No Work','Other NPTS','System Issues','Training','RAMP Hours','Quality Hours',
                             'Audits', 'QA Meeting','Adhoc-Business Name', 'Adhoc-Business Hours','Adhoc-Internal Name', 'Adhoc-Internal Hours',
                             'Comments','Productive Hours','NPT SUM','NPT Hours','Leave Sum','Total Hours']]

tracker_final= tracker_final.round({"Achieved Retails Weighted":2, "Retails %":2,'Retails Total Achieved': 2, 'Retails Total Target':2})

convert_dict = {'Achieved Retails Weighted':'float32','Retails Total Achieved':'float32',
            'Retails Total Target': 'float32','Retails %':'float32'}
  
tracker_final = tracker_final.astype(convert_dict)

tracker_final['Achieved Retails Weighted'] = tracker_final['Achieved Retails Weighted'].round(2)
tracker_final['Retails Total Achieved'] = tracker_final['Retails Total Achieved'].round(2)
tracker_final['Retails Total Target'] = tracker_final['Retails Total Target'].round(2)
tracker_final['Retails %'] = tracker_final['Retails %'].round(2)


retail=""

#DCR

dcr=pd.read_excel(filename2, sheet_name='DCR_Data')

dcr=pd.DataFrame(dcr)

dcr=dcr[['Login', 'Date', 'DCR Practice Count','DCR KAI(Live)']]

dcr.loc[:,'Date']=pd.to_datetime(dcr.loc[:,'Date']).dt.date

dcr['DCR Practice Count'].fillna(0, inplace=True)

dcr['DCR KAI(Live)'].fillna(0, inplace=True)

dcr.drop_duplicates(subset=['Date', 'Login'], inplace=True)

convert_dict = {'Login':'category', 'DCR Practice Count':'float16', 'DCR KAI(Live)':'float16'}
  
dcr = dcr.astype(convert_dict)

tracker_final=pd.merge(tracker_final, dcr, how='left', on=['Date', 'Login'])

tracker_final['DCR Practice Count'].fillna(0, inplace=True)

tracker_final['DCR KAI(Live)'].fillna(0, inplace=True)

tracker_final=pd.merge(tracker_final, target_data[['Week', 'DCR Target/hr']], how='left', on=['Week'])

convert_dict = {'Week':'int8','DCR Target/hr':'float16'}
  
tracker_final = tracker_final.astype(convert_dict)

tracker_final.loc[:, "temp1"]= [0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "DCR Practice File hours"],
                                                                          tracker_final.loc[:, "DCR Target/hr"], (tracker_final.loc[:, "DCR Practice Count"]/tracker_final.loc[:, "DCR Target/hr"]))]
   
tracker_final.loc[:, "temp2"]= [0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "DCR KAI EN(Live)"],
                                                                          tracker_final.loc[:, "DCR Target/hr"], (tracker_final.loc[:, "DCR KAI(Live)"]/tracker_final.loc[:, "DCR Target/hr"]))]
   
                    
tracker_final.loc[:, "Achieved DCR Weighted"]= tracker_final.loc[:, "temp1"]+tracker_final.loc[:, "temp2"]

tracker_final= tracker_final.drop(columns=['temp1', 'temp2'])

tracker_final.loc[:, "temp1"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "DCR Practice File hours"], tracker_final.loc[:, "DCR Practice Count"])]

tracker_final.loc[:, "temp2"]=[0 if x==0 else y for x,y in zip(tracker_final.loc[:, "DCR KAI EN(Live)"], tracker_final.loc[:, "DCR KAI(Live)"])]

tracker_final.loc[:, "DCR Total Achieved"]= tracker_final.loc[:, "temp1"]+tracker_final.loc[:, "temp2"]

tracker_final= tracker_final.drop(columns=['temp1', 'temp2'])

tracker_final.loc[:, "temp1"]= (tracker_final.loc[:,"DCR Practice File hours"]*tracker_final.loc[:, "DCR Target/hr"])

tracker_final.loc[:, "temp2"]= (tracker_final.loc[:,"DCR KAI EN(Live)"]*tracker_final.loc[:, "DCR Target/hr"])

tracker_final.loc[:, "DCR Total Target"]=tracker_final.loc[:, "temp1"]+tracker_final.loc[:, "temp2"]

tracker_final= tracker_final.drop(columns=['temp1', 'temp2'])

tracker_final.loc[:, "DCR %"]=[0 if x==0 or y==0 else z for x,y,z in zip(tracker_final.loc[:, "DCR Total Achieved"],tracker_final.loc[:, "DCR Total Target"],
                                                                         (tracker_final.loc[:, "DCR Total Achieved"]/tracker_final.loc[:, "DCR Total Target"]))]
tracker_final['DCR %']=tracker_final['DCR %']*100

tracker_final['DCR %'].fillna(0, inplace=True)

tracker_final=tracker_final.rename(columns={'DCR KAI(Live)': 'DCR KAI Live Count'})

tracker_final=tracker_final[['Login','Reporting Manager', 'Program', 'Year','Week','Date', 'End Date','Leave','Leave Type','Leave Hours',
                             'Instances','Bincheck EPCR Hours','Bincheck EPCR Count','Bincheck USN Hours','Bincheck USN Count',
                             'Bincheck IXD Hours', 'Bincheck IXD Count','Bincheck CCR Hours','Bincheck CCR Count','Bincheck Target/hr',
                             'Achieved Bincheck Weighted','Bincheck Total Achieved', 'Bincheck Total Target','Bincheck %','CCR EN','CCR EN Count','CCR Non EN','CCR Non EN Count','Brand Owner Audits',
                             'Brand Owner Audits Count','Non AVP Audits EN','Non AVP Audits EN Count','Non AVP Audits Non EN','Non AVP Audits Non EN Count',
                             'CCR Target/hr','BOA NA IN Target/hr', 'NON AVP Target/hr', 'Achieved CCR Weighted', 'CCR Total Achieved', 'CCR Total Target','CCR %',
                             'EIM FP Hours','EIM FP Count','EIM FP QA Hours','EIM FP QA Count',
                             'EIM Target/hr', 'Achieved EIM FP Weighted','EIM FP Total Achieved', 'EIM FP Total Target', 'EIM FP %','EIM Esc Hours','EIM Esc Count','EIM Esc QA Hours',
                             'EIM Esc QA Count','EIM ESC Target/hr', 'Achieved EIM Esc Weighted', 'EIM Esc Total Achieved', 'EIM Esc Total Target', 'EIM Escalations %',
                             'EPCR EN','EPCR EN Count','EPCR Non EN','EPCR Non EN Count','EACP Hours','EACP Count','EACP Appeals Hours','EACP Appeals Count',
                             'EACP Appeals Target/hr', 'EPCR ENG Target/hr', 'EPCR NON ENG Target/hr','EACP Target/hr','Achieved EPCR Weighted','EPCR Total Achieved',
                             'EPCR Total Target','EPCR %',
                             'MCM EN','MCM EN Count','MCM Non EN','MCM Non EN Count','MCM Live EN QA',
                             'MCM Live EN QA Count','MCM Live Non EN QA','MCM Live Non EN QA Count',
                             'MCM Others NE','MCM Others NE Count','MCM Practice File hours','MCM Practice Files Count',
                             'MCM ENG Target/hr', 'MCM NON ENG Target/hr','MCM Others NE Target/hr',
                             'Achieved MCM Weighted','MCM Total Achieved', 'MCM Total Target', 'MCM %',
                             'Safety Sensitivity Hours','Safety Sensitivity Count','​​Safety Sensitivity Target/hr','Achieved Safety Sensitivity Weighted',
                             'Safety Sensitivity Total Achieved', 'Safety Sensitivity Total Target','Safety Sensitivity %',
                             'Test Buy Hrs',
                             'Test Buy Count','GPV -Test Buy Target/hr','USN EN','USN EN Count','USN Non EN','USN Non EN Count','USN ENG Target/hr',
                             'USN NON ENG Target/hr','Achieved USN Weighted','USN Total Achieved','USN Total Target','USN %','DCR KAI EN(Live)','DCR KAI Live Count',
                             'DCR KAI Non EN (Live)','DCR Live EN QA','DCR Live Non EN QA','DCR Practice File hours','DCR Practice Count','DCR Target/hr','Achieved DCR Weighted','DCR Total Achieved','DCR Total Target','DCR %',
                             'WI KAI EN (Live)','WI KAI Non EN (Live)','WI Live EN QA',
                             'WI Live Non EN QA','WI Practice File Hours','WI Practice File Count','WI Target/hr','Achieved WI Weighted','WI Total Achieved','WI Total Target','WI %',
                             'SLIM Hours','SLIM Count','SLIM Target/hr', 'Achieved SLIM Weighted', 'SLIM Total Achieved', 'SLIM Total Target', 'SLIM %',
                             'PAD Hours','Coverage Hours','Coverage Count','Coverage Target/hr','Achieved Coverage Weighted','Coverage Total Achieved',
                             'Coverage Total Target','Coverage %','Retails Hours', 'Retails Count', 'Retails Target/hr','Achieved Retails Weighted','Retails Total Achieved','Retails Total Target','Retails %',
                             'Backlog & Ops Update','HTMS Data Reporting','Flash','Keyword Analysis',
                             'MOM','Ops Tracker','Other Reportings','PR Doc','Productivity Report',
                             'WBR Report','Work Allocation','EN QA Hours','QA Reporting',
                             'Non EN QA Hrs','QA Calibration','Deep Dive Hours','Deep Dive Name',
                             'Project Name','Project Hours',
                             'Other Task Name','Other Tasks Hours','Meeting',
                             'No Work','Other NPTS','System Issues','Training','RAMP Hours','Quality Hours',
                             'Audits', 'QA Meeting','Adhoc-Business Name', 'Adhoc-Business Hours','Adhoc-Internal Name', 'Adhoc-Internal Hours',
                             'Comments','Productive Hours','NPT SUM','NPT Hours','Leave Sum','Total Hours']]

tracker_final= tracker_final.round({"Achieved DCR Weighted":2, "DCR %":2,'DCR Total Achieved': 2, 'DCR Total Target':2})

convert_dict = {'Achieved DCR Weighted':'float32','DCR Total Achieved':'float32',
            'DCR Total Target': 'float32','DCR %':'float32'}
  
tracker_final = tracker_final.astype(convert_dict)

tracker_final['Achieved DCR Weighted'] = tracker_final['Achieved DCR Weighted'].round(2)
tracker_final['DCR Total Achieved'] = tracker_final['DCR Total Achieved'].round(2)
tracker_final['DCR Total Target'] = tracker_final['DCR Total Target'].round(2)
tracker_final['DCR %'] = tracker_final['DCR %'].round(2)




dcr=""
#computed hours, npt and productivity


tracker_final.replace(to_replace=np.nan, value=0, inplace=True)

tracker_final.loc[:, "Actual Cumputed Hours"]= (tracker_final.loc[:, "Bincheck EPCR Hours"]+tracker_final.loc[:, "Bincheck USN Hours"]+
                                                tracker_final.loc[:, "EIM Esc Hours"]+tracker_final.loc[:, "Bincheck IXD Hours"]+
                                                tracker_final.loc[:, "MCM EN"]+tracker_final.loc[:, "MCM Non EN"]+tracker_final.loc[:, "MCM Live EN QA"]+tracker_final.loc[:, "MCM Live Non EN QA"]+
                                                tracker_final.loc[:, "USN EN"]+tracker_final.loc[:, "USN Non EN"]+
                                                tracker_final.loc[:, "EIM Esc QA Hours"]+tracker_final.loc[:, "CCR EN"]+tracker_final.loc[:, "CCR Non EN"]+
                                                tracker_final.loc[:, "Brand Owner Audits"]+tracker_final.loc[:, "Non AVP Audits EN"]+tracker_final.loc[:, "Non AVP Audits Non EN"]+
                                                tracker_final.loc[:, "EIM FP Hours"]+tracker_final.loc[:, "EIM FP QA Hours"]+tracker_final.loc[:, "EPCR EN"]+
                                                tracker_final.loc[:, "EPCR Non EN"]+tracker_final.loc[:, "EACP Hours"]+tracker_final.loc[:, "EACP Appeals Hours"]+
                                                tracker_final.loc[:, "Safety Sensitivity Hours"]+
                                                tracker_final.loc[:, "Test Buy Hrs"]+
                                                tracker_final.loc[:, "SLIM Hours"]+tracker_final.loc[:, "WI Practice File Hours"]+
                                                tracker_final.loc[:, "DCR Practice File hours"]+tracker_final.loc[:, "Bincheck CCR Hours"]+
                                                tracker_final.loc[:, "DCR KAI EN(Live)"]+tracker_final.loc[:, "DCR KAI Non EN (Live)"]+
                                                tracker_final.loc[:, "DCR Live EN QA"]+tracker_final.loc[:, "DCR Live Non EN QA"]+
                                                tracker_final.loc[:, "WI KAI EN (Live)"]+tracker_final.loc[:, "WI KAI Non EN (Live)"]+
                                                tracker_final.loc[:, "WI Live EN QA"]+tracker_final.loc[:, "WI Live Non EN QA"]+
                                                tracker_final.loc[:, "Coverage Hours"]+tracker_final.loc[:, "MCM Others NE"])

tracker_final.loc[:, "Actual Cumputed NPT"]= (tracker_final.loc[:, "NPT SUM"]+tracker_final.loc[:, "MCM Practice File hours"]+
                                              tracker_final.loc[:, "Retails Hours"]+tracker_final.loc[:, "NPT Hours"]+tracker_final.loc[:, "PAD Hours"])

tracker_final.loc[:, "Weighted Producitve Sum"]= (tracker_final.loc[:, "Achieved EIM Esc Weighted"]+ tracker_final.loc[:, "Achieved MCM Weighted"]+
                                              tracker_final.loc[:, "Achieved USN Weighted"]+tracker_final.loc[:, "Achieved Bincheck Weighted"]+
                                              tracker_final.loc[:, "Achieved CCR Weighted"]+tracker_final.loc[:, "Achieved EPCR Weighted"]+
                                              tracker_final.loc[:, "Achieved EIM FP Weighted"]+tracker_final.loc[:, "Achieved Safety Sensitivity Weighted"]+
                                              tracker_final.loc[:, "Achieved WI Weighted"]+tracker_final.loc[:, "Achieved SLIM Weighted"]+
                                              tracker_final.loc[:, "Achieved Coverage Weighted"]+tracker_final.loc[:, "Achieved DCR Weighted"])

tracker_final['Actual Cumputed Hours'] = tracker_final['Actual Cumputed Hours'].round(2)
tracker_final['Actual Cumputed NPT'] = tracker_final['Actual Cumputed NPT'].round(2)
tracker_final['Weighted Producitve Sum'] = tracker_final['Weighted Producitve Sum'].round(2)


tracker_final.loc[:, "Actual Productivity %"]=""
for i in range(0, len(tracker_final['Login']), 1):
    if (tracker_final['Actual Cumputed NPT'][i]>=8):
        tracker_final.loc[:, "Actual Productivity %"][i]="No Target"
    elif (tracker_final['Leave Hours'][i]==8):
        tracker_final.loc[:, "Actual Productivity %"][i]="Leave"
    elif (tracker_final['Leave Hours'][i]==4) & (tracker_final['Actual Cumputed NPT'][i]==4):
        tracker_final.loc[:, "Actual Productivity %"][i]="No Target"
    elif (tracker_final['Leave Hours'][i]==4) & (tracker_final['Actual Cumputed NPT'][i]!=4):
        tracker_final.loc[:, "Actual Productivity %"][i]= ((tracker_final.loc[:, "Achieved EIM Esc Weighted"][i]+ tracker_final.loc[:, "Achieved MCM Weighted"][i]+
                                                          tracker_final.loc[:, "Achieved USN Weighted"][i]+tracker_final.loc[:, "Achieved Bincheck Weighted"][i]+
                                                          tracker_final.loc[:, "Achieved CCR Weighted"][i]+
                                                          tracker_final.loc[:, "Achieved EPCR Weighted"][i]+tracker_final.loc[:, "Achieved EIM FP Weighted"][i]+
                                                          tracker_final.loc[:, "Achieved Safety Sensitivity Weighted"][i]+
                                                          tracker_final.loc[:, "Achieved WI Weighted"][i]+tracker_final.loc[:, "Achieved SLIM Weighted"][i]+
                                                          tracker_final.loc[:, "Achieved Coverage Weighted"][i]+
                                                          tracker_final.loc[:, "Achieved DCR Weighted"][i])/(4-tracker_final.loc[:, "Actual Cumputed NPT"][i]))*100
    elif (tracker_final['Leave Hours'][i]==8) & (tracker_final['Actual Cumputed NPT'][i]==8):
        tracker_final.loc[:, "Actual Productivity %"][i]=0
    else:        
        tracker_final.loc[:, "Actual Productivity %"][i]= ((tracker_final.loc[:, "Achieved EIM Esc Weighted"][i]+ tracker_final.loc[:, "Achieved MCM Weighted"][i]+
                                                          tracker_final.loc[:, "Achieved USN Weighted"][i]+tracker_final.loc[:, "Achieved Bincheck Weighted"][i]+
                                                          tracker_final.loc[:, "Achieved CCR Weighted"][i]+
                                                          tracker_final.loc[:, "Achieved EPCR Weighted"][i]+tracker_final.loc[:, "Achieved EIM FP Weighted"][i]+
                                                          tracker_final.loc[:, "Achieved Safety Sensitivity Weighted"][i]+
                                                          tracker_final.loc[:, "Achieved WI Weighted"][i]+tracker_final.loc[:, "Achieved SLIM Weighted"][i]+
                                                          tracker_final.loc[:, "Achieved Coverage Weighted"][i]+
                                                          tracker_final.loc[:, "Achieved DCR Weighted"][i])/(8-tracker_final.loc[:, "Actual Cumputed NPT"][i]))*100
        
tracker_final['Login']=tracker_final['Login'].astype(str)


tracker_final.sort_values(['Date', 'Week','Login'],ascending=( True,True,True),inplace=True)

tracker_final= tracker_final.drop(tracker_final[(tracker_final['Week']>w1)].index)




#all weeks data

weekly_data=tracker_final.copy()

weekly_data['Login']=weekly_data['Login'].astype('category')

weekly_data.replace(to_replace="Leave", value=0, inplace=True)
    
weekly_data.replace(to_replace="No Target", value=0, inplace=True)

weekly_data1= weekly_data.groupby(['Login', 'Year', 'Week'])[['Leave Hours', 'Bincheck EPCR Hours','Bincheck EPCR Count', 'Bincheck USN Hours','Bincheck USN Count',
                                                              'Bincheck IXD Hours', 'Bincheck IXD Count','Bincheck CCR Hours','Bincheck CCR Count','Achieved Bincheck Weighted','Bincheck Total Achieved', 'Bincheck Total Target',
                                                              'CCR EN', 'CCR EN Count', 'CCR Non EN', 'CCR Non EN Count', 'Brand Owner Audits', 'Brand Owner Audits Count',
                                                              'Non AVP Audits EN', 'Non AVP Audits EN Count', 'Non AVP Audits Non EN', 'Non AVP Audits Non EN Count',
                                                              'Achieved CCR Weighted','CCR Total Achieved', 'CCR Total Target', 'EIM FP Hours', 'EIM FP Count', 'EIM FP QA Hours',
                                                              'EIM FP QA Count', 'Achieved EIM FP Weighted', 'EIM FP Total Achieved', 'EIM FP Total Target','EIM Esc Hours', 'EIM Esc Count',
                                                              'EIM Esc QA Hours', 'EIM Esc QA Count', 'EIM ESC Target/hr', 'Achieved EIM Esc Weighted', 'EIM Esc Total Achieved',
                                                              'EIM Esc Total Target', 'EPCR EN', 'EPCR EN Count', 'EPCR Non EN', 'EPCR Non EN Count', 'EACP Hours', 'EACP Count',
                                                              'EACP Appeals Hours','EACP Appeals Count','Achieved EPCR Weighted', 'EPCR Total Achieved', 'EPCR Total Target',
                                                              'MCM EN', 'MCM EN Count', 'MCM Non EN', 'MCM Non EN Count', 'MCM Live EN QA', 'MCM Live EN QA Count',
                                                              'MCM Live Non EN QA', 'MCM Live Non EN QA Count',
                                                              'MCM Others NE','MCM Others NE Count','MCM Practice File hours','MCM Practice Files Count',
                                                              'Achieved MCM Weighted', 'MCM Total Achieved', 'MCM Total Target','Safety Sensitivity Hours',
                                                              'Safety Sensitivity Count','Achieved Safety Sensitivity Weighted',
                                                              'Safety Sensitivity Total Achieved', 'Safety Sensitivity Total Target','Test Buy Hrs', 'Test Buy Count', 'USN EN', 'USN EN Count', 'USN Non EN',
                                                              'USN Non EN Count', 'Achieved USN Weighted','USN Total Achieved', 'USN Total Target', 'DCR KAI EN(Live)',
                                                              'DCR KAI Non EN (Live)', 'DCR Live EN QA', 'DCR Live Non EN QA','DCR Practice File hours','DCR Practice Count',
                                                              'Achieved DCR Weighted','DCR Total Achieved','DCR Total Target', 'WI KAI EN (Live)', 'WI KAI Non EN (Live)', 
                                                              'WI Live EN QA', 'WI Live Non EN QA', 'WI Practice File Hours','WI Practice File Count','Achieved WI Weighted',
                                                              'WI Total Achieved','WI Total Target','WI %','SLIM Hours','SLIM Count','SLIM Target/hr', 'Achieved SLIM Weighted',
                                                              'SLIM Total Achieved', 'SLIM Total Target',
                                                              'PAD Hours','Coverage Hours',
                                                              'Coverage Count','Achieved Coverage Weighted','Coverage Total Achieved',
                                                              'Coverage Total Target','Coverage %','Retails Hours', 'Retails Count','Achieved Retails Weighted',
                                                              'Retails Total Achieved','Retails Total Target','Retails %',
                                                              'Backlog & Ops Update','HTMS Data Reporting','Flash','Keyword Analysis',
                                                              'MOM','Ops Tracker','Other Reportings','PR Doc','Productivity Report',
                                                              'WBR Report','Work Allocation','EN QA Hours','QA Reporting',
                                                              'Non EN QA Hrs','QA Calibration','Deep Dive Hours','Deep Dive Name',
                                                              'Project Name','Project Hours',
                                                              'Other Task Name','Other Tasks Hours','Meeting',
                                                              'No Work','Other NPTS','System Issues','Training','RAMP Hours','Quality Hours',
                                                              'Audits', 'QA Meeting','Adhoc-Business Name', 'Adhoc-Business Hours','Adhoc-Internal Name', 'Adhoc-Internal Hours',
                                                              'Comments','Productive Hours','NPT SUM','NPT Hours','Leave Sum','Total Hours',
                                                              'Actual Cumputed Hours','Actual Cumputed NPT','Weighted Producitve Sum']].sum()
weekly_data1.reset_index(inplace=True)

weekly_data2=weekly_data.groupby(['Week', 'Year'])[['Bincheck Target/hr', 'CCR Target/hr',
                                                'BOA NA IN Target/hr', 'NON AVP Target/hr',
                                                'EIM Target/hr', 'EIM ESC Target/hr',
                                                'EACP Appeals Target/hr', 'EPCR ENG Target/hr',
                                                'EPCR NON ENG Target/hr', 'EACP Target/hr',
                                                'MCM ENG Target/hr', 'MCM NON ENG Target/hr',
                                                '​​Safety Sensitivity Target/hr',
                                                'GPV -Test Buy Target/hr','USN ENG Target/hr',
                                                'USN NON ENG Target/hr', 'WI Target/hr',
                                                'DCR Target/hr','Coverage Target/hr',
                                                'Retails Target/hr','SLIM Target/hr',
                                                'MCM Others NE Target/hr']].mean()

weekly_data2.reset_index(inplace=True)

weekly_data=pd.merge(weekly_data1, weekly_data2, on=['Week', 'Year'], how='left')

weekly_data1=""

weekly_data2=""

weekly_data.loc[:, "Bincheck %"]=[0 if x==0 or y==0 else z for x,y,z in zip(weekly_data.loc[:, "Bincheck Total Achieved"],weekly_data.loc[:, "Bincheck Total Target"],
                                                                            (weekly_data.loc[:, "Bincheck Total Achieved"]/weekly_data.loc[:, "Bincheck Total Target"]))]
weekly_data.loc[:, "Bincheck %"]=weekly_data.loc[:, "Bincheck %"]*100

weekly_data.loc[:, "CCR %"]=[0 if x==0 or y==0 else z for x,y,z in zip(weekly_data.loc[:, "CCR Total Achieved"],weekly_data.loc[:, "CCR Total Target"],
                                                                       (weekly_data.loc[:, "CCR Total Achieved"]/weekly_data.loc[:, "CCR Total Target"]))]
weekly_data.loc[:, "CCR %"]=weekly_data.loc[:, "CCR %"]*100

weekly_data['CCR %'].fillna(0, inplace=True)

weekly_data.loc[:, "EIM FP %"]=[0 if x==0 or y==0 else z for x,y,z in zip(weekly_data.loc[:, "EIM FP Total Achieved"],weekly_data.loc[:, "EIM FP Total Target"],
                                                                          (weekly_data.loc[:, "EIM FP Total Achieved"]/weekly_data.loc[:, "EIM FP Total Target"]))]
weekly_data.loc[:, "EIM FP %"]=weekly_data.loc[:, "EIM FP %"]*100

weekly_data['EIM FP %'].fillna(0, inplace=True)

weekly_data.loc[:, "EIM Escalations %"]=[0 if x==0 or y==0 else z for x,y,z in zip(weekly_data.loc[:, "EIM Esc Total Achieved"],weekly_data.loc[:, "EIM Esc Total Target"],
                                                                                   (weekly_data.loc[:, "EIM Esc Total Achieved"]/weekly_data.loc[:, "EIM Esc Total Target"]))]
weekly_data.loc[:, "EIM Escalations %"]=weekly_data.loc[:, "EIM Escalations %"]*100

weekly_data['EIM Escalations %'].fillna(0, inplace=True)

weekly_data.loc[:, "EPCR %"]=[0 if x==0 or y==0 else z for x,y,z in zip(weekly_data.loc[:, "EPCR Total Achieved"],weekly_data.loc[:, "EPCR Total Target"],
                                                                        (weekly_data.loc[:, "EPCR Total Achieved"]/weekly_data.loc[:, "EPCR Total Target"]))]
weekly_data.loc[:, "EPCR %"]=weekly_data.loc[:, "EPCR %"]*100

weekly_data['EPCR %'].fillna(0, inplace=True)

weekly_data.loc[:, "MCM %"]=[0 if x==0 or y==0 else z for x,y,z in zip(weekly_data.loc[:, "MCM Total Achieved"],weekly_data.loc[:, "MCM Total Target"],
                                                                       (weekly_data.loc[:, "MCM Total Achieved"]/weekly_data.loc[:, "MCM Total Target"]))]
weekly_data.loc[:, "MCM %"]=weekly_data.loc[:, "MCM %"]*100

weekly_data['MCM %'].fillna(0, inplace=True)

weekly_data.loc[:, "Safety Sensitivity %"]=[0 if x==0 or y==0 else z for x,y,z in zip(weekly_data.loc[:, "Safety Sensitivity Total Achieved"], weekly_data.loc[:, "Safety Sensitivity Total Target"],
                                                                                      (weekly_data.loc[:, "Safety Sensitivity Total Achieved"]/weekly_data.loc[:, "Safety Sensitivity Total Target"]))]
weekly_data.loc[:, "Safety Sensitivity %"]=weekly_data.loc[:, "Safety Sensitivity %"]*100

weekly_data['Safety Sensitivity %'].fillna(0, inplace=True)


weekly_data.loc[:, "USN %"]=[0 if x==0 or y==0 else z for x,y,z in zip(weekly_data.loc[:, "USN Total Achieved"],weekly_data.loc[:, "USN Total Target"],
                                                                       (weekly_data.loc[:, "USN Total Achieved"]/weekly_data.loc[:, "USN Total Target"]))]
weekly_data.loc[:, "USN %"]=weekly_data.loc[:, "USN %"]*100

weekly_data['USN %'].fillna(0, inplace=True)

weekly_data.loc[:, "WI %"]=[0 if x==0 or y==0 else z for x,y,z in zip(weekly_data.loc[:, "WI Total Achieved"],weekly_data.loc[:, "WI Total Target"],
                                                                      (weekly_data.loc[:, "WI Total Achieved"]/weekly_data.loc[:, "WI Total Target"]))]
weekly_data.loc[:, "WI %"]=weekly_data.loc[:, "WI %"]*100

weekly_data['WI %'].fillna(0, inplace=True)

weekly_data.loc[:, "DCR %"]=[0 if x==0 or y==0 else z for x,y,z in zip(weekly_data.loc[:, "DCR Total Achieved"],weekly_data.loc[:, "DCR Total Target"],
                                                                       (weekly_data.loc[:, "DCR Total Achieved"]/weekly_data.loc[:, "DCR Total Target"]))]
weekly_data.loc[:, "DCR %"]=weekly_data.loc[:, "DCR %"]*100

weekly_data['DCR %'].fillna(0, inplace=True)

weekly_data.loc[:, "Coverage %"]=[0 if x==0 or y==0 else z for x,y,z in zip(weekly_data.loc[:, "Coverage Total Achieved"],weekly_data.loc[:, "Coverage Total Target"],
                                                                            (weekly_data.loc[:, "Coverage Total Achieved"]/weekly_data.loc[:, "Coverage Total Target"]))]
weekly_data.loc[:, "Coverage %"]=weekly_data.loc[:, "Coverage %"]*100

weekly_data['Coverage %'].fillna(0, inplace=True)

weekly_data.loc[:, "Retails %"]=[0 if x==0 or y==0 else z for x,y,z in zip(weekly_data.loc[:, "Retails Total Achieved"],weekly_data.loc[:, "Retails Total Target"],
                                                                           (weekly_data.loc[:, "Retails Total Achieved"]/weekly_data.loc[:, "Retails Total Target"]))]
weekly_data.loc[:, "Retails %"]=weekly_data.loc[:, "Retails %"]*100

weekly_data['Retails %'].fillna(0, inplace=True)

weekly_data.loc[:, "SLIM %"]=[0 if x==0 or y==0 else z for x,y,z in zip(weekly_data.loc[:, "SLIM Total Achieved"],weekly_data.loc[:, "SLIM Total Target"],
                                                                        (weekly_data.loc[:, "SLIM Total Achieved"]/weekly_data.loc[:, "SLIM Total Target"]))]
weekly_data.loc[:, "SLIM %"]=weekly_data.loc[:, "SLIM %"]*100

weekly_data['SLIM %'].fillna(0, inplace=True)

convert_dict = {'Actual Cumputed Hours':'float64','Actual Cumputed NPT':'float64',
            'Weighted Producitve Sum': 'float64', 'Leave Hours':'int64'}
  
weekly_data = weekly_data.astype(convert_dict)

weekly_data['Actual Cumputed Hours'] = weekly_data['Actual Cumputed Hours'].round(2)
weekly_data['Actual Cumputed NPT'] = weekly_data['Actual Cumputed NPT'].round(2)
weekly_data['Weighted Producitve Sum'] = weekly_data['Weighted Producitve Sum'].round(2)


weekly_data.loc[:, "Total"]= weekly_data.loc[:, "Leave Hours"]+weekly_data.loc[:, "Actual Cumputed Hours"]+weekly_data.loc[:,"Actual Cumputed NPT"]

weekly_data.loc[:, "denomiator"]= [0 if x==0 else z for x,z in zip(weekly_data.loc[:, "Total"],(weekly_data.loc[:,"Total"]-(weekly_data.loc[:,"Leave Hours"]+weekly_data.loc[:,"Actual Cumputed NPT"])))]
    
weekly_data.loc[:, "Weighted Productivity %"]= [0 if x==0 else z for x,z in zip(weekly_data.loc[:, "Total"],(weekly_data.loc[:,"Weighted Producitve Sum"]/weekly_data.loc[:,"denomiator"]))]

weekly_data.loc[:, "Weighted Productivity %"]= weekly_data.loc[:,"Weighted Productivity %"]*100

weekly_data=weekly_data[['Login', 'Year', 'Week', 'Bincheck %', 'CCR %', 'EIM FP %',
                     'EIM Escalations %', 'EPCR %', 'MCM %', 'Safety Sensitivity %', 
                     'USN %','DCR %', 'WI %','SLIM %', 'Coverage %', 'Retails %',
                     'Total','Actual Cumputed Hours',
                     'Actual Cumputed NPT','Leave Hours','Weighted Producitve Sum',
                     'Weighted Productivity %']]

weekly_data['Bincheck %'] = weekly_data['Bincheck %'].round(2)
weekly_data['CCR %'] = weekly_data['CCR %'].round(2)
weekly_data['EIM FP %'] = weekly_data['EIM FP %'].round(2)
weekly_data['EIM Escalations %'] = weekly_data['EIM Escalations %'].round(2)
weekly_data['EPCR %'] = weekly_data['EPCR %'].round(2)
weekly_data['MCM %'] = weekly_data['MCM %'].round(2)
weekly_data['Safety Sensitivity %'] = weekly_data['Safety Sensitivity %'].round(2)
weekly_data['USN %'] = weekly_data['USN %'].round(2)
weekly_data['DCR %'] = weekly_data['DCR %'].round(2)
weekly_data['WI %'] = weekly_data['WI %'].round(2)
weekly_data['Coverage %'] = weekly_data['Coverage %'].round(2)
weekly_data['SLIM %'] = weekly_data['SLIM %'].round(2)
weekly_data['Retails %'] = weekly_data['Retails %'].round(2)
weekly_data['Total'] = weekly_data['Total'].round(2)
weekly_data['Coverage %'] = weekly_data['Coverage %'].round(2)
weekly_data['Actual Cumputed Hours'] = weekly_data['Actual Cumputed Hours'].round(2)
weekly_data['Actual Cumputed NPT'] = weekly_data['Actual Cumputed NPT'].round(2)
weekly_data['Weighted Producitve Sum'] = weekly_data['Weighted Producitve Sum'].round(2)
weekly_data['Weighted Productivity %'] = weekly_data['Weighted Productivity %'].round(2)
weekly_data['Leave Hours'] = weekly_data['Leave Hours'].round(2)

weekly_data['Total'] = weekly_data['Total'].round(0)

weekly_data['Login']=weekly_data['Login'].astype(str)

weekly_data= weekly_data.drop(weekly_data[(weekly_data['Week']>w1)].index)

weekly_data.sort_values(by=['Week', 'Login'], ascending=(True, True), inplace=True)

weekly_data.fillna(0, inplace=True)

weekly_data=weekly_data.round(2)

#current week data

current_week= weekly_data.copy()

current_week.fillna(0, inplace=True)

current_week= current_week.drop(current_week[(current_week['Week']!= w1)].index)

current_week=current_week.round(2)


#ytd data

ytd=tracker_final.copy()

ytd['Login']=ytd['Login'].astype('category')

ytd= ytd.drop(ytd[(ytd['Week']>w1)].index)

convert_dict = {'Actual Cumputed Hours':'float64','Actual Cumputed NPT':'float64',
            'Weighted Producitve Sum': 'float64', 'Leave Hours':'int64'}
  
ytd = ytd.astype(convert_dict)
    
ytd1= ytd.groupby(['Login', 'Year'])[['Leave Hours', 'Bincheck EPCR Hours','Bincheck EPCR Count', 'Bincheck USN Hours','Bincheck USN Count',
                                      'Bincheck IXD Hours','Bincheck IXD Count','Bincheck CCR Hours','Bincheck CCR Count','Achieved Bincheck Weighted','Bincheck Total Achieved', 'Bincheck Total Target',
                                      'CCR EN', 'CCR EN Count', 'CCR Non EN', 'CCR Non EN Count', 'Brand Owner Audits', 'Brand Owner Audits Count',
                                      'Non AVP Audits EN', 'Non AVP Audits EN Count', 'Non AVP Audits Non EN', 'Non AVP Audits Non EN Count',
                                      'Achieved CCR Weighted','CCR Total Achieved', 'CCR Total Target', 'EIM FP Hours', 'EIM FP Count', 'EIM FP QA Hours',
                                      'EIM FP QA Count', 'Achieved EIM FP Weighted', 'EIM FP Total Achieved', 'EIM FP Total Target','EIM Esc Hours', 'EIM Esc Count',
                                      'EIM Esc QA Hours', 'EIM Esc QA Count', 'EIM ESC Target/hr', 'Achieved EIM Esc Weighted', 'EIM Esc Total Achieved',
                                      'EIM Esc Total Target', 'EPCR EN', 'EPCR EN Count', 'EPCR Non EN', 'EPCR Non EN Count', 'EACP Hours', 'EACP Count',
                                      'EACP Appeals Hours','EACP Appeals Count','Achieved EPCR Weighted', 'EPCR Total Achieved', 'EPCR Total Target',
                                      'MCM EN', 'MCM EN Count', 'MCM Non EN', 'MCM Non EN Count', 'MCM Live EN QA', 'MCM Live EN QA Count',
                                      'MCM Live Non EN QA', 'MCM Live Non EN QA Count',
                                      'MCM Others NE','MCM Others NE Count','MCM Practice File hours','MCM Practice Files Count',
                                      'Achieved MCM Weighted', 'MCM Total Achieved', 'MCM Total Target','Safety Sensitivity Hours',
                                      'Safety Sensitivity Count','Achieved Safety Sensitivity Weighted','Safety Sensitivity Total Achieved', 'Safety Sensitivity Total Target',
                                      'Test Buy Hrs', 'Test Buy Count', 'USN EN', 'USN EN Count', 'USN Non EN',
                                      'USN Non EN Count', 'Achieved USN Weighted','USN Total Achieved', 'USN Total Target', 'DCR KAI EN(Live)',
                                      'DCR KAI Non EN (Live)', 'DCR Live EN QA', 'DCR Live Non EN QA','DCR Practice File hours','DCR Practice Count',
                                      'Achieved DCR Weighted','DCR Total Achieved','DCR Total Target', 'WI KAI EN (Live)', 'WI KAI Non EN (Live)', 
                                      'WI Live EN QA', 'WI Live Non EN QA', 'WI Practice File Hours','WI Practice File Count','Achieved WI Weighted',
                                      'WI Total Achieved','WI Total Target','WI %','SLIM Hours','SLIM Count','Achieved SLIM Weighted','SLIM Total Achieved',
                                      'SLIM Total Target','SLIM %',
                                      'PAD Hours','Coverage Hours',
                                      'Coverage Count','Achieved Coverage Weighted','Coverage Total Achieved',
                                      'Coverage Total Target','Coverage %','Retails Hours', 'Retails Count','Achieved Retails Weighted',
                                      'Retails Total Achieved','Retails Total Target','Retails %',
                                      'Backlog & Ops Update','HTMS Data Reporting','Flash','Keyword Analysis',
                                      'MOM','Ops Tracker','Other Reportings','PR Doc','Productivity Report',
                                      'WBR Report','Work Allocation','EN QA Hours','QA Reporting',
                                      'Non EN QA Hrs','QA Calibration','Deep Dive Hours','Deep Dive Name',
                                      'Project Name','Project Hours',
                                      'Other Task Name','Other Tasks Hours','Meeting',
                                      'No Work','Other NPTS','System Issues','Training','RAMP Hours','Quality Hours',
                                      'Audits', 'QA Meeting','Adhoc-Business Name', 'Adhoc-Business Hours','Adhoc-Internal Name', 'Adhoc-Internal Hours',
                                      'Comments','Productive Hours','NPT SUM','NPT Hours','Leave Sum','Total Hours','Actual Cumputed Hours',
                                      'Actual Cumputed NPT','Weighted Producitve Sum']].sum()
ytd1.reset_index(inplace=True)

ytd2=ytd.groupby(['Year'])[['Bincheck Target/hr', 'CCR Target/hr',
                        'BOA NA IN Target/hr', 'NON AVP Target/hr',
                        'EIM Target/hr', 'EIM ESC Target/hr',
                        'EACP Appeals Target/hr', 'EPCR ENG Target/hr',
                        'EPCR NON ENG Target/hr', 'EACP Target/hr',
                        'MCM ENG Target/hr', 'MCM NON ENG Target/hr',
                        '​​Safety Sensitivity Target/hr',
                        'GPV -Test Buy Target/hr','USN ENG Target/hr',
                        'USN NON ENG Target/hr','WI Target/hr',
                        'DCR Target/hr','Coverage Target/hr',
                        'Retails Target/hr', 'SLIM Target/hr',
                        'MCM Others NE Target/hr']].mean()

ytd2.reset_index(inplace=True)


ytd=pd.merge(ytd1, ytd2, on=['Year'], how='left')

#ytd1=""

#ytd2=""

ytd.loc[:, "Bincheck %"]=[0 if x==0 or y==0 else z for x,y,z in zip(ytd.loc[:, "Bincheck Total Achieved"],ytd.loc[:, "Bincheck Total Target"],
                                                                            (ytd.loc[:, "Bincheck Total Achieved"]/ytd.loc[:, "Bincheck Total Target"]))]
ytd.loc[:, "Bincheck %"]=ytd.loc[:, "Bincheck %"]*100

ytd.loc[:, "CCR %"]=[0 if x==0 or y==0 else z for x,y,z in zip(ytd.loc[:, "CCR Total Achieved"],ytd.loc[:, "CCR Total Target"],
                                                               (ytd.loc[:, "CCR Total Achieved"]/ytd.loc[:, "CCR Total Target"]))]
ytd.loc[:, "CCR %"]=ytd.loc[:, "CCR %"]*100

ytd['CCR %'].fillna(0, inplace=True)

ytd.loc[:, "EIM FP %"]=[0 if x==0 or y==0 else z for x,y,z in zip(ytd.loc[:, "EIM FP Total Achieved"],ytd.loc[:, "EIM FP Total Target"],
                                                                  (ytd.loc[:, "EIM FP Total Achieved"]/ytd.loc[:, "EIM FP Total Target"]))]
ytd.loc[:, "EIM FP %"]=ytd.loc[:, "EIM FP %"]*100

ytd['EIM FP %'].fillna(0, inplace=True)

ytd.loc[:, "EIM Escalations %"]=[0 if x==0 or y==0 else z for x,y,z in zip(ytd.loc[:, "EIM Esc Total Achieved"],ytd.loc[:, "EIM Esc Total Target"],
                                                                           (ytd.loc[:, "EIM Esc Total Achieved"]/ytd.loc[:, "EIM Esc Total Target"]))]
ytd.loc[:, "EIM Escalations %"]=ytd.loc[:, "EIM Escalations %"]*100

ytd['EIM Escalations %'].fillna(0, inplace=True)

ytd.loc[:, "EPCR %"]=[0 if x==0 or y==0 else z for x,y,z in zip(ytd.loc[:, "EPCR Total Achieved"],ytd.loc[:, "EPCR Total Target"],
                                                                (ytd.loc[:, "EPCR Total Achieved"]/ytd.loc[:, "EPCR Total Target"]))]
ytd.loc[:, "EPCR %"]=ytd.loc[:, "EPCR %"]*100

ytd['EPCR %'].fillna(0, inplace=True)

ytd.loc[:, "MCM %"]=[0 if x==0 or y==0 else z for x,y,z in zip(ytd.loc[:, "MCM Total Achieved"],ytd.loc[:, "MCM Total Target"],
                                                               (ytd.loc[:, "MCM Total Achieved"]/ytd.loc[:, "MCM Total Target"]))]
ytd.loc[:, "MCM %"]=ytd.loc[:, "MCM %"]*100

ytd['MCM %'].fillna(0, inplace=True)

ytd.loc[:, "Safety Sensitivity %"]=[0 if x==0 or y==0 else z for x,y,z in zip(ytd.loc[:, "Safety Sensitivity Total Achieved"], ytd.loc[:, "Safety Sensitivity Total Target"],
                                                                              (ytd.loc[:, "Safety Sensitivity Total Achieved"]/ytd.loc[:, "Safety Sensitivity Total Target"]))]
ytd.loc[:, "Safety Sensitivity %"]=ytd.loc[:, "Safety Sensitivity %"]*100

ytd['Safety Sensitivity %'].fillna(0, inplace=True)


ytd.loc[:, "USN %"]=[0 if x==0 or y==0 else z for x,y,z in zip(ytd.loc[:, "USN Total Achieved"],ytd.loc[:, "USN Total Target"],
                                                               (ytd.loc[:, "USN Total Achieved"]/ytd.loc[:, "USN Total Target"]))]
ytd.loc[:, "USN %"]=ytd.loc[:, "USN %"]*100

ytd['USN %'].fillna(0, inplace=True)

ytd.loc[:, "WI %"]=[0 if x==0 or y==0 else z for x,y,z in zip(ytd.loc[:, "WI Total Achieved"],ytd.loc[:, "WI Total Target"],
                                                              (ytd.loc[:, "WI Total Achieved"]/ytd.loc[:, "WI Total Target"]))]
ytd.loc[:, "WI %"]=ytd.loc[:, "WI %"]*100

ytd['WI %'].fillna(0, inplace=True)

ytd.loc[:, "DCR %"]=[0 if x==0 or y==0 else z for x,y,z in zip(ytd.loc[:, "DCR Total Achieved"],ytd.loc[:, "DCR Total Target"],
                                                               (ytd.loc[:, "DCR Total Achieved"]/ytd.loc[:, "DCR Total Target"]))]
ytd.loc[:, "DCR %"]=ytd.loc[:, "DCR %"]*100

ytd['DCR %'].fillna(0, inplace=True)

ytd.loc[:, "Coverage %"]=[0 if x==0 or y==0 else z for x,y,z in zip(ytd.loc[:, "Coverage Total Achieved"],ytd.loc[:, "Coverage Total Target"],
                                                                    (ytd.loc[:, "Coverage Total Achieved"]/ytd.loc[:, "Coverage Total Target"]))]
ytd.loc[:, "Coverage %"]=ytd.loc[:, "Coverage %"]*100

ytd['Coverage %'].fillna(0, inplace=True)

ytd.loc[:, "Retails %"]=[0 if x==0 or y==0 else z for x,y,z in zip(ytd.loc[:, "Retails Total Achieved"],ytd.loc[:, "Retails Total Target"],
                                                                   (ytd.loc[:, "Retails Total Achieved"]/ytd.loc[:, "Retails Total Target"]))]
ytd.loc[:, "Retails %"]=ytd.loc[:, "Retails %"]*100

ytd['Retails %'].fillna(0, inplace=True)

ytd.loc[:, "SLIM %"]=[0 if x==0 or y==0 else z for x,y,z in zip(ytd.loc[:, "SLIM Total Achieved"],ytd.loc[:, "SLIM Total Target"],
                                                                (ytd.loc[:, "SLIM Total Achieved"]/ytd.loc[:, "SLIM Total Target"]))]
ytd.loc[:, "SLIM %"]=ytd.loc[:, "SLIM %"]*100

ytd['SLIM %'].fillna(0, inplace=True)

ytd.loc[:, "Total"]= ytd.loc[:, "Leave Hours"]+ytd.loc[:, "Actual Cumputed Hours"]+ytd.loc[:,"Actual Cumputed NPT"]
    
ytd.loc[:, "Weighted Productivity %"]= [0 if x==0 else z for x,z in zip(ytd.loc[:, "Total"],
                                                                        (ytd.loc[:,"Weighted Producitve Sum"]/(ytd.loc[:,"Total"]-(ytd.loc[:,"Leave Hours"]+ytd.loc[:,"Actual Cumputed NPT"]))))]

ytd.loc[:, "Weighted Productivity %"]= ytd.loc[:,"Weighted Productivity %"]*100

ytd['Actual Cumputed Hours'] = ytd['Actual Cumputed Hours'].round(2)
ytd['Actual Cumputed NPT'] = ytd['Actual Cumputed NPT'].round(2)
ytd['Weighted Producitve Sum'] = ytd['Weighted Producitve Sum'].round(2)
ytd['Total'] = ytd['Total'].round(0)
ytd['Weighted Productivity %'] = ytd['Weighted Productivity %'].round(2)

ytd=ytd[['Login', 'Year', 'Bincheck %', 'CCR %', 'EIM FP %',
         'EIM Escalations %', 'EPCR %', 'MCM %', 'Safety Sensitivity %', 
         'USN %','DCR %', 'WI %', 'Coverage %', 'Retails %','SLIM %',
         'Total','Actual Cumputed Hours','Actual Cumputed NPT',
         'Leave Hours','Weighted Producitve Sum','Weighted Productivity %']]



ytd['Login']=ytd['Login'].astype(str)

ytd.fillna(0, inplace=True)

ytd.sort_values(by=['Login'], ascending=(True), inplace=True)




writer = pd.ExcelWriter(path1+"/USN_Productivity_Report_"+str(w1)+".xlsx", engine='xlsxwriter')

current_week.replace(to_replace=0, value=np.nan, inplace=True)
ytd.replace(to_replace=0, value=np.nan, inplace=True)
weekly_data.replace(to_replace=0, value=np.nan, inplace=True)
tracker_final.replace(to_replace=0, value=np.nan, inplace=True)

current_week.to_excel(writer, sheet_name='Current Week', index=False) 

ytd.to_excel(writer, sheet_name='YTD', index=False)

weekly_data.to_excel(writer, sheet_name='All Weeks', index=False) 
 
tracker_final.to_excel(writer, sheet_name='All Programs', index=False)  # Default position, cell A1.


writer.save()