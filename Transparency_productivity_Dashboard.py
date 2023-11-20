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
from tkinter import ttk
pd.options.mode.chained_assignment = None
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)
import win32com.client as win32
import pywintypes as py
from pywintypes import com_error


path=r'C:\Users\nkcho\Documents\Python\transparency_productivity report'

SP_Data= pd.read_excel(path+'\Data.xlsx', sheet_name=0)
SP_Data= pd.DataFrame(SP_Data)

SP_Data= SP_Data.drop(columns=['Ad-Hoc Description (If Applicable)', 'Week'])

SP_Data=SP_Data.rename(columns={'Time Spent in Minutes': 'Time on Tool (min)',
                                '# Cases Completed':'Actual Throughput'})
SP_Data.loc[:,'Login'].dropna(how='any',axis=0,inplace=True)
SP_Data.loc[:,'Task'].dropna(how='any',axis=0,inplace=True)
SP_Data.loc[:,'Date'].dropna(how='any',axis=0,inplace=True)

SP_Data= SP_Data.drop(SP_Data[(SP_Data['Login'].isnull()==True)].index)
SP_Data= SP_Data.drop(SP_Data[(SP_Data['Date'].isnull()==True)].index)
SP_Data= SP_Data.drop(SP_Data[(SP_Data['Task'].isnull()==True)].index)

cols=['Actual Throughput','Time on Tool (min)']

SP_Data[cols].fillna(0, inplace=True)

SP_Data[cols].replace(to_replace=np.nan, value=0, inplace=True)

SP_Data.loc[:, "Date"]=pd.to_datetime(SP_Data.loc[:,"Date"]).dt.date
SP_Data.loc[:, "Week"]=pd.to_datetime(SP_Data.loc[:,"Date"]).dt.week

convert_dict= {'Week':'int8','Login':'category', 'Task':'category',
               'Time on Tool (min)':'float32', 'Actual Throughput':'float32'}

SP_Data = SP_Data.astype(convert_dict)



QA_Data= pd.read_excel(path+'\Data.xlsx', sheet_name=1)
QA_Data= pd.DataFrame(QA_Data)

QA_Data= QA_Data.drop(columns=['Comments'])

QA_Data=QA_Data.rename(columns={'User_ID': 'Login',
                                '# Cases':'QA Sample',
                                'Process':'Task'})

QA_Data.loc[:,'Login'].dropna(how='any',axis=0,inplace=True)
QA_Data.loc[:,'Task'].dropna(how='any',axis=0,inplace=True)
QA_Data.loc[:,'Date'].dropna(how='any',axis=0,inplace=True)

QA_Data= QA_Data.drop(QA_Data[(QA_Data['Login'].isnull()==True)].index)
QA_Data= QA_Data.drop(QA_Data[(QA_Data['Week'].isnull()==True)].index)
QA_Data= QA_Data.drop(QA_Data[(QA_Data['Task'].isnull()==True)].index)

cols=['QA Sample','Errors Observed','Opportunities','Total Opportunities',
      'Error Opportunities','Quality %']

QA_Data[cols].fillna(0, inplace=True)

QA_Data['QA Sample'].replace(to_replace=np.nan, value=0, inplace=True)
QA_Data['Errors Observed'].replace(to_replace=np.nan, value=0, inplace=True)
QA_Data['Opportunities'].replace(to_replace=np.nan, value=0, inplace=True)
QA_Data['Total Opportunities'].replace(to_replace=np.nan, value=0, inplace=True)
QA_Data['Quality %'].replace(to_replace=np.nan, value=0, inplace=True)
QA_Data['Error Opportunities'].replace(to_replace=np.nan, value=0, inplace=True)

convert_dict= {'Week':'int8','Login':'category', 'Task':'category','Quality %':'float32','Opportunities':'float32',
               'QA Sample': 'int32','Errors Observed':'int16', 'Total Opportunities':'int16',
               'Error Opportunities':'int16'}
QA_Data = QA_Data.astype(convert_dict)



#final= SP_Data.groupby(['Date', 'Week','Login','Task'])[['Actual Throughput','Time on Tool (min)']].sum()

options= ['OPR Activation']

main_task1= SP_Data[SP_Data['Task'].isin(options)]

main_task1= main_task1[['Date', 'Login', 'Time on Tool (min)', 'Actual Throughput']]

main_task1= main_task1.groupby(['Date','Login'])[['Actual Throughput','Time on Tool (min)']].sum()

main_task1.reset_index(inplace=True)

main_task1=main_task1.rename(columns={'Actual Throughput': 'Actual Throughput(OPR Activation)',
                                      'Time on Tool (min)':'Time on Tool(OPR Activation) in min'})

options= ['OPR/IB De-activation']

main_task2= SP_Data[SP_Data['Task'].isin(options)]

main_task2= main_task2[['Date', 'Login', 'Time on Tool (min)', 'Actual Throughput']]

main_task2= main_task2.groupby(['Date','Login'])[['Actual Throughput','Time on Tool (min)']].sum()

main_task2.reset_index(inplace=True)

main_task2=main_task2.rename(columns={'Actual Throughput': 'Actual Throughput(OPR/IB De-activation)',
                                      'Time on Tool (min)':'Time on Tool(OPR/IB De-activation) in min'})

options= ['Outbound activation']

main_task3= SP_Data[SP_Data['Task'].isin(options)]

main_task3= main_task3[['Date', 'Login', 'Task', 'Time on Tool (min)', 'Actual Throughput']]

main_task3= main_task3.groupby(['Date','Login'])[['Actual Throughput','Time on Tool (min)']].sum()

main_task3.reset_index(inplace=True)

main_task3=main_task3.rename(columns={'Actual Throughput': 'Actual Throughput(Outbound activation)',
                                      'Time on Tool (min)':'Time on Tool(Outbound activation) in min'})

options= ['Adhoc Task','Training','Audits']

adhoc= SP_Data[SP_Data['Task'].isin(options)]

adhoc= adhoc[['Date', 'Login', 'Time on Tool (min)']]

adhoc=adhoc.rename(columns={'Time on Tool (min)': 'Adhoc Task in Min'})

adhoc= adhoc.groupby(['Date','Login'])[['Adhoc Task in Min']].sum()

adhoc.reset_index(inplace=True)

adhoc.drop_duplicates(subset=['Date', 'Login'], inplace=True)

options= ['Team Meeting','No Work Mode','Fun Session']

downtime= SP_Data[SP_Data['Task'].isin(options)]

downtime= downtime[['Date', 'Login','Time on Tool (min)']]

downtime=downtime.rename(columns={'Time on Tool (min)': 'Downtime in Min'})

downtime= downtime.groupby(['Date','Login'])[['Downtime in Min']].sum()

downtime.reset_index(inplace=True)

downtime.drop_duplicates(subset=['Date', 'Login'], inplace=True)

options= ['Metrics','Emails','Work Allocation']

reporting= SP_Data[SP_Data['Task'].isin(options)]

reporting= reporting[['Date', 'Login', 'Time on Tool (min)']]

reporting=reporting.rename(columns={'Time on Tool (min)': 'Reporting in Min'})

reporting= reporting.groupby(['Date','Login'])[['Reporting in Min']].sum()

reporting.reset_index(inplace=True)

reporting.drop_duplicates(subset=['Date', 'Login'], inplace=True)

final=pd.merge(main_task1, main_task2, how='left', on=['Date', 'Login'])

final=pd.merge(final,main_task3, how='left', on=['Date', 'Login'])

final=pd.merge(final,adhoc, how='left', on=['Date', 'Login'])

final=pd.merge(final, downtime, how='left', on=['Date', 'Login'])

final=pd.merge(final, reporting, how='left', on=['Date', 'Login'])

