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


root= tk.Tk()
root.minsize(100,100)
root.title("Generate Report")



def enter_year():
    global year_data
    #getting year entry from user
    try:
        year_data= int(year.get())
        year.delete(0, 4)
        
        return year_data
        
    except:
        tk.messagebox.showerror("Error", "Please enter a valid year", icon= "error")
        year.delete(0, 4)
        return
    
         
year= tk.Entry(root, width= 30)
year.grid(row=2, column= 1)

year_label= tk.Label(root, text= "Enter year:")  
year_label.grid(row=2, column=0)          


def enter_week():
    global w1
    #getting year entry from user
    #making customer enter the week till which it needs the report
    try:
        w1= int(Week.get())
        Week.delete(0, 4)
        
        return w1
    
    except:
        tk.messagebox.showerror("Error", "Please enter a valid week number", icon= "error")
        Week.delete(0, 4)
        return
    
Week= tk.Entry(root, width= 30)
Week.grid(row=3, column= 1)
Week_label= tk.Label(root, text= "Enter Week:")
Week_label.grid(row=3, column=0)   

def path2():
    global path1
    
    try:
        path1 =tk.filedialog.askdirectory()
        
        if path1=="":
            tk.messagebox.showerror("Error", "Please select folder where files are located", icon='error')
        
        else:
            tk.messagebox.showinfo("Info", "Your selection:"+path1)
            
    except:
        return
       

path_label= tk.Label(root, text= "Files location:")
path_label.grid(row=4, column=0)

def path_finder():
    global path
    try:
        path =tk.filedialog.askdirectory()
        
        if path1=="":
            tk.messagebox.showerror("Error", "Please select folder where the final file needs to be located", icon='error')
        
        else:
            tk.messagebox.showinfo("Info", "Your selection:"+path)
            
    except:
        return
        
        
    
        



path_label= tk.Label(root, text= "Final Report Location:")
path_label.grid(row=5, column=0) 

def generate_report():
               
    
    try:
        
        
        
        
        
        #importing all the files
        
        #importing login id QA for general data
        #pd.set_option(mode.chained)
                
        file_type = '\*xlsx'
        files = glob.glob(path1 + file_type)
        max_file = max(files, key=os.path.getctime)
        
        d = pd.read_excel(max_file, sheet_name=0)
        
        d=pd.DataFrame(d)
        
        #counting rows in order to repeat values based on week number
        index= d.index
        a= len(index)
        #creating a list from 1 to the week user wants
        
        
        c=(list(range(1, w1+2)))
        #storing the week numbers into the list 
        week= list(itertools.chain.from_iterable(itertools.repeat(x, a) for x in c))
                
        #converting list to dataframe column
        weekd=pd.DataFrame(week, columns= ['Week'])
                
        #repating rows with login and names the no. of times user has input week numbers
        df= pd.concat([d]*w1, ignore_index=True)
                
        #adding week column to the original data
        df['Week']=weekd
        
        #ntering year for all columns
        
        df['Year']= year_data
        
        
        
        #importing all the files

        #importing USN file for QA
        
        #folder_path = r'C:\Users\nkcho\Desktop\Report\USN'
        file_type = '\*xlsx'
        files = glob.glob(path1+'\\USN' + file_type)
        max_file = max(files, key=os.path.getctime)
        
        usno = pd.read_excel(max_file)
        
        usno=pd.DataFrame(usno)
        
        #importing team average file for CCR
        #folder_path = r'C:\Users\nkcho\Desktop\Report\TA'
        file_type = '\*xlsx'
        files = glob.glob(path1+'\\TA' + file_type)
        max_file = max(files, key=os.path.getctime)
        
        ta_ccro = pd.read_excel(max_file, sheet_name=0)
        
        ta_ccro=pd.DataFrame(ta_ccro)
        
        #importing team average file for USN
        
        
        #folder_path = r'C:\Users\nkcho\Desktop\Report\TA'
        file_type = '\*xlsx'
        files = glob.glob(path1+'\\TA' + file_type)
        max_file = max(files, key=os.path.getctime)
        
        ta_usno = pd.read_excel(max_file, sheet_name=1)
        
        ta_usno=pd.DataFrame(ta_usno)
        
        #import MCM file
        
        
        #folder_path = r'C:\Users\nkcho\Desktop\Report\MCM'
        file_type = '\*xlsx'
        files = glob.glob(path1+'\\MCM' + file_type)
        max_file = max(files, key=os.path.getctime)
        
        mcmo = pd.read_excel(max_file)
        
        mcmo=pd.DataFrame(mcmo)
        
        
        
        #import ccr file
        
        
        #folder_path = r'C:\Users\nkcho\Desktop\Report\CCR'
        file_type = '\*xlsx'
        files = glob.glob(path1+'\\CCR' + file_type)
        max_file = max(files, key=os.path.getctime)
        
        ccro = pd.read_excel(max_file)
        
        ccro=pd.DataFrame(ccro)
        
        #import adhoc projects
        
        
        #folder_path = r'C:\Users\nkcho\Desktop\Report\adhoc'
        file_type = '\*xlsx'
        files = glob.glob(path1+'\\adhoc' + file_type)
        max_file = max(files, key=os.path.getctime)
        
        adhoco = pd.read_excel(max_file)
        
        adhoco=pd.DataFrame(adhoco)
        
        #import EIM file
        
        
        #folder_path = r'C:\Users\nkcho\Desktop\Report\EIM'
        file_type = '\*xlsx'
        files = glob.glob(path1+'\\EIM' + file_type)
        max_file = max(files, key=os.path.getctime)
        
        eimo = pd.read_excel(max_file)
        
        eimo=pd.DataFrame(eimo)
        
        #define the sorter which is Month used in order to sort the data as per month
        
        Month_sorter= ['January', 'February', 'March', 'April', 'May', 'June', 
                       'July', 'August', 'September', 'October', 'November', 'December']
        
        #import productivity tracker report for 
        #tagged volume count for weekly and overall report
        
        #importing productivity tracker report with two sheets
        folder_path = r'C:\Users\nkcho\Desktop\Report\Tracker'
        file_type = '\*xlsx'
        files = glob.glob(folder_path + file_type)
        max_file = max(files, key=os.path.getctime)
        
        tracker1 = pd.read_excel(max_file, sheet_name=0)
        
        tracker1=pd.DataFrame(tracker1)
        
        #retrieving relevant columns
        tracker1= tracker1[['WK', 'User Login','Date','CCR EN Count', 'CCR Non EN Count', 'Brand Owner Audits Count', 
                            'Non AVP Audits EN Count', 'Non AVP Audits Non Eng Count',
                            'MCM EN Count', 'MCM Non EN Count', 'USN EN Count', 'USN Non EN Count',
                            'CCR Total Target', 'EIM Total Target', 'MCM Total Target', 
                            'USN Total Target','CCR Total Achieved','EIM Total Achieved', 'MCM Total Achieved', 'USN Total Achieved',
                            'Actual Productivity']]
        
        
        
        folder_path = r'C:\Users\nkcho\Desktop\Report\Tracker'
        file_type = '\*xlsx'
        files = glob.glob(folder_path + file_type)
        max_file = max(files, key=os.path.getctime)
        
        tracker2 = pd.read_excel(max_file, sheet_name=1)
        tracker2= pd.DataFrame(tracker2)
        
        #retrieving relevant columns
        tracker2= tracker2[['WK', 'User Login','CCR EN Count', 'CCR Non EN Count', 'Brand Owner Audits Count', 
                            'Non AVP Audits EN Count', 'Non AVP Audits Non Eng Count',
                            'MCM EN Count', 'MCM Non EN Count', 'USN EN Count', 'USN Non EN Count',
                            'CCR Total Target', 'EIM Total Target', 'MCM Total Target', 
                            'USN Total Target','CCR Total Achieved','EIM Total Achieved', 'MCM Total Achieved', 'USN Total Achieved',
                            'Actual Productivity']]
        
        #merging the two sheets into one dataframe
        dataframes= [tracker1, tracker2]
        tracker= pd.concat(dataframes, axis= 0, join= 'outer', ignore_index= True)
        tracker= tracker.rename(columns={'WK': 'Week', 'User Login':'Login'})
        tracker['Actual Productivity'] = pd.to_numeric(tracker['Actual Productivity'], errors='coerce')
        
        
        #getting month from date column
        
        
        t1= tracker.groupby(['Week', 'Login'])[['CCR EN Count', 'CCR Non EN Count', 
                                                        'Brand Owner Audits Count', 
                                                        'Non AVP Audits EN Count', 
                                                        'Non AVP Audits Non Eng Count',
                                                        'MCM EN Count', 'MCM Non EN Count', 
                                                        'USN EN Count', 'USN Non EN Count',
                                                        'CCR Total Target', 'EIM Total Target', 
                                                        'MCM Total Target', 'USN Total Target',
                                                        'CCR Total Achieved','EIM Total Achieved', 
                                                        'MCM Total Achieved', 'USN Total Achieved']].sum()
        t1=t1.reset_index()
        
        t2= tracker.groupby(['Week', 'Login'])[['Actual Productivity']].mean()
        t2=t2.reset_index()
        
        #merging all the data
        tracker_data= [t1,t2]
        
        tracker= reduce(lambda left,right: pd.merge(left,right,on=['Week', 'Login']), tracker_data)
        tracker=tracker.reset_index()
        
        
        
        #adding total tagged volume and target volume column to the dataframe
        
        tracker.loc[:,'Tagged Volume']= (tracker.loc[:,'CCR EN Count']+tracker.loc[:,'CCR Non EN Count']+
                                         tracker.loc[:,'Brand Owner Audits Count']+tracker.loc[:,'Non AVP Audits EN Count']+
                                         tracker.loc[:,'Non AVP Audits Non Eng Count']+tracker.loc[:,'MCM EN Count']+
                                         tracker.loc[:,'MCM Non EN Count']+tracker.loc[:,'USN EN Count']+
                                         tracker.loc[:,'USN Non EN Count'])
        
        tracker.loc[:,'Target Volume']= (tracker.loc[:,'CCR Total Target']+
                                         tracker.loc[:,'EIM Total Target']+
                                         tracker.loc[:,'MCM Total Target']+
                                         tracker.loc[:,'USN Total Target'])
        
        
        #USN File
        #changing the datatype of USN
        usno= usno.rename(columns={usno.columns[1]:'Login',
                                   usno.columns[4]:'Week',
                                   usno.columns[5]:'Month',
                                   usno.columns[10]:'USN QA %'})
        
        usno=usno.astype({usno.columns[4]:int, usno.columns[5]:str,
                          usno.columns[1]: str, usno.columns[10]:float})
                
        
        
        #adding total errors columns
        usno.loc[:,'USN Total errors']= usno.iloc[:,7]+usno.iloc[:,8]
        usno.loc[:,'USN QA Sample']= usno.iloc[:,6]
        
        #gettng the relevant columns from USN file
        usn1= usno[usno.columns[(1,4,5,10,11,12), ]].copy()
                
        #total USN count from tracker file
        tracker_usn= tracker[tracker.columns[(1,2,19), ]]
        
        tracker_usn=tracker_usn.rename(columns={'USN Total Achieved': "USN Tagged Volume"})
        
        #getting team average values
        ta_usno= ta_usno.rename(columns={ta_usno.columns[0]:'Week'})
        ta_usn1= ta_usno[ta_usno.columns[(0,7,9), ]].copy()
        ta_usn1.loc[:, 'USN Team Average']= ta_usn1.iloc[:, 2]/ta_usn1.iloc[:, 1]
        ta_usn1= ta_usn1[['Week', 'USN Team Average']]
                
        
        #merging USN with df file
        usn_final= pd.merge(df, usn1, how='left', on=['Week', 'Login'])
        
        #merging with tracker data
        usn_final= pd.merge(usn_final, tracker_usn,how='left', on= ['Week', 'Login'])
        
        #merging with team average data
        usn_final= pd.merge(usn_final, ta_usn1, how='left', on=['Week'])
        
        df1= usn_final.copy()
        
        
        #MCM file
        
        #converting object to float in MCM file
        mcmo.iloc[:,11] = pd.to_numeric(mcmo.iloc[:,11], errors='coerce')
        mcmo.iloc[:,9] = pd.to_numeric(mcmo.iloc[:,9], errors='coerce')
        mcmo.iloc[:,10] = pd.to_numeric(mcmo.iloc[:,10], errors='coerce')
                
                
        #getting team average for MCM
        team_average= mcmo.groupby(['Week'])[['QA Sample']].sum()
        team_average= pd.DataFrame(team_average)
        team_average1= mcmo.groupby(['Week'])[['Total Errors']].sum()
        team_average1= pd.DataFrame(team_average1)
                
        team_avg= pd.merge(team_average, team_average1, on=['Week'])
        team_avg= pd.DataFrame(team_avg)
                
        team_avg.loc[:,'MCM Team Average']= (team_avg['QA Sample']- team_avg['Total Errors'])/team_avg['QA Sample']
        team_avg= team_avg.reset_index()
        team_avg= pd.DataFrame(team_avg)
        team_avg= team_avg[['Week', 'MCM Team Average']]
        #merging with MCM file
        mcm= mcmo.copy()
        
        #rename columns
        mcm= mcm.rename(columns={mcm.columns[4]:'Login',
                                   mcm.columns[3]:'Week',
                                   mcm.columns[2]:'Month',
                                   mcm.columns[10]:'MCM Total Errors',
                                   mcm.columns[11]:'MCM QA %',
                                   mcm.columns[9]:'MCM QA Sample'})
        
        #deleting irrelevant columns for mcm file
                
        mcm= mcm[['Week', 'Month','Login', 'MCM Total Errors', 'MCM QA %', 'MCM QA Sample']]
                
        #total MCM count from tracker file
        tracker_mcm= tracker[tracker.columns[(1,2,18), ]].copy()
        tracker_mcm= tracker_mcm.rename(columns={'MCM Total Achieved':'MCM Tagged Volume'})
        
        
        #merging with mcm file to the final mcm file
        mcm_final= pd.merge(df1, mcm, how='left', on= ['Week', 'Login'])
        
        #merging the final mcm with team average data
        mcm_final= pd.merge(mcm_final, team_avg, how='left', on=['Week'])
        
        #merging with final mcm with tracker data
        mcm_final= pd.merge(mcm_final, tracker_mcm, how='left', on=['Week', 'Login'])
        
        df2= mcm_final.copy()
        
        #merging both month data
        df2['Month_x'].update(df2.pop('Month_y'))
        
        #renaming month_x into Month
        df2= df2.rename(columns={'Month_x':'Month'})
        
        
        #CCR file
        #CCR file renaming
        ccro= ccro.rename(columns={ccro.columns[1]:'Login',
                                   ccro.columns[4]:'Week',
                                   ccro.columns[5]:'Month',
                                   ccro.columns[10]:'CCR QA %',
                                   ccro.columns[6]:'CCR QA Sample'})
        
        ccro=ccro.astype({ccro.columns[4]:int, ccro.columns[5]:str,
                          ccro.columns[1]: str, ccro.columns[10]:float})
            
        
        #gettng the relevant columns from CCR file
        ccr1= ccro[ccro.columns[(1,4,5,6,10), ]].copy()
        
        #adding total errors columns
        ccr1.loc[:,'CCR Total errors']= ccro.iloc[:,7]+ccro.iloc[:,8]
        
        #getting relevant columns for TA
        ta_ccr1= ta_ccro[ta_ccro.columns[(0,8,9), ]].copy()
        ta_ccr1= ta_ccr1.rename(columns={ta_ccr1.columns[0]:'Week'})
        ta_ccr1.loc[:, 'CCR Team Average']= ta_ccr1.iloc[:, 1]/ ta_ccr1.iloc[:, 2]
        ta_ccr1= ta_ccr1[['Week','CCR Team Average']]
        
        
        
        #total MCM count from tracker file
        tracker_ccr= tracker[tracker.columns[(16,1,2), ]].copy()
        tracker_ccr= tracker_ccr.rename(columns={'CCR Total Achieved':'CCR Tagged Volume'})
        
        #merging ccr1 file with df2
        
        ccr_final= pd.merge(df2, ccr1, how='left', on= ['Week', 'Login'])
        
        #merging ccr final with the columns with the tracker report
        
        ccr_final= pd.merge(ccr_final, tracker_ccr, how= 'left', on=['Week', 'Login'])
        
        #merging ccr final with TA file for CCR
        ccr_final= pd.merge(ccr_final, ta_ccr1, how='left', on=['Week'])
        
        df3= ccr_final.copy()
        
        #merging month columns as one and overlappin NaN with the values if its present in other columns
        df3['Month_x'].update(df3.pop('Month_y'))
        
        #renaming month_x into Month
        df3= df3.rename(columns={'Month_x':'Month'})
        
        #EIM file
        eimo= eimo.rename(columns={eimo.columns[0]: 'Week',
                           eimo.columns[1]: 'Login',
                           eimo.columns[2]: 'Month',
                           eimo.columns[4]: 'EIM Tagged Volume',
                           eimo.columns[6]: 'EIM Total Errors',
                           eimo.columns[7]: 'EIM QA %',
                           eimo.columns[5]: 'EIM QA Sample',
                           eimo.columns[8]:'EIM Team Average'})

        eim1= eimo[['Week','Login', "Month", 'EIM Tagged Volume', 'EIM Total Errors', 'EIM QA %'
                    , 'EIM QA Sample']].copy()
        eim1.fillna(0, inplace=True)
        
        ta_eim= eimo[['Week', 'EIM Team Average']].copy()
        ta_eim= ta_eim.groupby("Week")['EIM Team Average'].mean()
        
        #merging with the df3
        eim_final= pd.merge(df3, eim1, how='left',on=['Week', 'Login'])
        
        #merging both month data
        eim_final['Month_x'].update(eim_final.pop('Month_y'))
        
        #renaming month_x into Month
        eim_final= eim_final.rename(columns={'Month_x':'Month'})
        eim_final= pd.merge(eim_final, ta_eim,how='left', on=['Week'])
        
        df4= eim_final.copy()
        
        #adhoc file
        #replacing values in adhoc
        adhoco.replace(to_replace= '-', value= 0, inplace= True)
        
        
        #converting object to float in MCM file
        adhoco.iloc[:,0] = pd.to_numeric(adhoco.iloc[:,0], errors='coerce')
        adhoco.iloc[:,4] = pd.to_numeric(adhoco.iloc[:,4], errors='coerce')
        adhoco.iloc[:,5] = pd.to_numeric(adhoco.iloc[:,5], errors='coerce')
        adhoco.iloc[:,6] = pd.to_numeric(adhoco.iloc[:,6], errors='coerce')
        adhoco.iloc[:,7] = pd.to_numeric(adhoco.iloc[:,7], errors='coerce')
        adhoco.iloc[:,8] = pd.to_numeric(adhoco.iloc[:,8], errors='coerce')
        adhoco.iloc[:,9] = pd.to_numeric(adhoco.iloc[:,9], errors='coerce')
        adhoco.iloc[:,10] = pd.to_numeric(adhoco.iloc[:,10], errors='coerce')
        adhoco.iloc[:,11] = pd.to_numeric(adhoco.iloc[:,11], errors='coerce')
        adhoco.iloc[:,12] = pd.to_numeric(adhoco.iloc[:,12], errors='coerce')
        
        #renaming log in columns
        adhoco= adhoco.rename(columns={adhoco.columns[3]: 'Login',
                                       adhoco.columns[0]: 'Month',
                                       adhoco.columns[1]: 'Week',
                                       adhoco.columns[4]: 'BO QA Sample Audits',
                                       adhoco.columns[5]: 'BO No. of errors',
                                       adhoco.columns[6]: 'BO QA %',
                                       adhoco.columns[7]: 'NON AVP QA Sample Audits',
                                       adhoco.columns[8]: 'NON AVP No. of errors',
                                       adhoco.columns[9]: 'NON AVP QA %',
                                       adhoco.columns[10]: 'Transparency QA Sample Audits',
                                       adhoco.columns[11]: 'Transparency No. of errors',
                                       adhoco.columns[12]: 'Transparency QA %'})
        
        adhoco= adhoco[['Week', 'Login','Month', 'BO QA Sample Audits', 'BO No. of errors',
                        'BO QA %', 'NON AVP QA Sample Audits', 'NON AVP No. of errors',
                        'NON AVP QA %', 'Transparency QA Sample Audits', 'Transparency No. of errors',
                        'Transparency QA %']]
        
        
        #brand owner tagged volume from tracker data
        tagged_volume= tracker.iloc[:,[1,2,5,6,7]].copy()
        
        tagged_volume= tagged_volume.rename(columns={'Brand Owner Audits Count': 'BO Audits Tagged Volume'})
        tagged_volume.loc[:,'NON AVP Tagged Volume']= tagged_volume.iloc[:,3]+tagged_volume.iloc[:,4]
        
        #adding bo column to the adhoc data
        
        adhoco= pd.merge(adhoco, tagged_volume, how='left', on=['Week','Login'])
        
        #calculating team average for adhoc
        tm_avg= adhoco.groupby(['Week'])[['BO Audits Tagged Volume', 'BO No. of errors']].sum()
        tm_avg.loc[:,'BO Team Average']= (tm_avg.iloc[:,0]-tm_avg.iloc[:,1])/(tm_avg.iloc[:,0])
        tm_avg= tm_avg.reset_index()
        tm_avg[tm_avg == -inf] = 0
        tm_avg[tm_avg == inf] = 0
        
        tm_avg1= adhoco.groupby(['Week'])[['NON AVP QA Sample Audits', 'NON AVP No. of errors']].sum()
        tm_avg1.loc[:,'NON AVP Team Average']= (tm_avg1.iloc[:,0]-tm_avg1.iloc[:,1])/(tm_avg1.iloc[:,0])
        tm_avg1= tm_avg1.reset_index()
        tm_avg1[tm_avg1 == -inf] = 0
        tm_avg1[tm_avg1 == inf] = 0
        tm_avg2= adhoco.groupby(['Week'])[['Transparency QA Sample Audits', 'Transparency No. of errors']].sum()
        tm_avg2.loc[:,'Transparency Team Average']= (tm_avg2.iloc[:,0]-tm_avg2.iloc[:,1])/(tm_avg2.iloc[:,0])
        tm_avg2=tm_avg2.reset_index()
        tm_avg2[tm_avg2 == -inf] = 0
        tm_avg2[tm_avg2 == inf] = 0
        
        avg_f= pd.merge(tm_avg, tm_avg1, on=['Week'])
        adhoc1= pd.merge(avg_f, tm_avg2, on=['Week'])
        
        adhoc1= adhoc1.reset_index()
        adhoc1= adhoc1.fillna(0)
        
        adhoc1=adhoc1[['Week', 'BO Team Average', 'NON AVP Team Average', 'Transparency Team Average']]
        
        #merging the adhoc with team average
        adhoc= adhoco.copy()
        
        
        
        
        #merging adhoc with the file
        
        # HAVE TO ADD MONTH COLUMN TO THE ORIGINAL SHEET ITSELF IN ORDER FOR SMOOTH EXECUTION
        
        df4= pd.merge(df4, adhoc,how='left', on=['Week', 'Login'])
        df4= pd.merge(df4, adhoc1,how='left', on=['Week'])
        
        #merging month columns as one and overlappin NaN with the values if its present in other columns
        df4['Month_x'].update(df4.pop('Month_y'))
        
        #renaming the month column 
        df4= df4.rename(columns={'Month_x':'Month'})
        
        #converting accuracy columns into percentage by multiplying by 100
        df4['MCM QA %']= df4['MCM QA %']*100
        df4['USN QA %']= df4['USN QA %']*100
        df4['CCR QA %']= df4['CCR QA %']*100
        df4['EIM QA %']= df4['EIM QA %']*100
        df4['BO QA %']= df4['BO QA %']*100
        df4['NON AVP QA %']= df4['NON AVP QA %']*100
        df4['Transparency QA %']= df4['Transparency QA %']*100
        
        
        df4['CCR Team Average']= df4['CCR Team Average']*100
        df4['USN Team Average']= df4['USN Team Average']*100
        df4['MCM Team Average']= df4['MCM Team Average']*100
        df4['EIM Team Average']= df4['EIM Team Average']*100
        df4['BO Team Average']= df4['BO Team Average']*100
        df4['NON AVP Team Average']= df4['NON AVP Team Average']*100
        df4['Transparency Team Average']= df4['Transparency Team Average']*100
        
        
        #filling NaN with 0 in df4 dataframe
            
        df4=df4.fillna(0)
        
        #adding weighted average column to the final report
           
        df4.loc[:,'Associate Goal(Weighted Average)']= ((df4.iloc[:,9]*df4.iloc[:,8])
                                                        +(df4.iloc[:,13]*df4.iloc[:,14])
                                                        +(df4.iloc[:,19]*df4.iloc[:,18])
                                                        +(df4.iloc[:,24]*df4.iloc[:,20])
                                                        +(df4.iloc[:,38]*df4.iloc[:,34])
                                                        +(df4.iloc[:,39]*df4.iloc[:,37])
                                                        +(df4.iloc[:,40]*df4.iloc[:,31]))/(df4.iloc[:,8]
                                                                                           +df4.iloc[:,14]+df4.iloc[:,18]+df4.iloc[:,34]
                                                                                           +df4.iloc[:,20]
                                                                                           +df4.iloc[:,37]+df4.iloc[:,31])
        
        
        
        df4.replace([np.inf, -np.inf], np.nan, inplace=True)
        df4= df4.fillna(0)
        #adding deviation in USN
        df4.loc[:,'USN Deviation'] = [0 if x == 0 else y for x,y in zip(df4.iloc[:,8],
            df4.iloc[:,5]-df4.iloc[:,9])]
        
        #adding deviation in MCM
        df4.loc[:,'MCM Deviation'] = [0 if x == 0 else y for x,y in zip(df4.iloc[:,14],
            df4.iloc[:,11]-df4.iloc[:,13])]
        
        
        #adding deviation in CCR
        df4.loc[:,'CCR Deviation'] = [0 if x == 0 else y for x,y in zip(df4.iloc[:,18],
            df4.iloc[:,16]-df4.iloc[:,19])]
        
        #adding deviation in EIM
        df4.loc[:,'EIM Deviation'] = [0 if x == 0 else y for x,y in zip(df4.iloc[:,20],
            df4.iloc[:,22]-df4.iloc[:,24])]
        
        #adding deviation in BO
        df4.loc[:,'BO Deviation'] = [0 if x == 0 else y for x,y in zip(df4.iloc[:,34],
            df4.iloc[:,27]-df4.iloc[:,38])]
        
        #adding deviation in NON AVP
        df4.loc[:,'NON AVP Deviation'] = [0 if x == 0 else y for x,y in zip(df4.iloc[:,37],
            df4.iloc[:,30]-df4.iloc[:,39])]
        
        #adding deviation in Transparency
        df4.loc[:,'Transparency Deviation'] = [0 if x == 0 else y for x,y in zip(df4.iloc[:,31],
            df4.iloc[:,33]-df4.iloc[:,40])]
        
        df4.replace([np.inf, -np.inf], np.nan, inplace=True)
        df4= df4.fillna(0)
        
        #net deviation column to the final report
        df4.loc[:,'Net Deviation']= ((df4.iloc[:,42]*df4.iloc[:,8])
                                     +(df4.iloc[:,43]*df4.iloc[:,14])
                                     +(df4.iloc[:,44]*df4.iloc[:,18])
                                     +(df4.iloc[:,45]*df4.iloc[:,20])
                                     +(df4.iloc[:,46]*df4.iloc[:,34])
                                     +(df4.iloc[:,47]*df4.iloc[:,37])
                                     +(df4.iloc[:,48]*df4.iloc[:,31]))/(df4.iloc[:,8]
                                                                        +df4.iloc[:,14]
                                                                        +df4.iloc[:,18]
                                                                        +df4.iloc[:,20]
                                                                        +df4.iloc[:,34]
                                                                        +df4.iloc[:,37]
                                                                        +df4.iloc[:,31])
        
        df4.replace([np.inf, -np.inf], np.nan, inplace=True)
        df4= df4.fillna(0)
                                                                                                                         #Final QA% (Weighed Average)
        df4.loc[:,'Final QA%(Weighted Average)']= df4.loc[:,'Associate Goal(Weighted Average)']- df4.loc[:,'Net Deviation']
        
        #converting NaN to 0
        df4.replace([np.inf, -np.inf], np.nan, inplace=True)
        df4= df4.fillna(0)
        
        #removing NA in month column
        df4['Month'].replace(to_replace= 0, value= np.nan, inplace= True)
        
        df4.loc[df4['Month'].isna(), 'Month'] = df4.groupby(['Week'])[['Month']].transform(lambda x: x.mode()[0] if any(x.mode()) else 'NA')
        
        #sort values based on month_sorter in the final data
        df4.Month= df4.Month.astype('category')
        
        df4.Month.cat.set_categories(Month_sorter, inplace=True)
        
        df4.sort_values(by=['Week'], inplace=True)
        
        #converting category to str in month column
        df4 = df4.astype({"Month": 'str'})
        
        
        #getting columns position as per the report
        
        df5= df4[['Name', 'Login', 'Week','Year', 'Month',
                  'USN Tagged Volume', 'USN QA %', 'USN Team Average',
                  'MCM Tagged Volume', 'MCM QA %', 'MCM Team Average',
                  'CCR Tagged Volume', 'CCR QA %', 'CCR Team Average',
                  'EIM Tagged Volume','EIM QA %', 'EIM Team Average',
                  'BO Audits Tagged Volume', 'BO QA %', 'BO Team Average',
                  'NON AVP Tagged Volume', 'NON AVP QA %', 'NON AVP Team Average',
                  'Transparency QA Sample Audits', 'Transparency QA %', 'Transparency Team Average',
                  'Associate Goal(Weighted Average)', 'USN Deviation',
                  'MCM Deviation','CCR Deviation','EIM Deviation','BO Deviation',
                  'NON AVP Deviation','Transparency Deviation','Net Deviation',
                  'Final QA%(Weighted Average)']]
        
        df5= df5.reset_index(drop= True)
        
        
        #MONTHLY REPORT
        
        
        data_1= df4
        data_1.reset_index(drop=True, inplace=True)
        #converting data into dataframe
        data_1= pd.DataFrame(data_1)
        
        #converting NA to 0 in month column
        
        data_1.replace(to_replace= 'NA', value= 0, inplace= True)
        
        
        #creating a column which will specify whether month is a string or numeric
        data_1['Isstring']= data_1.Month.str.contains(r'[0-9]')
        
        #keeping string rows for month column
        df_1= data_1[data_1['Isstring']==False].copy()
        
        #creating a new dataframe in the same sheet
        data4 = df_1.groupby(['Month', 'Login'])[['USN Tagged Volume',
                                                 'MCM Tagged Volume',
                                                 'CCR Tagged Volume',
                                                 'EIM Tagged Volume',
                                                 'BO Audits Tagged Volume',
                                                 'NON AVP Tagged Volume',
                                                 'USN Total errors','USN QA Sample',
                                                 'MCM Total Errors','MCM QA Sample',
                                                 'CCR Total errors','CCR QA Sample',
                                                 'EIM Total Errors', 'EIM QA Sample',
                                                 'BO QA Sample Audits','BO No. of errors',
                                                 'NON AVP QA Sample Audits', 'NON AVP No. of errors',
                                                 'Transparency QA Sample Audits', 'Transparency No. of errors']].sum()
        data4= pd.DataFrame(data4)
        data4= data4.reset_index()
        
        data4.loc[:,'USN QA %']= (data4.iloc[:,9]
                                  -data4.iloc[:,8])/data4.iloc[:,9]
        data4.loc[:,'MCM QA %']= (data4.iloc[:,11]
                                  -data4.iloc[:,10])/data4.iloc[:,11]
        data4.loc[:,'CCR QA %']= (data4.iloc[:,13]
                                  -data4.iloc[:,12])/data4.iloc[:,13]
        data4.loc[:,'EIM QA %']= (data4.iloc[:,15]
                                  -data4.iloc[:,14])/data4.iloc[:,15]
        data4.loc[:,'BO QA %']= (data4.iloc[:,16]
                                  -data4.iloc[:,17])/data4.iloc[:,16]
        data4.loc[:,'NON AVP QA %']= (data4.iloc[:,18]
                                  -data4.iloc[:,19])/data4.iloc[:,18]
        data4.loc[:,'Transparency QA %']= (data4.iloc[:,20]- data4.iloc[:,21])/data4.iloc[:,20]
        
        
        #converting accuracy columns into percentage by multiplying by 100
        data4['MCM QA %']= data4['MCM QA %']*100
        data4['USN QA %']= data4['USN QA %']*100
        data4['CCR QA %']= data4['CCR QA %']*100
        data4['EIM QA %']= data4['EIM QA %']*100
        data4['BO QA %']= data4['BO QA %']*100
        data4['NON AVP QA %']= data4['NON AVP QA %']*100
        data4['Transparency QA %']= data4['Transparency QA %']*100
        
        #sort values based on month_sorter in the final data
        data4.Month= data4.Month.astype('category')
        
        data4.Month.cat.set_categories(Month_sorter, inplace=True)
        
        data4.sort_values(by=['Month', 'Login'], inplace=True)
        
        #converting category to str in month column
        data4 = data4.astype({"Month": 'str'})
        
        #monthly team average
        
        x= df_1.groupby(['Month'])[['USN Team Average',
                                    'MCM Team Average',
                                    'CCR Team Average',
                                    'EIM Team Average',
                                    'BO Team Average',
                                    'NON AVP Team Average',
                                    'Transparency Team Average']].mean()
        x=x.reset_index()
        
        
        data4= pd.merge(data4, x, how= 'left', on=['Month'])
        
        #adding weighted average column to the final report
           
        data4.loc[:,'Associate Goal(Weighted Average)']= ((data4.iloc[:,29]*data4.iloc[:,2])
                                                        +(data4.iloc[:,30]*data4.iloc[:,3])
                                                        +(data4.iloc[:,31]*data4.iloc[:,4])
                                                        +(data4.iloc[:,32]*data4.iloc[:,5])
                                                        +(data4.iloc[:,33]*data4.iloc[:,6])
                                                        +(data4.iloc[:,34]*data4.iloc[:,7])
                                                        +(data4.iloc[:,35]*data4.iloc[:,20]))/(data4.iloc[:,2]
                                                                                           +data4.iloc[:,3]+data4.iloc[:,5]+data4.iloc[:,6]
                                                                                           +data4.iloc[:,7]
                                                                                           +data4.iloc[:,4]+data4.iloc[:,20])
        
        
        
        data4.replace([np.inf, -np.inf], np.nan, inplace=True)
        data4= data4.fillna(0)
        #adding deviation in USN
        data4.loc[:,'USN Deviation'] = [0 if x == 0 else y for x,y in zip(data4.iloc[:,2],
            data4.iloc[:,22]-data4.iloc[:,29])]
        
        #adding deviation in MCM
        data4.loc[:,'MCM Deviation'] = [0 if x == 0 else y for x,y in zip(data4.iloc[:,3],
            data4.iloc[:,23]-data4.iloc[:,30])]
        
        
        #adding deviation in CCR
        data4.loc[:,'CCR Deviation'] = [0 if x == 0 else y for x,y in zip(data4.iloc[:,4],
            data4.iloc[:,24]-data4.iloc[:,31])]
        
        #adding deviation in EIM
        data4.loc[:,'EIM Deviation'] = [0 if x == 0 else y for x,y in zip(data4.iloc[:,5],
            data4.iloc[:,25]-data4.iloc[:,32])]
        
        #adding deviation in BO
        data4.loc[:,'BO Deviation'] = [0 if x == 0 else y for x,y in zip(data4.iloc[:,6],
            data4.iloc[:,26]-data4.iloc[:,33])]
        
        #adding deviation in NON AVP
        data4.loc[:,'NON AVP Deviation'] = [0 if x == 0 else y for x,y in zip(data4.iloc[:,7],
            data4.iloc[:,27]-data4.iloc[:,34])]
        
        #adding deviation in Transparency
        data4.loc[:,'Transparency Deviation'] = [0 if x == 0 else y for x,y in zip(data4.iloc[:,20],
            data4.iloc[:,28]-data4.iloc[:,35])]
        
        data4.replace([np.inf, -np.inf], np.nan, inplace=True)
        data4= data4.fillna(0)
        
        #net deviation column to the final report
        data4.loc[:,'Net Deviation']= ((data4.iloc[:,37]*data4.iloc[:,2])
                                     +(data4.iloc[:,38]*data4.iloc[:,3])
                                     +(data4.iloc[:,39]*data4.iloc[:,4])
                                     +(data4.iloc[:,40]*data4.iloc[:,5])
                                     +(data4.iloc[:,41]*data4.iloc[:,6])
                                     +(data4.iloc[:,42]*data4.iloc[:,7])
                                     +(data4.iloc[:,43]*data4.iloc[:,20]))/(data4.iloc[:,2]
                                                                        +data4.iloc[:,3]
                                                                        +data4.iloc[:,4]
                                                                        +data4.iloc[:,5]
                                                                        +data4.iloc[:,6]
                                                                        +data4.iloc[:,7]
                                                                        +data4.iloc[:,20])
        
        data4.replace([np.inf, -np.inf], np.nan, inplace=True)
        data4= data4.fillna(0)
                                                                                                                         #Final QA% (Weighed Average)
        data4.loc[:,'Final QA%(Weighted Average)']= data4.loc[:,'Associate Goal(Weighted Average)']- data4.loc[:,'Net Deviation']
        
        #converting NaN to 0
        data4.replace([np.inf, -np.inf], np.nan, inplace=True)
        data4= data4.fillna(0)
        
        
        
        
        #getting columns position as per the report
        
        df6= data4[['Login', 'Month','USN Tagged Volume', 'USN QA %', 'USN Team Average',
                  'MCM Tagged Volume', 'MCM QA %', 'MCM Team Average',
                  'CCR Tagged Volume', 'CCR QA %', 'CCR Team Average',
                  'EIM Tagged Volume', 'EIM QA %', 'EIM Team Average',
                  'BO Audits Tagged Volume', 'BO QA %', 'BO Team Average',
                  'NON AVP Tagged Volume', 'NON AVP QA %', 'NON AVP Team Average',
                  'Transparency QA Sample Audits', 'Transparency QA %', 'Transparency Team Average',
                  'Associate Goal(Weighted Average)', 'USN Deviation',
                  'MCM Deviation','CCR Deviation','EIM Deviation','BO Deviation',
                  'NON AVP Deviation','Transparency Deviation','Net Deviation',
                  'Final QA%(Weighted Average)']]
        
        
        df6= df6.reset_index(drop= True)
        
        #QUATERLY REPORT
        
        qtr1= data4.copy()
        qtr1.reset_index(drop=True, inplace=True)
        qtr1.replace([np.inf, -np.inf], np.nan, inplace=True)
        qtr1= qtr1.fillna(0)
        
        #converting category to str in month column
        qtr1 = qtr1.astype({"Month": str,
                            'NON AVP QA Sample Audits':int,
                            'NON AVP No. of errors':int,
                            'Transparency QA Sample Audits':int})
        
        #removing NA from month
        qtr1.drop(qtr1[qtr1['Month']=="NA"].index, inplace= True)
        
        #converting month name to month int
        qtr1['month_num']=pd.to_datetime(qtr1.Month, format= '%B').dt.month
        
        #month int to quater
        qtr1['Qtr'] = pd.to_datetime(qtr1.month_num).dt.quarter
        qtr_1= qtr1.copy()
        
        qtr_final1 = qtr_1.groupby(['Qtr', 'Login'])[["USN Tagged Volume",
                                                    "MCM Tagged Volume", 
                                                   "CCR Tagged Volume",
                                                   'EIM Tagged Volume',
                                                   "BO Audits Tagged Volume",
                                                   "NON AVP Tagged Volume",
                                                   'USN Total errors','USN QA Sample',
                                                   'MCM Total Errors','MCM QA Sample',
                                                   'CCR Total errors','CCR QA Sample',
                                                   'EIM Total Errors','EIM QA Sample',
                                                   'BO QA Sample Audits','BO No. of errors',
                                                   'NON AVP QA Sample Audits', 'NON AVP No. of errors',
                                                   'Transparency QA Sample Audits', 'Transparency No. of errors']].sum()
        
        qtr_final1 = qtr_final1.astype({'NON AVP QA Sample Audits':float,
                            'NON AVP No. of errors':float,
                            'Transparency QA Sample Audits':float})
        qtr_final1= qtr_final1.reset_index()
        
        qtr_final1= pd.DataFrame(qtr_final1)
        
        qtr_final1.replace([np.inf, -np.inf], np.nan, inplace=True)
        qtr_final1= qtr_final1.fillna(0)
        
        
        
        qtr_final1.loc[:,'USN QA %']= (qtr_final1.iloc[:,9]
                                  -qtr_final1.iloc[:,8])/qtr_final1.iloc[:,9]
        qtr_final1.loc[:,'MCM QA %']= (qtr_final1.iloc[:,11]
                                  -qtr_final1.iloc[:,10])/qtr_final1.iloc[:,11]
        qtr_final1.loc[:,'CCR QA %']= (qtr_final1.iloc[:,13]
                                  -data4.iloc[:,12])/qtr_final1.iloc[:,13]
        qtr_final1.loc[:,'EIM QA %']= (qtr_final1.iloc[:,15]
                                  -data4.iloc[:,14])/qtr_final1.iloc[:,15]
        qtr_final1.loc[:,'BO QA %']= (qtr_final1.iloc[:,16]
                                  -qtr_final1.iloc[:,17])/qtr_final1.iloc[:,16]
        qtr_final1.loc[:,'NON AVP QA %']= (qtr_final1.iloc[:,18]
                                           -qtr_final1.iloc[:,19])/(qtr_final1.iloc[:,18])
        qtr_final1.loc[:,'Transparency QA %']= (qtr_final1.iloc[:,20]
                                                -qtr_final1.iloc[:,21])/qtr_final1.iloc[:,20]
        
        qtr_final1.replace([np.inf, -np.inf], np.nan, inplace=True)
        qtr_final1= qtr_final1.fillna(0)
        
        #converting accuracy columns into percentage by multiplying by 100
        qtr_final1['MCM QA %']= qtr_final1['MCM QA %']*100
        qtr_final1['USN QA %']= qtr_final1['USN QA %']*100
        qtr_final1['CCR QA %']= qtr_final1['CCR QA %']*100
        qtr_final1['EIM QA %']= qtr_final1['EIM QA %']*100
        qtr_final1['BO QA %']= qtr_final1['BO QA %']*100
        qtr_final1['NON AVP QA %']= qtr_final1['NON AVP QA %']*100
        qtr_final1['Transparency QA %']= qtr_final1['Transparency QA %']*100
        
        
        qtr_final3= qtr_1.groupby(['Qtr'])[["USN Team Average",
                                          "MCM Team Average",
                                          "CCR Team Average",
                                          "EIM Team Average",
                                          "BO Team Average",
                                          "NON AVP Team Average",
                                          "Transparency Team Average"]].mean()
        
        qtr_final3=qtr_final3.reset_index()
        
        qtr_final= pd.merge(qtr_final1, qtr_final3, on=['Qtr'])
        
        
        #adding weighted average column to the final report
           
        qtr_final.loc[:,'Associate Goal(Weighted Average)']= ((qtr_final.iloc[:,29]*qtr_final.iloc[:,2])
                                                        +(qtr_final.iloc[:,30]*qtr_final.iloc[:,3])
                                                        +(qtr_final.iloc[:,31]*qtr_final.iloc[:,4])
                                                        +(qtr_final.iloc[:,32]*qtr_final.iloc[:,5])
                                                        +(qtr_final.iloc[:,33]*qtr_final.iloc[:,6])
                                                        +(qtr_final.iloc[:,34]*qtr_final.iloc[:,7])
                                                        +(qtr_final.iloc[:,35]*data4.iloc[:,20]))/(qtr_final.iloc[:,2]
                                                                                           +qtr_final.iloc[:,3]+qtr_final.iloc[:,4]+qtr_final.iloc[:,5]
                                                                                           +qtr_final.iloc[:,6]+qtr_final.iloc[:,7]+qtr_final.iloc[:,20])
        
        qtr_final.replace([np.inf, -np.inf], np.nan, inplace=True)
        qtr_final= qtr_final.fillna(0)
        #adding deviaton in USN
        qtr_final.loc[:,'USN Deviation'] = [0 if x == 0 else y for x,y in zip(qtr_final.iloc[:,2],
            qtr_final.iloc[:,22]-qtr_final.iloc[:,29])]
        
        #adding deviation in MCM
        qtr_final.loc[:,'MCM Deviation'] = [0 if x == 0 else y for x,y in zip(qtr_final.iloc[:,3],
            qtr_final.iloc[:, 23]-qtr_final.iloc[:,30])]
        
        
        #adding deviation in CCR
        qtr_final.loc[:,'CCR Deviation'] = [0 if x == 0 else y for x,y in zip(qtr_final.iloc[:,4],
            qtr_final.iloc[:,24]-qtr_final.iloc[:,31])]
        
        #adding deviation in EIM
        qtr_final.loc[:,'EIM Deviation'] = [0 if x == 0 else y for x,y in zip(qtr_final.iloc[:,5],
            qtr_final.iloc[:,25]-qtr_final.iloc[:,32])]
        
        #adding deviation in BO
        qtr_final.loc[:,'BO Deviation'] = [0 if x == 0 else y for x,y in zip(qtr_final.iloc[:,6],
            qtr_final.iloc[:,26]-qtr_final.iloc[:,33])]
        
        #adding deviation in NON AVP
        qtr_final.loc[:,'NON AVP Deviation'] = [0 if x == 0 else y for x,y in zip(qtr_final.iloc[:,7],
            qtr_final.iloc[:,27]-qtr_final.iloc[:,34])]
        
        #adding deviation in Transparency
        qtr_final.loc[:,'Transparency Deviation'] = [0 if x == 0 else y for x,y in zip(qtr_final.iloc[:,20],
            qtr_final.iloc[:,28]-qtr_final.iloc[:,35])]
        
        
        qtr_final.replace([np.inf, -np.inf], np.nan, inplace=True)
        qtr_final= qtr_final.fillna(0)
        
        #net deviation column to the final report
        qtr_final.loc[:,'Net Deviation']= ((qtr_final.iloc[:,37]*qtr_final.iloc[:,2])
                                     +(qtr_final.iloc[:,38]*qtr_final.iloc[:,3])
                                     +(qtr_final.iloc[:,39]*qtr_final.iloc[:,4])
                                     +(qtr_final.iloc[:,40]*qtr_final.iloc[:,5])
                                     +(qtr_final.iloc[:,41]*qtr_final.iloc[:,6])
                                     +(qtr_final.iloc[:,42]*qtr_final.iloc[:,7])
                                     +(qtr_final.iloc[:,43]*qtr_final.iloc[:,20]))/(qtr_final.iloc[:,2]
                                                                        +qtr_final.iloc[:,3]
                                                                        +qtr_final.iloc[:,4]
                                                                        +qtr_final.iloc[:,5]
                                                                        +qtr_final.iloc[:,6]
                                                                        +qtr_final.iloc[:,7]
                                                                        +qtr_final.iloc[:,20])
                                                                                    
                                                                                    
        qtr_final.replace([np.inf, -np.inf], np.nan, inplace=True)
        qtr_final= qtr_final.fillna(0)
        
        #Final QA% (Weighed Average)
        qtr_final.loc[:,'Final QA%(Weighted Average)']= qtr_final.loc[:,'Associate Goal(Weighted Average)']- qtr_final.loc[:,'Net Deviation']
        
        #converting NaN to 0
        qtr_final.replace([np.inf, -np.inf], np.nan, inplace=True)
        qtr_final= qtr_final.fillna(0)
        
        qtr_final.sort_values(by=['Login'], inplace=True)
        
        
        
        #getting columns position as per the report
        
        df7= qtr_final[['Login', 'Qtr','USN Tagged Volume', 'USN QA %', 'USN Team Average',
                  'MCM Tagged Volume', 'MCM QA %', 'MCM Team Average',
                  'CCR Tagged Volume', 'CCR QA %', 'CCR Team Average',
                  'EIM Tagged Volume', 'EIM QA %', 'EIM Team Average',
                  'BO Audits Tagged Volume', 'BO QA %', 'BO Team Average',
                  'NON AVP Tagged Volume', 'NON AVP QA %', 'NON AVP Team Average',
                  'Transparency QA Sample Audits', 'Transparency QA %', 'Transparency Team Average',
                  'Associate Goal(Weighted Average)', 'USN Deviation',
                  'MCM Deviation','CCR Deviation','EIM Deviation','BO Deviation',
                  'NON AVP Deviation','Transparency Deviation','Net Deviation',
                  'Final QA%(Weighted Average)']]
        
        df7= df7.reset_index(drop= True)
        
        
        #PRODUCTIVITY and QUALITY REPORT OVERALL
        
        
        #getting original data from every sheet for total of everything
        #getting login data which is defined as "df"
        
        #getting target and tagged volume from tracker sheet and saving it in a dataframe
        
        f1= tracker.copy()
        
        f1= f1[['Login', 'Week', 'Tagged Volume', 'Target Volume', 'Actual Productivity']]
        
        f1.fillna(0, inplace=True)
        f1.iloc[:,4]= f1.iloc[:,4]*100
        
        f2= df4.copy()
        f2.fillna(0, inplace=True)
        
        f2.loc[:, 'Total QA Sample']= (f2.iloc[:,7]+f2.iloc[:,12]+f2.iloc[:,15]+f2.iloc[:,23]
                                       +f2.iloc[:,25]+f2.iloc[:,28]+f2.iloc[:,31])
        
        f2.loc[:,'Total Errors']=(f2.iloc[:,6]+f2.iloc[:,10]+f2.iloc[:,17]
                                  +f2.iloc[:,21]+f2.iloc[:,26]+f2.iloc[:,29]+f2.iloc[:,32])
        
        f2.loc[:,'QA %']= ((f2.iloc[:,51]-f2.iloc[:,52])/f2.iloc[:,51])*100
        
        
        f2=f2[['Name','Month', 'Login', 'Week', 'Total QA Sample', 'Total Errors', 'QA %']]
        
        
        final_report= pd.merge(df,f1,how='left',on=['Week','Login'])
        
        final_report= pd.merge(final_report, f2, how='left', on=['Week', 'Login'])
        final_report.fillna(0, inplace=True)
        #merging month columns as one and overlappin NaN with the values if its present in other columns
        final_report['Name_x'].update(final_report.pop('Name_y'))
        
        #renaming the month column 
        final_report= final_report.rename(columns={'Name_x':'Name'})
        final_report.fillna(0, inplace=True)
        
        final_report.sort_values(by=['Week'], inplace=True)
        
        df8= final_report[['Name', 'Login','Week','Year','Tagged Volume',
                                    'Target Volume','Actual Productivity',
                                    'Total QA Sample', 'Total Errors', 'QA %']]
        df8= df8.reset_index(drop= True)
        
        
        
        #monthwise productivity and quality report
        
        f3= final_report[['Name', 'Login', 'Week', 'Month', 'Tagged Volume',
                          'Target Volume', 'Total QA Sample', 'Total Errors']].copy()
        #converting category to str in month column
        f3 = f3.astype({"Month": 'str'})
        f4= f3.groupby(['Login', 'Month'])[['Target Volume', 'Tagged Volume',
                                            'Total QA Sample', 'Total Errors']].sum()
        
        f4= f4.reset_index()
        
        #sort values based on month_sorter in the final data
        f4.Month= f4.Month.astype('category')
        
        f4.Month.cat.set_categories(Month_sorter, inplace=True)
        
        f4.sort_values(by=['Month', 'Login'], inplace=True)
        
        #calculate productivity and QA %
        
        f4.loc[:,'Actual Productivity']= (f4.iloc[:,3]/f4.iloc[:,2])*100
        
        f4.loc[:, 'QA %']=((f4.iloc[:,4]-f4.iloc[:,5])/f4.iloc[:,4])*100
        #converting category to str in month column
        f4 = f4.astype({"Month": 'str'})
        
        
        f4.fillna(0, inplace=True)
        
        
        df9= f4[['Login','Month','Tagged Volume',
                                    'Target Volume','Actual Productivity',
                                    'Total QA Sample', 'Total Errors', 'QA %']].copy()
        df9= df9.reset_index(drop= True)
        
        #quaterwise productivity and quality report
        
        f5=f4[['Login','Month', 'Tagged Volume',
                          'Target Volume', 'Total QA Sample', 'Total Errors']].copy()
        #converting category to str in month column
        f5 = f5.astype({"Month": 'str'})
        f5.fillna(0, inplace=True)
        
        
        #removing NA from month
        f5.drop(f5[f5['Month']=="NA"].index, inplace= True)
        
        #converting month name to month int
        f5['month_num']=pd.to_datetime(f5.Month, format= '%B').dt.month
        
        #month int to quater
        f5['Quater'] = pd.to_datetime(f5.month_num).dt.quarter
        
        f5= f5.groupby(['Login', 'Quater'])[['Target Volume', 'Tagged Volume',
                                            'Total QA Sample', 'Total Errors']].sum()
        
        f5= f5.reset_index()
        
        #calculate productivity and QA %
        
        f5.loc[:,'Actual Productivity']= (f5.iloc[:,3]/f5.iloc[:,2])*100
        
        f5.loc[:, 'QA %']=((f5.iloc[:,4]-f5.iloc[:,5])/f5.iloc[:,4])*100
        
        f5.sort_values(by=['Quater', 'Login'], inplace=True)
        
        f5.fillna(0, inplace=True)
        
        df10= f5.copy()
        df10= df10.reset_index(drop= True)
        
        
        """Top Performers"""
        
        #weekly productivity and quality
        
        weekly= df8[['Week', 'Login', 'Actual Productivity','QA %']].copy()
        
        weekly.sort_values(['Week','Actual Productivity','QA %'],ascending=(True, False, False),inplace= True)
        
        weekly= weekly.reset_index(drop= True)
        
        
        #monthly productivity and quality
        
        monthly= df9[['Month', 'Login', 'Actual Productivity','QA %']].copy()
        
        monthly.Month= monthly.Month.astype('category')
        
        monthly.Month.cat.set_categories(Month_sorter, inplace=True)
        
        
        monthly.sort_values(by=['Month', 'Actual Productivity','QA %'],ascending=(True, False, False), inplace=True)
        
        monthly= monthly.reset_index(drop= True)
        
        
        #quaterly productivity and quality
        
        quaterly= df10[['Quater', 'Login', 'Actual Productivity','QA %']].copy()
        
        quaterly.sort_values(['Quater','Actual Productivity','QA %'],ascending=(True, False, False), inplace= True)

        quaterly= quaterly.reset_index(drop= True)
        
        #weekly productivity and quality for manager file
        
        weekly1= df8[['Week', 'Login', 'Actual Productivity','QA %']].copy()
        
        weekly1.sort_values(['Week','Actual Productivity','QA %'],ascending=(True, False, False),inplace= True)
        
        weekly1= weekly1.reset_index(drop= True)
        
        
        #monthly productivity and quality
        
        monthly1= df9[['Month', 'Login', 'Actual Productivity','QA %']].copy()
        
        monthly1.Month= monthly1.Month.astype('category')
        
        monthly1.Month.cat.set_categories(Month_sorter, inplace=True)
        
        monthly1.sort_values(by=['Month', 'Actual Productivity','QA %'],ascending=(True, False, False), inplace=True)
        
        
        monthly1= monthly1.reset_index(drop= True)
        
        
        #quaterly productivity and quality
        
        quaterly1= df10[['Quater', 'Login', 'Actual Productivity','QA %']].copy()
        
        quaterly1.sort_values(['Quater','Actual Productivity','QA %'],ascending=(True, False, False),inplace= True)

        quaterly1= quaterly1.reset_index(drop= True)
        
        
        writer = pd.ExcelWriter(path+"/Productivity_Quality_Report_Till_WK_"+str(w1)+".xlsx", engine='xlsxwriter')
        writer1= pd.ExcelWriter(path+"/Consolidated_Report_Till_WK_"+ str(w1)+".xlsx", engine= 'xlsxwriter')
        # Position the dataframes in the worksheet.
        df5.to_excel(writer, sheet_name='Weekly Report', index=False)  # Default position, cell A1.
        df6.to_excel(writer, sheet_name='Monthly Report', index=False)
        df7.to_excel(writer, sheet_name='Quaterly Report', index=False)
        df8.to_excel(writer, sheet_name='Productivity & Quality Report', startcol=0, index=False)
        df9.to_excel(writer, sheet_name='Productivity & Quality Report', startcol= 12, index=False)
        df10.to_excel(writer, sheet_name='Productivity & Quality Report', startcol= 22, index=False)
        weekly.to_excel(writer, sheet_name='Top Performer', startcol=0, index=False)
        monthly.to_excel(writer, sheet_name='Top Performer', startcol= 6, index=False)
        quaterly.to_excel(writer, sheet_name='Top Performer', startcol= 12, index=False)
        
        weekly1.to_excel(writer1, sheet_name='Consolidated', startcol=0, index=False)
        monthly1.to_excel(writer1, sheet_name='Consolidated', startcol= 6, index=False)
        quaterly1.to_excel(writer1, sheet_name='Consolidated', startcol= 12, index=False)
        
        
        writer.save()
        writer1.save()
        
    except NameError:
        tk.messagebox.showerror("Error", "All fields are required")
            
        
    except:
        tk.messagebox.showerror("Error", "Please check your files for correct data")
        
        
    else:
        tk.messagebox.showinfo("Info", "Productivity and Quality report has been generated. Check:"+path)
        
        
def send_email():
    
    global email_body, subject
    
    app= tk.Tk()
    app.minsize(100,100)
    app.title("send email")
    
    email_body= ("PFA the productivity and quality scores for Week no.-"+ str(w1))
    subject= ("Weekly Report no.-"+str(w1))
    
        
    def attach_file():
        global filename1
        try:
            filename1= tk.filedialog.askopenfilename(title='Select a file')
            
            
            
           
        except ValueError:
            tk.messagebox.showerror("Error","File not attached", icon="error")
            
        else:
            tk.messagebox.showinfo("Info","File attached", icon="info")
        
            
    def sending_email():
        global a, b, d
            
            
        try:
            a= str(sender_name_entry.get())
            b= str(receiver_name_entry.get())
            d= str(receiver_email_entry.get())
            
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = d
            mail.Subject = str(subject)
            mail.Body = ("Hello"+ " "+  b + "," + "\n"+ "\n"+ str(email_body) + "."+ "\n" + "\n"+ "Regards" + "\n"+ a)
            

                # To attach a file to the email (optional):
            mail.Attachments.Add(filename1)

            mail.Send()
                
                
            
        except py.com_error as e:
            print(e)
            tk.messagebox.showerror("Error", "Error sending email", icon="error")
                
        else:
            tk.messagebox.showinfo("Info", "Email sent successfully", icon="info")
         
        
    
        
    def close1():
        app.destroy()
         
         
    #storage
     
    sender_email= tk.StringVar()
    sender_name= tk.StringVar()
    receiver_email= tk.StringVar()
    receiver_name= tk.StringVar()
     
    #entries
    
    sender_email_entry= tk.Entry(app,textvariable= sender_email, width=30)
    sender_email_entry.grid(row=2, column=0)
     
    sender_name_entry= tk.Entry(app,textvariable= sender_name, width=30)
    sender_name_entry.grid(row=4, column=0)
    
    receiver_email_entry= tk.Entry(app,textvariable= receiver_email, width=30)
    receiver_email_entry.grid(row=6, column=0)
     
    receiver_name_entry= tk.Entry(app,textvariable= receiver_name, width=30)
    receiver_name_entry.grid(row=8, column=0)
    
    
    
    #labels
     
    sender_label= tk.Label(app, text= "Sender's Email:")
    sender_label.grid(row=1, column=0)
     
    sender_name_label= tk.Label(app, text= "Sender's Name:")
    sender_name_label.grid(row=3, column=0)
     
    receiver_label= tk.Label(app, text= "Recipent's Email:")
    receiver_label.grid(row=5, column=0)
    
    receiver_name_label= tk.Label(app, text= "Recipent's Name:")
    receiver_name_label.grid(row=7, column=0)
     
    
    #buttons
     
    button_send= tk.Button(app, text= "Send Email", command= sending_email)
    button_send.grid(row=2, column=1)
    
    button_attachment= tk.Button(app, text= "Add attachments", command= attach_file)
    button_attachment.grid(row=4, column=1)
    
    button_close= tk.Button(app, text= "close", command= close1)
    button_close.grid(row=6, column=1)
        
    app.mainloop()
    
            
   
 

def window_destroy():
    root.destroy()
    

    

button2= tk.Button(root, text="Select folder", command= path2)
button2.grid(row=4, column=1)   

button2= tk.Button(root, text="Select folder", command= path_finder)
button2.grid(row=5, column=1)   
b1 = tk.Button(root, text= "Submit", command= lambda : [enter_year(), enter_week()])
b1.grid(row=6, column= 0)

button_report= tk.Button( root, text= "Generate Report", command= generate_report)
button_report.grid(row=6, column= 1)


button_exit= tk.Button(root, text="Quit", command= window_destroy)
button_exit.grid(row= 7, column= 1)

button_email= tk.Button(root, text= "Send email", command= send_email)
button_email.grid(row=7, column=0)

root.mainloop()


        

 
