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


root= tk.Tk()
root.minsize(300,300)
root.title("Generate Work Allocation Report")

#functions
def path2():
    global path1
    
    try:
        path1 =tk.filedialog.askdirectory()
        
        
        
        if path1=="":
            tk.messagebox.showerror("Error", "Please select folder where files are located", icon='error')
        
        else:
            tk.messagebox.showinfo("Info", "Your selection:"+path1)
            
        return path1
            
    except:
        return
       


def path_finder():
    global path
    try:
        path =tk.filedialog.askdirectory()
        
        
        
        if path=="":
            tk.messagebox.showerror("Error", "Please select folder where the final file needs to be located", icon='error')
        
        else:
            tk.messagebox.showinfo("Info", "Your selection:"+path)
            
        return path
            
    except:
        return
        
def generate_report():
    try:
        data= pd.read_csv(path1+'/raw_data.csv',low_memory=False)
        data=pd.DataFrame(data)
        
        login= pd.read_excel(path1+'/Data.xlsx', sheet_name=0)
        login=pd.DataFrame(login)
        
        asin=pd.read_excel(path1+'/Da''ta.xlsx', sheet_name=1)
        asin=pd.DataFrame(asin)
        
        login= login[login.columns[(0,5), ]]
        
        asin_list= asin.iloc[:, 0].to_list()

        sample_list=asin.iloc[:, 1].to_list()
        
        df1= data[data['Asin'].isin(asin_list)]
        
        
        a=0
        b=0
        c=len(asin['Asins'])
        #c=0
        count=0
        final_data=pd.DataFrame([])
        while count<c:
            if asin_list[a]==np.nan:
                break
            else:
                d1=df1[df1['Asin']==asin_list[a]]
                d1= d1.iloc[0:sample_list[b],:]
                final_data=final_data.append(d1)
            
            count=count+1
            a=a+1
            b=b+1
            if count>c:
                break 
        
               
        final_data.reset_index(inplace=True)
        
        login_list=login.iloc[:, 0].to_list()
        
        series=pd.Series(login_list)
        
        repeat_list=login.iloc[:, 1].to_list()
        
        s=series.repeat(repeat_list)
        s=s.reset_index()
        
        s=s[[0]]
        
        s.columns.values[0]='Login'
        
        final_data=pd.merge(final_data, s, left_index=True, right_index=True)
        
        
        final_data.loc[:, "Vote"]=""
        
        final_data.loc[:, "reason_code"]=""
        
        final_data.loc[:, "time_taken"]=""
        
        final_data= final_data.drop(columns=['index'])
        
        writer1 = pd.ExcelWriter(path+'/final_report.xlsx', engine='xlsxwriter')
        
        final_data.to_excel(writer1, sheet_name='Final_Work_Allocation', index=False)
        
        writer1.save()
        

    except:
        tk.messagebox.showerror("Error", "Error generating report. Check files for errors.", icon="error")
        
    else:
        tk.messagebox.showinfo("Info", "Report Generated Successfully", icon='info')
        
def window_destroy():
    root.destroy()
    

#labels
headline= tk.Label(root, text= "Work Allocation for DCR:", width=20)
headline.grid(row=0, column=0, padx=10, pady=10)

path_label1= tk.Label(root, text= "Files location:")
path_label1.grid(row=1, column=0,padx=5, pady=5)

path_label= tk.Label(root, text= "Final Report Location:")
path_label.grid(row=2, column=0,padx=5, pady=5)  


#buttons
button2= tk.Button(root, text="Select folder", command= path2)
button2.grid(row=1, column=1, padx=5, pady=5)   

button3= tk.Button(root, text="Select folder", command= path_finder)
button3.grid(row=2, column=1,padx=5, pady=5)     

report_button=tk.Button(root, text="Generate File", command=generate_report, width=10)
report_button.grid(row=3, column=0, padx=5, pady=5)

exit_button=tk.Button(root, text="Quit", command=window_destroy, width=10)
exit_button.grid(row=3, column=1, padx=5, pady=5)















root.mainloop()