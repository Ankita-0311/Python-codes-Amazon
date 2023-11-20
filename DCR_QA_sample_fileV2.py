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
root.minsize(350,200)
root.title("Generate QA File")



#functions

def enter_week():
    global w1
    #getting year entry from user
    #making customer enter the week till which it needs the report
    try:
        w1= int(Week.get())
        Week.delete(0, 4)
        
        
    
    except:
        tk.messagebox.showerror("Error", "Please enter a valid week number", icon= "error")
        Week.delete(0, 4)
        return

def select_file():
    global filename1
    
    filename1= tk.filedialog.askopenfilename(title='Select a file')
        
    if filename1=="":
        tk.messagebox.showerror("Error", "Please select file for generating QA file", icon='error')
        
    else:
        tk.messagebox.showinfo("Info", " File selected. Your selection:"+filename1)


def path_finder():
    global path
    try:
        path =tk.filedialog.askdirectory()
        
        
        
        if path=="":
            tk.messagebox.showerror("Error", "Please select folder where the QA file needs to be located", icon='error')
        
        else:
            tk.messagebox.showinfo("Info", "Your selection:"+path)
            
        return path
            
    except:
        return
    
def generate_report():
    try:


        df= pd.read_excel(filename1, sheet_name=0)
        df=pd.DataFrame(df)
        
        df1=df.groupby(["login",'defect_type'])[['vote']].count()
        
        df1=df1.reset_index()
        df1
        
        login_data= (df1.login.value_counts())
        login_data=pd.DataFrame(login_data)
        
            
        login_data=login_data.reset_index()
        
        
        login_list=login_data.loc[:, 'index'].to_list()
        
        c=len(df1.login.value_counts())
        #c=1
        a=0
        count=0
        
        
        final_data=pd.DataFrame([])
        
        
        for i in login_list:
            
            while count<c:
                
                d1=df1[df1['login']==login_list[a]]
                d1.sort_values(['vote'],ascending=( True),inplace=True)
                d1=d1.reset_index(drop=True)
                d1
                for x in d1.login.unique():
                    z=0
                    column=0
                    d1.loc[:, "col1"]=0
                    while z<=10:
                       
                        
                        d1.loc[:,"col1"]=[z if x==0 else y for z, x,y in zip(d1.loc[:, "col1"],d1.loc[:, "vote"],
                                                                             (d1.loc[:,"col1"]+1))]
                        d1.loc[:, "vote"]= d1.loc[:, "vote"]-1
                        d1.loc[:, "vote"][d1.loc[:, "vote"]<=0]=0
                        
                        #print(d1)
                        d1= d1.reset_index(drop=True)
                       
                        z=d1.loc[:, "col1"].sum()
                        #print(z)
                        d1.sort_values(['vote'], ascending=False, inplace=True)
                        #print(d1)
                        if z==11:
                            d1.loc[0, "col1"]=d1.loc[0, "col1"]-1
                        elif z==12:
                            d1.loc[0, "col1"]=d1.loc[0, "col1"]-2
                        elif z==13:
                            d1.loc[0, "col1"]=d1.loc[0, "col1"]-1
                            d1.loc[1, "col1"]=d1.loc[1, "col1"]-1
                            d1.loc[2, "col1"]=d1.loc[2, "col1"]-1
                        elif z==14:
                            d1.loc[0, "col1"]=d1.loc[0, "col1"]-1
                            d1.loc[1, "col1"]=d1.loc[1, "col1"]-1
                            d1.loc[2, "col1"]=d1.loc[2, "col1"]-1
                            d1.loc[3, "col1"]=d1.loc[3, "col1"]-1
                        elif z==15:
                            d1.loc[0, "col1"]=d1.loc[0, "col1"]-1
                            d1.loc[1, "col1"]=d1.loc[1, "col1"]-1
                            d1.loc[2, "col1"]=d1.loc[2, "col1"]-1
                            d1.loc[3, "col1"]=d1.loc[3, "col1"]-1
                            d1.loc[4, "col1"]=d1.loc[4, "col1"]-1
                        elif z==16:
                            d1.loc[0, "col1"]=d1.loc[0, "col1"]-1
                            d1.loc[1, "col1"]=d1.loc[1, "col1"]-1
                            d1.loc[2, "col1"]=d1.loc[2, "col1"]-1
                            d1.loc[3, "col1"]=d1.loc[3, "col1"]-1
                            d1.loc[4, "col1"]=d1.loc[4, "col1"]-1
                            d1.loc[5, "col1"]=d1.loc[5, "col1"]-1
                        elif z==17:
                            d1.loc[0, "col1"]=d1.loc[0, "col1"]-1
                            d1.loc[1, "col1"]=d1.loc[1, "col1"]-1
                            d1.loc[2, "col1"]=d1.loc[2, "col1"]-1
                            d1.loc[3, "col1"]=d1.loc[3, "col1"]-1
                            d1.loc[4, "col1"]=d1.loc[4, "col1"]-1
                            d1.loc[5, "col1"]=d1.loc[5, "col1"]-1
                            d1.loc[6, "col1"]=d1.loc[6, "col1"]-1
                        elif z==18:
                            d1.loc[0, "col1"]=d1.loc[0, "col1"]-1
                            d1.loc[1, "col1"]=d1.loc[1, "col1"]-1
                            d1.loc[2, "col1"]=d1.loc[2, "col1"]-1
                            d1.loc[3, "col1"]=d1.loc[3, "col1"]-1
                            d1.loc[4, "col1"]=d1.loc[4, "col1"]-1
                            d1.loc[5, "col1"]=d1.loc[5, "col1"]-1
                            d1.loc[6, "col1"]=d1.loc[6, "col1"]-1
                            d1.loc[7, "col1"]=d1.loc[7, "col1"]-1
                        elif z==19:
                            d1.loc[0, "col1"]=d1.loc[0, "col1"]-1
                            d1.loc[1, "col1"]=d1.loc[1, "col1"]-1
                            d1.loc[2, "col1"]=d1.loc[2, "col1"]-1
                            d1.loc[3, "col1"]=d1.loc[3, "col1"]-1
                            d1.loc[4, "col1"]=d1.loc[4, "col1"]-1
                            d1.loc[5, "col1"]=d1.loc[5, "col1"]-1
                            d1.loc[6, "col1"]=d1.loc[6, "col1"]-1
                            d1.loc[7, "col1"]=d1.loc[7, "col1"]-1
                            d1.loc[8, "col1"]=d1.loc[8, "col1"]-1
                        elif z==20:
                            d1.loc[0, "col1"]=d1.loc[0, "col1"]-1
                            d1.loc[1, "col1"]=d1.loc[1, "col1"]-1
                            d1.loc[2, "col1"]=d1.loc[2, "col1"]-1
                            d1.loc[3, "col1"]=d1.loc[3, "col1"]-1
                            d1.loc[4, "col1"]=d1.loc[4, "col1"]-1
                            d1.loc[5, "col1"]=d1.loc[5, "col1"]-1
                            d1.loc[6, "col1"]=d1.loc[6, "col1"]-1
                            d1.loc[7, "col1"]=d1.loc[7, "col1"]-1
                            d1.loc[8, "col1"]=d1.loc[8, "col1"]-1
                            d1.loc[9, "col1"]=d1.loc[9, "col1"]-1
                        
                        else:
                            pass
                            
                            
                        
                        
                        #print(z)
                        #final_data=final_data.append(d1)
                        
                        
                        
                        z=z+1
                        
                        
                        if z>10:
                            
                            break
                
                    final_data= final_data.append(d1)
                    final_data.reset_index(drop=True)
                
                    
                        
                   
                   
                
                count=count+1
                a=a+1
                if count>c:
                    break 
            
        
        
        #print(final_data)
        final_sample=pd.DataFrame([])
        #login_list= ['jhamnani']
        for i in login_list:
            #print(i)
            
            d2= final_data[final_data.loc[:, "login"]==i]
            temp= df[df.loc[:, "login"]==i]
            #print(temp)
            
            
            
            for x in d2.defect_type.unique():
                #print(x)
                n= d2.loc[d2["defect_type"]==x]["col1"].values[0]
                #print(n)
                
                temp2=temp[temp.loc[:,"defect_type"]==x].sample(n)
                
                final_sample= final_sample.append(temp2)
                
                #final_sample.to_excel("sample_file.xlsx", index=False)
                
            
            
            
        
        
        
        writer1 = pd.ExcelWriter(path+'/Sample_File_'+str(w1)+'.xlsx', engine='xlsxwriter')
        
        final_sample.to_excel(writer1, sheet_name='Sample', index=False)
        
        writer1.save()

    except:
        tk.messagebox.showerror("Error", "Error generating file. Check raw files for errors.", icon="error")
            
    else:
        tk.messagebox.showinfo("Info", "File Generated Successfully", icon='info')
            
        
def window_destroy():
    root.destroy()
    
#entry box

Week= tk.Entry(root, width= 15)
Week.grid(row=1, column= 1)

#labels


Week_label= tk.Label(root, text= "Enter Week:")
Week_label.grid(row=1, column=0, padx=5, pady=5)   

path_label1= tk.Label(root, text= "Select file:")
path_label1.grid(row=2, column=0,padx=5, pady=5)

path_label= tk.Label(root, text= "Final File Location:")
path_label.grid(row=3, column=0,padx=5, pady=5) 

generate_file_label= tk.Label(root, text="Click to generate file:")
generate_file_label.grid(row=4, column=0, padx=5, pady=5)

exit_label= tk.Label(root, text="Click to exit:")
exit_label.grid(row=5, column=0, padx=5, pady=5)

#buttons
button2= tk.Button(root, text="Choose file", command= select_file)
button2.grid(row=2, column=1, padx=5, pady=5)   

button3= tk.Button(root, text="Select folder", command= path_finder)
button3.grid(row=3, column=1,padx=5, pady=5)     

report_button=tk.Button(root, text="Generate File", command=lambda : [enter_week(),generate_report()], width=10)
report_button.grid(row=4, column=1, padx=5, pady=5)

exit_button=tk.Button(root, text="Quit", command=window_destroy, width=10)
exit_button.grid(row=5, column=1, padx=5, pady=5)

root.mainloop()