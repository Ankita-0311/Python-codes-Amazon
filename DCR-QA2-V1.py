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
        
        #print(df)
        
        df=df.groupby(["login",'tagging_type']).sample(frac=0.1, random_state=1)
        
        df= df.drop(columns=['time_taken'])
        
        df.loc[:, "QA"]=""
        
        df.loc[:, "QA Comments"]=""
        
        
        writer1 = pd.ExcelWriter(path+'/Final_QA_File_'+str(w1)+'.xlsx', engine='xlsxwriter')
        
        df.to_excel(writer1, sheet_name='QA_data', index=False)
        
        writer1.save()
        
    except:
        tk.messagebox.showerror("Error", "Error generating file. Check raw files for errors.", icon="error")
        
    else:
        tk.messagebox.showinfo("Info", "File Generated Successfully", icon='info')
        
def sender_name():
    global sender_name1
    #getting year entry from user
    try:
        sender_name1= str(sender_name_entry.get())
        sender_name_entry.delete(0, 100)
        
        return sender_name1
        
    except:
        tk.messagebox.showerror("Error", "Please enter a valid name", icon= "error")
        sender_name_entry.delete(0, 100)
        return

def sender_email():
    global sender_email1
    #getting year entry from user
    try:
        sender_email1= str(sender_email_entry.get())
        sender_email_entry.delete(0, 100)
        
        return sender_email1
        
    except:
        tk.messagebox.showerror("Error", "Please enter a valid email", icon= "error")
        sender_email_entry.delete(0, 100)
        return    

def receiver_name():
    global receiver_name1
    #getting year entry from user
    try:
        receiver_name1= str(receiver_name_entry.get())
        receiver_name_entry.delete(0, 100)
        
        return receiver_name1
        
    except:
        tk.messagebox.showerror("Error", "Please enter a valid name", icon= "error")
        receiver_name_entry.delete(0, 100)
        return

def receiver_email():
    global receiver_email1
    #getting year entry from user
    try:
        receiver_email1= str(receiver_email_entry.get())
        receiver_email_entry.delete(0, 100)
        
        return receiver_email1
        
    except:
        tk.messagebox.showerror("Error", "Please enter a valid email", icon= "error")
        receiver_email_entry.delete(0, 100)
        return

        
def attach_file():
    global filename2
    try:
        filename2= path+'/Final_QA_File_'+str(w1)+'.xlsx'
      
    except FileNotFoundError:
        
        tk.messagebox.showinfo("Error","File not found", icon="error")
    except:
        
        tk.messagebox.showinfo("Info", "File not attached", icon="Info")
        
def sending_email():
    
    email_body= ("PFA the QA file for the week-"+ str(w1))
    subject= ("Weekly QA file no.-"+str(w1))
    
            
            
    try:
        if filename1=="":
            tk.messagebox.showerror("Error", "Please select file to be attached with mail", icon='error')
        
        else:
            print(filename2)
               
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = receiver_email1
        mail.Subject = str(subject)
        mail.Body = ("Hello"+ " "+ receiver_name1  + "," + "\n"+ "\n"+ str(email_body) + "."+ "\n" + "\n"+ "Regards" + "\n"+ sender_name1)
            

                # To attach a file to the email (optional):
        mail.Attachments.Add(filename2)

        mail.Send()
                
                
            
    except py.com_error as e:
        print(e)
        tk.messagebox.showerror("Error", "Error sending email", icon="error")
                
    else:
        tk.messagebox.showinfo("Info", "Email sent successfully", icon="info")
         

    
        
def window_destroy():
    root.destroy()
    


#entry box

Week= tk.Entry(root, width= 15)
Week.grid(row=1, column= 1)

sender_name_entry= tk.Entry(root, width=30)
sender_name_entry.grid(row=5, column=1)

sender_email_entry= tk.Entry(root, width=30)
sender_email_entry.grid(row=6, column=1)

receiver_name_entry= tk.Entry(root, width=30)
receiver_name_entry.grid(row=7, column=1)

receiver_email_entry= tk.Entry(root, width=30)
receiver_email_entry.grid(row=8, column=1)
    


#labels


Week_label= tk.Label(root, text= "Enter Week:")
Week_label.grid(row=1, column=0, padx=5, pady=5)   

path_label1= tk.Label(root, text= "Select file:")
path_label1.grid(row=2, column=0,padx=5, pady=5)

path_label= tk.Label(root, text= "Final File Location:")
path_label.grid(row=3, column=0,padx=5, pady=5) 

generate_file_label= tk.Label(root, text="Click to generate file:")
generate_file_label.grid(row=4, column=0, padx=5, pady=5)

sender_name_label= tk.Label(root, text= "Sender's Name:")
sender_name_label.grid(row=5, column=0,padx=5, pady=5)

sender_label= tk.Label(root, text= "Sender's Email:")
sender_label.grid(row=6, column=0,padx=5, pady=5)
 
receiver_name_label= tk.Label(root, text= "Recipent's Name:")
receiver_name_label.grid(row=7, column=0,padx=5, pady=5)

receiver_label= tk.Label(root, text= "Recipent's Email:")
receiver_label.grid(row=8, column=0,padx=5, pady=5) 


#buttons
button2= tk.Button(root, text="Choose file", command= select_file)
button2.grid(row=2, column=1, padx=5, pady=5)   

button3= tk.Button(root, text="Select folder", command= path_finder)
button3.grid(row=3, column=1,padx=5, pady=5)     

report_button=tk.Button(root, text="Generate File", command=lambda : [enter_week(),generate_report()], width=10)
report_button.grid(row=4, column=1, padx=5, pady=5)



button_email= tk.Button(root, text= "Send email", command= lambda:[ sender_name(), sender_email(),
                                                                   receiver_name(),receiver_email(),
                                                                   attach_file(),sending_email()])
button_email.grid(row=9, column=0)    

exit_button=tk.Button(root, text="Quit", command=window_destroy, width=10)
exit_button.grid(row=9, column=1, padx=5, pady=5)


root.mainloop()
