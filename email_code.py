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


app= tk.Tk()
app.minsize(100,100)
app.title("Sending email to Manager")

    
        
        
        
email_body= str("PFA the productivity and quality scores for Week:"+ str(10))
subject= str("Weekly Report no.-"+str(10))
        
def attach_file():
    global filename, attachment
            
    filename= tk.filedialog.askopenfilename(title='Select a file')
            
            
    tk.messagebox.showinfo("Info","File attached", icon="info")
            
def sending_email():
            
    a= sender_name_entry.get()
    b= receiver_name_entry.get()
            
            
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = b
        mail.Subject = subject
        mail.Body = ("Hello"+ " "+ str(b) +"," + "\n"+ "\n"+ email_body +"\n"+ "\n" "Regards" + "\n"+ "\n"+ str(a))
               

               # To attach a file to the email (optional):
                
        mail.Attachments.Add(filename)

        mail.Send()
                
                
            
    except:
        tk.messagebox.showerror("Error", "Error sending email", icon="error")
                
    else:
        tk.messagebox.showinfo("Info", "Email sent successfully", icon="info")
'''
except:
        tk.messagebox.showerror("Error", "Error Occurred",icon="")
'''    
def close():
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
    
button_close= tk.Button(app, text= "close", command= close)
button_close.grid(row=6, column=1)
        
app.mainloop()
    