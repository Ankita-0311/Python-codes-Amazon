import pandas as pd
import glob
import os.path
import tkinter as tk
from tkinter import filedialog
pd.options.mode.chained_assignment = None
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)


root= tk.Tk()
root.minsize(200,100)
root.title("Merge Files")

def path_finder():
    global path
    try:
        path =tk.filedialog.askdirectory()
        
        
        
        if path=="":
            tk.messagebox.showerror("Error", "Please select folder where the files are located", icon='error')
        
        else:
            tk.messagebox.showinfo("Info", "Your selection:"+path)
            
        return path
            
    except:
        return
    
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
    
    
def generate_file():
    try:
        # reading all the excel files
        filenames = glob.glob(path + "\*.xlsx")
        print('File names:', filenames)
        
        # Initializing empty data frame
        finalexcelsheet = pd.DataFrame()
        
        # to iterate excel file one by one 
        # inside the folder
        for file in filenames:
          
            # combining multiple excel worksheets 
            # into single data frames
            df = pd.concat(pd.read_excel(file, sheet_name=None),
                           ignore_index=True, sort=False)
              
            # Appending excel files one by one
            finalexcelsheet = finalexcelsheet.append(
              df, ignore_index=True)
            
        
        # save combined data
        finalexcelsheet.to_excel(path+'/Consolidated_'+str(w1)+'.xlsx',index=False)
            
        tk.messagebox.showinfo("Info", "File has been generated.", icon="info")
    except:
        tk.messagebox.showerror("Error", "Error preparing the file.", icon="error")
        
def window_destroy():
    root.destroy()
    


#entry box

Week= tk.Entry(root, width= 30)
Week.grid(row=0, column= 1, padx=5, pady=5)


#labels  

Week_label= tk.Label(root, text= "Enter Week:")
Week_label.grid(row=0, column=0, padx=5, pady=5)   

path_label= tk.Label(root, text= "Files Location:")
path_label.grid(row=1, column=0, padx=5, pady=5)   



#buttons

button3= tk.Button(root, text="Select folder", command= path_finder)
button3.grid(row=1, column=1) 

generate_button= tk.Button(root, text="Generate Consolidated File",command= lambda : [enter_week(),generate_file()], width=20)
generate_button.grid(row=2, column=0, padx=5, pady=5)

exit_button= tk.Button(root, text="Quit", command= window_destroy, width=20)
exit_button.grid(row=2, column=1, padx=5, pady=5)


root.mainloop()