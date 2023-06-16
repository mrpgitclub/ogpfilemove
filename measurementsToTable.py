import pandas as pd
import tkinter as tk    
from tkinter import *
from tkinter import filedialog
import os
import shutil
src_dir="G:\SHARED\QA\SPC Toolbox\Raw Data Submission - Small - blank.xlsx"

def SFOLDataFormat(location):
    myDataframe = pd.read_excel(location, sheet_name = "Sheet1")
    myDataframe.columns = [column.replace(" ", "_") for column in myDataframe.columns]
    part = myDataframe['Part'].iloc[0]
    piv= myDataframe.pivot_table(index=['Date/Time','Head_No'],columns='Variable_Type',values='Value')
    piv.to_csv('temp.csv')
    piv2 = pd.read_csv('temp.csv')
    #piv2.insert((len(piv2.columns)-1),'Work Order',piv2.pop(piv2.columns[0]))    
    return part,piv2

def writer():
    location = browseFiles()
    part,frame =SFOLDataFormat(location)
    dst_dir=f"G:\SHARED\QA\SPC Toolbox\{part} raw data.xlsx"
    shutil.copy(src_dir,dst_dir)
    with pd.ExcelWriter(dst_dir,mode='a',if_sheet_exists='overlay') as writer:
        frame.to_excel(writer, sheet_name='Sheet1',startcol=0,startrow=15, index=False) 

mainGUI = tk.Tk()
mainGUI.title("SFOL Data converter")
mainGUI.attributes('-topmost', 'true')
for num in range(1, 5): [mainGUI.columnconfigure(num, minsize = 15), mainGUI.rowconfigure(num, minsize = 15)]


tk.Frame(mainGUI).grid(column = 1, row = 1)
tk.Frame(mainGUI).grid(column = 4, row = 1)
tk.Frame(mainGUI).grid(column = 1, row = 4)
tk.Frame(mainGUI).grid(column = 4, row = 4)
def browseFiles():
    filename = filedialog.askopenfilename(initialdir = "/",title = "Select a File",filetypes = (("Excel Files","*.xlsx*"),("all files","*.*")))
    return filename

tk.Button(mainGUI, text = "Select SFOL Data", command = writer).grid(column = 2, row = 2)
tk.Button(mainGUI, text = "Exit",command= exit).grid(column=3,row=2)
mainGUI.mainloop()