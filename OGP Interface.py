import sys
import os
import pandas as pd
import tkinter as tk
import tkinter.ttk as ttk
import sqlite3
import time
from sqlite3 import connect
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

#import matplotlib.pyplot as plt    #not implemented yet
#import numpy as np                 #not implemented yet

###
#   Functions
###

def submitshots(dfObject,filename,outputDir):
    dfObject.to_csv(str(outputDir + '\\' + filename), header = False, index = False)
    wdEventHandler.mostRecentShot = filename
    print(f'Submitted:{outputDir}\\{filename}', )
    return

def grabData(location,num):
    dfObject = pd.read_excel(location, sheet_name = num, header = 0, index_col = None, usecols = None, dtype=str) #reads export file and takes data from specified sheet
    dfObject.columns = [column.replace(" ", "_") for column in dfObject.columns] #replace spaces with underscores for formatting
    lastRow = dfObject.iloc[-1] #grab the last row 
    partType = lastRow["Product_Code"] #reads product code from last row
    workOrder = lastRow["Work_Order"] #grabs the correct work order from the last row
    return dfObject,workOrder,lastRow,partType

def formatQCtoDF(dataframe,lastRow,workOrder):
    dataframe.query("Work_Order == @workOrder", inplace=True) #selects only the rows with the workorder
    dataframe.drop_duplicates(keep = 'last', inplace = True, ignore_index = True, subset = 'Cavity') #remove extra lines from partial shots
    dataframe.dropna(axis = 1, how = 'all', inplace = True)
    dataframe.pop('Fails') #delete fails column
    dateTime = dataframe.pop('Date_Time') # assigns datetime column to a variable
    dataframe.insert(len(dataframe.columns),'Date_Time',dateTime) # Replaces the date time column to the correct location
    return dataframe

def twoPartCRC(dfPartone,dfParttwo): #to executre when the part number correlates to a two part crc inner program 
    topOD = dfParttwo.pop('Top_OD_DIA') #next three lines assign the columns we wish to move to variables
    hZ = dfParttwo.pop('HUL_ZD')
    weight = dfParttwo.pop('Weight_RES')
    dfPartone.insert(6,'Top_OD_DIA',topOD)
    dfPartone.insert(7,'HUL_ZD',hZ)
    dfPartone.insert(8,'Weight_RES',weight)
    return dfPartone

def twoPartOllyOuter(dfPartone,dfParttwo):
    topOD = dfParttwo.pop('Top_OD_DIA')     #need to test
    dfPartone.insert(6,'Top_OD_DIA',topOD) 
    return dfPartone

def twoDosage(dfPartone,dfParttwo):
    bW = dfParttwo.pop('BW_RES')
    weight = dfParttwo.pop('WEIGHT_RES')     #THIS IS DONE, MAYBE? I NEED TO TEST THE INDEX POSITIONS
    dfPartone.insert(4,'BW_RES',bW)
    dfPartone.insert(5,'WEIGHT_RES',weight)
    return dfPartone

def twoPartOllyInner(dfPartone,dfParttwo):
    domeHeight = dfParttwo.pop('Dome_Height_RES')
    weight = dfParttwo.pop('Part_Weight')     #THIS IS DONE, MAYBE? I NEED TO TEST THE INDEX POSITIONS
    dfPartone.insert(3,'Dome_Height_RES',domeHeight)
    dfPartone.insert(4,'Part_Weight',weight) 
    return dfPartone

def checkPartno(part):
    sql = """SELECT Part_number, Part_Type FROM Part_Numbers2 WHERE Part_number = ?""" #provides SQL queury statement with option for parameter
    confirmedPartType = False
    while confirmedPartType is False:
        partDB = pd.read_sql_query(sql, conn,params=[part])  #fetchs the line item in the DB file matching the part #
        partConfirmationCheck = partDB["Part_number"].loc[0] #extracts only the part type, to check for two part program
        if partConfirmationCheck == part: confirmedPartType = True
        else: 
            part = input('The Given part number is not recognized, please re-enter the part number:')
            continue    #add fallthru logic if a user cancels this step. 
        partnosql = partDB["Part_Type"].loc[0] #extracts only the part type, to check for two part program
    return partnosql

def grabfilenameData(location,workOrder):   #works
    trackerData = pd.read_excel(location, sheet_name = 'Production',dtype=str, engine = 'openpyxl')
    trackerData.columns = [column.replace(" ", "_") for column in trackerData.columns]
    trackerData.query("Work_Order == @workOrder", inplace=True)
    while trackerData.empty:
        newWo = str(input('The entered work order is not in the daily tracker, please reenter the work order number:'))
        trackerData = pd.read_excel(location,'Production',dtype=str)
        trackerData.columns = [column.replace(" ", "_") for column in trackerData.columns]
        trackerData.query("Work_Order == @newWo", inplace=True)        
    else:
        return trackerData

def namer(trackerData):
    sql = """SELECT Part_number, Part_Type, Naming_Specific FROM Part_Numbers2 WHERE Part_number = ?"""
    part = trackerData['Product_Code'].iloc[0]
    partDB = pd.read_sql_query(sql, conn,params=[part])
    specific = partDB['Naming_Specific'].iloc[0]
    if specific == None:
        filename = str(str(trackerData['Work_Order'].iloc[0]) + ' ' + str(trackerData['Product_Code'].iloc[0]) + ' ' + str(trackerData['Cav'].iloc[0]) + 'cav ' + str(trackerData['Mold_#'].iloc[0]) + '.csv')
        return filename
    elif specific == 'Resin Specific':
        if trackerData['Product_Code'].iloc[0] == 'CI038' and trackerData['Material'].iloc[0] == 'CP0001':
            filename = str(str(trackerData['Work_Order'].iloc[0]) + ' ' + str(trackerData['Product_Code'].iloc[0]) + ' ' + str(trackerData['Cav'].iloc[0]) + 'cav ' + str(trackerData['Mold_#'].iloc[0]) + '.csv')
            return filename
        else:
            resinCode = resins[trackerData['Material'].iloc[0]]
            filename = str(str(trackerData['Work_Order'].iloc[0]) + ' ' + str(trackerData['Product_Code'].iloc[0]) + resinCode + ' ' + str(trackerData['Cav'].iloc[0]) + 'cav ' + str(trackerData['Mold_#'].iloc[0]) + '.csv') 
            return filename
    elif specific == 'Mold Specific':
        filename = str(str(trackerData['Work_Order'].iloc[0]) + ' ' + str(trackerData['Product_Code'].iloc[0]) + '-mold-' + str(trackerData['Mold_#']) + ' ' + str(trackerData['Cav'].iloc[0]) + 'cav ' + str(trackerData['Mold_#'].iloc[0]) + '.csv')
        return filename
    elif specific == 'Customer Specific':
        filename = str(str(trackerData['Work_Order'].iloc[0]) + ' ' + str(trackerData['Product_Code'].iloc[0]) + ' ' + str(trackerData['Cav'].iloc[0]) + 'cav ' + str(trackerData['Mold_#'].iloc[0]) + '.csv')
        return filename

def main():
    global shotCounter
    dailyTracker ='G:\\SHARED\\QA\\SPC Daily Tracker\\2023 SPC Daily Tracker.xlsm'
    excFileLocation = "\\\\beowulf.mold-rite.local\\spc\\ogptest.xls"
    outputDir = '\\\\lighthouse2020\\Data Import\\Production\\Testing'
    while shotCounter < 60:
        shotCounter = shotCounter + 1
        try:
            print("Grabbing Data")
            dfObject,workOrder,lastRow,partType = grabData(excFileLocation,shotCounter)
            print("Checking PartNo")
            partnosql = checkPartno(partType)
            print(f"Grabbing filenamedata {workOrder}")
            trackerData = grabfilenameData(dailyTracker, workOrder)
            print("Formatting QC to DF")
            dfObject = formatQCtoDF(dfObject,lastRow,workOrder)
            print("Fetching name from Namer")
            filename = namer(trackerData)
            print("Submitting shot to B drive")
            submitshots(dfObject, filename, outputDir)
#            while(wdEventHandler.uploadDispatchState is True):
#                print("Waiting on watchdog")
#                time.sleep(1)
            
        except:
            pass

    return
###
#   GUI
###

#to-do: move all GUI initialization to class system, these don't belong in the global scope

mainGUI = tk.Tk()
mainGUI.title("OGP Interface")
for num in range(1, 5): [mainGUI.columnconfigure(num, minsize = 15), mainGUI.rowconfigure(num, minsize = 15)]

#redefine the mainGUI grid layout. 
#maybe add a checkbox for validations? rather than production shots
tk.Frame(mainGUI).grid(column = 1, row = 1)
tk.Frame(mainGUI).grid(column = 4, row = 1)
tk.Frame(mainGUI).grid(column = 1, row = 4)
tk.Frame(mainGUI).grid(column = 4, row = 4)

tk.Button(mainGUI, text = "Submit Shot", command = main).grid(column = 2, row = 2)

###
#   Classes
###
class ogpHandler(FileSystemEventHandler):
    def __init__(self):
        self.mostRecentShot = ''
        self.uploadDispatchState = False    #indicates that a CSV file is being saved to the B drive for upload
        self.uploadStatus = False           #indicates that SFOL accepted or rejected the CSV file
    def on_modified(self, event):
        mRS = self.mostRecentShot
        print(event.src_path)
        if self.uploadDispatchState is not True: self.uploadDispatchState = True
        match event.src_path:
            case 'backup':
                self.uploadStatus = True
                self.uploadDispatchState = False
            case 'suspect':
                self.uploadStatus = False
                self.uploadDispatchState = False
            case _:
                pass

###
#   Global variables 
###

shotCounter = 0
outputDir = '\\\\lighthouse2020\\Data Import\\Production\\Testing'
file_path = os.path.abspath(os.path.dirname(__file__))
conn = sqlite3.connect(str(file_path + '\\Part_Numbers2.db')) #small database of partnumbers for verification and checking for two part programs
c = conn.cursor()
 #to be read for up to date part data
testWatchDog = Observer()
wdEventHandler = ogpHandler()
testWatchDog.schedule(wdEventHandler, path = outputDir, recursive = True)
testWatchDog.start()
resins = {'MRP-PP30-1':'PP','PS3101':'PS','CP0001':'CP','PPSR549M':'CP','HDPE 5618':'HD','PA68253 ULTRAMID':'-Nylon'}

###
#   Entrypoint
###

#tkinter's *.mainloop() function fires off a blocking event loop. use tkinter's .after() method to schedule a function call with tkinter's event loop.
mainGUI.mainloop()