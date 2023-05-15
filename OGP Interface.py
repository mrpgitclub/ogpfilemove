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
import pyodbc





import openpyxl

#import matplotlib.pyplot as plt    #not implemented yet
#import numpy as np                 #not implemented yet

###
#   Functions
###

def dataVerify(dataframe,trackerdata):
    dataframe.iloc[0,-8] = trackerdata['Mold_#'].iloc[0]
    dataframe.iloc[0,-7] = trackerdata['Work_Order'].iloc[0]
    dataframe.iloc[0,-2] = trackerdata['Product_Code'].iloc[0]

    return dataframe

def submitshots(dfObject,filename,outputDir):
    global wdEventHandler #remove? need to confirm
    dfObject.to_csv(str(outputDir + '\\' + filename), header = False, index = False)
    wdEventHandler.mostRecentShot = filename
    wdEventHandler.uploadDispatchState = True

    return

def grabData(location,num):
    dfObject = pd.read_excel(location, sheet_name = num, header = 0, index_col = None, usecols = None, dtype=str) #reads export file and takes data from specified sheet
    dfObject.columns = [column.replace(" ", "_") for column in dfObject.columns] #replace spaces with underscores for formatting
    lastRow = dfObject.iloc[-1] #grab the last row 
    partType = lastRow["Product_Code"] #reads product code from last row
    workOrder = lastRow["Work_Order"] #grabs the correct work order from the last row

    return dfObject,workOrder,partType

def formatQCtoDF(dataframe,workOrder):
    workOrder = dataframe.iloc[-1]["Work_Order"]
    dataframe.query("Work_Order == @workOrder", inplace=True) #selects only the rows with the workorder
    dataframe.drop_duplicates(keep = 'last', inplace = True, ignore_index = True, subset = 'Cavity') #remove extra lines from partial shots
    dataframe.dropna(axis = 1, how = 'all', inplace = True)
    dataframe.pop('Fails') #delete fails column
    dataframe.insert(len(dataframe.columns),'Date_Time', dataframe.pop('Date_Time')) # Replaces the date time column to the correct location

    return dataframe

def mergeTwoDataframes(dfObject, second_dfObject, partnoSql):
    match partnoSql:
        case 'CRC Inner': dfObject = twoPartCRC(dfObject, second_dfObject)
        case 'Olly Outer': dfObject = twoPartOllyOuter(dfObject, second_dfObject)
        case 'Olly Inner': dfObject = twoPartOllyInner(dfObject, second_dfObject)
        case 'dosage cup': dfObject = twoDosage(dfObject, second_dfObject)
        case _: pass
    return dfObject

def twoPartCRC(dfPartone,dfParttwo): #to execute when the part number correlates to a two part crc inner program 
    dfPartone.insert(6,'TOP_OD_DIA', dfParttwo.pop('TOP_OD_DIA'))   #todo- index columns by number rather than name. 
    dfPartone.insert(7,'HUL_ZD',dfParttwo.pop('HUL_ZD'))
    dfPartone.insert(8,'Weight_RES',dfParttwo.pop('Weight_RES'))
    return dfPartone

def twoPartOllyOuter(dfPartone,dfParttwo):
    dfPartone.insert(6,'Top_OD_DIA', dfParttwo.pop('Top_OD_DIA')) #todo- index columns by number rather than name. 
    return dfPartone

def twoDosage(dfPartone,dfParttwo):
    dfPartone.insert(4,'BW_RES',dfParttwo.pop('BW_RES'))
    dfPartone.insert(5,'WEIGHT_RES', dfParttwo.pop('WEIGHT_RES'))
    return dfPartone

def twoPartOllyInner(dfPartone,dfParttwo):
    dfPartone.insert(2,'Dome_Height_RES',dfParttwo.pop('Dome_Height_RES'))
    dfPartone.insert(3,'Part_Weight_RES',dfParttwo.pop('Part_Weight_RES'))
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
    second_dfObject = None
    dailyTracker ='G:\\SHARED\\QA\\SPC Daily Tracker\\2023 SPC Daily Tracker.xlsm'
    dbFileLocation = "\\\\beowulf.mold-rite.local\\spc\\ogptest.mdb"
    outputDir = '\\\\lighthouse2020\\Data Import\\Production\\Testing'
    tableIndex = 1
    #improve error handling in the code block below. replace try/except
    try:
        dfObject,workOrder,lastRow,partType = grabData(dbFileLocation, tableIndex)
        partnoSql = checkPartno(partType)   #check for two part programs here
        trackerData = grabfilenameData(dailyTracker, workOrder)
        dfObject = formatQCtoDF(dfObject,lastRow,workOrder) #implement two part programs here
        if partnoSql is not None: 
            second_dfObject = grabData(dbFileLocation, tableIndex + 1)
            dfObject = mergeTwoDataframes(dfObject, second_dfObject, partnoSql)
        filename = namer(trackerData)
        submitshots(dfObject, filename, outputDir)
        while(wdEventHandler.uploadDispatchState is True):
            time.sleep(1)
    except:
        pass
    finally:
        pass #delete worksheets here
    return
###
#   GUI
###

#to-do: move all GUI initialization to class system, these don't belong in the global scope

mainGUI = tk.Tk()
mainGUI.title("OGP Interface")
for num in range(1, 5): [mainGUI.columnconfigure(num, minsize = 15), mainGUI.rowconfigure(num, minsize = 15)]

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
        if self.uploadDispatchState is not True: self.uploadDispatchState = True
        if event.src_path.find('backup') > -1:
            self.uploadStatus = True
            self.uploadDispatchState = False
        elif event.src_path.find('suspect') > -1:
            self.uploadStatus = False
            self.uploadDispatchState = False
        else: pass

###
#   Global variables 
###

shotCounter = 0
outputDir = '\\\\lighthouse2020\\Data Import\\Production\\Testing'
file_path = os.path.abspath(os.path.dirname(__file__))
conn = sqlite3.connect(str(file_path + '\\Part_Numbers2.db')) #small database of partnumbers for verification and checking for two part programs
c = conn.cursor() #to be read for up to date part data
testWatchDog = Observer()
wdEventHandler = ogpHandler()
testWatchDog.schedule(wdEventHandler, path = outputDir, recursive = True)
testWatchDog.start()
resins = {'MRP-PP30-1':'PP','PS3101':'PS','CP0001':'CP','PPSR549M':'CP','HDPE 5618':'HD','PA68253 ULTRAMID':'-Nylon'}

conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=S:\ogptest.mdb;'
    )
cnxn = pyodbc.connect(conn_str)
crsr = cnxn.cursor()




###
#   Entrypoint
###

mainGUI.mainloop()