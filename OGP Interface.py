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
    print(f'Submitted:{outputDir}\\{filename}', )
    return

def grabData(location,num):
    dfObject = pd.read_excel(location, sheet_name = num, header = 0, index_col = None, usecols = None, dtype=str) #reads export file and takes data from specified sheet
    dfObject.columns = [column.replace(" ", "_") for column in dfObject.columns] #replace spaces with underscores for formatting
    lastRow = dfObject.iloc[-1] #grab the last row 
    partType = lastRow["Product_Code"] #reads product code from last row
    workOrder = lastRow["Work_Order"] #grabs the correct work order from the last row
    return dfObject,workOrder,partType

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
    dfPartone.insert(6,'Top_OD_DIA',topOD) #take the previous three variables and place them in the correct column positions for shopfloor to read them
    dfPartone.insert(7,'HUL_ZD',hZ)
    dfPartone.insert(8,'Weight_RES',weight)
    return dfPartone

def twoPartOllyOuter(dfPartone,dfParttwo):
    topOD = dfParttwo.pop('Top_OD_DIA')     #THIS IS NOT DONE
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
    sql = """SELECT Part_number, Part_Type FROM Part_Numbers WHERE Part_number = ?""" #provides SQL queury statement with option for parameter
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

def namer(dfObject, trackerData):    #this needs logic to determine materical composition of a jar
    filename = str(str(dfObject['Work_Order'].iloc[0]) + ' ' + str(trackerData['Product_Code'].iloc[0]) + ' ' + str(trackerData['Cav'].iloc[0]) + 'cav ' + str(trackerData['Mold_#'].iloc[0]) + '.csv')
    return filename

def main():
    global shotCounter
    dailyTracker ='G:\\SHARED\\QA\\SPC Daily Tracker\\2023 SPC Daily Tracker.xlsm'
    excFileLocation = "\\\\beowulf.mold-rite.local\\spc\\ogptest.xls"
    outputDir = '\\\\lighthouse2020\\Data Import\\Production\\Testing'
    while shotCounter < 60:
        shotCounter = shotCounter + 1
        try:
            dfObject,lastRow,workOrder,partType = grabData(excFileLocation,shotCounter)
            partnosql = checkPartno(partType)
            trackerData = grabfilenameData(dailyTracker, workOrder)
            dfObject = formatQCtoDF(dfObject,lastRow,workOrder)
            filename = namer(dfObject, trackerData)
            submitshots(dfObject, filename, outputDir)
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
    def on_modified(self, event):
        print(f'Modification Detected {event.src_path}')


###
#   Global variables 
###

shotCounter = 0
outputDir = '\\\\lighthouse2020\\Data Import\\Production\\Testing'
file_path = os.path.abspath(os.path.dirname(__file__))
conn = sqlite3.connect(str(file_path + '\\Part_Numbers.db')) #small database of partnumbers for verification and checking for two part programs
c = conn.cursor()
 #to be read for up to date part data
testWatchDog = Observer()
wdEventHandler = ogpHandler()
testWatchDog.schedule(wdEventHandler, path = outputDir)
testWatchDog.start()

###
#   Entrypoint
###

#tkinter's *.mainloop() function fires off a blocking event loop. use tkinter's .after() method to schedule a function call with tkinter's event loop.
mainGUI.mainloop()