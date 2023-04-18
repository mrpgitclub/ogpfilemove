
import pandas as pd
import tkinter as tk
import tkinter.ttk as ttk
import os
import sqlite3
from sqlite3 import connect


#import matplotlib.pyplot as plt    #not implemented yet
#import numpy as np                 #not implemented yet

###
#   Global variables 
###

dfList = dict()
excFileLocation = "\\\\beowulf.mold-rite.local\\spc\\ogptest.xls"
conn = sqlite3.connect('Part_Numbers.db') #small database of partnumbers for verification and checking for two part programs
c = conn.cursor()
dailyTracker ='G:\\SHARED\\QA\\SPC Daily Tracker\\SPC Daily Tracker.xlsm' #to be read for up to date part data

###
#   Functions
###

def submitshots():
    #filename = str(str(int(dfObject.at[0, 'Work Order'])) + ' ' + str(dfObject.at[0,'Product Code']) + ' ' + str(len(dfObject)) + 'cav ' + str(int(dfObject.at[0,'MOLD Number'])) + '.csv')
    #dfObject.to_csv(str(dir + filename), header = False, index = False)

    return

excFileLocation = "\\\\beowulf.mold-rite.local\\spc\\ogptest.xls"

def grabData(location,num):
    dfObject = pd.read_excel(location, sheet_name = num, header = 0, index_col = None, usecols = None, dtype=str) #reads export file and takes data from specified sheet
    dfObject.columns = [column.replace(" ", "_") for column in dfObject.columns] #replace spaces with underscores for formatting
    lastRow = dfObject.iloc[-1] #grab the last row 
    partType = lastRow["Product_Code"] #reads product code from last row
    workOrder = lastRow["Work_Order"] #grabs the correct work order from the last row
    return dfObject,lastRow,workOrder,partType

def formatQCtoDF(dataframe,lastRow,workOrder):    
    dataframe.query("Work_Order == @workOrder", inplace=True) #selects only the rows with the workorder
    dataframe.drop_duplicates(keep = 'last', inplace = True, ignore_index = True, subset = 'Cavity') #remove extra lines from partial shots
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
    topOD = dfParttwo.pop('Top_OD_DIA') #THIS IS NOT DONE
    dfPartone.insert(6,'Top_OD_DIA',topOD) 
    return dfPartone

def checkPartno(part):
    sql = """SELECT Part_number, Part_Type FROM Part_Numbers WHERE Part_number = ?""" #provides SQL queury statement with option for parameter
    partDB = pd.read_sql_query(sql, conn,params=[part])  #fetchs the line item in the DB file matching the part #
    partnosql = partDB["Part_Type"].loc[0] #extracts only the part type, to check for two part program
    return partnosql


def grabfilenameData(location,workOrder):
    trackerData = pd.read_excel(location,'Production',dtype=str)
    trackerData.columns = [column.replace(" ", "_") for column in trackerData.columns]
    trackerData.query("Work_Order == @workOrder", inplace=True)
    while trackerData.empty:
        newWo = input('The entered work order is not in the daily tracker, please reenter the work order number:')
        str(newWo)
        print(newWo)
        trackerData = pd.read_excel(location,'Production',dtype=str)
        trackerData.columns = [column.replace(" ", "_") for column in trackerData.columns]
        trackerData.query("Work_Order == @newWo", inplace=True)        
    else:
        return trackerData
    #use dataframe workorder number to lookup the file name items in SPC daily tracker, return this as well
    
def main(excFileLocation):
    mainshot,msLast,msWo,msPartno = grabData(excFileLocation,1)

    return
###
#   GUI
###

mainGUI = tk.Tk()
mainGUI.title("OGP Interface")
for num in range(1, 5): [mainGUI.columnconfigure(num, minsize = 15), mainGUI.rowconfigure(num, minsize = 15)]

#redefine the mainGUI grid layout. 
#maybe add a checkbox for validations? rather than production shots
tk.Frame(mainGUI).grid(column = 1, row = 1)
tk.Frame(mainGUI).grid(column = 4, row = 1)
tk.Frame(mainGUI).grid(column = 1, row = 4)
tk.Frame(mainGUI).grid(column = 4, row = 4)

tk.Button(mainGUI, text = "Submit Shot", command = submitshots).grid(column = 2, row = 2)


#tkinter's *.mainloop() function fires off a blocking event loop. use tkinter's .after() method to schedule a function call with tkinter's event loop.
main(excFileLocation)
mainGUI.mainloop()