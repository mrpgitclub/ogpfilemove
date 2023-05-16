import pandas as pd 
import os
import openpyxl
import sqlite3
from sqlite3 import connect
import pyodbc

conn_str = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        r'DBQ=S:\\ogptest - Copy.mdb;'
        )
cnxn = pyodbc.connect(conn_str)
crsr = cnxn.cursor()
tableList = list()
for table_info in crsr.tables(tableType = 'TABLE'): 
    tableList.append(table_info.table_name)

excFileLocation = "\\\\beowulf.mold-rite.local\\spc\\ogptest.xls"
dailyTracker ='G:\\SHARED\\QA\\SPC Daily Tracker\\2023 SPC Daily Tracker.xlsm' #to be read for up to date part data
file_path = os.path.abspath(os.path.dirname(__file__))
conn = sqlite3.connect(str(file_path + '\\Part_Numbers2.db')) #small database of partnumbers for verification and checking for two part programs
c = conn.cursor()

resins = {'MRP-PP30-1':'PP','PS3101':'PS','CP0001':'CP','PPSR549M':'CP','HDPE 5618':'HD','PA68253 ULTRAMID':'-Nylon'}


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

def namer(dfObject):
    sql = """SELECT Part_number, Part_Type, Naming_Specific FROM Part_Numbers2 WHERE Part_number = ?"""
    part = dfObject['Product_Code'].iloc[0]
    partDB = pd.read_sql_query(sql, conn,params=[part])
    specific = partDB['Naming_Specific'].iloc[0]
    if specific == None:
        filename = str(str(dfObject['Work_Order'].iloc[0]) + ' ' + str(dfObject['Product_Code'].iloc[0]) + ' ' + str(dfObject['Cav'].iloc[0]) + 'cav ' + str(dfObject['Mold_#'].iloc[0]) + '.csv')
        return filename
    elif specific == 'Resin Specific':
        if dfObject['Product_Code'].iloc[0] == 'CI038' and dfObject['Material'].iloc[0] == 'CP0001':
            filename = str(str(dfObject['Work_Order'].iloc[0]) + ' ' + str(dfObject['Product_Code'].iloc[0]) + ' ' + str(dfObject['Cav'].iloc[0]) + 'cav ' + str(dfObject['Mold_#'].iloc[0]) + '.csv')
            return filename
        else:
            resinCode = resins[dfObject['Material'].iloc[0]]
            filename = str(str(dfObject['Work_Order'].iloc[0]) + ' ' + str(dfObject['Product_Code'].iloc[0]) + resinCode + ' ' + str(dfObject['Cav'].iloc[0]) + 'cav ' + str(dfObject['Mold_#'].iloc[0]) + '.csv') 
            return filename
    elif specific == 'Mold Specific':
        filename = str(str(dfObject['Work_Order'].iloc[0]) + ' ' + str(dfObject['Product_Code'].iloc[0]) + '-mold-' + str(dfObject['Mold_#']) + ' ' + str(dfObject['Cav'].iloc[0]) + 'cav ' + str(dfObject['Mold_#'].iloc[0]) + '.csv')
        return filename
    elif specific == 'Customer Specific':
        filename = str(str(dfObject['Work_Order'].iloc[0]) + ' ' + str(dfObject['Product_Code'].iloc[0]) + ' ' + str(dfObject['Cav'].iloc[0]) + 'cav ' + str(dfObject['Mold_#'].iloc[0]) + '.csv')
        return filename

def grabData(location,num):
    dfObject = pd.read_excel(location, sheet_name = num, header = 0, index_col = None, usecols = None, dtype=str) #reads export file and takes data from specified sheet
    dfObject.columns = [column.replace(" ", "_") for column in dfObject.columns] #replace spaces with underscores for formatting
    lastRow = dfObject.iloc[-1] #grab the last row 
    partType = lastRow["Product_Code"] #reads product code from last row
    workOrder = lastRow["Work_Order"] #grabs the correct work order from the last row
    return dfObject,lastRow,workOrder

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
    topOD = dfParttwo.pop('Top_OD_DIA') #
    dfPartone.insert(6,'Top_OD_DIA',topOD)
    return dfPartone


