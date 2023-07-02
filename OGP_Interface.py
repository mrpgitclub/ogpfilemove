import sys
import os
import pandas as pd
import tkinter as tk
from tkinter import simpledialog
import tkinter.ttk as ttk
import sqlite3
import time
from tkinter import messagebox
from sqlite3 import connect
import pyodbc
#import matplotlib.pyplot as plt    #not implemented yet
#import numpy as np                 #not implemented yet

###
#   Functions
###

def dropTable(crsr, tableName):
    crsr.execute(f'DROP TABLE "{tableName}"') #TEST THIS A LOT
    crsr.commit()
    print(f"Table Dropped {tableName}")

    return

def watchdog(filename):
    if os.path.isfile(outputDir + '\\backup\\' + filename) == True: return [True, outputDir + '\\backup\\' + filename]
    elif os.path.isfile(outputDir + '\\suspect\\' + filename) == True: return [False, outputDir + '\\suspect\\' + filename]
    else: pass
    return [None,]

def dataVerify(dataframe,trackerdata):
    dataframe.iloc[0,-8] = trackerdata['Mold_#'].iloc[0]
    dataframe.iloc[0,-7] = trackerdata['Work_Order'].iloc[0]
    dataframe.iloc[0,-2] = trackerdata['Product_Code'].iloc[0]

    return dataframe

def submitshots(dfObject, filename, opCode): 
    #assign different output directories depending on opcode?
    if opCode == 0: saveFolder = outputDir
    else: saveFolder = os.environ['USERPROFILE'] + '\\Desktop'
    dfObject.to_csv(str(saveFolder + '\\' + filename), header = False, index = False, date_format = '%m/%d/%Y %H:%M')

    return

def grabData(crsr, tableList, tableIndex = 0): #use twoPartProgIndicator to fetch the first table or the second table
    dfObject = pd.read_sql_query(f'SELECT * FROM "{tableList[tableIndex]}"', crsr)
    if dfObject.size > 0: dfObject.columns = [column.replace(" ", "_") for column in dfObject.columns] #replace spaces with underscores for formatting
    dfObject = dfObject.applymap(lambda x: x.strip() if isinstance(x, str) else x) #strip whitespace in a dataframe
    dfObject = dfObject.applymap(lambda x: round(x, 4) if isinstance(x, float) else x) #strip whitespace in a dataframe

    return dfObject

def formatQCtoDF(dataframe):
    workOrder = dataframe['Work_Order'].iloc[-1]
    dataframe.query("Work_Order == @workOrder", inplace=True) #selects only the rows with the workorder
    dataframe.drop_duplicates(keep = 'last', inplace = True, ignore_index = True, subset = 'Cavity') #remove extra lines from partial shots
    dataframe.dropna(axis = 1, how = 'all', inplace = True)
    
    for aCol in dataframe.columns:
        if str(dataframe[aCol].iloc[0]).isspace(): dataframe.drop(aCol, axis = 1, inplace = True)

    dataframe.pop('Fails') #delete fails column
    dataframe.insert(len(dataframe.columns) - 1,'Date_Time', dataframe.pop('Date_Time')) # Replaces the date time column to the correct location
    dataframe["Date_Time"] = pd.to_datetime(dataframe["Date_Time"], format = 'ISO8601')

    return dataframe

def rawDataformatQCtoDF(dataframe): #make a condensed version of this for raw data export
    if 'Cavity' in dataframe.columns: dataframe.query("Cavity != 0", inplace=True)
    if len(dataframe) == 0: return dataframe
    dataframe.dropna(axis = 1, how = 'all', inplace = True)
    
    for aCol in dataframe.columns:
        if str(dataframe[aCol].iloc[0]).isspace(): dataframe.drop(aCol, axis = 1, inplace = True)

    dataframe.pop('Fails') #delete fails column
    dataframe.insert(len(dataframe.columns) - 1,'Date_Time', dataframe.pop('Date_Time')) # Replaces the date time column to the correct location
    dataframe["Date_Time"] = pd.to_datetime(dataframe["Date_Time"], format = 'ISO8601')

    return dataframe

def mergeTwoDataframes(dfObject, second_dfObject, partnoSql):
    if partnoSql == 'CRC Inner': dfObject = twoPartCRC(dfObject, second_dfObject)
    elif partnoSql == 'Olly Outer': dfObject = twoPartOllyOuter(dfObject, second_dfObject)
    elif partnoSql == 'Olly Inner': dfObject = twoPartOllyInner(dfObject, second_dfObject)
    elif partnoSql == 'dosage cup': dfObject = twoDosage(dfObject, second_dfObject)
    else: pass

    return dfObject

def twoPartCRC(dfPartone,dfParttwo): #to execute when the part number correlates to a two part crc inner program 
    dfPartone.insert(6,'TOP_OD_DIA',dfParttwo.pop(dfParttwo.columns[0]))   #swapped to index call- testing needed
    dfPartone.insert(7,'HUL_ZD',dfParttwo.pop(dfParttwo.columns[0]))
    dfPartone.insert(8,'Weight_RES',dfParttwo.pop(dfParttwo.columns[0]))
    return dfPartone

def twoPartOllyOuter(dfPartone,dfParttwo):
    dfPartone.insert(0,'Top_OD_DIA', dfParttwo.pop(dfParttwo.columns[0]))  #swapped to index call- testing needed
    return dfPartone

def twoDosage(dfPartone,dfParttwo):
    dfPartone.insert(4,'BW_RES',dfParttwo.pop(dfParttwo.columns[0]))    #swapped to index call- testing needed
    dfPartone.insert(5,'WEIGHT_RES',dfParttwo.pop(dfParttwo.columns[0]))
    return dfPartone

def twoPartOllyInner(dfPartone,dfParttwo):
    dfPartone.insert(2,'Dome_Height_RES',dfParttwo.pop(dfParttwo.columns[0])) #swapped to index call- testing needed
      #todo- index columns by number rather than name. 
    dfPartone.insert(3,'Part_Weight_RES',dfParttwo.pop(dfParttwo.columns[0]))
    return dfPartone

def checkPartno(part, conn):
    sql = "SELECT Part_Type FROM Part_Numbers2 WHERE Part_number = ?" #provides SQL queury statement with option for parameter
    confirmedPartExists = False
    partnoSql = None

    #confirm part exists in the DB
    while confirmedPartExists is False:
        partDB = pd.read_sql_query("SELECT Part_number FROM Part_Numbers2 WHERE Part_number = ?", conn,params=[part])  #fetchs the line item in the DB file matching the part
        if partDB.size == 0:
            part = tk.simpledialog.askstring('OGP Interface', f'The Given part number "{part}" is not recognized, please re-enter the part number:')
            if part is None: 
                partnoSql = False
                break
        else: confirmedPartExists = True

    #then fetch the part type
    if partnoSql != False:
        partDB = pd.read_sql_query(sql, conn,params=[part])  #fetchs the line item in the DB file matching the part
        if partDB.size == 0: partnoSql = None #extracts only the part type, to check for two part program
        else: partnoSql = partDB["Part_Type"].loc[0] #extracts only the part type, to check for two part program
    
    return partnoSql

def grabfilenameData(location,workOrder):   #works
    trackerData = pd.read_excel(location, sheet_name = 'Production',dtype=str)
    trackerData.columns = [column.replace(" ", "_") for column in trackerData.columns]
    trackerData.query("Work_Order == @workOrder", inplace=True)
    while trackerData.empty:
        workOrder = tk.simpledialog.askstring('Wrong Workorder', f'Unable to find WO#: {workOrder}, try again')

        if workOrder is None:
            trackerData = None
            break

        trackerData = pd.read_excel(location, sheet_name = 'Production',dtype=str)
        trackerData.columns = [column.replace(" ", "_") for column in trackerData.columns]
        trackerData.query("Work_Order == @workOrder", inplace=True)
    return trackerData

def namer(trackerData, conn):
    filename = None
    specific = pd.read_sql_query("SELECT Part_number, Part_Type, Naming_Specific FROM Part_Numbers2 WHERE Part_number = ?", conn,params=[trackerData['Product_Code'].iloc[0]])['Naming_Specific'].iloc[0]

    if specific == 'Resin Specific':
        if trackerData['Product_Code'].iloc[0] == 'CI038' and trackerData['Material'].iloc[0] == 'CP0001': filename = str(str(trackerData['Work_Order'].iloc[0]) + ' ' + str(trackerData['Product_Code'].iloc[0]) + ' ' + str(trackerData['Cav'].iloc[0]) + 'cav ' + str(trackerData['Mold_#'].iloc[0]) + '.csv')
        elif trackerData['Product_Code'].iloc[0] == 'CI038' and trackerData['Material'].iloc[0] != 'CP0001': filename = str(str(trackerData['Work_Order'].iloc[0]) + ' ' + str(trackerData['Product_Code'].iloc[0]) + '-' + resins[str(trackerData['Material'].iloc[0])] + ' ' + str(trackerData['Cav'].iloc[0]) + 'cav ' + str(trackerData['Mold_#'].iloc[0]) + '.csv')
        else: filename = str(str(trackerData['Work_Order'].iloc[0]) + ' ' + str(trackerData['Product_Code'].iloc[0]) + resins[str(trackerData['Material'].iloc[0])] + ' ' + str(trackerData['Cav'].iloc[0]) + 'cav ' + str(trackerData['Mold_#'].iloc[0]) + '.csv')
    elif specific == 'Mold specific':
        filename = str(str(trackerData['Work_Order'].iloc[0]) + ' ' + str(trackerData['Product_Code'].iloc[0]) + '-mold-' + str(trackerData['Mold_#'].iloc[0]) + ' ' + str(trackerData['Cav'].iloc[0]) + 'cav ' + str(trackerData['Mold_#'].iloc[0]) + '.csv')
    elif specific == 'Customer Specific':
        filename = str(str(trackerData['Work_Order'].iloc[0]) + ' ' + str(trackerData['Product_Code'].iloc[0]) + ' ' + str(trackerData['Cav'].iloc[0]) + 'cav ' + str(trackerData['Mold_#'].iloc[0]) + '.csv')
    else: 
        filename = str(str(trackerData['Work_Order'].iloc[0]) + ' ' + str(trackerData['Product_Code'].iloc[0]) + ' ' + str(trackerData['Cav'].iloc[0]) + 'cav ' + str(trackerData['Mold_#'].iloc[0]) + '.csv')

    return filename

def main(opCode = 0): 
    dailyTracker ='G:\\SHARED\\QA\\SPC Daily Tracker\\2023 SPC Daily Tracker.xlsm'
    twoPartProgramPartTypes = ['CRC Inner', 'Olly Outer', 'Olly Inner', 'dosage cup']
    conn_str = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        r'DBQ=S:\\ogptest.mdb;'
        )
    filename = f'rawdata {time.time()}.csv'
    partnoSql = None
    conn = sqlite3.connect(str(file_path + '\\Part_Numbers2.db')) #small database of partnumbers for verification and checking for two part programs
    cnxn = pyodbc.connect(conn_str)
    crsr = cnxn.cursor()

    #confirm files exist
    if not os.path.isfile(dailyTracker): return messagebox.showinfo(title = "OGP Interface", message = "Unable to locate SPC Daily Tracker")
    if not os.path.isfile(str(file_path + '\\Part_Numbers2.db')): return messagebox.showinfo(title = "OGP Interface", message = "Unable to connect to database: 0x0001")
    if not os.path.isfile(str('S:\\ogptest.mdb')): return messagebox.showinfo(title = "OGP Interface", message = "Unable to connect to database: 0x0002")
    
    tableList = list()
    for table_info in crsr.tables(tableType = 'TABLE'): tableList.append(table_info.table_name)
    
    if len(tableList) < 1: return messagebox.showinfo(title = "OGP Interface", message = "Unable to find measurements in the OGP: 0x0003")

    dfObject = grabData(cnxn, tableList)
    if dfObject.size == 0: return messagebox.showinfo(title = "OGP Interface", message = f"Unable to find measurements in the OGP: 0x0004\r\n{dfObject.iloc[-1:-2]}")

    if opCode == 0:
        if 'Work_Order' not in dfObject.columns: return messagebox.showinfo(title = "OGP Interface", message = f"Unable to find a work order in the measurements: 0x0008\r\n{dfObject.iloc[-1:-2]}")
        trackerData = grabfilenameData(dailyTracker, dfObject.iloc[-1]["Work_Order"].strip())
        if trackerData is None: return messagebox.showinfo(title = "OGP Interface", message = f"Unable to find data in SPC Daily Tracker: 0x0006\r\n{str(dfObject.iloc[-1:-2])}")

        partnoSql = checkPartno(str(trackerData['Product_Code'].iloc[0]), conn)   #check for two part programs here
        if partnoSql is False: return messagebox.showinfo(title = "OGP Interface", message = f"Unable to find measurements in the OGP: 0x0005\r\n{str(dfObject.iloc[1:2])}")

        dfObject = formatQCtoDF(dfObject)
        if dfObject.size == 0: return messagebox.showinfo(title = "OGP Interface", message = f"Unable to find measurements in the OGP: 0x0007")

        if partnoSql in twoPartProgramPartTypes: 
            second_dfObject = grabData(cnxn, tableList, 1)
            dfObject = mergeTwoDataframes(dfObject, second_dfObject, partnoSql)

        filename = namer(trackerData, conn)
        if filename is None: 
            wonum = trackerData['Work_Order'].iloc[0]
            filename = tk.simpledialog.askstring('OGP Interface', f'Couldn\'t create a filename based on this workorder number: "{wonum}". Input a filename for the B drive: ')
    if opCode != 0:
        dfObject = rawDataformatQCtoDF(dfObject)

    if (dfObject.size) > 0:
        submitshots(dfObject, filename, opCode)  #implement raw data export, cherry pick certain functions from main()
    else: 
        dropTable(crsr, tableList.pop(0))
        return

    if opCode != 0: 
        dropTable(crsr, tableList.pop(0))
        return

    timeout = 0

    while((timeout < 10) and (watchdog(filename)[0] is None)):
        time.sleep(1)
        timeout += 1
    else: 
        if (watchdog(filename)[0] is True):
            dropTable(crsr, tableList.pop(0))
            if opCode == 0 and partnoSql in twoPartProgramPartTypes:
                dropTable(crsr, tableList.pop(0))
            #os.remove(watchdog(filename))  #delete the file from backup folder
        elif (watchdog(filename)[0] is False):
            messagebox.showinfo(title = "OGP Interface", message = "Failed to upload.")
            os.remove(outputDir + '\\suspect\\' + filename)  #delete the file from suspect folder
        else: messagebox.showinfo(title = "OGP Interface", message = "Operation timed out. Try again.")

    conn.close()
    crsr.close()
    cnxn.close()

    return
###
#   GUI
###

mainGUI = tk.Tk()
mainGUI.title("OGP Interface")
mainGUI.attributes('-topmost', 'true')
for num in range(1, 5): [mainGUI.columnconfigure(num, minsize = 15), mainGUI.rowconfigure(num, minsize = 15)]

tk.Frame(mainGUI).grid(column = 1, row = 1)
tk.Frame(mainGUI).grid(column = 4, row = 1)
tk.Frame(mainGUI).grid(column = 1, row = 4)
tk.Frame(mainGUI).grid(column = 4, row = 4)

tk.Button(mainGUI, text = "Production", command = lambda: main(0)).grid(column = 2, row = 2)
tk.Button(mainGUI, text = "Non Production", command = lambda: main(1)).grid(column = 3, row = 2)

###
#   Global variables 
###

outputDir = '\\\\lighthouse2020\\Data Import\\Production\\Testing'
file_path = os.path.abspath(os.path.dirname(__file__))
resins = {'MRP-PP30-1':'PP','PS3101':'PS','CP0001':'CP','PPSR549M':'CP','HDPE 5618':'HD','PA68253 ULTRAMID':'-Nylon'}

###
#   Entrypoint
###

mainGUI.mainloop()