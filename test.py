import pandas as pd 
import os
import openpyxl

excFileLocation = "\\\\beowulf.mold-rite.local\\spc\\ogptest.xls"

dfObject = pd.read_excel(excFileLocation, sheet_name = 1, header = 0, index_col = None, usecols = None, dtype=str)
dfObject.columns = [column.replace(" ", "_") for column in dfObject.columns]
lastRow = dfObject.iloc[-1] #grab the last row to pull the 
partType = lastRow["Product_Code"]
workOrder = lastRow["Work_Order"] #grabs the correct wo #
dfObject.query("Work_Order == @workOrder", inplace=True) #selects only the rows with the workorder
dfObject.drop_duplicates(keep = 'last', inplace = True, ignore_index = True, subset = 'Cavity') #remove extra lines from partial shots
dfObject.pop('Fails')
dateTime = dfObject.pop('Date_Time')
dfObject.insert(len(dfObject.columns),'Date_Time',dateTime)
dfObject.to_csv('test.csv')
