import pandas as pd 
import os
import openpyxl

excFileLocation = "\\\\beowulf.mold-rite.local\\spc\\ogptest.xls"
twoPartprogramsCRC = ['CI020','CI022','CI024','CI028','CI033','CI038','CI045','CI053','CI063','CI058','CI070','CI089'] #Criteria lists for comparing part # 
dosagePrograms = ['DF030','DF020','DF024','DF028','DG024','DI024','DJ024']
ollyOuter = ['YY058','YY073','YO073','YO058']

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

def twoPartCRC(dfPartone,dfParttwo):
    topOD = dfParttwo.pop('Top_OD_DIA')
    hZ = dfParttwo.pop('HUL_ZD')
    weight = dfParttwo.pop('Weight_RES')
    dfPartone.insert()
    return

dfObject,lastRow,workOrder = grabData(excFileLocation,1)
csvExport = formatQCtoDF(dfObject,lastRow,workOrder)
csvExport.to_csv('test.csv')
