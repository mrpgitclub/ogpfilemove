"""
"OGP Interface" Application Project overview

Overview of feature requirements
1. GUI
2. Pandas Python Library
3. Parsing script
4. Graph

Detailed features
1. GUI
    1.1 Allow user to view measurements of current shot in a 2D grid (emulate QC-Calc)
        1.1.1 User can "review" the shot of data before sending it to SFOL
    1.2 Allow user to omit, remove individual measurements from the shot before sending it to SFOL
        1.2.1 Allows for filtering of step edits, random testing of new routines, etc
    1.3 Tkinter [In Process]
        1.3.1 Treeview widget to represent data in a grid [In Process]
        1.3.2 Button widget to "Submit shot", "Delete individual measurement(s)"

2. Pandas
    2.1 Collect individual measurements into a Dataframe
        2.1.1 Each dataframe created per OGP program
        2.1.2 Header information collected per part by the given work order number as user input, product codes need to be marshalled against available product codes in SFOL
    2.2 Ingest measurements from OGP data export xls file, delete tables (? Or delete entire workbook)

3. Parsing script [Complete]
    3.1 Continuous loop to check for QC.STA file in the working directory [Complete]
    3.2 Interpret the plain text in QC.STA (measurement data, header information for the shot) [Complete]
        3.2.1 str.split(sep=None, maxsplit=- 1)
            3.2.1.1 Call this on the text file to delimit text blocks by "END!". This text block is considered one set of measurements on a given part for each dimension
            3.2.1.2 Split the text block by carriage return/newline. Each row is distinct and pipe-delimited "|". A match statement can be used to process each row accordingly
            3.2.1.3 Brief row definitions: "Name" can be used to define individual tables similar to how QC-Calc establishes individual *.qcc files. "Date" and "Time" can be combined to provide a timestamp of the measurement record. "DATA" rows provide either shot identifying information, or measurement results. The second field in "DATA" rows indicate which one it is. "DATA|FACTOR" rows are header rows, while the other "DATA|DimensionName" rows are measurement records.
    3.3 Convert the text to equivalent 2D array "measurement records" similar to QC-Calc [Complete]

4. Graph [In Process] (2.0 feature?)
    4.1 Replicate the graphing display of QC-Calc [In Process]
        4.1.1 Leverage Matplotlib [In Process]
        4.1.2 Allow user to adjust the # of consecutive measurements to render (default = 10, up to 96?)
    4.2 Render measurements as in or out of spec
        4.2.1 Fetch specs from DaedriVictus based on product code
        4.2.2 Refresh the graphs upon ingest of another measurement record 
            4.2.2.1 This should be fairly easily leveraged with matplotlibs built in event handler.


        
Notes:
    -Difficulty parsing QC.STA into Measurement Values table format, then from Measurement Values table format to 2D grid. 
    Consider translating the existing QC.STA file output directly into a 2D grid (QC-Calc's format) [Rejected]
    
    -If measuring consecutive parts and the DATA|FACTOR rows aren't present, use a query to find the most recent set of 
    data factors and use those for each consecutive part measurement [In Process]

    - Tkinter's .mainloop() function is blocking, multi threading is highly discouraged due to the nature of Tkinter's event loop. Developers advise others to explore other options such as:
        - bypassing tkinter's .mainloop() entirely in favor of .update() and update_idletasks() which are manual function calls to push the event loop forward
            - .update() and update_idletasks() might make the GUI appear to be frozen and unresponsive in between these function calls, as GUI events will pile up in the event queue
        - exiting the .mainloop() event loop with .quit() to call other functions, and then calling .mainloop() again to re-enter the event loop
            - same issue here as with .update*() function calls. This might be negligible based on the scope of this project. This may also require extensive testing/error handling to ensure the program always returns to the event loop no matter what.
        - calling tkinter's built-in .after() function, which schedules a function to be called within the event loop, this allows tkinter to handle background tasks within a single threaded application by scheduling a function to be called after a timed delay

"""

import pandas as pd
import tkinter as tk
import tkinter.ttk as ttk
import os
import openpyxl #utilize openpyxl to reorgnize columns for 2-part programs
#import matplotlib.pyplot as plt    #not implemented yet
#import numpy as np                 #not implemented yet

###
#   Global variables 
###

dfList = dict()
excFileLocation = "\\\\beowulf.mold-rite.local\\spc\\ogptest.xls"

###
#   Functions
###

def submitshots():
    global dfList

    for nameOfDF, dfObject in dfList.items():
        print(len(dfObject.columns))
        numberOfHeads = 11
        numberOfMeasurements = len(dfObject.columns)
        reorganizedDataFrame = pd.DataFrame([])
        #dir = "C:\\Users\\tmartinez\\Documents\\"
        #filename = str(str(int(dfObject.at[0, 'Work Order'])) + ' ' + str(dfObject.at[0,'Product Code']) + ' ' + str(len(dfObject)) + 'cav ' + str(int(dfObject.at[0,'MOLD Number'])) + '.csv')
        #dfObject.to_csv(str(dir + filename), header = False, index = False)

    #all dataframes have been processed (saved to B drive). Wipe dfList and start fresh.
    dfList = dict()
    return

def truncateDataFrames():
    #hard reset on any pending shots that are measured. this clears any data frames left in memory in case analyst decides we simply need to discard and start fresh measuring any shots
    global dfList
    dfList = dict()
    return

excFileLocation = "\\\\beowulf.mold-rite.local\\spc\\ogptest.xls"

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

'''
A note on workflow for myself
to properly function, first the "sheet one" or main shot must be read into the script to gather the critical info from the last line
In this case, the first call is to grabdata() with the parameters of the sheet in either position 1 or 0, depending on if the OGP export requires sheet one to remain.
From there we can now query both the sql database of part #'s to determine if it is a two-part shot and which type, as well as query powershop using the work order number 
to retrieve header info and ensure that all of the data matches up, then store the necessary info for a filename at the end of the process. 
(NOTE1, this step should also verify that the part # in fact exists as well and should it not, generate a dialogue box to gather the correct part #)
Should there be a two part shot then grabdata() will need to be executed a second time with the sheet positioned after the first shot. After we have both shots attached to variables
as dataframes we will then need to parse both through formatQCtoDF(). Once they are formatted, there will need to be a logic set(if statement or for loop?)
to assign the correct two part function and execute this. 
After this step, in the case of both single and two part shots, we will need to take the previous powershop data, and use it to confirm all entered data is correct and for a filename
Once that is completed the file will need to be moved to the appropriate folder. NOTE2, perhaps we should leverage an event listener at this point, to verify that the file has been placed into the backup folder. 
once we confirm it has been backed up, then we can finish the program by deleting the sheets of the ogptest.xls file.
'''

def main(excFileLocation):
    
    return
###
#   GUI
###

#to-do: move all GUI initialization to class system, these don't belong in the global scope


mainGUI = tk.Tk()
mainGUI.title("OGP Interface")
for num in range(1, 5): [mainGUI.columnconfigure(num, minsize = 15), mainGUI.rowconfigure(num, minsize = 15)]

#redefine the mainGUI grid layout. 
tk.Frame(mainGUI).grid(column = 1, row = 1)
tk.Frame(mainGUI).grid(column = 4, row = 1)
tk.Frame(mainGUI).grid(column = 1, row = 4)
tk.Frame(mainGUI).grid(column = 4, row = 4)

tk.Button(mainGUI, text = "Submit Shot", command = submitshots).grid(column = 2, row = 2)
tk.Button(mainGUI, text = "Start Over", command = truncateDataFrames).grid(column = 3, row = 2)

#tkinter's *.mainloop() function fires off a blocking event loop. use tkinter's .after() method to schedule a function call with tkinter's event loop.
main(excFileLocation)
mainGUI.mainloop()