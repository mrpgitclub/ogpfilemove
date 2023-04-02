"""
"OGP Interface" Application Project overview

Overview of feature requirements
1. GUI
2. SQL Database
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

2. SQL Database
    2.1 Collect individual measurements and group them as a single measurement set [Complete]
    2.2 Guarantees users cannot modify data, only omit or rearrange measurements prior to upload to SFOL [Complete]
    2.3 Can either run an in-memory database or file-based database
        2.3.1 Pro: in-memory database is much faster [Rejected]
        2.3.2 Con: in-memory database won't persist if the tool is restarted, or the computer is restarted [Rejected]
        2.3.3 Pro: file-based database is ACID, and will persist across restarts, crashes, etc [In Process]
        2.3.4 Con: file-based database is much slower, but that might be negligible in this application [In Process] --likely wont be a factor to use in file in this scenario-- At least at first, -AR 
    2.4 Clear out measurement records after submission to SFOL
        2.4.1 Optionally send to a file-based archive DB to retain a history for tracking purposes
    2.5 SQL Statement definitions
        2.5.1 CREATE TABLE #routineName# (dim1 text, dim2 text, ..., "MOLD Number" text, "Work Order" text, "Operator" text, "Machine" text, "Color" text, "Resin Formula" text, "Color Code" text, "Product Code" text)  [Complete]
        2.5.2 INSERT INTO #routineName# (dim1, dim2, ..., "MOLD Number", "Work Order", "Operator", "Machine", "Color", "Resin Formula", "Color Code", "Product Code") values(msmntValue1, msmntValue2, ..., header1, header2, ...) [Complete]
        2.5.3 SELECT [In Process]

3. Parsing script [Complete]
    3.1 Continuous loop to check for QC.STA file in the working directory [Complete]
    3.2 Interpret the plain text in QC.STA (measurement data, header information for the shot) [Complete]
        3.2.1 str.split(sep=None, maxsplit=- 1)
            3.2.1.1 Call this on the text file to delimit text blocks by "END!". This text block is considered one set of measurements on a given part for each dimension
            3.2.1.2 Split the text block by carriage return/newline. Each row is distinct and pipe-delimited "|". A match statement can be used to process each row accordingly
            3.2.1.3 Brief row definitions: "Name" can be used to define individual tables similar to how QC-Calc establishes individual *.qcc files. "Date" and "Time" can be combined to provide a timestamp of the measurement record. "DATA" rows provide either shot identifying information, or measurement results. The second field in "DATA" rows indicate which one it is. "DATA|FACTOR" rows are header rows, while the other "DATA|DimensionName" rows are measurement records.
    3.3 Convert the text to equivalent 2D array "measurement records" similar to QC-Calc [Complete]

4. Graph [In Process]
    4.1 Replicate the graphing display of QC-Calc [In Process]
        4.1.1 Leverage Matplotlib [In Process]
        4.1.2 Allow user to adjust the # of consecutive measurements to render (default = 10, up to 96?)
    4.2 Render measurements as in or out of spec
        4.2.1 Fetch specs from DaedriVictus based on product code
        4.2.2 Refresh the graphs upon ingest of another measurement record 
            4.2.2.1 This should be fairly easily leveraged with matplotlibs built in event handler.
            4.2.2.1  

        
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

import sqlite3
import tkinter as tk
import tkinter.ttk as ttk
import os
#import matplotlib.pyplot as plt    #not implemented yet
#import numpy as np                 #not implemented yet
#from time import sleep             #unused, replaced sleep with tkinter's .after() method

###
#   Global variables 
###

qcFile = 'QC.STA' #need to confirm that this will search within the program's current directory for QC.STA
dbFile = 'TestDB.db'

###
#   Functions
###

def loadOGPRoutines(treeTables, CUR):
    print("Entering loadOGPRoutines")
    list = CUR.execute("select name from sqlite_master where type = 'table';").fetchall()
    print(list)
    treeTables.insert(parent = '', index = 'end', values = (list))


    return

def redrawMeasurementsTable(mTable):
    #add a large number of rows to test scrollbar
    for colName in ("OD", "T", "E"): 
        mTable.heading(colName, text = colName)

    for rowNum in range(1, 50): 
        treeTables.insert(parent = '', index = 'end', values = (rowNum))
        mTable.insert(parent = '', index = 'end', values = (rowNum))


def parseAndIngest(qcFile, CON, CUR):
    #open qc.sta, if for some reason it can't open it then return immediately
    try: QCFobject = open(qcFile, mode = 'r')
    except OSError: return
    #Upon successfully opening and reading QC.STA, delete the file
    contents = QCFobject.read()
    QCFobject.close()
    #os.remove(qcFile)  #omitted during testing, uncomment for official release

    #split text blocks by "END!". This text block is considered one set of measurements on a given part for each dimension. If there aren't any blocks to process, return immediately. May need to refine this error processing to ensure that 'len(textblocks) < 1' is appropriate
    textblocks = contents.split("END!")
    if len(textblocks) < 1: return

    #split text blocks by row
    for currentblock in textblocks:
        rows = currentblock.splitlines()
        if len(rows) < 1: return

        #define headers early, in order to determine if this text block provides the headers or not. If not, at the end of the text block, we will fetch the previously used headers from the DB and use them here, making the assumption that these are consecutive parts being measured in the same routine.
        headers = {"Cavity":        {"Position":1, "Value": None},
                    "MOLD Number":  {"Position":2, "Value": None},
                    "Work Order":   {"Position":3, "Value": None},
                    "Operator":     {"Position":4, "Value": None},
                    "Machine":      {"Position":5, "Value": None},
                    "Color":        {"Position":6, "Value": None},
                    "Resin Formula": {"Position":7, "Value": None},
                    "Color Code":   {"Position":8, "Value": None},
                    "Product Code": {"Position":9, "Value": None}}
        measurements = {}
        position = 1
        #Establish accumulation strings, to be used for error checking in text blocks. Assist in detecting incomplete measurement records
        createcolumnnames = str('')
        insertcolumnnames = str('')
        insertvalues = str('')

        #split row into individual fields, delimited by "|", the order of rows in QC.STA dictates the order which SFOL will receive them. This is an area for improvement down the road but for now, we will emulate this behavior between OGP -> QC-Calc 
        for currentrow in rows:
            fields = currentrow.split("|")
            if len(fields) < 1: continue

            #parse each row, attempt to validate OGP output and convert data to the data format that SFOL would expect to receive. Remove trailing and leading 0's, common numeral notation
            match fields[0]:
                case "NAME": tablename = fields[1]
                case "DATE": datetime = fields[1].replace(":", "-")
                case "TIME": datetime += ' ' + fields[1] + '.000'
                case "DATA":
                    match fields[1]:
                        case "FACTOR": 
                            if fields[4] == 'Cavity': headers[fields[4]]["Value"] = int(fields[9].lstrip('+'))
                            else: headers[fields[4].lstrip('+')]["Value"] = str(fields[9].lstrip('+'))
                        case _: 
                            measurements[fields[1]] = {"Position": position, "Value": fields[6].lstrip("+")}
                            position += 1

        #we are done processing the text block and are ready to begin assembling SQL statements to send to the DB
        #the headers and measurements dictionaries contain all the information we need

        #Assembling SQL statements
        for KEY, VAL in measurements.items():
            createcolumnnames = str(createcolumnnames + "\'" + KEY + "\' text, ")
            insertvalues = str(insertvalues + "\'" + VAL["Value"] + "\', ")
        insertcolumnnames = createcolumnnames.replace('text', '')

        #Assembling SQl statement. Detect if headers are provided. Headers are only provided in the beginning of a shot (cavity 1). Subsequent measurements will be missing these headers. 
        #if headers found in the current block, then assign as normal. If NOT found, fetch from tablename later after the table has been created
        #convert this to a list comprehension. Count the occur
        headersDetected = False
        for KEY, VAL in headers.items():
            if VAL["Value"] is not None:
                headersDetected = True
                break
        if headersDetected is True:
            for KEY, VAL in headers.items(): insertvalues += str("\'" + VAL["Value"] + "\', ")

        #Assembling SQL statements
        insertcolumnnames += " 'Cavity', 'MOLD Number', 'Work Order', 'Operator', 'Machine', 'Color', 'Resin Formula', 'Color Code', 'Product Code', 'Datetime'"

        #Assembling the individual parts of SQL statements into the {tablename} body of each statement.
        #createSQL will always query, which will attempt to create the table if it doesn't exist. This will fall through if it already exists.
        createSQL = f'CREATE TABLE IF NOT EXISTS \"{tablename}\" ({createcolumnnames}"Cavity" text, "MOLD Number" text, "Work Order" text, "Operator" text, "Machine" text, "Color" text, "Resin Formula" text, "Color Code" text, "Product Code" text, "Datetime" text)'

        #create the table ONLY if it doesn't exist
        CUR.execute(createSQL)

        #attempt to fetch last known headers from tablename, else if this is a brand new table AND there are no headers in the table, or in the text block, then finally fall through and force headers to be empty placeholder values
        if headersDetected is False:
            countofRows = CUR.execute(f"SELECT Count(*) FROM '064089RSPS 4 CAV 890409.RTN'").fetchone()
            if countofRows[0] < 1: results = ['0', '0', '0', '0', '0', '0', '0', '0', '0']
            else: results = CUR.execute(f"SELECT COALESCE('{tablename}'.'Cavity' + 1, 'Unknown'), COALESCE('{tablename}'.'MOLD Number', 'Unknown'), COALESCE('{tablename}'.'Work Order', 'Unknown'), COALESCE('{tablename}'.'Operator', 'Unknown'), COALESCE('{tablename}'.'Machine', 'Unknown'), COALESCE('{tablename}'.'Color', 'Unknown'), COALESCE('{tablename}'.'Resin Formula',  'Unknown'), COALESCE('{tablename}'.'Color Code', 'Unknown'), COALESCE('{tablename}'.'Product Code', 'Unknown') FROM '{tablename}' ORDER BY '{tablename}'.'Datetime' DESC LIMIT 1").fetchone()

            for VAL in results: insertvalues += str("\'" + str(VAL) + "\', ")
        
        insertvalues += str("\'" + datetime + "\'")
        #insertSQL will always query, and attempt to insert measurement values into the named table. This should allow for partial measurements to be taken, which shouldn't traditionally happen in a real-world production shot but will be useful for testing, step edits, and developing new OGP routines
        insertSQL = f'INSERT INTO \"{tablename}\" ({insertcolumnnames}) values({insertvalues})'

        CUR.execute(insertSQL)
        CON.commit()
        #break after first text block. remove to allow for the full QC.STA file to run, remove this after a GUI is implemented in order to observe proper workflow
        #break

    return

#Clear the currently select program and refresh it. This is triggered when parseAndIngest has been called and new measurements were found
def regenerateTablesTreeView(tree, list):

    pass

# not implemented yet. used for rendering measurements in a graph
# def renderGraphs(db):
#     plt.style.use('_mpl-gallery')
# 
#     # make data
#     x = np.linspace(0, 10, 100)
#     y = 4 + 2 * np.sin(2 * x)
#     countOfDimensions = 9
#     # plot
#     fig, ax = plt.subplots(nrows = 2, ncols = countOfDimensions)
# 
#     for row in ax:
#         for col in row:
#             col.plot(x, y)
#     plt.show()

def main():
    try: CON = sqlite3.connect(dbFile)
    except: return
    if CON:
        CUR = CON.cursor()
        #parseAndIngest(qcFile, CON, CUR)
        print('Ready to enter loadOGProutines')
        loadOGPRoutines(treeTables, CUR)
        CUR.close()
    CON.close()
    rootTk.after(1000, main)

###
#   GUI
###

#to-do: move all GUI initialization to class system, these don't belong in the global scope
rootTk = tk.Tk()
mainGUI = tk.Toplevel(master = rootTk)
mainGUI.title("OGP Interface")

#default sizing options for the entire grid
#move these config calls to a redrawing method to be called during parseAndIngest. This is to allow for dynamic sizing of the Measurements table based on variable # of dimensions
for colNum in range(1,9): mainGUI.columnconfigure(colNum, minsize = 6)
for rowNum in range(1,15): mainGUI.rowconfigure(rowNum, minsize = 10)

#redefine the mainGUI grid layout. 
tk.Frame(mainGUI).grid(column = 1, row = 1)
tk.Frame(mainGUI).grid(column = 7, row = 1)
tk.Frame(mainGUI).grid(column = 1, row = 14)
tk.Frame(mainGUI).grid(column = 7, row = 14)

mainGUI.deiconify()

treeTables = ttk.Treeview(mainGUI, show = "headings", selectmode = "browse", height = 30, columns = ("OGP Routine"), displaycolumns=(1))
treeTables.heading(1, text = "OGP Routine")
treeTables.grid(column= 1, row = 2, rowspan = 11, sticky=tk.EW)

treeTablesScrollBar = ttk.Scrollbar(mainGUI, command = treeTables.yview, orient = tk.VERTICAL)
treeTablesScrollBar.grid(column=2, row =2, rowspan = 11, sticky = tk.NS)
treeTables['yscrollcommand'] = treeTablesScrollBar.set

#test scroll bar functionality by adding random columns and random numbers into the rows
treeMeasurements = ttk.Treeview(mainGUI, show = "headings", selectmode = "extended", height = 30, columns = ("measurements"))
treeMeasurements.heading(0, text = "Measurements")
treeMeasurements.grid(column = 3, row = 2, columnspan = 3, rowspan = 11, sticky=tk.EW)
treeMeasurementsVerticalScrollBar = ttk.Scrollbar(mainGUI, command = treeMeasurements.yview, orient = tk.VERTICAL)
treeMeasurementsVerticalScrollBar.grid(column = 7, row = 2, rowspan = 11, sticky = tk.NS)
treeMeasurementsHorizontalScrollBar = ttk.Scrollbar(mainGUI, command = treeMeasurements.xview, orient = tk.HORIZONTAL)
treeMeasurementsHorizontalScrollBar.grid(column = 3, row = 13, columnspan = 3, sticky = tk.EW)
treeMeasurements['yscrollcommand'] = treeMeasurementsVerticalScrollBar.set
treeMeasurements['xscrollcommand'] = treeMeasurementsHorizontalScrollBar.set

tk.Button(mainGUI, name = "submit", text = 'Submit').grid(column =3, row =14, sticky = tk.NSEW)
tk.Button(mainGUI, name = "delete", text = 'Delete').grid(column =4, row =14, sticky = tk.NSEW)
tk.Button(mainGUI, name = "rename", text = 'Rename').grid(column =5, row =14, sticky = tk.NSEW)
tk.Button(mainGUI, name = "graphs", text = 'Graphs').grid(column =6, row =14, sticky = tk.NSEW)

#tkinter's *.mainloop() function fires off a blocking event loop. use tkinter's .after() method to schedule a function call with tkinter's event loop.
main()
mainGUI.mainloop()