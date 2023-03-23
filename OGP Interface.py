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
    1.3 Tkinter
        1.3.1 Treeview widget to represent data in a grid
        1.3.2 Button widget to "Submit shot", "Delete individual measurement(s)"

2. SQL Database
    2.1 Collect individual measurements and group them as a single measurement set
    2.2 Guarantees users cannot modify data, only omit or rearrange measurements prior to upload to SFOL
    2.3 Can either run an in-memory database or file-based database
        2.3.1 Pro: in-memory database is much faster
        2.3.2 Con: in-memory database won't persist if the tool is restarted, or the computer is restarted
        2.3.3 Pro: file-based database is ACID, and will persist across restarts, crashes, etc
        2.3.4 Con: file-based database is much slower, but that might be negligible in this application
    2.4 Clear out measurement records after submission to SFOL
        2.4.1 Optionally send to a file-based archive DB to retain a history for tracking purposes
    2.5 SQL Statement definitions
        2.5.1 CREATE TABLE #routineName# (dim1 text, dim2 text, ..., "MOLD Number" text, "Work Order" text, "Operator" text, "Machine" text, "Color" text, "Resin Formula" text, "Color Code" text, "Product Code" text) 
        2.5.2 INSERT INTO #routineName# (dim1, dim2, ..., "MOLD Number", "Work Order", "Operator", "Machine", "Color", "Resin Formula", "Color Code", "Product Code") values(msmntValue1, msmntValue2, ..., header1, header2, ...)
        2.5.3 SELECT

3. Parsing script
    3.1 Continuous loop to check for QC.STA file in the working directory
    3.2 Interpret the plain text in QC.STA (measurement data, header information for the shot)
        3.2.1 str.split(sep=None, maxsplit=- 1)
            3.2.1.1 Call this on the text file to delimit text blocks by "END!". This text block is considered one set of measurements on a given part for each dimension
            3.2.1.2 Split the text block by carriage return/newline. Each row is distinct and pipe-delimited "|". A match statement can be used to process each row accordingly
            3.2.1.3 Brief row definitions: "Name" can be used to define individual tables similar to how QC-Calc establishes individual *.qcc files. "Date" and "Time" can be combined to provide a timestamp of the measurement record. "DATA" rows provide either shot identifying information, or measurement results. The second field in "DATA" rows indicate which one it is. "DATA|FACTOR" rows are header rows, while the other "DATA|DimensionName" rows are measurement records.
    3.3 Convert the text to equivalent 2D array "measurement records" similar to QC-Calc

4. Graph
    4.1 Replicate the graphing display of QC-Calc
        4.1.1 Leverage Matplotlib
        4.1.2 Allow user to adjust the # of consecutive measurements to render (default = 10, up to 96?)
    4.2 Render measurements as in or out of spec
        4.2.1 Fetch specs from DaedriVictus based on product code
        4.2.2 Refresh the graphs upon ingest of another measurement record

        
Notes:
    -Difficulty parsing QC.STA into Measurement Values table format, then from Measurement Values table format to 2D grid. 
    Consider translating the existing QC.STA file output directly into a 2D grid (QC-Calc's format)
    
    -If measuring consecutive parts and the DATA|FACTOR rows aren't present, use a query to find the most recent set of 
    data factors and use those for each consecutive part measurement

"""

import matplotlib.pyplot as plt
import numpy as np
import sqlite3
from tkinter import messagebox

###
#   Global variables 
###

qcFile = 'C:\\Users\\tmartinez\\Downloads\\QC.STA'
dbFile = ":memory:" #in-memory database during testing

###
#   Functions
###

def mainloop(qcFile):
    #open qc.sta, if for some reason it can't open it then return immediately
    try: QCFobject = open(qcFile, mode = 'r')
    except OSError: return

    #split text blocks by "END!". This text block is considered one set of measurements on a given part for each dimension. If there aren't any blocks to process, return immediately. May need to refine this error processing to ensure that 'len(textblocks) < 1' is appropriate
    textblocks = QCFobject.split("END!")
    if len(textblocks) < 1: return

    #split text blocks by row
    for currentblock in textblocks:
        rows = currentblock.split('\r')
        if len(rows) < 1: return

        #define headers early, in order to determine if this text block provides the headers or not. If not, at the end of the text block, we will fetch the previously used headers from the DB and use them here, making the assumption that these are consecutive parts being measured in the same routine.
        headers = {"MOLD Number": {"Position":1, "Value": None},
                   "Work Order": {"Position":2, "Value": None},
                   "Operator": {"Position":3, "Value": None},
                   "Machine": {"Position":4, "Value": None},
                   "Color": {"Position":5, "Value": None},
                   "Resin Formula": {"Position":6, "Value": None},
                   "Color Code": {"Position":7, "Value": None},
                   "Product Code": {"Position":8, "Value": None}}
        measurements = {}
        position = 1

        #split row into individual fields, delimited by "|", the order of rows in QC.STA dictates the order which SFOL will receive them. This is an area for improvement down the road but for now, we will emulate this behavior between OGP -> QC-Calc 
        for currentrow in rows:
            fields = currentrow.split("|")
            if len(fields) < 1: continue

            match fields[0]:
                case "NAME": tablename = fields[1]
                case "DATE": timestamp = fields[1].replace(":", "/")
                case "TIME": timestamp = timestamp + ' ' + fields[1]
                case "DATA":
                    match fields[1]:
                        case "FACTOR": headers[fields[4].lstrip("+")]["Value"] = fields[9]
                        case _: 
                            measurements[fields[1]] = {"Position": position, "Value": fields[6]}
                            position += 1

        #we are done processing the text block and are ready to begin assembling SQL statements to send to the DB
        #the headers and measurements dictionaries contain all the information we need


    return


def renderGraphs(db):
    plt.style.use('_mpl-gallery')

    # make data
    x = np.linspace(0, 10, 100)
    y = 4 + 2 * np.sin(2 * x)
    countOfDimensions = 9
    # plot
    fig, ax = plt.subplots(nrows = 2, ncols = countOfDimensions)

    for row in ax:
        for col in row:
            col.plot(x, y)
    plt.show()

