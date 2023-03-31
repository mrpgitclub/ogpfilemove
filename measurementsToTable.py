import pandas as pd

myDataframe = pd.read_excel("Measurement Values History.xlsx", sheet_name = "Sheet1")

#how to access something from a dataframe
#row['Work Order Number']

#dataframe names are (for example) 'w2509360'

listofdataframes = []
for index, row in myDataframe.iterrows():
    if str('w' + row['Work Order Number']) in listofdataframes:
        currentDataFrame = listofdataframes[str('w' + row['Work Order Number'])]
    else: 
        workorder = pd.DataFrame()
        globals()[workorder] = str('w' + row['Work Order Number'])
        listofdataframes.append(workorder)
    
    iD =  row['ID']
    cavity =  row['Head No']
    measurement =  row['Value']
    dimension =  row['Variable Type']

    currentDataframe = 

    break