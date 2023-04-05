import pandas as pd

myDataframe = pd.read_excel("Measurement Values History.xlsx", sheet_name = "Sheet1")

#how to access something from a dataframe
#row['Work Order Number']

#   series is a column, created by a list of elements 
# srs = [1, 2, 3]

#   dataframe is multiple series joined by index 
# data = {"column_header": [value1, value2, value3], "column_header2": [value1, value2, value3]}
# df = pd.DataFrame(data)

#dataframe names are (for example) 'w2509360'
# df['Merged'] = [{key: val} for key, val in zip(df.Stage_Name, df.Metrics)] to merge cols
listofdataframes = []
#myDataframe['Merged'] = [{key: val} for key, val in zip(myDataframe.Variable_Type, myDataframe.Value)]

myDataframe['merged'] = myDataframe.apply(lambda row: {row['Variable Type']:row['Value']}, axis=1)
print(myDataframe)


"""for index, row in myDataframe.iterrows():
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



    break"""