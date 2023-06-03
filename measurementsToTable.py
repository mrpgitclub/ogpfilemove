import pandas as pd


myDataframe = pd.read_excel("Measurement Values History.xlsx", sheet_name = "Sheet1")
myDataframe.columns = [column.replace(" ", "_") for column in myDataframe.columns]
woSet = myDataframe['Work_Order_Number'].unique()
print(woSet)

def createframes(df,set):
    dfList = []
    for x in set:
        df.query('Work_Order_Number= @x')
        dfList.append(exec(f'{x} = df.copy()'))
    return dfList

list = createframes(myDataframe,woSet)
print(list)
