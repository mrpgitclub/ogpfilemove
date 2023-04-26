import pyodbc as pydb

server = '10.100.1.96'
db = 'Plattsburgh_Owner'
username = 'tmartinez'
pwd = 'tenoneninenine1!'

cnxn = pydb.connect('DRIVER={ODBC Driver 18 for SQL Server};SERVER='+server+';DATABASE='+db+';ENCRYPT=yes;UID='+username+';PWD='+ pwd)
cursor = cnxn.cursor()
