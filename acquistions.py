import pandas as pd
import os
import openpyxl
import sqlite3
from sqlite3 import connect

conn = sqlite3.connect('Part_Numbers.db')
c = conn.cursor()

"""acquire = pd.read_csv('Acquisitions.csv', dtype=str)

acquire.to_sql('Part_Numbers',con = conn,dtype='str')
"""