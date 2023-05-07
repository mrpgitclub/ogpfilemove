import pandas as pd
import os
import openpyxl
import sqlite3
from sqlite3 import connect

conn = sqlite3.connect('Part_Numbers2.db')
c = conn.cursor()


acquire = pd.read_csv('Acquisitions2.csv', dtype=str)

acquire.to_sql('Part_Numbers2',con = conn,dtype='str')
