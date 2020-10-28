_author_ = 'arichland'

import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
import pymysql.cursors
from pydict import sql_dict
import pandas as pd
import numpy as np


# SQL Fields
user = sql_dict.get('user')
password = sql_dict.get('password')
host = sql_dict.get('host')
database = sql_dict.get('database')
charset = sql_dict.get('charset')
con = pymysql.connect(user=user,
                       password=password,
                       host=host,
                       database=database,
                       charset=charset)

# Workbook Fields
filename = "Inventory Analysis.xlsx"
wb = openpyxl.load_workbook(filename=filename)
wb.active = wb.sheetnames.index("Analysis")
ws = wb.active
df_rows = dataframe_to_rows



# Get monthly inventory reports
try:
    with con.cursor() as cur:
        qry_rpt_inven = "SELECT sku, location, abc_class, dates, qty FROM pyInven_Report;"

        # SQL query to pandas dataframe
        read_df = pd.read_sql(qry_rpt_inven, con)
finally:
        con.commit()
        cur.close()
        con.close()


df = pd.DataFrame(read_df)
#print(df)

pivotTableDF = df.filter(items=['sku', 'location', 'qty'])
pivotTableDF = pd.pivot_table(df,
                              index='sku',
                              columns='dates',
                              values='qty',
                              aggfunc=np.sum,
                              margins=True)
pivotTableDF.sort_values(by=['All'],
                         inplace=True,
                         ascending=False)
print(pivotTableDF)
# Save to workbook
wb.save(filename=filename)