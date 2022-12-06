_author_ = 'arichland'

import pandas as pd
import numpy as np
import pprint
import sql
import xlwings as xw
pp = pprint.PrettyPrinter(indent=1)
def skus(self):
    print("SKUs")

    # Select active sheet and fetch data
    ws = self.sheet[10]
    query = "SELECT concat(sku_id, location) as concat, sku_id, name, description, location, make_buy FROM tbl_skus;"
    data = self.sql.fetch(query)

    # Load to Pandas dataframe
    df = pd.DataFrame(data[1], columns=data[0])
    df = df.set_index(data[0][0])

    # Write dataframe to workbook
    ws["A3"].options(pd.DataFrame, header=1, index=True, expand='table').value = df
    self.wb.save(self.filename)