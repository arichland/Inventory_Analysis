_author_ = 'arichland'

import pandas as pd
import numpy as np
import pprint
import query
import xlwings as xw
pp = pprint.PrettyPrinter(indent=1)

class InvenHistory:
    def __init__(self, **kargs):
        self.filename = kargs.get('filename')
        #self.wb = xw.Book(self.filename)
        #self.sheet = xw.sheets
        self.qry = query.query(qry='inventory history')
        self.data = self.qry['inventory history']
        self.excel = kargs.get('excel')
        self.calc = kargs.get('calc')

    def arguments(self):
        data = []
        if self.calc == True:
            data.append(self.calculations())
        else: pass

        if self.excel == True:
            self.write_to_excel(data[0])
            return data[0]
        else:
            return data[0]

    def calculations(self):
        # Load data to Pandas dataframe
        df = pd.DataFrame(self.data, columns=['inven id', 'location', 'dates', 'qty'])

        #  Convert dataframe to pivot table
        df = pd.pivot_table(df,
                            values='qty',
                            index=['inven id', 'location'],
                            columns=['dates'],
                            aggfunc=np.sum,
                            fill_value=0)
        num_cols = len(df.columns)

        # Add calculations to dataframe
        df['Average'] = df.iloc[:, 0:num_cols].mean(axis=1)  # Avg inventory per SKU, per location
        df['Min'] = df.iloc[:, 0:num_cols].min(axis=1)  # Min inventory quantity per SKU, per location
        df['Max'] = df.iloc[:, 0:num_cols].max(axis=1)  # Max inventory quantity per SKU, per location
        df["Total"] = df.iloc[:, 0:num_cols].sum(axis=1)  # Sum of inventory per month, per SKU per location
        df["Months w/ Inven"] = df.iloc[:, 0:num_cols].astype(bool).sum(
            axis=1)  # Count of months with inventory quantities greater than 0

        df["Avg Inven when >0"] = (df.iloc[:, 0:num_cols].sum(axis=1)) / (
            df.iloc[:, 0:num_cols].astype(bool).sum(axis=1))  # Average inventory levels in months w/ inventory > 0

        df["Min Inven when >0"] = df.iloc[:, 0:num_cols][df.gt(0)].min(1)  # Min inventory level when inventory > 0
        return df

    def write_to_excel(self, data):
        ws = self.sheet[2]
        ws["A3"].options(pd.DataFrame,
                         header=1,
                         index=True,
                         expand='table').value = data
        self.wb.save(self.filename)

def history(**kargs):
    go = InvenHistory(**kargs)
    return go.arguments()

if __name__ == '__main__':
    history(calc=True)