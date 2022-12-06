_author_ = 'arichland'

import pandas as pd
import numpy as np
import pprint
import query
import xlwings as xw
pp = pprint.PrettyPrinter(indent=1)

class InvenCurrent:
    def __init__(self, **kargs):
        self.filename = kargs.get('filename')
        #self.wb = xw.Book(self.filename)
        #self.sheet = xw.sheets
        self.qry = query.query(qry='current inventory')
        self.data = self.qry['current inventory']
        self.excel = kargs.get('excel')
        self.calc = kargs.get('calc')

    def arguments(self):
        data = []
        if self.calc == True:
            data.append(self.calculations())
        else:
            pass

        if self.excel == True:
            self.write_to_excel(data[0])
            return data[0]
        else:
            return data[0]

    def calculations(self):
        df = pd.DataFrame(self.data, columns=['inven id',
                                              'location',
                                              'dates',
                                              'qty',
                                              'unit cost',
                                              'unit price',
                                              'total cogs',
                                              'total val'])
        #df.fillna(0)
        df = pd.pivot_table(df, values='qty', index=['inven id', 'location', 'unit cost', 'unit price'], fill_value=0)
        return df

    def write_to_excel(self, data):
        ws = self.sheet[3]
        ws["A3"].options(pd.DataFrame, header=1, index=True, expand='table').value = data
        self.wb.save(self.filename)

def current(**kargs):
    go = InvenCurrent(**kargs)
    return go.arguments()

if __name__ == '__main__':
    print(current(calc=True))

