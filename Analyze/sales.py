_author_ = 'arichland'

import pandas as pd
import numpy as np
import pprint
import query
import xlwings as xw
pp = pprint.PrettyPrinter(indent=1)

class SalesHistory:
    def __init__(self, **kargs):
        self.filename = kargs.get('filename')
        self.wb = xw.Book(self.filename)
        self.sheet = xw.sheets
        self.qry = query.query(qry='sales')
        self.data = self.qry['sales']
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
        headers = []
        # Create base dataframe
        df1 = pd.DataFrame(self.data, columns=['inven id', 'location', 'dates', 'sales'])
        df1 = pd.pivot_table(df1,
                            index=['inven id', 'location'],
                            columns=['dates'],
                            values='sales',
                            aggfunc=np.sum,
                            fill_value=0)

        df2 = pd.DataFrame(self.data, columns=['inven id', 'location', 'dates', 'sales'])
        df2 = pd.pivot_table(df2,
                             index=['inven id', 'location'],
                             columns=['dates'],
                             values='sales',
                             aggfunc=np.sum,
                             fill_value=0)

        df3 = pd.DataFrame(self.data, columns=['inven id', 'location', 'dates', 'sales'])
        df3 = pd.pivot_table(df3,
                             index=['inven id', 'location'],
                             columns=['dates'],
                             values='sales',
                             aggfunc=np.sum,
                             fill_value=0)
        num_cols = len(df1.columns)
        col = df1.iloc[:, 0:num_cols]

        # Primary Calculations
        df1["Total Demand"] = col.sum(axis=1)
        df1["Active Months"] = col.astype(bool).sum(axis=1)
        df1['Average'] = col.mean(axis=1).round()
        df1['Min'] = col.min(axis=1)
        df1['Max'] = col.max(axis=1)
        df1["Avg Demand when >0"] = ((col.sum(axis=1)) / (col.astype(bool).sum(axis=1))).round()
        df1["Min Demand when >0"] = col[df1.gt(0)].min(1).round()

        # Calc Standard Deviation
        df1["StDev"] = col.std(axis=1).round()
        col_num_stdev1 = len(df1.columns) + 4

        df2 = df2.groupby(pd.PeriodIndex(df2.columns, freq='Q'), axis=1).std().round()
        pd.set_option("display.max_rows", None, "display.max_columns", None)

        for i in df2.columns: headers.append(str(i))
        df = df1.merge(df2, how='inner', right_index=True, left_index=True)
        end_stdev_qtrs = len(df.columns) + len(headers)

        # Calc Standard Deviation for months w/ demand >0
        df3 = df3.replace(0, np.NaN)
        df['Stdev >0'] = col[df.gt(0)].std(1).round()
        col_num_stdev2 = len(df.columns) + 4
        end_col = col_num_stdev2 + len(headers)
        df3 = df3.groupby(pd.PeriodIndex(df3.columns, freq='Q'), axis=1).std().round()
        df = df.merge(df3, how='inner', right_index=True, left_index=True)
        return df

    def write_to_excel(self, data):
        ws = self.sheet[4]
        # Write dataframe to workbook & reformat headers
        ws["A3"].options(pd.DataFrame, header=1, index=True, expand='table').value = data
        #ws.range((3, col_num_stdev1), (3, end_stdev_qtrs)).value = headers
        #ws.range((3, col_num_stdev2), (3, end_col)).value = headers
        self.wb.save(self.filename)

def sales(**kargs):
    go = SalesHistory(**kargs)
    go.arguments()

if __name__ == '__main__':
    sales(calc=True)