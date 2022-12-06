_author_ = 'arichland'

import pandas as pd
import numpy as np
import pprint
import query
import xlwings as xw
pp = pprint.PrettyPrinter(indent=1)

class AnalyzeForecasts:
    def __init__(self, **kargs):
        self.filename = kargs.get('filename')
        #self.sheet = xw.sheets
        self.qry = query.query(qry='forecasts')
        self.data = self.qry['forecasts']
        self.excel = kargs.get('excel')
        self.calc = kargs.get('calc')

    def arguments(self):
        data = []
        if self.calc == True:
            return self.calculations()
        else:
            pass

        if self.excel == True:
            self.write_to_excel(data[0])
            return data[0]
        else:
            return data[0]

    def calculations(self):
        # ws = self.sheet[5]  # active sheet
        results = {}
        base_df = pd.DataFrame(self.data, columns=['inven_id',
                                              'dates',
                                              'location',
                                              'sales',
                                              'forecast',
                                              'fc_error',
                                              'error_sqrd',
                                              'mape'])

        index = ['inven_id', 'location']
        dframes = ['Sales', 'Forecasts', 'Forecast Error', 'Forecast Error Squared', 'MAPE']
        frame_format = ['#,##0', '#,##0', '#,##0', '#,##0', '0%']

        def sales_dataframe(base_df):
            return pd.pivot_table(base_df,
                                      index=index,
                                      columns=['dates'],
                                      values='sales',
                                      aggfunc=np.sum,
                                      fill_value=0).round()
        sales_df = sales_dataframe(base_df)
        num_periods = len(sales_df.columns)
        headers = sales_df.columns

        def forecast_dataframe(base_df):
            return pd.pivot_table(base_df,
                                   index=index,
                                   columns=['dates'],
                                   values='forecast',
                                   aggfunc=np.sum,
                                   fill_value=0).round()
        fc_df = forecast_dataframe(base_df)

        def error_dataframe(base_df):
            return pd.pivot_table(base_df,
                                      index=index,
                                      columns=['dates'],
                                      values='fc_error',
                                      aggfunc=np.sum,
                                      fill_value=0).round()
        error_df = error_dataframe(base_df)

        def bias_dataframe(base_df):
            df1 = pd.pivot_table(base_df,
                                     index=index,
                                     values='fc_error',
                                     aggfunc=np.sum,
                                     fill_value=0)

            df1['Avg Bias'] = (df1['fc_error'] / num_periods).round()

            return df1
        bias_df = bias_dataframe(base_df)

        def error_sqrt_dataframe(base_df):
            return pd.pivot_table(base_df,
                                     index=index,
                                     columns=['dates'],
                                     values='error_sqrd',
                                     aggfunc=np.sum,
                                     fill_value=0).round()
        sqrd_df = error_sqrt_dataframe(base_df)
        sqrd_col = sqrd_df.iloc[:, 0:len(sqrd_df.columns)]

        def mean_sqrd_error_dataframe(base_df):
            rmse_df = pd.pivot_table(base_df,
                                     index=index,
                                     values='error_sqrd',
                                     aggfunc=np.sum,
                                     fill_value=0).round()

            rmse_df['RMSE'] = np.sqrt(rmse_df['error_sqrd'] / num_periods).round()
            # rmse_df['RMSE Active'] = math.sqrt(rmse_df['error_sqrd'] / num_periods)
            rmse_df['RMSE Active'] = np.sqrt(rmse_df['error_sqrd'] / sqrd_col.astype(bool).sum(axis=1)).round()
            return rmse_df
        rmse_df = mean_sqrd_error_dataframe(base_df)

        def mape_dataframe(base_df):
            df1 = pd.pivot_table(base_df,
                                 index=index,
                                 columns=['dates'],
                                 values='mape',
                                 aggfunc=np.sum,
                                 fill_value=0).round()

            df2 = pd.pivot_table(base_df,
                                 index=index,
                                 values='mape',
                                 aggfunc=np.sum,
                                 fill_value=0).round()
            df2['Avg MAPE'] = (df2['mape'] / num_periods).round()

            df1 = df1.merge(df2,
                            how='inner',
                            right_index=True,
                            left_index=True)
            return df1
        mape_df = mape_dataframe(base_df)

        # Merge dataframes
        base_df = sales_df.merge(fc_df, how='inner', right_index=True, left_index=True)
        base_df = base_df.merge(error_df, how='inner', right_index=True, left_index=True)
        base_df = base_df.merge(sqrd_df, how='inner', right_index=True, left_index=True)
        base_df = base_df.merge(mape_df, how='inner', right_index=True, left_index=True)

        results.update({'results': base_df})
        results.update({'dataframes': {
                       'sales': sales_df,
                       'forecast': fc_df,
                       'error': error_df,
                       'bias': bias_df,
                       'error squared': sqrd_df,
                       'mean squared error': rmse_df,
                       'mape': mape_df}})
        return results

        # ws["A3"].options(pd.DataFrame, header=1, index=True, expand='table').value = df
    def write_to_excel(self, data):
        # Format headers for each dataframe
        dframes = ['Sales', 'Forecasts', 'Forecast Error', 'Forecast Error Squared', 'MAPE']
        frame_format = ['#,##0', '#,##0', '#,##0', '#,##0', '0%']
        col = 4
        count = 0
        """
        for i in dframes:
            count += 1
            format = frame_format[count - 1]
            if count == 1:
                col1 = col + num_periods - 1
                frame_num = count - 1
                ws.range(2, col).value = dframes[frame_num]
                ws.range((3, col), (3, col1)).value = headers
                ws.range((4, col), (len(df.index), col1)).number_format = format
            else:
                col = col + num_periods
                col2 = col + num_periods
                frame_num = count - 1
                ws.range(2, col).value = dframes[frame_num]
                ws.range((3, col), (3, col2)).value = headers
                ws.range((4, col), (len(df.index), col2)).number_format = format

        # ws.range(3, len(df.columns)+3).value = "Sum of Error%"

        # Write dataframe to workbook
        self.wb.save(self.filename)"""

def forecasts(**kargs):
    go = AnalyzeForecasts(**kargs)
    return go.arguments()

if __name__ == '__main__':
    forecasts(calc=True)