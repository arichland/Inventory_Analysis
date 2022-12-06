_author_ = 'arichland'

import pandas as pd
import numpy as np
import pprint
import sql
import xlwings as xw
pp = pprint.PrettyPrinter(indent=1)

class Calc:
    def __init__(self, filename):
        self.filename = filename
        self.wb = xw.Book(self.filename)
        self.sheet = xw.sheets
        self.sql = sql.Database()
        self.sql = sql.Database()

    def analysis(self):
        pass

    def inventory_history(self):
        print("Inventory History")

        # Select active sheet and fetch data
        ws = self.sheet[2]
        query = "SELECT concat(sku, location) as concat, sku, location, abc_class, dates, qty FROM tbl_inventory_history ORDER BY id ASC;"
        data = self.sql.fetch(query)

        # Load data to Pandas dataframe
        df = pd.DataFrame(data[1], columns=data[0])
        #df = df.set_index(data[0][0])
        df = pd.pivot_table(df, values='qty', index=['concat', 'sku', 'location'], columns=['dates'], aggfunc=np.sum, fill_value=0)
        num_cols = len(df.columns)

        # Calculations
        df['Average'] = df.iloc[:, 0:num_cols].mean(axis=1)  # Avg inventory per SKU, per location
        df['Min'] = df.iloc[:, 0:num_cols].min(axis=1)  # Min inventory quantity per SKU, per location
        df['Max'] = df.iloc[:, 0:num_cols].max(axis=1)  # Max inventory quantity per SKU, per location
        df["Total"] = df.iloc[:, 0:num_cols].sum(axis=1)  # Sum of inventory per month, per SKU per location
        df["Months w/ Inven"] = df.iloc[:, 0:num_cols].astype(bool).sum(axis=1)  # Count of months with inventory quantities greater than 0
        df["Avg Inven when >0"] = (df.iloc[:, 0:num_cols].sum(axis=1)) / (df.iloc[:, 0:num_cols].astype(bool).sum(axis=1))  # Average inventory levels in months w/ inventory > 0
        df["Min Inven when >0"] = df.iloc[:, 0:num_cols][df.gt(0)].min(1) #Min inventory level when inventory > 0

        # Write dataframe to workbook
        ws["A3"].options(pd.DataFrame, header=1, index=True, expand='table').value = df
        self.wb.save(self.filename)

    def current_inven(self):
        print("Current Inventory")

        # Select active sheet and fetch data
        ws = self.sheet[3]
        query = """SELECT 
                    concat(sku, location) AS concat,
                    tbl_inventory_history.sku,
                    tbl_inventory_history.location,
                    tbl_inventory_history.dates,
                    tbl_inventory_history.qty AS onhand_qty,
                    tbl_materials.unit_cost AS cogs,
                    tbl_materials.unit_price,
                    format(tbl_inventory_history.qty * tbl_materials.unit_cost, 2) AS total_cogs,
                    format(tbl_inventory_history.qty * tbl_materials.unit_price, 2) AS total_value 
                    FROM inven_mngt.tbl_inventory_history
                    INNER JOIN tbl_materials 
                    ON tbl_inventory_history.sku = tbl_materials.material
                    WHERE dates = (SELECT max(dates) FROM inven_mngt.tbl_inventory_history);"""
        data = self.sql.fetch(query)

        # Load data to Pandas dataframe
        df = pd.DataFrame(data[1], columns=data[0])
        df.fillna(0)
        df.reset_index(drop=True, inplace=True)
        df.set_index([data[0][0], data[0][1]])
        #ToDo: totals are formatted as text, change to int
        #df = pd.pivot_table(df, values='qty', index=['concat', 'sku', 'location', 'cogs', 'msrp'], fill_value=0)

        # Write dataframe to workbook
        ws["A3"].options(pd.DataFrame, header=1, index=True, expand='table').value = df
        self.wb.save(self.filename)

    def sales(self):
        print("Sales")

        # Select active sheet and fetch data
        headers = []
        ws = self.sheet[4]
        query = """SELECT
                    concat(sku, location) as concat,
                    sku,
                    location,
                    date_format(concat(year(dates), '-', Month(dates), '-', 1), '%Y-%m-%d' ) as dates,
                    sales
                    FROM tbl_sales;"""
        data = self.sql.fetch(query)

        # Load to Pandas dataframe
        df = pd.DataFrame(data[1], columns=data[0])
        df = pd.pivot_table(df, index=['concat', 'sku', 'location'], columns=['dates'], values='sales', aggfunc=np.sum, fill_value=0)

        df2 = pd.DataFrame(data[1], columns=data[0])
        df2 = pd.pivot_table(df2, index=['concat', 'sku', 'location'], columns=['dates'], values='sales', aggfunc=np.sum,fill_value=0)

        df3 = pd.DataFrame(data[1], columns=data[0])
        df3 = pd.pivot_table(df3, index=['concat', 'sku', 'location'], columns=['dates'], values='sales', aggfunc=np.sum, fill_value=0)
        num_cols = len(df.columns)
        col = df.iloc[:, 0:num_cols]

        # Calculations
        df["Total Demand"] = col.sum(axis=1)
        df["Active Months"] = col.astype(bool).sum(axis=1)
        df['Average'] = col.mean(axis=1)
        df['Min'] = col.min(axis=1)
        df['Max'] = col.max(axis=1)
        df["Avg Demand when >0"] = (col.sum(axis=1)) / (col.astype(bool).sum(axis=1))
        df["Min Demand when >0"] = col[df.gt(0)].min(1)

        # Standard Deviation
        df["StDev"] = col.std(axis=1)
        col_num_stdev1 = len(df.columns)+4

        df2 = df2.groupby(pd.PeriodIndex(df2.columns, freq='Q'), axis=1).std()
        pd.set_option("display.max_rows", None, "display.max_columns", None)

        for i in df2.columns: headers.append(str(i))
        df = df.merge(df2, how='inner', right_index=True, left_index=True)
        end_stdev_qtrs = len(df.columns) + len(headers)

        # Standard Deviation for months w/ demand >0
        df3 = df3.replace(0, np.NaN)
        df['Stdev >0'] = col[df.gt(0)].std(1)
        col_num_stdev2 = len(df.columns)+4
        end_col = col_num_stdev2 + len(headers)
        df3 = df3.groupby(pd.PeriodIndex(df3.columns, freq='Q'), axis=1).std()
        df = df.merge(df3, how='inner', right_index=True, left_index=True)
        print("len of df after merge with df3", len(df.columns))

        # Write dataframe to workbook & reformat headers
        ws["A3"].options(pd.DataFrame, header=1, index=True, expand='table').value = df
        ws.range((3, col_num_stdev1), (3, end_stdev_qtrs)).value = headers
        ws.range((3, col_num_stdev2), (3, end_col)).value = headers
        self.wb.save(self.filename)

    def forecasts(self):
        print("Forecasts")

        # Select active sheet and fetch data
        ws = self.sheet[5] # active sheet
        query = """SELECT 
                    sq1.concad,
                    sq1.sku,
                    sq1.dates,
                    sq1.location,
                    sq1.sales,
                    sq1.forecast,
                    sq1.fc_error,
                    sq1.error_sqrd,
                    sq1.mape                    
                    FROM
                    (SELECT
                    concat(tbl_sales.sku, tbl_sales.location) as concad,
                    tbl_sales.sku,
                    date_format(concat(year(tbl_forecasts.dates), '-', Month(tbl_forecasts.dates), '-', 1), '%Y-%m-%d' ) as dates,
                    tbl_sales.location,
                    tbl_sales.sales,
                    tbl_forecasts.forecast,
                    sales-forecast as fc_error,
                    POWER(sales-forecast,2) as error_sqrd,
                    ABS((sales-forecast)/forecast) as mape
                    FROM inven_mngt.tbl_sales
                    INNER JOIN tbl_forecasts
                    ON tbl_sales.sku = tbl_forecasts.sku AND 
                    tbl_sales.location = tbl_forecasts.location AND 
                    extract(year_month from tbl_sales.dates) = extract(year_month from tbl_forecasts.dates)) AS sq1;"""
        data = self.sql.fetch(query)

        # Create Dataframes
        df = pd.DataFrame(data[1], columns=data[0])
        index = ['concad', 'sku', 'location']
        dframes = ['Sales', 'Forecasts', 'Forecast Error', 'Forecast Error Squared', 'MAPE']
        frame_format = ['#,##0', '#,##0', '#,##0', '#,##0', '0%']

        # Sales Dataframe
        sales_df = pd.pivot_table(df, index=index, columns=['dates'], values='sales', aggfunc=np.sum, fill_value=0)
        num_periods = len(sales_df.columns)
        headers = sales_df.columns

        # Forecast Dataframe
        fc_df = pd.pivot_table(df, index=index, columns=['dates'], values='forecast', aggfunc=np.sum, fill_value=0)

        # Error Dataframes
        error_df = pd.pivot_table(df, index=index, columns=['dates'], values='fc_error', aggfunc=np.sum, fill_value=0)
        bias_df = pd.pivot_table(df, index=index, values='fc_error', aggfunc=np.sum, fill_value=0)

        # Error Squared Dataframes
        sqrd_df = pd.pivot_table(df, index=index, columns=['dates'], values='error_sqrd', aggfunc=np.sum, fill_value=0)
        sqrd_col = sqrd_df.iloc[:, 0:len(sqrd_df.columns)]

        rmse_df = pd.pivot_table(df, index=index, values='error_sqrd', aggfunc=np.sum, fill_value=0)
        rmse_df['RMSE'] = np.sqrt(rmse_df['error_sqrd']/num_periods)
        #rmse_df['RMSE Active'] = math.sqrt(rmse_df['error_sqrd'] / num_periods)
        rmse_df['RMSE Active'] = np.sqrt(rmse_df['error_sqrd']/sqrd_col.astype(bool).sum(axis=1))

        # MAPE Dataframes
        mape_df = pd.pivot_table(df, index=index, columns=['dates'], values='mape', aggfunc=np.sum, fill_value=0)
        mape_df2 = pd.pivot_table(df, index=index, values='mape', aggfunc=np.sum, fill_value=0)
        mape_df2['Avg MAPE'] = mape_df2['mape'] / num_periods
        mape_df = mape_df.merge(mape_df2, how='inner', right_index=True, left_index=True)

        print(rmse_df)

        # Merge dataframes
        df = sales_df.merge(fc_df, how='inner', right_index=True, left_index=True)
        df = df.merge(error_df, how='inner', right_index=True, left_index=True)
        df = df.merge(sqrd_df, how='inner', right_index=True, left_index=True)
        df = df.merge(mape_df, how='inner', right_index=True, left_index=True)

        ws["A3"].options(pd.DataFrame, header=1, index=True, expand='table').value = df

        # Format headers for each dataframe
        col = 4
        count = 0
        for i in dframes:
            count += 1
            format = frame_format[count - 1]
            if count == 1:
               col1 = col + num_periods -1
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

        #ws.range(3, len(df.columns)+3).value = "Sum of Error%"


        # Write dataframe to workbook
        self.wb.save(self.filename)

    def inventory_turns(self):
        print("Inventory Turns")

        # Select active sheet and fetch data
        ws = self.sheet[6]
        query = "SELECT material, unit_cost, currency, category, subcategory, org_code FROM tbl_materials;"
        data = self.sql.fetch(query)

        # Load to Pandas dataframe
        df = pd.DataFrame(data[1], columns=data[0])
        df = df.set_index(data[0][0])

        # Write dataframe to workbook
        ws["A3"].options(pd.DataFrame, header=1, index=True, expand='table').value = df
        self.wb.save(self.filename)

    def segments(self):
        # Select active sheet and fetch data
        ws = self.sheet[7]

    def segment_calc(self):
        # Select active sheet and fetch data
        ws = self.sheet[8]

    def materials(self):
        print("Materials")

        # Select active sheet and fetch data
        ws = self.sheet[9]
        query = "SELECT material, description, unit_cost, currency, category, subcategory, org_code FROM tbl_materials;"
        data = self.sql.fetch(query)

        # Load to Pandas dataframe
        df = pd.DataFrame(data[1], columns=data[0])
        df = df.set_index(data[0][0])

        # Write dataframe to workbook
        ws["A3"].options(pd.DataFrame, header=1, index=True, expand='table').value = df
        self.wb.save(self.filename)

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