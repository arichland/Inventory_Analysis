_author_ = 'arichland'

import pymysql
import pprint
import credentials
pp = pprint.PrettyPrinter(indent=1)

class SQLQuery:
    def __init__(self, **kargs):
        self.user = credentials.user
        self.pw = credentials.pw
        self.host = credentials.host
        self.db = credentials.db
        self.con = pymysql.connect(user=self.user, password=self.pw, host=self.host, database=self.db)
        self.qry = kargs.get('qry')

    def arguments(self):
        data_dict = {}
        if self.qry == True:
            data_dict.update({'inventory history': self.qry_inventory_history()})
            data_dict.update({'current inventory': self.qry_current_inventory()})
            data_dict.update({'sales': self.qry_sales()})
            data_dict.update({'forecasts': self.qry_forecasts()})

        elif self.qry == 'inventory history':
            data = self.qry_inventory_history()
            data_dict.update({self.qry: data})

        elif self.qry == 'current inventory':
            data = self.qry_current_inventory()
            data_dict.update({self.qry: data})

        elif self.qry == 'sales':
            data = self.qry_sales()
            data_dict.update({self.qry: data})

        elif self.qry == 'forecasts':
            data = self.qry_forecasts()
            data_dict.update({self.qry: data})

        return data_dict

    def qry_inventory_history(self):
        with self.con.cursor() as cur:
            data = []
            qry = """
                SELECT
                    inven_id,
                    location,
                    dates,
                    quantity
                FROM tbl_inventory_history
                ORDER BY id ASC;"""
            cur.execute(qry)
            rows = cur.fetchall()
            for row in rows:
                data.append([row[0],
                            row[1],
                            row[2],
                            row[3]])
            return data

    def qry_current_inventory(self):
        with self.con.cursor() as cur:
            data = []
            qry = """
                SELECT
                    inven_id,
                    location,
                    dates,
                    quantity,
                    unit_cost,
                    unit_price
                FROM inven_mngt.tbl_inventory_history
                INNER JOIN tbl_materials 
                ON tbl_inventory_history.inven_id = tbl_materials.material
                WHERE dates = (SELECT max(dates) FROM inven_mngt.tbl_inventory_history);"""
            cur.execute(qry)
            rows = cur.fetchall()
            for row in rows:
                total_cogs = round(row[3] * row[4], 2)
                total_value = round(row[3] * row[5], 2)
                data.append([row[0],
                            row[1],
                            row[2],
                            row[3],
                            row[4],
                            row[5],
                            total_cogs,
                            total_value])
        return data

    def qry_sales(self):
        with self.con.cursor() as cur:
            data = []
            qry = """
                SELECT
                    inven_id,
                    location,
                    DATE_FORMAT(concat(year(dates), '-', Month(dates), '-', 1), '%Y-%m-%d' ) as dates,
                    sales
                FROM tbl_sales;"""
            cur.execute(qry)
            rows = cur.fetchall()
            for row in rows:
                data.append([row[0],
                            row[1],
                            row[2],
                            row[3]])
        return data

    def qry_forecasts(self):
        with self.con.cursor() as cur:
            data = []
            qry = """
                SELECT 
                    sq1.inven_id,
                    sq1.dates,
                    sq1.location,
                    sq1.sales,
                    sq1.forecast,
                    sq1.fc_error,
                    sq1.error_sqrd,
                    sq1.mape                    
                FROM (SELECT
                        tbl_sales.inven_id,
                        date_format(concat(year(tbl_forecasts.dates), '-', Month(tbl_forecasts.dates), '-', 1), '%Y-%m-%d' ) as dates,
                        tbl_sales.location,
                        tbl_sales.sales,
                        tbl_forecasts.forecast,
                        sales-forecast as fc_error,
                        POWER(sales-forecast,2) as error_sqrd,
                        ABS((sales-forecast)/forecast) as mape
                    FROM inven_mngt.tbl_sales
                    INNER JOIN tbl_forecasts ON tbl_sales.inven_id = tbl_forecasts.sku_id AND 
                        tbl_sales.location = tbl_forecasts.location AND 
                        extract(year_month from tbl_sales.dates) = extract(year_month from tbl_forecasts.dates)) AS sq1;"""
            cur.execute(qry)
            rows = cur.fetchall()
            for row in rows:
                data.append([row[0],
                            row[1],
                            row[2],
                            row[3],
                            row[4],
                            row[5],
                            row[6],
                            float(row[7])])
        return data

def query(**kargs):
    go = SQLQuery(**kargs)
    return go.arguments()

if __name__ == '__main__':
    query()
