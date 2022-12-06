_author_ = 'arichland'

import pymysql
import pprint
import credentials
pp = pprint.PrettyPrinter(indent=1)

class SQLTables:
    def __init__(self, **kargs):
        self.user = credentials.user
        self.pw = credentials.pw
        self.host = credentials.host
        self.db = credentials.db
        self.con = pymysql.connect(user=self.user, password=self.pw, host=self.host, database=self.db)
        self.tables = [kargs.get('tables')] if type(kargs.get('tables')) == type(str()) else kargs.get('tables')

    def create_tables(self):
        if self.tables == True:
            self.tbl_current_inventory()
            self.tbl_inventory_history()
            self.tbl_sales()
            self.tbl_materials()
            self.tbl_skus()
            self.tbl_forecasts()
            self.tbl_locations()
            self.tbl_orders()
        else:
            for table in self.tables:
                if table == 'current inventory':
                    self.tbl_current_inventory()

                elif table == 'inventory history':
                    self.tbl_inventory_history()

                elif table == 'inventory history':
                    self.tbl_sales()

                elif table == 'material':
                    self.tbl_materials()

                elif table == 'skus':
                    self.tbl_skus()

                elif table == 'forecasts':
                    self.tbl_forecasts()

                elif table == 'location':
                    self.tbl_locations()

                elif table == 'orders':
                    self.tbl_orders()

    def tbl_current_inventory(self):
        with self.con.cursor() as cur:
                qry = """
                CREATE TABLE IF NOT EXISTS tbl_current_inventory(
                id INT AUTO_INCREMENT PRIMARY KEY,
                cogs DOUBLE,
                currency TINYTEXT,
                dates DATE,
                inven_id TINYTEXT,
                inven_type TINYTEXT,
                location TINYTEXT,
                msrp DOUBLE,
                quantity INT)
                ENGINE=INNODB;"""
                cur.execute(qry)

    def tbl_inventory_history(self):
        with self.con.cursor() as cur:
            qry = """
                CREATE TABLE IF NOT EXISTS tbl_inventory_history(
                id INT AUTO_INCREMENT PRIMARY KEY,
                abc_class TINYTEXT,
                cogs DOUBLE,
                currency TINYTEXT,
                dates DATE,
                inven_type TINYTEXT,
                inven_id TINYTEXT,
                location TINYTEXT,
                msrp DOUBLE,
                quantity INT)
                ENGINE=INNODB;"""
            cur.execute(qry)

    def tbl_sales(self):
        con = pymysql.connect(user=self.user, password=self.pw, host=self.host, database=self.db)
        with con.cursor() as cur:
            try:
                qry_create_table = """
                CREATE TABLE IF NOT EXISTS tbl_sales(
                id INT AUTO_INCREMENT PRIMARY KEY,
                sku_id TEXT,
                date DATE,
                location TEXT,
                sales INT)
                ENGINE=INNODB;"""
                cur.execute(qry_create_table)
            finally:
                con.commit()
                cur.close()
                con.close()

    def tbl_materials(self):
        con = pymysql.connect(user=self.user, password=self.pw, host=self.host, database=self.db)
        with con.cursor() as cur:
            try:
                qry_create_table = """
                       CREATE TABLE IF NOT EXISTS tbl_materials(
                       id INT AUTO_INCREMENT PRIMARY KEY,
                       category TINYTEXT,
                       currency TINYTEXT,
                       description TEXT,
                       dim_uom TINYTEXT,
                       height FLOAT,
                       length FLOAT,
                       location TEXT,
                       make_buy TEXT,
                       name TINYTEXT,          
                       material_id TINYTEXT,
                       subcategory TINYTEXT,
                       unit_cost FLOAT,
                       unit_price FLOAT,
                       weight FLOAT,
                       weight_uom TINYTEXT,
                       width FLOAT))
                       ENGINE=INNODB;"""
                cur.execute(qry_create_table)
            finally:
                con.commit()
                cur.close()
                con.close()

    def tbl_skus(self):
        con = pymysql.connect(user=self.user, password=self.pw, host=self.host, database=self.db)
        with con.cursor() as cur:
            try:
                qry_create_table = """
                       CREATE TABLE IF NOT EXISTS tbl_skus(
                       id INT AUTO_INCREMENT PRIMARY KEY,
                       category TINYTEXT,
                       cogs FLOAT,
                       currency TINYTEXT,
                       description TEXT,
                       dim_uom TINYTEXT,
                       height FLOAT,                       
                       length FLOAT,
                       location TEXT,
                       make_buy TEXT,
                       name TINYTEXT,
                       retail_price FLOAT,            
                       sku_id TINYTEXT,
                       subcategory TINYTEXT,
                       weight FLOAT,
                       weight_uom TINYTEXT,
                       unit_price FLOAT
                       width FLOAT,
                       width_uom TINYTEXT)
                       ENGINE=INNODB;"""
                cur.execute(qry_create_table)
            finally:
                con.commit()
                cur.close()
                con.close()

    def tbl_forecasts(self):
        con = pymysql.connect(user=self.user, password=self.pw, host=self.host, database=self.db)
        with con.cursor() as cur:
            try:
                qry_create_table = """
                       CREATE TABLE IF NOT EXISTS tbl_forecasts(
                       id INT AUTO_INCREMENT PRIMARY KEY,
                       dates DATE,
                       forecast INT,
                       location TINYTEXT,
                       sku_id TINYTEXT)
                       ENGINE=INNODB;"""
                cur.execute(qry_create_table)
            finally:
                con.commit()
                cur.close()
                con.close()

    def tbl_locations(self):
        con = pymysql.connect(user=self.user, password=self.pw, host=self.host, database=self.db)
        with con.cursor() as cur:
            try:
                query = """
                       CREATE TABLE IF NOT EXISTS tbl_locations(
                       id INT AUTO_INCREMENT PRIMARY KEY,
                       name TINYTEXT,
                       address1 TINYTEXT,
                       address2 TINYTEXT,
                       city TINYTEXT,
                       category TINYTEXT,
                       state TINYTEXT,
                       country TINYTEXT,
                       longitude DOUBLE,
                       latitude DOUBLE)
                       ENGINE=INNODB;"""
                cur.execute(query)
            finally:
                con.commit()
                cur.close()
                con.close()

    def tbl_orders(self):
        con = pymysql.connect(user=self.user, password=self.pw, host=self.host, database=self.db)
        with con.cursor() as cur:
            try:
                query = """
                       CREATE TABLE IF NOT EXISTS tbl_orders(
                       id INT AUTO_INCREMENT PRIMARY KEY,
                       order_id TINYTEXT,
                       customer_id INT,
                       customer_name TINYTEXT,
                       ship_origin TINYTEXT,
                       ship_date DATE,
                       ship_cost FLOAT,
                       ship_method TINYTEXT,
                       carrier TINYTEXT,
                       address1 TINYTEXT,
                       address2 TINYTEXT,
                       city TINYTEXT,
                       state TINYTEXT,
                       country TINYTEXT,
                       longitude FLOAT,
                       latitude FLOAT)
                       ENGINE=INNODB;"""
                cur.execute(query)
            finally:
                con.commit()
                cur.close()
                con.close()

def create_tables(**kargs):
    go = SQLTables(**kargs)
    go.create_tables()

if __name__ == '__main__':
    create_tables()