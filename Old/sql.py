_author_ = 'arichland'

import pymysql
import pydict
import pprint
pp = pprint.PrettyPrinter(indent=1)

class Database:
    def __init__(self):
        self.user = pydict.localhost.get('user')
        self.pw = pydict.localhost.get('password')
        self.host = pydict.localhost.get('host')
        self.db = pydict.localhost.get('database')

    def fetch(self, query):
        con = pymysql.connect(user=self.user, password=self.pw, host=self.host, database=self.db)
        cols = []
        with con.cursor() as cur:
            cur.execute(query)
            con.commit()
            rows = cur.fetchall()
            col = cur.description
            for i in col:
                cols.append(i[0])
        cur.close()
        con.close()
        return cols, rows

    def tbl_current_inventory(self):
        con = pymysql.connect(user=self.user, password=self.pw, host=self.host, database=self.db)
        with con.cursor() as cur:
            try:
                qry_create_table = """
                CREATE TABLE IF NOT EXISTS tbl_current_inventory(
                id INT AUTO_INCREMENT PRIMARY KEY,
                sku TEXT,
                dates DATE,
                location TEXT,
                cogs DOUBLE,
                msrp DOUBLE,
                qty INT)
                ENGINE=INNODB;"""
                cur.execute(qry_create_table)
            finally:
                con.commit()
                cur.close()
                con.close()

    def tbl_inventory_history(self):
        con = pymysql.connect(user=self.user, password=self.pw, host=self.host, database=self.db)
        with con.cursor() as cur:
            try:
                qry_create_table = """
                CREATE TABLE IF NOT EXISTS tbl_inventory_history(
                id INT AUTO_INCREMENT PRIMARY KEY,
                sku TEXT,
                dates DATE,
                location TEXT,
                sku_type TEXT,
                abc_class TEXT,
                qty INT)
                ENGINE=INNODB;"""
                cur.execute(qry_create_table)
            finally:
                con.commit()
                cur.close()
                con.close()

    def tbl_sales(self):
        con = pymysql.connect(user=self.user, password=self.pw, host=self.host, database=self.db)
        with con.cursor() as cur:
            try:
                qry_create_table = """
                CREATE TABLE IF NOT EXISTS tbl_sales(
                id INT AUTO_INCREMENT PRIMARY KEY,
                sku TEXT,
                dates DATE,
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
                       material TEXT,
                       description TEXT,
                       unit_cost FLOAT,
                       currency TEXT,
                       unit_price FLOAT,
                       category TEXT,
                       subcategory TEXT,
                       org_code TEXT)
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
                       concat TEXT,
                       sku_id TEXT,
                       location TEXT,
                       make_buy TEXT,
                       name TEXT,
                       description TEXT)
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
                       concat TEXT,
                       sku TEXT,
                       location TEXT,
                       dates DATE,
                       forecast INT)
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
                       name TEXT,
                       address1 TEXT,
                       address2 TEXT,
                       city TEXT,
                       state TEXT,
                       country TEXT,
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
                       order_id TEXT,
                       customer_id INT,
                       customer_name TEXT,
                       ship_origin TEXT,
                       ship_date DATE,
                       ship_cost FLOAT,
                       ship_method TEXT,
                       carrier TEXT,
                       address1 TEXT,
                       address2 TEXT,
                       city TEXT,
                       state TEXT,
                       country TEXT,
                       longitude FLOAT,
                       latitude FLOAT)
                       ENGINE=INNODB;"""
                cur.execute(query)
            finally:
                con.commit()
                cur.close()
                con.close()

def create_database():
    db = Database()
    db.tbl_current_inventory()
    db.tbl_forecasts()
    db.tbl_inventory_history()
    db.tbl_locations()
    db.tbl_materials()
    db.tbl_orders()
    db.tbl_sales()
    db.tbl_skus()
create_database()