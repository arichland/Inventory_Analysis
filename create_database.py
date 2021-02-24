_author_ = 'arichland'

import pymysql
import pydict

class tables:
    local = pydict.localhost.get
    user = local('user')
    pw = local('password')
    host = local('host')
    db = local('database')

    def current_inventory():
        con = pymysql.connect(user=tables.user, password=tables.pw, host=tables.host, database=tables.db)
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

    def inventory_history():
        con = pymysql.connect(user=tables.user, password=tables.pw, host=tables.host, database=tables.db)
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

    def sales():
        con = pymysql.connect(user=tables.user, password=tables.pw, host=tables.host, database=tables.db)
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

    def materials():
        con = pymysql.connect(user=tables.user, password=tables.pw, host=tables.host, database=tables.db)
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

    def skus():
        con = pymysql.connect(user=tables.user, password=tables.pw, host=tables.host, database=tables.db)
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

    def forecasts():
        con = pymysql.connect(user=tables.user, password=tables.pw, host=tables.host, database=tables.db)
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

    def create():
        tables.current_inventory()
        tables.inventory_history()
        tables.sales()
        tables.materials()
        tables.skus()
        tables.forecasts()

tables.create()


