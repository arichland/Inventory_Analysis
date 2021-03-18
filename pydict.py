_author_ = 'arichland'

localhost = {
    'user': 'root',
    'password': 'RichlanD0530)(',
    'host': '127.0.0.1',
    'database': 'inven_mngt',
    'charset': 'utf8mb4'
}

queries = {
    "inventory history": "SELECT id, concat(sku, location) as concat, sku, location, abc_class, dates, qty FROM tbl_inventory_history ORDER BY id ASC;",
    "current inventory": "SELECT id, concat(sku, location) as concat, sku, location, cogs, msrp, qty FROM tbl_current_inventory ORDER BY id ASC;",
    "sales history": "SELECT id, concat(sku, location) as concat, sku, location, dates, sales, '' as col7 FROM tbl_sales ORDER by id ASC;",
    "forecasts": "SELECT id, concat(sku, location) as concat, sku, location, dates, forecast, '' as col7 FROM tbl_forecasts ORDER by id ASC;",
    "materials": "SELECT id, material, unit_cost, currency, category, subcategory, org_code FROM tbl_materials ORDER by id ASC;",
    "skus": "SELECT id, concat(sku_id, location) as concat, sku_id, name, description, location, make_buy FROM tbl_skus ORDER by id ASC;"
}