_author_ = 'arichland'

#from openpyxl import Workbook

import openpyxl as xl

# Create new worksbook 'Inventory Analysis.xlsx'
wb = xl.Workbook()
sheet = xl.Workbook.active
filename = "Inventory_Analysis.xlsx"

# Create sheets
params = wb.create_sheet("Parameters", 0)
analysis = wb.create_sheet("Analysis", 1)
inven_hist = wb.create_sheet("Inventory History", 2)
cur_inven = wb.create_sheet("Current Inventory", 3)
sales_hist = wb.create_sheet("Sales History", 4)
forecast_hist = wb.create_sheet("Forecast History", 5)
inven_turns = wb.create_sheet("Inventory Turns", 6)
segmen = wb.create_sheet("Segmentation", 7)
segmen_calc = wb.create_sheet("Segmentation Calculation", 8)
material = wb.create_sheet("Material Data", 9)
skus = wb.create_sheet("SKU Data", 10)
locale = wb.create_sheet("Location Data", 11)
conversion = wb.create_sheet("Conversion", 12)

# Save
wb.save(filename=filename)