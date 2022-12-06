_author_ = 'arichland'

import calculations as cals

filename = "Inventory Analysis.xlsx"
cal = cals.Calc(filename)

def functions():
    #cal.inventory_history()
    #cal.inventory_turns()
    #cal.current_inven()
    #cal.sales()
    cal.forecasts()
    #cal.materials()
    #cal.skus()
functions()