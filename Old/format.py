_author_ = 'Gary.Manley'
import xlwings as xw
from xlwings import constants

# setting up RGB Colours
def rgb(col):
    if col == 'Turq':
        R = 0
        G = 165
        B = 185
    elif col == 'Blue':
        R = 55
        G = 70
        B = 150
    elif col == 'Pink':
        R = 230
        G = 30
        B = 105
    elif col == 'Oran':
        R = 245
        G = 150
        B = 20
    elif col == 'Yell':
        R = 255
        G = 220
        B = 50
    elif col == 'Pur':
        R = 130
        G = 55
        B = 140
    elif col == 'DBlu':
        R = 64
        G = 86
        B = 106
    elif col == 'Grey':
        R = 128
        G = 128
        B = 128
    elif col == 'White':
        R = 255
        G = 255
        B = 255
    elif col == 'Black':
        R = 0
        G = 0
        B = 0
    return(R,G,B)

# formats range in XLwings

def formatrange(xlrange,fontcol,backcol,border_weight,border_style,bot_border,align):
    # set up background colour
    xw.Range(xlrange).color = rgb(backcol)

    # Set up Font Colour
    xw.Range(xlrange).api.Font.Color = xw.utils.rgb_to_int(rgb(fontcol))

    # set up border weight
    if border_weight > 0:
        xw.Range(xlrange).api.Borders(1).Weight = border_weight ### sets vertical
        xw.Range(xlrange).api.Borders(2).Weight = border_weight ### sets vertical
        xw.Range(bot_border).api.Borders(3).Weight = border_weight ### sets horizontal top needs to be one row more

    # set up border style
    if border_style > 0:
        xw.Range(xlrange).api.Borders(1).LineStyle = border_style
        xw.Range(xlrange).api.Borders(2).LineStyle = border_style
        xw.Range(bot_border).api.Borders(3).LineStyle = border_style

    # set up alignment
    if align =='Center':
        xw.Range(xlrange).api.HorizontalAlignment = constants.HAlign.xlHAlignCenter
    if align =='Right':
        xw.Range(xlrange).api.HorizontalAlignment = constants.HAlign.xlHAlignRight
    if align =='Left':
        xw.Range(xlrange).api.HorizontalAlignment = constants.HAlign.xlHAlignLeft