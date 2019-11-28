from openpyxl import *
from datetime import date
workbook = Workbook()


def createTemplate():
    sheet["A1"] = 'Date'
    sheet["C1"] = 'Mileage'
    sheet["E1"] = 'How many liters'
    sheet["G1"] = 'How many liters Computer say'
    sheet['I1'] = 'Amount for refueling'
    sheet['K1'] = 'Sum distance in KM'
    sheet['L1'] = 'Money sum'
    sheet["N1"] = 'AVG combustion'
    sheet["P1"] = 'AVG price per day'
    sheet.freeze_panes = "A1"

def takeDate():
    today = date.today()
    sheet["A2"] = str(today)

try:
    workbook = load_workbook(filename="Raport.xlsx")
    sheet = workbook.active
    print('exist')
except:
    workbook.save(filename="Raport.xlsx")
    sheet = workbook.active
    print('created')
    createTemplate()
    takeDate()
    workbook.save(filename="Raport.xlsx")





