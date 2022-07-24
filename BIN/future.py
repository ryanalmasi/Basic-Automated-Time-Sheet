from openpyxl import Workbook
from openpyxl import load_workbook

wb = load_workbook("test.xlsx")
ws = wb["Sheet1"] 

for row in ws.rows:
    for cell in row:
        print(cell.comment)