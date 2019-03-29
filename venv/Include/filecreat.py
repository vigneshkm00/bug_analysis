import openpyxl
import os
import openpyxl.styles
os.chdir("C:\\Users\\vikm\Documents\\testbuganalysis\\report")
wb = openpyxl.Workbook()
sheet = wb.active
print(sheet.title)
sheet.title = "725714"
print(sheet.title)
sheet['A1']='Project ID'
sheet['B1']='Before(count)'
sheet['C1']='After(count)'
sheet['D1']='Before(blank_count)'
sheet['D1']='Before(blank_count)'
sheet['E1']='Missing(count)'
sheet['F1']='Comment missing'
sheet['A']='test'
wb.save("Report.xlsx")