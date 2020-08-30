import xlwings as xw
import openpyxl as pyxl

wb = xw.Book('ProjectTemp.xlsx')
sht1 = wb.sheets['Sheet1']

#sht1.range('B2').value = 35

row_count = pyxl.load_workbook('ProjectTemp.xlsx').worksheets[0].max_row

for i in range(2,row_count+1):
    print(i)




