import os
import openpyxl

os.chdir('F:\git_practice_dev1')
wb = openpyxl.load_workbook('0WDAGAPP_F130D11_PIT_REPORT.xlsx')

Sheet_name = wb.get_sheet_names()
sheet = wb.get_sheet_by_name(Sheet_name[0])
print sheet.cell(row=1,column=1).value
