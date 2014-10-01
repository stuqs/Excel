#!/usr/bin/python
# -*- coding: utf-8 -*-









# """
# Вариант для XLSX
# """
#
# import os
# import glob
# import csv
# from xlsxwriter.workbook import Workbook
#
#
# for csvfile in glob.glob(os.path.join('.', '*.csv')):
#     workbook = Workbook(csvfile + '.xlsx')
#     worksheet = workbook.add_worksheet()
#     with open(csvfile, 'rb') as f:
#         reader = csv.reader(f)
#         for r, row in enumerate(reader):
#             for c, col in enumerate(row):
#                 worksheet.write(r, c, col.decode("utf-8"))
#     workbook.close()



import csv
import xlwt
import glob
import os


for csvfile in glob.glob(os.path.join('.', '*.csv')):
    with open(csvfile, 'rb') as f:
        reader = csv.reader(f, delimiter=",")
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet(u"Лист 1")
        for r, row in enumerate(reader):
            for c, col in enumerate(row):
                sheet.write(r, c, col.decode("utf-8"))
        workbook.save(csvfile + '.xls')












# import csv
# from openpyxl import Workbook
# from openpyxl.cell import get_column_letter
#
# f = open(r'C:\Users\Asus\Desktop\herp.csv', "rU")
#
# csv.register_dialect('colons', delimiter=':')
#
# reader = csv.reader(f, dialect='colons')
#
# wb = Workbook()
# dest_filename = r"C:\Users\Asus\Desktop\herp.xlsx"
#
# ws = wb.worksheets[0]
# ws.title = "A Snazzy Title"
#
# for row_index, row in enumerate(reader):
#     for column_index, cell in enumerate(row):
#         column_letter = get_column_letter((column_index + 1))
#         ws.cell('%s%s'%(column_letter, (row_index + 1))).value = cell
#
# wb.save(filename = dest_filename)









