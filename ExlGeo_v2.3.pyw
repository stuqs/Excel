#!/usr/bin/python
# -*- coding: utf-8 -*-

from xlrd import open_workbook
from xlutils.copy import copy
import xlwt
from getxlwtstylelist import get_xlwt_style_list
from tkFileDialog import askopenfilenames
from tkMessageBox import showerror
import os
import csv


def geodata(geodata_name):
    '''
    Open the file with geo data and converts them to strings of the form [name longitude latitude]
    :param geodata_name: Path to the file with geo data
    :return: List with processed data
    '''
    data = []
    if not os.path.isfile(geodata_name):
        showerror("Файл данных не найден", "Файл данных не найден, выберите файл c гео данными.")
        geodata_name = askopenfilenames(initialdir=os.path.abspath(os.getcwd()), filetypes=[("Файл данных txt", ".txt")], title="Выберите файл данных txt")[0]
    with open(geodata_name) as geodata_f:
        geodata_f.readline()
        for d in geodata_f:
            data_row = []
            data_row.append(d.split(",")[0].strip().decode('utf-8'))
            data_row.append(float(d.split(",")[1].strip()))
            data_row.append(float(d.split(",")[2].strip()))
            data.append(data_row)
    return data

def get_input_name():
    """
    Set the Excel files for conversion
    :return: List of paths to the opened files
    """
    xlsTypes = [("Книга Excel 97 - 2003 / csv", ".xls .csv")]
    return askopenfilenames(initialdir=os.path.abspath(os.getcwd()), filetypes=xlsTypes, title="Выберите файлы Excel или CSV")

def get_output_name(input_path):
    """
    Set name to the output file
    :param input_path: Path to the input file
    :return: Name of the output file
    """
    file_name, file_ext = os.path.splitext(os.path.basename(input_path))
    return "out" + os.path.sep + file_name + "_geo" + file_ext

def check_for_csv(input_paths):
    path_wo_csv = []
    for path in input_paths:
        file_name, file_ext = os.path.splitext(os.path.basename(path))
        if file_ext == ".csv":
            path_wo_csv.append(convert_csv(path))
        else:
            path_wo_csv.append(path)
    return path_wo_csv

def convert_csv(path_csv):
    with open(path_csv, 'rb') as f:
        reader = csv.reader(f, delimiter=",")
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet(u"Лист 1")
        for r, row in enumerate(reader):
            for c, col in enumerate(row):
                sheet.write(r, c, col.decode("utf-8"))
        file_name, file_ext = os.path.splitext(os.path.basename(path_csv))
        if not os.path.isdir("out"):
            os.mkdir("out")
        workbook.save("out" + str(os.path.sep) + file_name + '.xls')
    return "out" + str(os.path.sep) + file_name + '.xls'

def process_files(geodata_name, input_paths):
    """
    Process Excel files, if the search request is valid, then writes the geo data in the copy file.
    :param geodata_name: Name of the file with geo data
    :param input_paths: List of paths to Excel files
    """
    try:
        data = geodata(geodata_name)
    except UnicodeDecodeError:
        showerror("Ошибка кодирования", "Файл данных должен быть закодирован в utf-8")
        data = geodata(askopenfilenames(initialdir=os.path.abspath(os.getcwd()), filetypes=[("Файл данных txt", ".txt")], title="Выберите файл данных txt")[0])
    for book in input_paths:
        book_flag = False
        with open_workbook(book, on_demand=True, formatting_info=True) as rb:
            header = False
            wb = copy(rb)
            for numb, sheet in enumerate(rb.sheets()):
                for row in range(sheet.nrows):
                    for col in range(sheet.ncols):
                        for data_row in data:
                            if sheet.cell(row, col).value == data_row[0]:
                                book_flag = True
                                sheet_wb = wb.get_sheet(numb)
                                sheet_wb.write(row, sheet.ncols, data_row[1])
                                sheet_wb.write(row, sheet.ncols+1, data_row[2])
                                if not header:
                                    header = True
                                    style_list = get_xlwt_style_list(rb)
                                    wb.get_sheet(numb).write(0, sheet.ncols, u"Широта", style=style_list[sheet.cell_xf_index(0, 0)])
                                    wb.get_sheet(numb).write(0, sheet.ncols+1, u"Долгота", style=style_list[sheet.cell_xf_index(0, 0)])
                                break
        if book_flag:
            if not os.path.isdir("out"):
                os.mkdir("out")
            wb.save(get_output_name(book))

def main():
    geodata_name = "data.txt"
    process_files(geodata_name, check_for_csv(get_input_name()))

if __name__ == '__main__':
    main()
