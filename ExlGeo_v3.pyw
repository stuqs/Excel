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
from win32com.client.gencache import EnsureDispatch


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
    Set the Excel or CVS files for conversion
    :return: List of paths to the opened files
    """
    xlsTypes = [("Файлы Excel или csv", ".xls .csv .xlsx")]
    return askopenfilenames(initialdir=os.path.abspath(os.getcwd()), filetypes=xlsTypes, title="Выберите файлы Excel или CSV")

def separate(input_paths):
    """
    Create dict with lists of paths to csv, xls and xlsx files
    :param input_paths: List of paths to the opened files
    :return: Dict with separated paths
    """
    sep_paths = {".csv": [], ".xls": [], ".xlsx": [], "del": [], "out": []}
    for path in input_paths:
        file_name, file_ext = os.path.splitext(os.path.basename(path))
        sep_paths[file_ext].append(path)
    return sep_paths

def get_output_name(input_path):
    """
    Set name to the output file
    :param input_path: Path to the input file
    :return: Name of the output file
    """
    file_name, file_ext = os.path.splitext(os.path.basename(input_path))
    return os.path.abspath("out" + os.path.sep + file_name + "_geo" + file_ext)

def check_for_csv(inp_dict):
    """
    Convert files if they exist, write changes to output dict
    :param inp_dict: Dict with separated paths
    :return: Updated dict with separated paths
    """
    if inp_dict[".csv"]:
        for path in inp_dict[".csv"]:
            csv_path = convert_csv(path)
            inp_dict[".xls"].append(csv_path)
            inp_dict["del"].append(csv_path)
            inp_dict["out"].append(csv_path)
        inp_dict[".csv"] = []
    return inp_dict

def convert_csv(path_csv):
    """
    Convert csv file to xls ans save it to the same folder
    :param path_csv: Path to csv file
    :return: Path to converted xls file
    """
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

def convert_xls_to_xlsx(inp_dict):
    """
    Convert xls files in output list to xlsx files
    :param inp_dict: Dict with separated paths
    :return: Updated dict with separated paths
    """
    if inp_dict["out"]:
        for fname in inp_dict["out"]:
            excel = EnsureDispatch('Excel.Application')
            fname = os.path.abspath(fname.encode("utf-8"))
            fname = os.path.abspath(fname.decode("utf-8"))
            wb = excel.Workbooks.Open(fname)
            excel.DisplayAlerts = False
            wb.SaveAs(fname+"x", FileFormat=51)    #FileFormat = 51 is for .xlsx extension
            wb.Close()                             #FileFormat = 56 is for .xls extension
            excel.Application.Quit()
            excel.DisplayAlerts = True
            inp_dict[".xlsx"].append(fname+"x")
        inp_dict["out"] = []
    return inp_dict

def convert_xlsx_to_xls(inp_dict):
    """
    Convert xlsx files in .xlsx list to xls files
    :param inp_dict: Dict with separated paths
    :return: Updated dict with separated paths
    """
    if inp_dict[".xlsx"]:
        for fname in inp_dict[".xlsx"]:
            fname = os.path.abspath(fname.encode("utf-8"))
            fname = os.path.abspath(fname.decode("utf-8"))
            excel = EnsureDispatch('Excel.Application')
            wb = excel.Workbooks.Open(fname)
            excel.DisplayAlerts = False
            wb.SaveAs(fname[:-1], FileFormat=56)    #FileFormat = 51 is for .xlsx extension
            wb.Close()                              #FileFormat = 56 is for .xls extension
            excel.Application.Quit()
            excel.DisplayAlerts = True
            inp_dict[".xls"].append(fname[:-1])
            inp_dict["del"].append(fname[:-1])
    return inp_dict

def delete_xls(inp_dict):
    """
    Delete files from del list
    :param inp_dict: Dict with separated paths
    """
    if inp_dict["del"]:
        for del_f in inp_dict["del"]:
            os.remove(os.path.abspath(del_f))

def process_files(geodata_name, inp_dict):
    """
    Process Excel files from .xls list, if the search request is valid, then writes the geo data in the copy file.
    :param geodata_name: Name of the file with geo data
    :param input_paths: Dict with separated paths
    :return: Updated dict with separated paths
    """
    input_paths = inp_dict[".xls"][:]
    try:
        data = geodata(geodata_name)
    except UnicodeDecodeError:
        showerror("Ошибка кодирования", "Файл данных должен быть закодирован в utf-8")
        data = geodata(askopenfilenames(initialdir=os.path.abspath(os.getcwd()), filetypes=[("Файл данных txt", ".txt")], title="Выберите файл данных txt")[0])


    for book in input_paths:
        book_flag = False
        with open_workbook(book, formatting_info=True) as rb:
            header = False
            wb = copy(rb)
        for numb, sheet in enumerate(rb.sheets()):
            column = "False"
            for row in range(sheet.nrows):
                if column != "False":
                    for data_row in data:
                        if sheet.cell(row, column).value == data_row[0]:
                            sheet_wb = wb.get_sheet(numb)
                            sheet_wb.write(row, sheet.ncols, data_row[1])
                            sheet_wb.write(row, sheet.ncols+1, data_row[2])
                            break
                else:
                    for col in range(sheet.ncols):
                        for data_row in data:
                            if sheet.cell(row, col).value == data_row[0]:
                                column = col
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
            f_out = get_output_name(book)
            wb.save(f_out)
            inp_dict["del"].append(f_out)
            inp_dict["out"].append(f_out)
    return inp_dict

def main():
    geodata_name = "data.txt"
    inp_dict = separate(get_input_name())
    inp_dict = check_for_csv(inp_dict)
    inp_dict = convert_xlsx_to_xls(inp_dict)
    inp_dict = process_files(geodata_name, inp_dict)
    inp_dict = convert_xls_to_xlsx(inp_dict)
    delete_xls(inp_dict)



if __name__ == '__main__':
    main()
