import xlwings as xw
import pandas as pd
import tkinter as tk
from tkinter import filedialog

# Importation


def get_file():
    file_path = filedialog.askopenfilename(title="Select File To Import")
    data = xw.books.open(file_path)
    return data

def import_report():
    checklists = [4, 7, 10, 13]
    sheets_description = ['D4:D127', 'D4:D317', 'D3:D151', 'D3:D73']
    data = get_file()
    wb = xw.Book.caller()
    date = report_date(data)
    for i in range(len(checklists)):
        sheet_index = checklists[i]
        range_sheet = sheets_description[i]
        if wb.sheets[i].range('A1').value == None :
            names = initial_import(data, sheet_index)
            wb.sheets[i].range('A1').value = names
        values = import_sheet(data, sheet_index, range_sheet)
        write_sheet(wb, i, values, date)


def import_sheet(data, n_sheet, range):
    values = read_sheet(data, n_sheet, range)
    return values


def read_sheet(data, n_sheet, range):
    sheet = data.sheets[n_sheet]
    values = sheet.range(range).options(pd.DataFrame).value
    return values


def report_date(data):
    date = data.sheets[0].range('H6').value
    return date


def initial_import(data, n_sheet):
    sheet = data.sheets[n_sheet]
    last_line = sheet.range('A4').end('down').row
    values = sheet.range('A3:C'+str(last_line)).options(pd.DataFrame).value
    return values


def write_sheet(wb, n_sheet, values, date):
    if wb.sheets[n_sheet].range('D1').value  == None :
        new_column = 4
    if wb.sheets[n_sheet].range('D1').value  != None :
        last_column = wb.sheets[n_sheet].range('D1').end('right').column
        if last_column == 16384:
            last_column = 4
        new_column = last_column + 1
    wb.sheets[n_sheet].range(2, new_column).value = values
    wb.sheets[n_sheet].range(1, new_column).value = date
