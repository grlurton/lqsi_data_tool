import xlwings as xw
import pandas as pd
from tkinter import filedialog

# Importation
def get_file():
    file_path = filedialog.askopenfilename(title="Select File To Import")
    data = xw.books.open(file_path)
    return data

def import_line(sheet, line):
    check_id = 'A' + str(line)
    criterions_adress = 'D' + str(line)
    check = sheet.range(check_id).value
    criterions = sheet.range(criterions_adress).expand('right').value
    return (check, criterions)


def write_line(sheet, line, n_phase , n_cl):
    last_line = sheet.range('A1').end('down').row
    if last_line == 1048576:
        last_line = 1
    new_line = last_line + 1
    if type(line[1]) is list:
        n = len(line[1])
        sheet.range((new_line, 1), (last_line+len(line[1]), 1)).options(transpose=True).value = ['CL' + str(n_phase) + '.' + str(i) for i in list(range(n_cl , n_cl + n))]
        sheet.range((new_line, 2), (last_line+len(line[1]), 2)).value = line[0]
        sheet.range((new_line, 4)).options(transpose=True).value = line[1]
    if type(line[1]) is str:
        sheet.range((new_line, 1)).value = 'CL' + str(n_phase) + '.' + str(n_cl)
        sheet.range((new_line, 2)).value = line[0]
        sheet.range((new_line, 4)).options(transpose=True).value = line[1]
