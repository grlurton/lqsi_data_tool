import xlwings as xw
import pandas as pd
from tkinter import filedialog

# Importation
def get_file():
    file_path = filedialog.askopenfilename(title="Select File To Import")
    data = xw.books.open(file_path)
    return data


def indicators_to_checklist():
    wb = xw.Book.caller()
    check_sheet = wb.sheets[15]
    check_sheet.range('A2:A1000').clear_contents()
    check_sheet.range('B2:B1000').clear_contents()
    check_sheet.range('D2:D1000').clear_contents()
    n_phase = 1
    for n_sheet in [2, 5, 8, 11]:
        n_cl = 1
        sheet = wb.sheets[n_sheet]
        for line in range(3, sheet.api.UsedRange.Rows.Count):
            new_line = import_line(sheet, line)
            write_line(check_sheet, new_line , n_phase , n_cl)
            if type(new_line[1]) is list :
                n_cl = n_cl + len(new_line[1])
            if type(new_line[1]) is str :
                n_cl = n_cl + 1
        n_phase = n_phase + 1


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


def import_equipment():
    data = get_file()
    wb = xw.Book.caller()
    sheet_out = wb.sheets[4]
    sheet = data.sheets[0]
    max_lines = max(81, sheet.api.UsedRange.Rows.Count)
    equipments = {}
    for index in range(7, max_lines):
        line = sheet.range((index, 2), (index, 16)).value
        if len(set(line)) > 1:
            equipments[index] = line
        table_equipments = pd.DataFrame(equipments)
        out = table_equipments.count(1) / len(table_equipments.columns)
        for i in range(0, len(table_equipments)):
            sheet_out.range((88+i, 5)).value = out[i]


'CL' + str([1,2,3])
