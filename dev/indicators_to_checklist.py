# -*- coding: utf-8 -*-
"""
Created on Fri Jan 26 10:45:32 2018

@author: grlurton
"""

import xlwings as xw
from simplified_data_collection import *

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
        
        
if __name__ == "__main__":
    # Used for frozen executable
    indicators_to_checklist()