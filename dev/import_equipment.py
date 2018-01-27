# -*- coding: utf-8 -*-
"""
Created on Wed Jan 24 22:09:44 2018

@author: grlurton
"""
from simplified_data_collection import *

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
            
if __name__ == "__main__":
    # Used for frozen executable
    import_equipment()