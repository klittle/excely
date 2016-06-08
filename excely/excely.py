#!/usr/bin/env python

import sys, os

# read
from openpyxl import load_workbook
# write
from openpyxl import Workbook

# http://openpyxl.readthedocs.io/en/default/usage.html#write-a-workbook
from openpyxl.compat import range
from openpyxl.cell import get_column_letter

#sys.path.append(os.path.abspath(os.path.join('..', 'excely')))
#from excely import blah

# http://stackoverflow.com/questions/419163/what-does-if-name-main-do
if __name__ == '__main__':

    """
    Open excel file and get sheet names
    """
    # this works with file in.xlsx in root directory
    in_workbook = load_workbook('in.xlsx')
    # ['Sheet 1']
    print(in_workbook.get_sheet_names())
    in_worksheet = in_workbook.active
    # x
    print(in_worksheet['b3'].value)
    # 88
    print(in_worksheet['c3'].value)


    out_workbook = Workbook()
    dest_filename = 'out.xlsx'

    out_worksheet = out_workbook.active
    out_worksheet.title = "range names"

    for row in range(1, 40):
        out_worksheet.append(range(600))

    ws2 = out_workbook.create_sheet(title="Pi")
    ws2['F5'] = 3.14
    ws3 = out_workbook.create_sheet(title="Data")
    for row in range(10, 20):
        for col in range(27, 54):
            _ = ws3.cell(column=col, row=row, value="%s" % get_column_letter(col))
    print(ws3['AA10'].value)
    out_workbook.save(filename = dest_filename)
