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
    # project root directory, filename in.xlsx
    in_workbook = load_workbook('in.xlsx')
    # ['Sheet 1']
    print(in_workbook.get_sheet_names())
    in_sheet = in_workbook.active

    out_workbook = Workbook()
    out_filename = 'out.xlsx'
    out_sheet = out_workbook.active
    out_sheet.title = "my_sheet"

    # http://stackoverflow.com/questions/37440855/how-do-i-iterate-through-cells-in-a-specific-column-using-openpyxl-1-6
    first_non_header_row = 3
    in_column = 2
    out_column = 1

    for row in range(first_non_header_row, in_sheet.max_row + 1):
        in_word = in_sheet.cell(row=row, column=in_column).value
        out_word = in_word + 'foo'
        print(in_word + ', ' + out_word)
        # out_sheet set cell value
        _ = out_sheet.cell(column=out_column, row=row, value=out_word)

    out_workbook.save(filename = out_filename)
