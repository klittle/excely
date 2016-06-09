#!/usr/bin/env python

import sys, os

# read
from openpyxl import load_workbook
# write
from openpyxl import Workbook

from openpyxl.compat import range
from openpyxl.cell import get_column_letter

#sys.path.append(os.path.abspath(os.path.join('..', 'excely')))
#from excely import blah

# http://stackoverflow.com/questions/419163/what-does-if-name-main-do
if __name__ == '__main__':

    def spelled_word(in_word):
        """
        returns spelled_word
        """
        return in_word + 'foo'

    """
    Read from Excel file in.xlsx and write to out.xlsx
    """
    # input file in project root directory
    in_filename = 'in.xlsx'
    in_workbook = load_workbook(in_filename)
    # ['Sheet 1']
    print(in_workbook.sheetnames)
    in_sheet = in_workbook.active
    print('in_sheet name ', in_workbook.sheetnames[0])

    out_workbook = Workbook()
    out_sheet = out_workbook.active
    out_sheet.title = "my_sheet"
    # output file in project root directory
    out_filename = 'out.xlsx'

    # http://stackoverflow.com/questions/37440855/how-do-i-iterate-through-cells-in-a-specific-column-using-openpyxl-1-6
    first_non_header_row = 2
    # column to read in in_sheet
    in_column = 2
    # column to write in out_sheet
    out_column = 1

    for row in range(first_non_header_row, in_sheet.max_row + 1):
        in_word = in_sheet.cell(row=row, column=in_column).value
        out_word = spelled_word(in_word)
        print(in_word + ', ' + out_word)
        # out_sheet set cell value
        _ = out_sheet.cell(column=out_column, row=row, value=out_word)

    out_workbook.save(filename = out_filename)

