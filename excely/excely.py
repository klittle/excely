#!/usr/bin/env python

import sys, os
from openpyxl import load_workbook

#sys.path.append(os.path.abspath(os.path.join('..', 'excely')))
#from excely import blah

# http://stackoverflow.com/questions/419163/what-does-if-name-main-do
if __name__ == '__main__':

    """
    Open excel file and get sheet names
    """
    # this works with in.xls in root directory
    in_workbook = load_workbook('in.xlsx')
    # ['Sheet 1']
    print(in_workbook.get_sheet_names())
    in_worksheet = in_workbook.active
    # x
    print(in_worksheet['b3'].value)
    # 88
    print(in_worksheet['c3'].value)
