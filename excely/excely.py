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
    in_xlsx = load_workbook('in.xlsx')
    # ['Sheet 1']
    print(in_xlsx.get_sheet_names())
