#!/usr/bin/env python3

from excely import excely

# in_filename = 'data/input/in.xlsx'
# out_filename = 'data/out.xlsx'
# excely.read_in_write_out(in_filename, out_filename)

spellings_file_name = 'data/input/spellings.xlsx'
misspelled_file_name = 'data/misspelled.xlsx'

excely.write_spelling_matches_to_file(spellings_file_name, misspelled_file_name)
