#!/usr/bin/env python3

from excely import excely

#excely.read_in_write_out()

spellings_file_name = 'data/spellings.xlsx'
misspelled_file_name = 'data/misspelled.xlsx'

excely.write_spelling_matches_to_file(spellings_file_name, misspelled_file_name)
