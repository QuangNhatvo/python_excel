#!/usr/bin/env python3

import xlsxwriter
import openpyxl

input_file = 'filtered_data.xlsx'
workbook = openpyxl.load_workbook(input_file)
sheet = workbook.active

def format_name (name):
    if isinstance(name,str):
        return ' '.join(word.capitalize() for word in name.split())
    return name

for row in sheet.iter_rows(min_row=2, min_col=2, max_col=2):
    for cell in row:
        cell.value = format_name(cell.value)

output_file = 'formated.xlsx'
workbook.save(output_file)
