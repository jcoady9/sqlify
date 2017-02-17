import openpyxl

import sqlite

workbook = openpyxl.load('file.xlsx')
worksheet = workbook.active

#TODO: CREATE
for col_num in range(1, worksheet.max_column + 1):
    header = worksheet.cell(row=1,column=col_num)

    if header.data_type == cell.TYPE_NUMERIC:
        #TODO: INTEGER COLUMN
        pass
    elif header.data_type == cell.TYPE_STRING:
        #TODO: TEXT COLUMN
        pass
    elif header.is_date:
        #TODO: DATE COLUMN (ISO8601 Text)
        pass
    else:
        #TODO: ERROR

#TODO: INSERT

#TODO: UPDATE

