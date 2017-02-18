import openpyxl

import sqlite

workbook = openpyxl.load('file.xlsx')
worksheet = workbook.active

#TODO: CREATE

conn = sqlite3.connect('db.sqlite3')
cursor = conn.cursor()

create_cmd = 'CREATE TABLE IF NOT EXISTS %s' % 'table_name'

for col_num in range(1, worksheet.max_column + 1):
    header = worksheet.cell(row=1,column=col_num)

    if header.data_type == cell.TYPE_NUMERIC:
        #TODO: INTEGER COLUMN
        create_cmd += '%s INTEGER' % header[col_num]
        pass
    elif header.data_type == cell.TYPE_STRING:
        #TODO: TEXT COLUMN
        create_cmd += '%s TEXT' % header[col_num]
        pass
    elif header.is_date:
        #TODO: DATE COLUMN (ISO8601 Text)
        pass
    else:
        #TODO: ERROR
    create_cmd += ','

#TODO: INSERT

#TODO: UPDATE

