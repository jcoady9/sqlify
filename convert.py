import openpyxl

import sqlite3

workbook = openpyxl.load_workbook('test.xlsx')
worksheet = workbook.active

#TODO: CREATE
def create_db():
    conn = sqlite3.connect('test_db.sqlite3')
    cursor = conn.cursor()

    create_cmd = 'CREATE TABLE IF NOT EXISTS %s (' % 'table_name'

    for col_num in range(1, worksheet.max_column + 1):
        header = worksheet.cell(row=1,column=col_num)
        cell = worksheet.cell(row=2,column=col_num)

        if cell.data_type == cell.TYPE_NUMERIC:
            #TODO: INTEGER COLUMN
            create_cmd += '%s INTEGER' % header.value.strip().replace(' ','_')
            #pass
        elif cell.data_type == cell.TYPE_STRING:
            #TODO: TEXT COLUMN
            create_cmd += '%s TEXT' % header.value.strip().replace(' ','_')
            #pass
        elif cell.is_date:
            #TODO: DATE COLUMN (ISO8601 Text)
            pass
        else:
            #TODO: ERROR
            pass
        if col_num < worksheet.max_column:
            create_cmd += ','
    #creat_cmd = create_cmd.rstrip()
    create_cmd += ')'
    print(create_cmd)

    cursor.execute(create_cmd)
    conn.commit()
    conn.close()

#TODO: INSERT

#TODO: UPDATE

create_db()
