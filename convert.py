import openpyxl

import sqlite3

def create(db_name, spreadsheet_file):
    workbook = openpyxl.load_workbook(spreadsheet_file)
    worksheet = workbook.active

    conn = sqlite3.connect(db_name)
    cursor = conn.cursor()

    create_cmd = 'CREATE TABLE IF NOT EXISTS {} '.format('table_name')

    columns = list()
    for col_num in range(1, worksheet.max_column + 1):
        header = worksheet.cell(row=1,column=col_num)
        col_name = header.value.strip().replace(' ','_')
        cell = worksheet.cell(row=2,column=col_num)

        if cell.data_type == cell.TYPE_NUMERIC:
            columns.append('{} INTEGER'.format(col_name))
        elif cell.data_type == cell.TYPE_STRING:
            columns.append('{} TEXT'.format(col_name))
        elif cell.is_date:
            #TODO: DATE COLUMN (ISO8601 Text)
            raise NotImplementedError('Dates are not yet supported.')
        else:
            raise NotImplementedError('Datatype not supported.')

    create_cmd += '({})'.format(','.join(col for col in columns))
    print(create_cmd)

    cursor.execute(create_cmd)
    conn.commit()

    #Populate database
    rows = list()
    for row_num in range(2, worksheet.max_row + 1):
        row = list()
        for col_num in range(1, worksheet.max_column + 1):
            cell = worksheet.cell(row=row_num,column=col_num)
            row.append(cell.value)
        print(row)
        rows.append(tuple(row))

    insert_cmd = 'INSERT INTO {} VALUES ({})'.format('table_name', ','.join(['?' for x in range(0,worksheet.max_column)]))

    cursor.executemany(insert_cmd, rows)
    conn.commit()

    conn.close()

def insert(db_name, spreadsheet_file):
    conn = sqlite3.connect(db_name)
    cursor = conn.cursor()
    workbook = openpyxl.load_workbook(spreadsheet_file)
    worksheet = workbook.active

    rows = list()
    for row_num in range(2, worksheet.max_row + 1):
        row = list()
        for col_num in range(1, worksheet.max_column + 1):
            cell = worksheet.cell(row=row_num,column=col_num)
            row.append(cell.value)
        print(row)
        rows.append(tuple(row))

    insert_cmd = 'INSERT INTO {} VALUES ({})'.format('table_name', ','.join(['?' for x in range(0,worksheet.max_column)]))

    cursor.executemany(insert_cmd, rows)
    conn.commit()
    conn.close()

#TODO: UPDATE

#TODO: MAIN

create('test_db.sqlite3', 'test.xlsx')
insert('test_db.sqlite3', 'test2.xlsx')
