# Sqlify

Turn your spreadsheet into a SQLite database.

## Install

```

git clone https://github.com/jcoady9/sqlify.git

cd sqlify

pip install -r requirements.txt

```

## Usage

```
usage: sqlify.py [-h] [--db DB] file

A simple python script to create a sqlite database from a xlsx spreadsheet
file.

positional arguments:
  file        Path of the Excel file.

optional arguments:
  -h, --help  show this help message and exit
  --db DB     Path to the SQLite database. If no path is provided, a new
              database named [file] will be created.

```

By default, sqlify will create a brand new SQLite database file with the same name as your spreadsheet. If you want to create the new table with an existing SQLite database use the '--db' argument.

```
# This will create a new SQLite database file called 'myspreadsheet.sqlite3'

python sqlify.py myspreadsheet.xlsx


# Use this if you already have a SQLite database created or you do not want the database
# to have the same name as your spreadsheet.

python sqlify.py --db my_db.sqlite3 myspreadsheet.xlsx
``` 

## License

This project is licensed under the MIT License. See LICENSE for details.
