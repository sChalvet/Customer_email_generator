import xlrd
from datetime import datetime

PROJECT_FILE_PATH = '../Core_Projects_Majime.xlsx'
HOURS_FILE_PATH = '../Core_Hours_Majime.xlsx'
SHEET_NAME = 'Sheet1'

# Column numbers from HOURS_FILE_PATH workbook
COL_ENTRY_DATE = 31
COL_HOURS = 32
COL_PROJECT_ID = 33
COL_PROJECT_LEAD_EMAIL = 49
COL_LAST_UPDATE_DATE = 57


def get_number_rows(file_path, column):
    workbook = xlrd.open_workbook(file_path)
    worksheet = workbook.sheet_by_name(SHEET_NAME)
    num_cells = 0
    for cell in worksheet.col(column):
        if cell.value is not None:
            num_cells += 1
    return num_cells


def get_column_data(column, numrows, filepath):
    workbook = xlrd.open_workbook(filepath)
    worksheet = workbook.sheet_by_name(SHEET_NAME)
    array = []
    for x in range(1, numrows):
        array.append(worksheet.cell(x, column).value)
    return array


def get_hours_worked(start_date, end_date, project_id):
    workbook = xlrd.open_workbook(HOURS_FILE_PATH)
    worksheet = workbook.sheet_by_name(SHEET_NAME)

    num_rows = get_number_rows(HOURS_FILE_PATH, COL_ENTRY_DATE)

    #print(num_rows)

# TODO there is a gap in IDs so we should ignore (pop() all of the entries with no hub ID)
    total_hours = 0

    for x in range(1, num_rows):
        date_obj = xlrd.xldate_as_datetime(worksheet.cell(x, COL_ENTRY_DATE).value, workbook.datemode)
        if worksheet.cell(x, COL_PROJECT_ID).value == int(project_id) and end_date >= date_obj >= start_date:
            total_hours += float(worksheet.cell(x, COL_HOURS).value)

    return total_hours


