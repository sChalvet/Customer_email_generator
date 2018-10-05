# Created by: Samuel Chalvet
# Email: samuelchalvet@gmail.com
# On: 10/05/2018

import xlrd
from datetime import datetime

PROJECT_FILE_PATH = '../Services+%7C+CORE%3A+Projects+-+Majime.xlsx'
HOURS_FILE_PATH = '../Services+%7C+CORE%3A+Hours+-+Majime.xlsx'
SHEET_NAME = 'Sheet1'

# Column numbers from HOURS_FILE_PATH workbook
COL_ENTRY_DATE = 31
COL_HOURS = 32
COL_PROJECT_ID = 33
COL_PROJECT_LEAD_EMAIL = 49
COL_LAST_UPDATE_DATE = 57


def get_hours_workbook():
    try:
        workbook = xlrd.open_workbook(HOURS_FILE_PATH)
    except FileNotFoundError:
        raise Warning("Make sure the PROJECT HOURS file is in the same directory.\n They should be: "
                      "Services+%7C+CORE%3A+Projects+-+Majime and Services+%7C+CORE%3A+Hours+-+Majime")
        exit()
    return workbook


def get_hours_worked(start_date, end_date, project_id):
    workbook = get_hours_workbook()
    worksheet = workbook.sheet_by_name(SHEET_NAME)
    total_hours = 0

    for x in range(1, worksheet.nrows):
        date_obj = xlrd.xldate_as_datetime(worksheet.cell(x, COL_ENTRY_DATE).value, workbook.datemode).date()
        if worksheet.cell(x, COL_PROJECT_ID).value == int(project_id) and end_date >= date_obj >= start_date:
            total_hours += float(worksheet.cell(x, COL_HOURS).value)

    return total_hours


def get_lead_email(project_id):
    workbook = get_hours_workbook()
    worksheet = workbook.sheet_by_name(SHEET_NAME)

    lead_email = "noEmailFound"

    for x in range(1, worksheet.nrows):
        if worksheet.cell(x, COL_PROJECT_ID).value == int(project_id):
            lead_email = worksheet.cell(x, COL_PROJECT_LEAD_EMAIL).value
            break

    return lead_email


def comment_exist(start_date, end_date, last_comment_date, datemode):

    # if the cell isn't empty and within the date range return True
    if last_comment_date != '' and last_comment_date is not None and \
            end_date >= xlrd.xldate_as_datetime(last_comment_date, datemode).date() >= start_date:
        return True
    else:
        return False

