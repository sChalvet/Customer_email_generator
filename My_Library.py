# Created by: Samuel Chalvet
# Email: samuelchalvet@gmail.com
# On: 10/05/2018

import xlrd
import Config
from datetime import datetime

columns_dict = {}


def get_hours_workbook():
    try:
        workbook = xlrd.open_workbook(Config.HOURS_FILE_PATH)
        worksheet = workbook.sheet_by_name(Config.SHEET_NAME)
    except FileNotFoundError:
        raise Warning("Make sure the PROJECT HOURS file is in the same directory.\n They should be: "
                      "Services+%7C+CORE%3A+Projects+-+Majime and Services+%7C+CORE%3A+Hours+-+Majime")
        exit()

    # Column index and names from HOURS_FILE_PATH workbook
    for col_num in range(0, worksheet.ncols):
        columns_dict[worksheet.cell(0, int(col_num)).value] = col_num

    return workbook


def get_hours_worked(start_date, end_date, project_id):
    workbook = get_hours_workbook()
    worksheet = workbook.sheet_by_name(Config.SHEET_NAME)
    total_hours = 0

    for x in range(1, worksheet.nrows):
        date_obj = datetime.strptime(worksheet.cell(x, columns_dict["Entry Date"]).value, '%Y-%m-%d').date()
        if worksheet.cell(x, columns_dict["Project ID"]).value == int(project_id) and end_date >= date_obj >= start_date:
            total_hours += float(worksheet.cell(x, columns_dict["Hours"]).value)

    return total_hours


def get_lead_email(project_id):
    workbook = get_hours_workbook()
    worksheet = workbook.sheet_by_name(Config.SHEET_NAME)

    lead_email = "noEmailFound"

    for x in range(1, worksheet.nrows):
        if worksheet.cell(x, columns_dict["Project ID"]).value == int(project_id):
            lead_email = worksheet.cell(x, columns_dict["Project Lead Email"]).value
            break

    return lead_email


def comment_exist(start_date, end_date, last_comment_date, datemode):

    # if the cell isn't empty and within the date range return True
    if last_comment_date != '' and last_comment_date is not None and \
            end_date >= datetime.strptime(last_comment_date, '%Y-%m-%d  %H:%M:%S').date() >= start_date:
        return True
    else:
        return False

