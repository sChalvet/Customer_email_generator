import xlrd

FILE_PATH = '../Core_Projects_Majime.xlsx'

def get_column_data(colrum, numrows):
    workbook = xlrd.open_workbook(FILE_PATH)
    worksheet = workbook.sheet_by_name('Sheet1')
    array = []
    for x in range(numrows):
        array.append(worksheet.cell(x, colrum).value)
    return array
