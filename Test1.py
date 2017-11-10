__author__ = '569882'

# import openpyxl
from openpyxl.cell import get_column_letter, column_index_from_string
c = Sheet['B1']

# for i in range(1,8,1):
# print(i ,Sheet.cell(row=i , column=2).value)
row_length = Sheet.max_row
col_length = Sheet.max_column
print(col_length)
#import openpyxl

# print("Writing selected result")

# wb1 = openpyxl.Workbook()  # creates a new workbo
# wb = openpyxl.load_workbook('incident.xlsx')  # Load the incident xlsx
# Sheet = wb.get_sheet_by_name('Page 1')  # pulls the sheet with name as passed in arguments
# row_length = Sheet.max_row
# col_length = Sheet.max_column
s = wb1.active
s.title ='Main'
# wb1.create_sheet("Tasks")
# sheet = wb1.get_active_sheet()
# sheet.title = 'Incidents'
# for i in range(1, row_length+1):
    # for j in range(1, 11):
            # x = Sheet.cell(row=i, column=j).value
                    # sheet.cell(row=i, column=j).value = x
                    sheet['A1'] = 'NEERAJ'
# sheet['D1'] = 'Application'
# wb1.save('FOOBAR.xlsx')
# print("Done")

