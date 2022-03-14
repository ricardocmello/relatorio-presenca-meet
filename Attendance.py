import csv
import os

TITLE, DATE, TIME, TOTAL, START = 0, 1, 2, 3, 4

def append_value(dict_obj, key, value): # from https://thispointer.com/python-how-to-add-append-key-value-pairs-in-dictionary-using-dict-update/
    if key in dict_obj:
        if not isinstance(dict_obj[key], list):
            dict_obj[key] = [dict_obj[key]]
        dict_obj[key].append(value)
    else:
        dict_obj[key] = value

def get_attendance_dict():
    attendance_dict = {}
    files = [f for f in os.listdir('.') if os.path.isfile(f) and f.endswith(".csv")]
    for f in files:
        with open(f, newline='') as csvfile:
            day_attendance = csv.reader(csvfile, delimiter=';', quotechar='|')
            header = list(day_attendance)
            if header[DATE][0].strip() == 'Data': date = header[DATE][1]
            if header[TOTAL][0].strip() == 'Confirmados': confirmados = header[TOTAL][1]
            for row in header[START:]:
                student = row[0]
                append_value(attendance_dict, student, date)   
    return attendance_dict

# Dict relating names and class attendance
attendance = get_attendance_dict()

# List of names
sortednames = sorted(attendance.keys(), key=lambda x:x.lower()) #https://stackoverflow.com/questions/24728933/sort-dictionary-alphabetically-when-the-key-is-a-string-name

# List of dates
dates = list(sorted({ele for val in attendance.values() for ele in val if type(val)==type([])}))

import xlsxwriter

workbook = xlsxwriter.Workbook('Attendance.xlsx')
worksheet = workbook.add_worksheet()

present = workbook.add_format()
present.set_bg_color('#C6EFCE')
absent = workbook.add_format()
absent.set_bg_color('#FFC7CE')

row, col = 0, 1
for date in dates:
    worksheet.write(row,col,date)
    col += 1

row, col = 1, 0
for name in sortednames:
    worksheet.write(row, col,name)
    for date in dates:
        col += 1
        if date in attendance[name]: 
            worksheet.write(row, col, '.', present)
        else:
            worksheet.write(row, col, 'F', absent)
    col = 0
    row += 1

workbook.close()