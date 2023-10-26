from openpyxl import load_workbook, Workbook
import os
file_names = os.listdir('data-input')
rows_to_delete = []
for i in file_names:
    if i != '.gitignore':
        wb = load_workbook(filename='data-input/'+i)
        ws = wb.active
        for row in ws.iter_rows():
            start_timestamp = row[4].value.split(":")[1]
            end_timestamp = row[5].value.split(":")[1]
            if start_timestamp == end_timestamp:
                rows_to_delete.append(row[0].row)
        for row_num in rows_to_delete:
            ws.delete_rows(row_num)
        wb.save('data-output/'+i)


