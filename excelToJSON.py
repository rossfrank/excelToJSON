import xlrd
import csv
import json

#input file
input_file = ''

#output file
output_file = 'python_out.txt'

def col_num_string(n):
    div = n
    string = ""
    while div > 0:
        module = (div-1) % 26
        string = chr(65+module) + string
        div = int((div-module) / 26)
    return string

result = {}
result[input_file] = {}

if input_file.endswith('.csv'):
    f = open(input_file, 'rb')
    csv_file = csv.reader(f)
    array = []
    for row in csv_file:
        array.append(row)
    result[input_file]['csv'] = array

elif input_file.endswith('.xls') or input_file.endswith('.xlsx'):
    book = xlrd.open_workbook(filename=input_file)
    sheets = book.sheet_names()
    result[input_file]['size'] = {}
    for sheet_name in sheets:
        sheet = book.sheet_by_name(sheet_name)
        ar = []
        row_cnt = 0
        col_cnt = 0
        for row in sheet.get_rows():
            ar.append(row)
            temp_cnt = 0
            for cell in row:
                if cell.ctype == 0:
                    temp_cnt += 1
                    continue
                elif col_cnt is 0:
                    col_cnt = temp_cnt
                else:
                    col_cnt = min(col_cnt, temp_cnt)
                break
        for row in ar:
            if all(item.ctype is 0 for item in row):
                row_cnt += 1
                ar.pop(0)
            else:
                break
        for row in ar:
            for x in range(col_cnt):
                row.pop(0)
        size = col_num_string(col_cnt + 1) + str(row_cnt + 1) + ":" + col_num_string(len(ar[0]) + col_cnt) + str(len(ar) + row_cnt)
        result[input_file]['size'][sheet_name] = size
        for x in range(len(ar)):
            for y in range(len(ar[x])):
                cell = ar[x][y]
                if cell.ctype is 5:
                    ar[x][y] = xlrd.biffh.error_text_from_code[cell.value]
                else:
                    ar[x][y] = cell.value
        result[input_file][sheet_name] = ar
else:
    print "error, not a csv or excel file"
    import sys
    sys.exit()

with open(output_file, 'w') as outfile:
    json.dump(result, outfile)
