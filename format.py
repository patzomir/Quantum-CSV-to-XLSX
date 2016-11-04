# -*- coding: UTF-8 -*-
from datetime import datetime
import csv
import sys
import format_lib as fl
import Table as table
import io
from Sheet import Sheet

t1 = datetime.now()

# USER CHANGED VARIABLES
title_row_num = 3
input_document = 'Amber_11-03-16_uwt_v4.csv'
many_sheets = True
output_excel = 'formatted.xlsx'
decoding = 'cp1251'
#####

# CHANGED VIA IN-LINE ARGUMENTS
if len(sys.argv) > 1: input_document = sys.argv[1]
if len(sys.argv) > 2: output_excel = sys.argv[2]
if len(sys.argv) > 3: title_row_num = int(sys.argv[3])
if len(sys.argv) > 4:
    try:
        if int(sys.argv[4]) < 1:
            many_sheets = False
    except ValueError as verr:
        print "The argument for many sheets should be integer! The conversion will continue by default: to many sheets."
        many_sheets = True


# create output object
out = fl.Output(output_excel, title_row_num, many_sheets)


def decode_from_csv(row):
    out = []
    for x in row:
        try:
            temp=x.decode(decoding, 'strong')
        except:
            # temp = remove_odd_chars(x)
            print "Failed decoding! Weird symbols will be ignored for: " + x
            temp=x.decode(decoding, 'ignore')
        out.append(temp)
    return out


# out.set_current_ws(Sheet(out, "Tables"))

# READING TABLES FROM CSV
row_count = 1
frow = True
f = io.open(input_document, 'rb')
reader = csv.reader(f)
for utf8_row in reader:
    row = decode_from_csv(utf8_row)
    if len(row) == 0: continue
    if row[0].find('Proportions/Means') == 0: continue
    if row[0].find('_________') == 0:         continue
    if row[0].find('* small base') == 0:      continue
    if row[0].find('** very small') == 0:     continue
    if row[0].find('Overlap formulae used.') == 0:     continue
    if row[0] == "#page":
        if not frow:
            tab.loop_recorded_rows()
            tab.print_footer()
            if many_sheets: tab.close_file()
            tab.append_to_table_of_content()
            if not many_sheets: out.get_current_ws().add_to_current_row(2 + tab.get_current_row())
        tab = table.Table(out, out.get_sheet_count())
        out.increment_sheet_count()
        frow = False
        row_count = 0
    elif row_count == title_row_num:
        tname = fl.TableName(tab, row)
    elif (row[0].find("Base ") >= 0 or
          row[0].find("Base:") >= 0 or
          row[0].find("Base-") >= 0) and (len(row) == 1 or len(row[1]) == 0):
        base_text = fl.BaseText(tab, row)
    elif (row[0].find("Total") == 0 or row[0].find("Base") == 0
          or row[0].find("Weighted") == 0) and len(row) > 1:
        total = fl.Total(tab, row)
    else:
        tab.fill_data(row)
    row_count += 1
tab.loop_recorded_rows()
tab.print_footer()
tab.append_to_table_of_content()
f.close()

out.close_toc()
out.add_toc()
out.close()

print (datetime.now() - t1)
print ("FINISH")

