# -*- coding: UTF-8 -*-
import xlsxwriter
import os
import csv
import xml.etree.ElementTree as ET
from xml.etree.ElementTree import ElementTree
from Sheet import Sheet

class Output(xlsxwriter.Workbook):
    TableOfContent = ""
    one_sheet_ws = ""
    sheet_count = 0
    output_excel = ""
    title_row_num = ""
    many_sheets = ""
    n23_background = ""
    blued_style = ""
    blued_style_pc = ""
    hyperlink = ""
    borders = ""
    tstat = ""
    bold = ""
    toc = ""
    toc_header = ""
    toc_hyperlink = ""
    banner = ""
    percentage = ""
    number = ""
    current_ws = ""

    def __init__(self, output_excel, title_row_num, many_sheets):
        super(self.__class__, self).__init__(output_excel, {'constant_memory': True, 'strings_to_numbers': True})
        if os.path.isfile("TableOfContent.txt"):
            os.remove("TableOfContent.txt")
        self.sheet_count = 1
        self.one_sheet_ws = None
        self.TableOfContent = open("TableOfContent.txt","ab")
        self.TableOfContent.write("<?xml  version='1.0' encoding='UTF-8'?>\n")
        self.TableOfContent.write("<tables>\n")
        self.add_styles()
        self.many_sheets = many_sheets
        self.current_row = 1
        if not many_sheets:
            # self.one_sheet_ws = self.add_worksheet("Tables")
            self.one_sheet_ws = Sheet(self, "Tables")
            self.current_ws = self.one_sheet_ws
            
    def add_styles(self):
        # Formatting styles
        self.n23_background = self.add_format({ 'bg_color': "#99CCFF", 'border': 1 })
        self.blued_style = self.add_format({'font_name': 'Calibri', 'font_size': 11, 'bold': True, 'font_color': 'white',
                            'align': 'center', 'valign': 'vcenter', 'bg_color': "#376091", 'border': 1, 'text_wrap': True})
        self.hyperlink = self.add_format({ 'underline': 'single', 'font_color': '#0000EE' })
        self.borders = self.add_format({ 'border': 1, 'align': 'left', 'valign': 'vcenter'})
        self.tstat = self.add_format({ 'bg_color':  "#FFFF99", 'border': 1, 'align': 'center' })
        self.bold = self.add_format({ 'bold': True })
        self.toc = self.add_format({ 'text_wrap': True, 'border': 1 })
        self.toc_header = self.add_format({ 'font_size': 14, 'bold': True, 'align': 'center', 'valign': 'vcenter'})
        self.toc_hyperlink = self.add_format({ 'underline': 'single', 'font_color': '#0000EE', 'text_wrap': True, 'border': 1 })
        self.banner = self.add_format({ 'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'border': 1})
        self.percentage = self.add_format({ 'border': 1, 'align': 'center', 'valign': 'vcenter',
                                   'num_format': '0%'})
        self.blued_style_pc = self.add_format({'font_name': 'Calibri', 'font_size': 11, 'bold': True, 'font_color': 'white',
                            'align': 'center', 'valign': 'vcenter', 'bg_color': "#376091", 'border': 1, 'text_wrap': True,
                                               'num_format': '0%'})
        self.number = self.add_format({ 'border': 1, 'align': 'center', 'valign': 'vcenter'})

    def add_toc(self):
        # ADDING TOC
        i = 3
        output_ws = self.add_worksheet("Contents")

        output_ws.write(0, 1, "TABLE OF CONTENTS", self.toc_header)
        output_ws.set_row(0, 75)

        output_ws.write(2, 0, "Sheet", self.blued_style)
        output_ws.write(2, 1, "Question label", self.blued_style)
        output_ws.write(2, 2, "Base text", self.blued_style)
        output_ws.write(2, 3, "Base", self.blued_style)
        output_ws.set_row(2, 30)

        tree = ET.parse('TableOfContent.txt')
        root = tree.getroot()

        for table in root.iter('table'):
            x = 0
            output_ws.write(i, 0, table.find('sheet_name').text, self.toc)
            link = "#" + table.find('sheet_name').text + "!A" + table.find('row_start').text
            output_ws.write_url(i, 1, link, self.toc_hyperlink, table.find('name').text)
            output_ws.write(i, 2, table.find('b_text').text, self.toc)
            output_ws.write(i, 3, table.find('total').text, self.toc)
            i += 1

        # FORMATTING OF TOC
        output_ws.set_column(0, 0, 30)
        output_ws.set_column(1, 1, 100)
        output_ws.set_column(2, 2, 40)
        output_ws.set_column(3, 3, 10)

        def worksheet_compare_sort(x, y):
            if x == "Contents":
                return -1
            if y == "Contents":
                return 1
            return 0

        self.worksheets_objs.sort(cmp=worksheet_compare_sort, key=lambda x: x.name)

    def increment_sheet_count(self):
        self.sheet_count += 1

    def get_sheet_count(self):
        return self.sheet_count

    def close_toc(self):
        self.TableOfContent.write("\n</tables>")
        self.TableOfContent.close()
        
    def set_current_ws(self, ws):
        self.current_ws = ws
        
    def get_current_ws(self):
        return self.current_ws


class BaseText:
    def __init__(self, table, row):
        self.table = table
        self.row = row
        table.set_btext_obj(self)  
        
    def process(self):
        btext = self.row[0]
        self.table.update_base_text_row()
        self.table.set__base_text(btext)
        self.table.print_bold(self.row)
        self.table.increment_current_row()
        if len(self.row) >= 2 and len(self.row[1]) > 0:
            self.table.set__total(self.row[1])


class Total:
    def __init__(self, table, row):
        self.table = table
        self.row = row
        table.set_total_obj(self)      
        self.row_pos = table.get_data_rows()

    def get_total_row_position(self):
        return self.row_pos
    
    def process(self):
        self.table.increment_current_row()
        self.table.set__total(self.row[1])
        self.table.print_total_row(self.row)


class TableName:
    def __init__(self, table, row):
        self.table = table
        self.ttext = row[0]
        self.row = row
        table.set_tableName_obj(self)  
    
    def process(self):
        self.table.print_bold(self.row)
        self.table.increment_current_row()
        self.table.set__table_name(self.ttext)

    def get_row(self):
        return self.row
