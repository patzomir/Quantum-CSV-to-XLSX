# -*- coding: UTF-8 -*-
import xlsxwriter
import os
import csv
import xml.etree.ElementTree as ET
from xml.etree.ElementTree import ElementTree

class Output(xlsxwriter.Workbook):
    TableOfContent = ""
    one_sheet_ws = ""
    sheet_count = 0
    current_row = ""
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
            self.one_sheet_ws = self.add_worksheet("Tables")

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

        output_ws.write(2, 0, "Question", self.blued_style)
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

    def get_current_row(self):
        return self.current_row

    def add_to_current_row(self, offset):
        if offset > 0:
            self.current_row = offset


class BaseText:
    def __init__(self, table, row):
        self.__table = table
        btext = row[0]
        table.update_base_text_row()
        table.set__base_text(btext)
        table.print_bold(row)
        table.increment_current_row()
        if len(row) >= 2 and len(row[1]) > 0:
            table.set__total(row[1])


class Total:
    def __init__(self, table, row):
        self.__table = table
        table.set__total(row[1])
        table.print_total_row(row)


class TableName:
    def __init__(self, table, row):
        self.__table = table
        ttext = row[0]
        table.print_bold(row)
        table.increment_current_row()
        table.set__table_name(ttext)
