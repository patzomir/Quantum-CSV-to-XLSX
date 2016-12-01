# -*- coding: UTF-8 -*-
import re
import xml.etree.ElementTree as ET
from Sheet import Sheet
import format_lib

class Table:
    __current_row = 0
    __BaseText = ""
    __base_text_row = 0
    __Total = ""
    __empty_row = True
    __max_cell = 0
    __last_ban_row = False
    __row_start = 1
    __first_break = True
    __TableName = ""
    banner = True
    workbook = ""
    out = ""
    __out_ws = ""
    row_types = []
    large_row = ""
    current_row_type = ""
    footer = []
    data = []

    def __init__(self, output, sheet_count=1):
        self.out = output
        self.__sheet_count = sheet_count
        if self.out.many_sheets:
            self.__SheetName = "T{0}".format(sheet_count)
            self.__out_ws = Sheet(output, self.__SheetName)
        else:
            self.__SheetName = "Tables"
            self.__out_ws = self.out.get_current_ws()
        self.__out_ws.get_sheet().set_column(0,0,25)
        self.row_types = []
        for i in range(0,9):
            self.row_types.append(0)
        self.large_row = False
        self.data = []
        self.baseTextObj = format_lib.BaseText(self, [''])
        self.totalObj = format_lib.Total(self, [''])
        self.tableNameObj = format_lib.TableName(self, [''])
        self.__row_start = self.__out_ws.get_current_row()
        self.btxt = 0

    def write(self, *args):
        lst = list(args)
        if re.search('^[0-9.]+%$', str(lst[2].encode('utf8','replace'))):
            lst[2] = str(lst[2]).encode('utf8','replace').replace('%','').decode('utf8')
            lst[2] = float(lst[2])/100
            if not (len(lst) > 4 and lst[4] == "total row"):
                if len(lst) == 3:
                    lst.append(self.out.percentage)
                else:
                    lst[3] = self.out.percentage
            else:
                lst[3] = self.out.blued_style_pc
        elif re.search('^[0-9.]+$', str(lst[2].encode('utf8','replace'))) and self.current_row_type in (4,5):
            lst[3] = self.out.number
        elif lst[2].encode('utf8','replace') in [ "-", "- " ]:
            lst[3] = self.out.center
        if len(lst) > 4:
            del lst[4:]
        t = tuple(lst)
        self.__out_ws.write(*t)

    @staticmethod
    def wrap_write_to_xml(string):
        string = string.replace('&', '&amp;')
        string = string.replace('>', '&gt;')
        string = string.replace('<', '&lt;')
        return string

    def append_to_table_of_content(self):
        self.out.TableOfContent.write('<table>\n'
                                      + '<table_id>' + self.wrap_write_to_xml(str(self.__sheet_count)) + '</table_id>\n'
                                      + '<sheet_name>' + self.wrap_write_to_xml(self.__out_ws.get_sheetname().encode('utf8')) + '</sheet_name>\n'
                                      + '<name>' + self.wrap_write_to_xml(self.__TableName.encode('utf8')) + '</name>\n'
                                      + '<b_text>' + self.wrap_write_to_xml(self.__BaseText.encode('utf8')) + '</b_text>\n'
                                      + '<total>' + self.wrap_write_to_xml(self.__Total.encode('utf8')) + '</total>\n'
                                      + '<row_start>' + self.wrap_write_to_xml(str(self.__row_start + 1)) + '</row_start>\n'
                                      + '</table>\n')

    def set__table_name(self, table_name):
        self.__TableName = table_name

    def set__base_text(self, base_text):
        temp = base_text.replace("Base: ","")
        temp = temp.replace("Base - ","")
        self.__BaseText = temp

    def get__base_text(self):
        return self.__BaseText

    def set__total(self, total):
        if self.__Total == "":
            self.__Total = total
            #if (self.out.many_sheets): self.freeze_pane()

    def get__total(self):
        return self.__Total

    def freeze_pane(self):
        self.__out_ws.freeze_panes(self.__out_ws.get_current_row()+1, 2)

    def print_link_to_contents(self):
        link = "internal:#'Contents'!B" + str(self.out.get_table_number()+3)
        self.__out_ws.get_sheet().write_url(self.__out_ws.get_current_row(), 0, link,
                                            self.out.hyperlink,  "Table of content")
        self.__out_ws.add_to_current_row(1)

    def update_base_text_row(self):
        if self.__base_text_row == 0: self.__base_text_row = self.__out_ws.get_current_row()

    def n23(self, cell):
        self.write(self.__out_ws.get_current_row(), 0, str(cell), self.out.n23_background)
        for i in range(1, self.__max_cell+1):
            self.write(self.__out_ws.get_current_row(), i, None, self.out.n23_background)

    def get_row_type(self, row):
        # Returns the row type:
        # Row Type 0 = blank
        # Row Type 1 = first column only - Title
        # Row Type 2 = first column only - n23
        # Row Type 3 = not first column only - cross-break
        # Row Type 4 = not first column only - tstat
        # Row Type 5 = data on first and other columns
        # Row Type 6 = tstat letters in cross-break
        # Row Type 7 = sub-title
        # Row Type 8 = foot note
        rtype = 0
        temp = 0
        tstat_cb = True

        if row[0].find("bot:") == 0:
            return 8
        
        for cell in row:
            pattern = re.compile(ur'[^ ]', re.UNICODE)
            if len(cell) > 0 and pattern.search(cell):
                temp += 1
        if temp == 0:
            return 0
        if temp == 1:
            if len(row[0]) > 0:
                check = 0
                for i in range(2,6):
                    if self.row_types[i] > 0:
                        check += 1
                if check > 0:
                    return 2
                else:
                    if len(self.__TableName) > 0:
                        return 7
                    return 1
        if len(row[0]) == 0:
            if self.row_types[5] > 0:
                return 4
            else:
                for cell in row:
                    pattern = re.compile(ur'^[A-Za-z]$', re.UNICODE)
                    if len(cell) > 0 and not pattern.search(cell):
                        tstat_cb = False
                        break
                if tstat_cb:
                    return 6
                return 3
        return 5

    def print_bold(self, row):
        for i in range(0, len(row)):
            self.write(self.__out_ws.get_current_row(), i, row[i], self.out.bold)

    def print_title(self, row):
        for cell in row:
            self.write(self.__out_ws.get_current_row(), 0, cell)

    def print_n23(self, row):
        self.write(self.__out_ws.get_current_row(), 0, row[0], self.out.n23_background)
        for i in range(1,self.__max_cell):
            self.write(self.__out_ws.get_current_row(), i, "", self.out.n23_background)

    def print_cross_break(self, row):
        check = self.row_types[3] + self.row_types[6]
        if check == 1:
            self.__out_ws.add_to_current_row(2)
        temp = 0
        if len(row) > self.__max_cell:
            self.__max_cell = len(row)
        for i in range(0, len(row)):
            if not row[i] == "" and not row[i] == " ":
                if not temp == 0:
                    if temp + 1 < i:
                        self.__out_ws.get_sheet().merge_range(self.__out_ws.get_current_row(),temp,self.__out_ws.get_current_row(),i-1,row[temp], self.out.banner)
                    else:
                        self.write(self.__out_ws.get_current_row(),i-1,row[temp], self.out.banner)
                        self.write(self.__out_ws.get_current_row(), i, row[i], self.out.banner)
                else:   self.write(self.__out_ws.get_current_row(), i, row[i], self.out.banner)
                temp = i
            else:
                self.write(self.__out_ws.get_current_row(), i, '', self.out.banner)
        if temp + 1 <= i and temp > 0 and len(row[temp]) > 1:
            self.__out_ws.get_sheet().merge_range(self.__out_ws.get_current_row(),temp,self.__out_ws.get_current_row(),i,row[temp], self.out.banner)
        self.__out_ws.get_sheet().set_row(self.__out_ws.get_current_row(), 25)

    def print_tstat(self, row):
        i = 0
        for cell in row:
            pattern = re.compile(ur'[^A-Za-z]', re.UNICODE)
            if len(cell) > 0 and not pattern.search(cell):
                self.write(self.__out_ws.get_current_row(), i, cell, self.out.tstat)
            else:
                self.write(self.__out_ws.get_current_row(), i, cell, self.out.borders)
            i += 1

    def print_regular(self,row):
        i = 0
        for cell in row:
            self.write(self.__out_ws.get_current_row(), i, cell, self.out.borders)
            i += 1

    def print_total_row(self,row):
        i = 0
        self.__out_ws.add_to_current_row(-1)
        if not self.large_row:
            self.__out_ws.get_sheet().set_row(self.__out_ws.get_current_row() - 1, 100)
            self.large_row = True
        for cell in row:
            self.write(self.__out_ws.get_current_row(), i, cell, self.out.blued_style, "total row")
            i += 1
        self.row_types[5] += 1
        self.__out_ws.add_to_current_row(1)

    def print_sub_title(self, row):
        self.add_sub_title(row[0])
        self.print_title(row)

    def print_footer(self):
        self.__out_ws.add_to_current_row(1)
        for row in self.footer:
            self.print_title(row)
            self.__out_ws.add_to_current_row(1)

    def print_center(self, row):
        i = 0
        for cell in row:
            self.write(self.__out_ws.get_current_row(), i, cell, self.out.center)
            i += 1

    def append_to_footer(self, row):
        self.footer.append([row[0][5:]])

    def switcher(self, arg):
        switch = {
            1: self.print_title,
            2: self.print_n23,
            3: self.print_cross_break,
            4: self.print_tstat,
            5: self.print_regular,
            6: self.print_center,
            7: self.print_sub_title,
            8: self.append_to_footer,
        }
        return switch.get(arg)

    def print_content(self, row):
        r_type = self.get_row_type(row)
        self.current_row_type = r_type        
        if r_type == 0 and self.row_types[5] == 0:
            return
        if not r_type in [ 1, 7 ] and self.btxt == 0: 
            self.baseTextObj.process()
            self.btxt = 1
        if r_type == 0:
            r_type = 6
        if r_type in (2, 4, 5, 6) and not self.large_row:
            self.__out_ws.get_sheet().set_row(self.__out_ws.get_current_row() - 1, 100)
            self.large_row = True
        self.row_types[r_type] += 1
        func = self.switcher(r_type)
        func(row)
        if r_type == 6:
            self.__out_ws.get_sheet().set_row(self.__out_ws.get_current_row(), 15)
        if not r_type == 8:
            self.__out_ws.add_to_current_row(1)

    def close_file(self):
        self.__out_ws.get_sheet()._opt_close()

    def add_sub_title(self, string):
        if len(string) > 0:
            self.__TableName = self.__TableName + " - " + string

    def get_current_row(self):
        return self.__current_row

    def increment_current_row(self):
        self.__out_ws.add_to_current_row(1)
    
    def fill_data(self, row):
        if row[0].find("$$sheet_name$$") >= 0:
            process_sheet_name_row(row)
            return 0
        self.data.append(row)

    def process_sheet_name_row(self, row):
        row[0] = row[0].replace("$$sheet_name$$", "")
        self.__out_ws = Sheet(self.out, row[0])
        self.__row_start = 0
        self.out.set_current_ws(self.__out_ws)
        
    def loop_recorded_rows(self):
        self.print_link_to_contents()
        i = 1
        for row in self.data:
            if i == self.tableNameObj.get_total_row_position()+1:
                self.tableNameObj.process()
            if i == self.totalObj.get_total_row_position():
                self.totalObj.process()
                i += 1; continue
            self.print_content(row)
            i += 1
        
    def set_btext_obj(self, obj):
        self.baseTextObj = obj
        
    def set_total_obj(self, obj):
        self.totalObj = obj
    
    def set_tableName_obj(self, obj):
        self.tableNameObj = obj
        
    def get_data_rows(self):
        return len(self.data)
        
