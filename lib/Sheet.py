# -*- coding: utf-8 -*-
"""
Created on Thu Nov  3 11:39:19 2016

@author: plamen.tarkalanov
"""
class Sheet():
    current_row = ""   
    sheet_name = ""
    sheet = ""
    
    def __init__(self, output_excel, sheet_name):
        self.sheet = output_excel.add_worksheet(sheet_name) 
        self.current_row = 0
        self.sheet_name = sheet_name
        
    def get_current_row(self):
        return self.current_row

    def add_to_current_row(self, offset):
        self.current_row += offset
        
    def write(self, *args):
        self.sheet.write(*args)    
    
    def get_sheet(self):
        return self.sheet
        
    def get_sheetname(self):
        return self.sheet_name
        
        
        