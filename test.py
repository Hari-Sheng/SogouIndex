# -*- coding: utf-8 -*-
__author__ = 'Hari Sheng'
import xlrd, os

WorkbookTemp = xlrd.open_workbook('iphone.xlsx')
print WorkbookTemp.get_sheet(0).cell(0,0).value