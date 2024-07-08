from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Alignment
from openpyxl.utils import column_index_from_string,get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

import logging
import os


"""
os.chdir("..")
badir =  os. getcwd()  
os.chdir("02_FEM tables")
femdir = os.getcwd()
# entered the FEM save space"""

LC1dir ={}
LC2dir ={}
LC3dir ={}
maindir = {}
maindir ['LC1'] = LC1dir
maindir ['LC2'] = LC2dir
maindir ['LC3'] = LC3dir


hey = 'LC1'
value = {}
for i in range (1,4):
    maindir [f'LC{i}'] ['ficken'] = value
    maindir [f'LC{i}'] ['spast'] = value

print ('hello')