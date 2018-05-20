import xlrd
import xlwt
from xlutils.copy import copy

INPUT_XLS = "pattern.xls"

def createPattern(table=INPUT_XLS):
    rb = xlrd.open_workbook(table, 
                       on_demand=True,
                       formatting_info=True)
    wb = copy(rb)
    return wb
