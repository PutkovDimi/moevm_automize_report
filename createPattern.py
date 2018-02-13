import xlrd
import xlwt
from xlutils.copy import copy


INPUT_XLS = 'input.xls'

    def createPattern():
    rb = xlrd.open_workbook(INPUT_XLS, 
                       on_demand=True,
                       formatting_info=True)
    wb = copy(rb)
    return wb
