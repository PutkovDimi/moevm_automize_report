#!/usr/bin/python3

import xlrd, xlwt
from xlutils.copy import copy
from createPattern import createPattern

GROUP = "ГРУППЫ_МОЭВМ.xls"
SUBJ_TABLE = "groups"
LECT_TABLE = "lections"
TT_PATTERN = "timetablePattern.xls"

def open_read_wrksht():
    rb = xlrd.open_workbook(GROUP,
                          on_demand=True,
                          formatting_info=True)
    return rb.sheet_by_index(0)


def get_data_from_groups_files():
    rb = open_read_wrksht() 
    wb = createPattern(TT_PATTERN)
    wb_lection = createPattern(TT_PATTERN)
    for sem in range(0,2):
        worksheet = wb.get_sheet(sem) #change sem
        worksheet_lection = wb_lection.get_sheet(sem)
        r_row = 0
        w_lect_row = 1
        w_row = 1
        while r_row < 4:
            worksheet,worksheet_lection,w_row,w_lect_row = get_subj(worksheet,worksheet_lection, rb, w_row, w_lect_row,r_row, sem)
            w_row += 1
            r_row += 1
        create_border(worksheet, w_row) # write END for valid stopping
        create_border(worksheet_lection, w_lect_row) # of program
    wb.save("pract.xls")
    wb_lection.save("lect.xls")


def create_border(sheet, row):
    for col in range(0,5):
        sheet.write(row, col, "END")
    return


def get_read_sheet(rb, r_row):
    read_wb = xlrd.open_workbook(rb.cell(r_row,4).value,
                                        on_demand=True,
                                        formatting_info=True)
    course = str(int(rb.cell(r_row,0).value))
    print("course", course)
    return read_wb.sheet_by_name("Курс" + course), course



def get_subj(worksheet,worksheet_lection, rb,w_row,w_lect_row, r_row, sem):
    read_sheet, course = get_read_sheet(rb, r_row)
    if course == '1':
        border = 'Элективные курсы по физической культуре'
        start = 16
    elif course == '2':
        border  = 'Элективные курсы по физической культуре'
        start = 28
    elif course == '3':
        border = 'Производственная практика'
        start = 45
    else:
        border = 'Военная подготовка (Обучение граждан по программе военной подготовки офицеров запаса на факультете военного обучения (военной кафедре))'
        start = 73
    for row in range(start, read_sheet.nrows):
        if read_sheet.cell(row,4).value == border:
            break
        if sem == 0:
            lab_col = 10
            flag = (read_sheet.cell(row, 6).value != '')
        else:
            lab_col = 20
            flag = (read_sheet.cell(row, 16).value != '')
        if str(read_sheet.cell(row,36).value) == "14" and flag:

            if read_sheet.cell(row,lab_col-1).value != '':
                worksheet_lection.write(w_lect_row, 0, course)
                worksheet_lection.write(w_lect_row, 1, 
                                        read_sheet.cell(row,4).value)
                worksheet_lection.write(w_lect_row, 2, 
                                        rb.cell(r_row,1).value)

            if read_sheet.cell(row,lab_col).value != '':
                worksheet.write(w_row, 0, course)
                worksheet.write(w_row, 1, 
                                read_sheet.cell(row,4).value)
                worksheet.write(w_row, 2, 
                                rb.cell(r_row,1).value)
                worksheet.write(w_row, 3, "Лаб")

            if read_sheet.cell(row,lab_col+1).value != '':
                w_row += 1
                worksheet.write(w_row, 0, course)
                worksheet.write(w_row, 1, 
                                read_sheet.cell(row,4).value)
                worksheet.write(w_row, 2, 
                                rb.cell(r_row,1).value)
                worksheet.write(w_row, 3, "Пр")
            w_row += 1
            w_lect_row += 1
    return worksheet,worksheet_lection,w_row, w_lect_row



if __name__ == '__main__':
    get_data_from_groups_files()
