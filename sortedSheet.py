#!/usr/bin/python3
import xlrd, xlwt
from createPattern import createPattern
from read_groups import get_data_from_groups_files

PRACT = "pract.xls"
LECT = "lect.xls"
#SUBJ_TABLE = "timetablePattern"
SUBJ_COL = 1


def open_read_wrksht(spreadsheet):
    rb = xlrd.open_workbook(spreadsheet,
                            on_demand=True,
                            formatting_info=True)
    return rb


def choose_semester(spreadsheet, lection=False):
    rb = open_read_wrksht(spreadsheet)
    wb = createPattern(spreadsheet)
    for semester in range(0,2):
        write_sheet = wb.get_sheet(semester)
        read_sheet = rb.sheet_by_index(semester)
        write_sheet = sorted_sheet(write_sheet=write_sheet,
                                   read_sheet=read_sheet,
                                   lection=lection)
    wb.save(spreadsheet)
    return


def sorted_sheet(write_sheet, read_sheet, lection):
    read_row = 1
    write_row = 0 
    subj_list = []
    while read_sheet.cell(read_row, SUBJ_COL).value != 'END':
        found_subject = read_sheet.cell(read_row, SUBJ_COL).value
        if found_subject in subj_list:
            read_row += 1
            continue
        subj_list.append(found_subject)
        write_row += 1
        write_sheet.write(write_row, 0, read_sheet.cell(read_row, 0).value)
        write_sheet.write(write_row, 1, read_sheet.cell(read_row, 1).value)
        write_sheet.write(write_row, 2, read_sheet.cell(read_row, 2).value)
        if lection and read_sheet.cell(read_row, 2).value != '':
            write_sheet.write(write_row, 3, "Лекция")
        else:
            write_sheet.write(write_row, 3, read_sheet.cell(read_row, 3).value)
        r_row = read_row + 1
        while read_sheet.cell(r_row, SUBJ_COL).value != 'END':
            if found_subject == read_sheet.cell(r_row, SUBJ_COL).value:
                write_row += 1
                write_sheet.write(write_row, 0, 
                                  read_sheet.cell(r_row, 0).value)
                write_sheet.write(write_row, 1, 
                                  read_sheet.cell(r_row, 1).value)
                write_sheet.write(write_row, 2, 
                                  read_sheet.cell(r_row, 2).value)
                if lection and read_sheet.cell(r_row, 2).value != '':
                    write_sheet.write(write_row, 3, "Лекция")
                else:
                    write_sheet.write(write_row, 3,
                                      read_sheet.cell(r_row, 3).value)
            r_row += 1
        read_row += 1
    return write_sheet


if __name__ == '__main__':
#def generate_tables():
    get_data_from_groups_files()
    choose_semester(LECT, lection=True)
    choose_semester(PRACT)
