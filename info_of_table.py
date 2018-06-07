#!/usr/bin/python3
import xlwt
import xlrd
import xlutils
from createPattern import createPattern
from get_data import get_data_from_input_table
#from sortedSheet import generate_tables
file_name = "input.txt"

LECT = "lect.xls"
PRACT = "pract.xls"
T_COL = 4 # teacher's column
L = 'l'   # flag for  lection's rows
W = 'w'   # flag for lab work
P = 'p'   # flag for practic's work

def get_sheet_data(read_table, teacher_subj, lection=False):
    if read_table == LECT:
        lection = True
    rb = xlrd.open_workbook(read_table,
                             on_demand=True,
                             formatting_info=True)
    for semester in range(0,2): # open autumn or summer semester's sheet
        sheet = rb.sheet_by_index(semester)
        teacher_subj.update(push_in_list(sheet=sheet,
                                    teacher_subj=teacher_subj,
                                    lection=lection,
                                    semester=semester))
    return teacher_subj


def push_in_list(sheet, teacher_subj, semester, lection=False):
    row = 1
    sem = str(semester + 1) #creation of key like s1 - autumn sem or s2 - spring sem
    while sheet.cell(row, T_COL).value != 'END':
        teacher = sheet.cell(row, T_COL).value # we find all rows of this teacher
        if teacher == '': 
            row += 1
            continue
        if teacher not in teacher_subj:
            if lection:
                teacher_subj.update({teacher:[sem + str(row) + L + str(int(sheet.cell(row, 0).value))]}) # it points that it's a lection table's row
            elif 'Пр' in sheet.cell(row, 3).value:
                teacher_subj.update({teacher:[sem + str(row) + P + str(int(sheet.cell(row, 0).value))]}) # it's points that there is a practic
            else:
                teacher_subj.update({teacher:[sem + str(row) + W + str(int(sheet.cell(row, 0).value))]}) # it points that there is a lab work
            row += 1
            continue
        if lection:
            teacher_subj[teacher].append(sem + str(row) + L + str(int(sheet.cell(row, 0).value)))
        elif 'Пр' in sheet.cell(row, 3).value:
            teacher_subj[teacher].append(sem + str(row) + P + str(int(sheet.cell(row, 0).value)))
        else:
            teacher_subj[teacher].append(sem + str(row) + W + str(int(sheet.cell(row, 0).value)))
        row += 1
    return teacher_subj


def info_of_table():
    #generate_tables()  # this functions generate tables for lections and practise
    teacher_subj = {}   # dictionary which contains teachers and row of their subjects with keys like semesters and kind of work
    print("!!", teacher_subj)
    for table in [LECT, PRACT]: # creation dict with information of all kinds of subject
        teacher_subj.update(get_sheet_data(table, teacher_subj))
    for key in teacher_subj:
        print(teacher_subj[key])
        new_table = createPattern()
        new_table = get_data_from_input_table(new_table, teacher_subj[key])
        print(key)
        new_table.save(key + ".xls")


if __name__ == "__main__":
    info_of_table()

