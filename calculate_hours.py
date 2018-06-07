import xlwt
import xlrd
from createPattern import createPattern

RESULT_COL = 32
EXAM_HOURS = 11.5
EXAM_HOURS_COL = 20
EXAM_POINT = 7
SUBJECT_ROW = 9
name = "Пример"
COL_OF_STUDENTS = 4
COL_TF = 8
COL_TF_COl = 19
COL_LECT_HOURS = 15
COL_PRACT_HOURS = 16
COL_LAB_HOURS = 17
RESULT_ROW = 26


def iterate_subjects(wb, rb):
    row_dict = {COL_OF_STUDENTS: 0, EXAM_HOURS_COL: EXAM_POINT, COL_TF_COl: 0, COL_LECT_HOURS: 0, COL_PRACT_HOURS: 0,
                COL_LAB_HOURS: 0, RESULT_COL: 0}
    col_dict = row_dict.copy()
    col_dict[EXAM_HOURS_COL] = 0
    row = SUBJECT_ROW
    while rb.cell(row, 0).value:
        print(rb.cell(row, 0).value)
        row_dict.update({COL_OF_STUDENTS: rb.cell(row, COL_OF_STUDENTS).value if rb.cell(row, COL_OF_STUDENTS).value else 0,
                         EXAM_HOURS_COL: EXAM_HOURS if rb.cell(row, EXAM_POINT).value else 0,
                         COL_TF_COl: rb.cell(row, COL_OF_STUDENTS).value if rb.cell(row, COL_TF).value else 0,
                         COL_LECT_HOURS: rb.cell(row, COL_LECT_HOURS).value if rb.cell(row,
                                                                                       COL_LECT_HOURS).value else 0,
                         COL_PRACT_HOURS: rb.cell(row, COL_PRACT_HOURS).value if rb.cell(row,
                                                                                         COL_PRACT_HOURS).value else 0,
                         COL_LAB_HOURS: rb.cell(row, COL_LAB_HOURS).value if rb.cell(row,
                                                                                     COL_LAB_HOURS).value else 0})
        print(row_dict)
        for key in row_dict:
            col_dict[key] += row_dict[key]
            wb.write(row, key, row_dict[key])
        row_dict[COL_OF_STUDENTS] = 0
        wb.write(row, RESULT_COL, sum(row_dict.values()))
        print(col_dict)
        row += 1
    col_dict[COL_OF_STUDENTS] = 0
    col_dict[RESULT_COL] = sum(col_dict.values())
    for key in col_dict:
        wb.write(RESULT_ROW, key, col_dict[key])


def iterate_semesters(name=name):
    final_table = createPattern(name+".xls")
    data_table = xlrd.open_workbook(name+".xls")
    for semester in range(1, 3):
        wb = final_table.get_sheet(semester)
        rb = data_table.sheet_by_index(semester)
        iterate_subjects(wb, rb)
    final_table.save(name + "2.xls")


if __name__ == "__main__":
    iterate_semesters()
