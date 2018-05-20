import xlrd, xlwt, xlutils
from create_group_coll import create_coll
INPUT_TABLE = 'dataTable.xls'
LECT = 'lect.xls'
PRACT = 'pract.xls'
COL = 4
GROUPS_TABLE = "ГРУППЫ_МОЭВМ.xls"

def get_data_from_input_table(wb, subj_list):
    read_wb = xlrd.open_workbook(INPUT_TABLE)
    input_pract = xlrd.open_workbook(PRACT)
    input_lect = xlrd.open_workbook(LECT)
    numb_of_row = 0
    numb_of_sec_sheet_row = 0
    for subject in subj_list:
        if subject[0] == '1':# and subject[-2] == 'l':
            if subject[-2] == 'l':
                info_table = input_lect.sheet_by_index(0)
            else:
                info_table = input_pract.sheet_by_index(0)
        else:
            if subject[-2] == 'l':
                info_table = input_lect.sheet_by_index(1)
            else:
                info_table = input_pract.sheet_by_index(1)
        read_sheet = read_wb.sheet_by_name("Курс" + subject[-1])
        print(info_table.cell(int(subject[1:-2]), 1).value)
        r_row = find_row(read_sheet, info_table.cell(int(subject[1:-2]), 1).value)
        if subject[0] == '1':
            get_row_data(wb, read_sheet, info_table, r_row, subject, numb_of_row)
            numb_of_row += 1
        else:
            get_row_data(wb,read_sheet, info_table,r_row,subject,numb_of_sec_sheet_row)
            numb_of_sec_sheet_row += 1
            
    return wb
        

def find_row(sheet, subject):
    ROW = 0
    for row in range(sheet.nrows):
        cell = sheet.cell(row, COL)
        if cell.value.lower() == subject.lower():
            return ROW
        ROW += 1
    return ROW


def get_row_data(wb, read_sheet, info_table, r_row, subject, numb_of_row):
    groups_count = create_coll()
    #find a row of interested subj
    if subject[0] == '1':
        cell_diff = 0
        write_sheet = wb.get_sheet(1)
    else:
        cell_diff = 10
        write_sheet = wb.get_sheet(2)
    print(info_table.cell(int(subject[1:-2]), 1).value, 'gr1')
    print(info_table.cell(int(subject[1:-2]), 0).value, 'gr')
    print(info_table.cell(int(subject[1:-2]), 2).value, 'gr4')
    write_sheet.write(9+numb_of_row, 0, info_table.cell(int(subject[1:-2]), 1).value)
    write_sheet.write(9+numb_of_row, 3, info_table.cell(int(subject[1:-2]), 0).value)
    write_sheet.write(9+numb_of_row, 1, info_table.cell(int(subject[1:-2]), 2).value) 
    write_sheet.write(9+numb_of_row, 4, groups_count[info_table.cell(int(subject[1:-2]), 2).value])
## place for groups and studs
    if "Экз" in read_sheet.cell(r_row,6).value:
        write_sheet.write(9+ numb_of_row,7,"1")
    else:
        write_sheet.write(9+numb_of_row,8,"1")
    if subject[-2] == 'l':
        write_sheet.write(9+numb_of_row,15,
                read_sheet.cell(r_row,9+cell_diff).value)#lekcii
        print(read_sheet.cell(r_row, 9+cell_diff).value, 'f')
    if subject[-2] == 'w':
        print(read_sheet.cell(r_row, 10+cell_diff).value, 'f')
        write_sheet.write(9+numb_of_row,17,
                read_sheet.cell(r_row, 10+cell_diff).value)#lab
    if subject[-2] == 'p':
        write_sheet.write(9+numb_of_row,16,
                read_sheet.cell(r_row, 11+cell_diff).value)#practic
        print(read_sheet.cell(r_row, 11+cell_diff).value, 'f')
    return
