import xlrd

TABLE = "ГРУППЫ_МОЭВМ.xls"

def create_coll():
    groups_students = {}
    rb = xlrd.open_workbook(TABLE,
                            on_demand=True,
                            formatting_info=True)
    read_sheet = rb.sheet_by_index(0)
    for row in range(0, 20):
        if read_sheet.cell(row, 0).value == 'END':
            break
        print(read_sheet.cell(row,1).value)
        groups_students.update({read_sheet.cell(row,1).value:
            read_sheet.cell(row,3).value})
    return groups_students


if __name__ == '__main__':
    print(create_coll())
