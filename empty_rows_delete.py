import openpyxl


def empty_rows_delete(bom_path):
    wb = openpyxl.load_workbook(bom_path)
    type(wb)
    arkusze = wb.sheetnames
    sheet = wb[arkusze[0]]
    max_row = sheet.max_row
    print(max_row)

    for i in range(max_row, 1, -1):
        value = sheet.cell(row=i, column=2).value
        print(i, ". - ", value)
        if value != None:
            value = value.lstrip()
        print(value)

        if value == "" or value == None:
            sheet.delete_rows(i, 1)
            print(" ====>  usuwamy pusty wiersz")

    wb.save(bom_path)
    wb.close()