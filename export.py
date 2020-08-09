# Reading an excel file using Python
import xlrd
import glob

f = open('VN.php', 'wb')

files = glob.glob("excel_files/*.xls")

line = "$places['VN'] = array(\n"
f.write(line.encode('UTF-8'))

for item in files:
    # Give the location of the file
    # loc = ("excel_files/Danh sách quận huyện phường xã thuộc Thành phố Cần Thơ.xls")
    print(item)

    # To open Workbook
    wb = xlrd.open_workbook(item)
    sheet = wb.sheet_by_index(0)

    province_id = sheet.cell_value(1, 1)

    line = "\"" + province_id + "\" => array(\n"
    f.write(line.encode('UTF-8'))

    for i in range(1, sheet.nrows):
        line = ''
        # line = '\'' + sheet.cell_value(i, 5) + '\' => '
        line = line + '\"' + sheet.cell_value(i, 2) + ' - '
        line = line + sheet.cell_value(i, 4) + '\",\n'
        f.write(line.encode('UTF-8'))
        # f.write(u'\n')

    line = "),\n"
    f.write(line.encode('UTF-8'))

line = ");\n"
f.write(line.encode('UTF-8'))
f.close()  # you can omit in most cases as the destructor will call it
