# Reading an excel file using Python
import xlrd
import glob
import json
from slugify import slugify

files = glob.glob("excel_files/*.xls")

vn_data = []

for item in files:
    # Give the location of the file
    # loc = ("excel_files/Danh sách quận huyện phường xã thuộc Thành phố Cần Thơ.xls")
    print(item)

    # To open Workbook
    wb = xlrd.open_workbook(item)
    sheet = wb.sheet_by_index(0)

    vnsfw_code = sheet.cell_value(1, 1)
    vnsfw_parent_code = None
    vnsfw_name_with_type = sheet.cell_value(1, 0)
    vnsfw_path_with_type = vnsfw_name_with_type

    if u'Thành phố ' in vnsfw_name_with_type:
        vnsfw_type = 'thanh-pho'
        vnsfw_path = vnsfw_name_with_type.replace('Thành phố ', '')
        vnsfw_name = vnsfw_name_with_type.replace('Thành phố ', '')
        vnsfw_slug = slugify(vnsfw_name)
    elif u'Tỉnh ' in vnsfw_name_with_type:
        vnsfw_type = 'tinh'
        vnsfw_path = vnsfw_name_with_type.replace('Tỉnh ', '')
        vnsfw_name = vnsfw_name_with_type.replace('Tỉnh ', '')
        vnsfw_slug = slugify(vnsfw_name)
    else:
        raise Exception("Error")

    current_province = vnsfw_code
    current_province_name = vnsfw_name
    current_district = None
    current_district_name = None

    vn_data.append({
        'vnsfw_code': vnsfw_code,
        'vnsfw_parent_code': vnsfw_parent_code,
        'vnsfw_name': vnsfw_name,
        'vnsfw_name_with_type': vnsfw_name_with_type,
        'vnsfw_type': vnsfw_type,
        'vnsfw_slug': vnsfw_slug,
        'vnsfw_path': vnsfw_path,
        'vnsfw_path_with_type': vnsfw_path_with_type
    })

    for i in range(1, sheet.nrows):
        if (sheet.cell_value(i, 3) == current_district):
            # Populate ward data for current district
            vnsfw_code = sheet.cell_value(i, 5)  # Ward code
            vnsfw_parent_code = current_district
            vnsfw_name_with_type = sheet.cell_value(i, 4)
            vnsfw_type_name = sheet.cell_value(i, 6)
            vnsfw_type = slugify(vnsfw_type_name)
            vnsfw_path_with_type = vnsfw_name_with_type + ', ' + \
                sheet.cell_value(i, 2) + ', ' + sheet.cell_value(i, 0)
            vnsfw_name = vnsfw_name_with_type.replace(
                vnsfw_type_name + ' ', '')
            vnsfw_path = vnsfw_name + ', ' + current_district_name + ', ' + \
                current_province_name
            vnsfw_slug = slugify(vnsfw_name)

            vn_data.append({
                'vnsfw_code': vnsfw_code,
                'vnsfw_parent_code': vnsfw_parent_code,
                'vnsfw_name': vnsfw_name,
                'vnsfw_name_with_type': vnsfw_name_with_type,
                'vnsfw_type': vnsfw_type,
                'vnsfw_slug': vnsfw_slug,
                'vnsfw_path': vnsfw_path,
                'vnsfw_path_with_type': vnsfw_path_with_type
            })
        else:
            # New district section
            vnsfw_code = sheet.cell_value(i, 3)
            vnsfw_parent_code = current_province
            vnsfw_name_with_type = sheet.cell_value(i, 2)
            if u'Quận ' in vnsfw_name_with_type:
                vnsfw_type = 'quan'
                vnsfw_name = vnsfw_name_with_type.replace('Quận ', '')
                vnsfw_slug = slugify(vnsfw_name)
            elif u'Huyện ' in vnsfw_name_with_type:
                vnsfw_type = 'huyen'
                vnsfw_name = vnsfw_name_with_type.replace('Huyện ', '')
                vnsfw_slug = slugify(vnsfw_name)
            elif u'Thị xã ' in vnsfw_name_with_type:
                vnsfw_type = 'thi-xa'
                vnsfw_name = vnsfw_name_with_type.replace('Thị xã ', '')
                vnsfw_slug = slugify(vnsfw_name)
            elif u'Thành phố ' in vnsfw_name_with_type:
                vnsfw_type = 'thanh-pho'
                vnsfw_name = vnsfw_name_with_type.replace('Thành phố ', '')
                vnsfw_slug = slugify(vnsfw_name)
            else:
                raise Exception("Error %s" % vnsfw_name_with_type)
            vnsfw_path = vnsfw_name + ', ' + current_province_name
            vnsfw_path_with_type = vnsfw_name_with_type + \
                ', ' + sheet.cell_value(i, 0)

            current_district = vnsfw_code
            current_district_name = vnsfw_name
            vn_data.append({
                'vnsfw_code': vnsfw_code,
                'vnsfw_parent_code': vnsfw_parent_code,
                'vnsfw_name': vnsfw_name,
                'vnsfw_name_with_type': vnsfw_name_with_type,
                'vnsfw_type': vnsfw_type,
                'vnsfw_slug': vnsfw_slug,
                'vnsfw_path': vnsfw_path,
                'vnsfw_path_with_type': vnsfw_path_with_type
            })

        # line = '\'' + sheet.cell_value(i, 5) + '\' => '
        # line = line + '\"' + sheet.cell_value(i, 2) + ' - '
        # line = line + sheet.cell_value(i, 4) + '\",\n'
        # f.write(line.encode('UTF-8'))
        # f.write(u'\n')

    # line = "),\n"
    # f.write(line.encode('UTF-8'))

# line = ");\n"
# f.write(line.encode('UTF-8'))
# f.close()  # you can omit in most cases as the destructor will call it

data = {}
data['VN'] = vn_data
with open('vn.json', 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=4)
