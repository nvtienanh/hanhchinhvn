# Reading an excel file using Python
import xlrd
import glob
import json
import requests
from slugify import slugify

files = glob.glob("excel_files/*.xls")

vn_data = []


response = requests.get(
    'https://partner.viettelpost.vn/v2/categories/listProvinceById',
    params={
        'provinceId': -1
    }
)
result = response.text
json_data = json.loads(result)
province_list = json_data['data']
provice_dict = {}
for item in province_list:
    key = slugify(item['PROVINCE_NAME'].rstrip())
    value = item['PROVINCE_ID']
    provice_dict[key] = value

idx = 1
for item in files:
    print(idx)
    idx = idx + 1
    # To open Workbook
    wb = xlrd.open_workbook(item)
    sheet = wb.sheet_by_index(0)

    state_name = sheet.cell_value(1, 0)
    state_code = sheet.cell_value(1, 1)
    state_name_short = state_name.replace(u'Thành phố ', '')
    state_name_short = state_name_short.replace(u'Tỉnh ', '')
    state_name_slug = slugify(state_name_short)

    if state_name_slug in provice_dict:
        vtp_state_code = provice_dict[state_name_slug]
    else:
        raise Exception(
            "Sorry, no numbers below zero %s", state_name_slug)

    response = requests.get(
        'https://partner.viettelpost.vn/v2/categories/listDistrict',
        params={
            'provinceId': vtp_state_code
        }
    )
    result = response.text
    json_data = json.loads(result)
    district_list = json_data['data']
    district_dict = {}
    
    for item in district_list:
        key = slugify(item['DISTRICT_NAME'].rstrip())
        key = key.replace('huyen-', '')
        key = key.replace('quan-', '')
        key = key.replace('thi-xa-', '')
        key = key.replace('thanh-pho-', '')

        key = key.replace('dao-cat-hai', 'cat-hai')
        value = item['DISTRICT_ID']
        district_dict[key] = value
    
    print(state_name_slug)

    current_district = None

    for i in range(1, sheet.nrows):
        
        district_slug = slugify(sheet.cell_value(i, 2))
        district_slug = district_slug.replace('huyen-', '')
        district_slug = district_slug.replace('quan-', '')
        district_slug = district_slug.replace('thi-xa-', '')
        district_slug = district_slug.replace('thanh-pho-', '')

        if (district_slug != current_district):
            if district_slug in district_dict:
                vtp_district_code = district_dict[district_slug]
            else:
                print(district_dict.keys())
                raise Exception(
                    "Sorry, error district", district_slug)
            current_district = district_slug
            print('Quan/huyen: %s' % current_district)
            response = requests.get(
                'https://partner.viettelpost.vn/v2/categories/listWards',
                params={
                    'districtId': vtp_district_code
                }
            )
            result = response.text
            json_data = json.loads(result)
            ward_list = json_data['data']
            ward_dict = {}

            for item in ward_list:
                key = slugify(item['WARDS_NAME'].rstrip())
                key = key.replace('xa-', '')
                key = key.replace('phuong-', '')
                key = key.replace('thi-tran-', '')
                key = key.replace('tt-', '')
                # Stub
                # Correct: Viettelpost => vnsfw_
                # CHƯƠNG DƯƠNG ĐỘ => CHUONG DUONG
                # My dinh i -> My dinh 1
                # My dinh ii -> My dinh 2

                key = key.replace('chuong-duong-do', 'chuong-duong')                
                key = key.replace('my-dinh-ii', 'my-dinh-2')
                key = key.replace('my-dinh-i', 'my-dinh-1')

                # Bình Dương, Bàu Bàng, cay-truong -> cay-truong-ii
                #  dat-quoc -> dat-cuoc
                if current_district == 'bau-bang':
                    key = key.replace('cay-truong', 'cay-truong-ii')
                if current_district == 'bac-tan-uyen':
                    key = key.replace('dat-quoc', 'dat-cuoc')
                if current_district == 'bu-gia-map':
                    key = key.replace('dachia', 'da-kia')
                if current_district == 'hon-quan':
                    key = key.replace('tan-quang', 'tan-quan')
                if current_district == 'bu-dang':
                    key = key.replace('duong-muoi', 'duong-10')
                    key = key.replace('daknhau', 'dak-nhau')
                if current_district == 'ham-thuan-nam':
                    key = key.replace('ham-thuan-nam', 'thuan-nam')
                    key = key.replace('thuan-quy', 'thuan-qui')
                if current_district == 'an-lao':
                    key = key.replace('anh-quang', 'an-quang')
                if current_district == 'hoai-nhon':
                    key = key.replace('hoai-bao', 'hoai-hao')
                    key = key.replace('hoa-xuan', 'hoai-xuan')
                if current_district == 'hoai-an':
                    key = key.replace('dakmang', 'dak-mang')
                    key = key.replace('boktoi', 'bok-toi')
                if current_district == 'gia-rai':
                    key = key.replace('gia-rai', '1')
                if current_district == 'son-dong':
                    key = key.replace('thanh-son', 'tay-yen-tu')
                    key = key.replace('thach-son', 'phuc-son')
                    key = key.replace('chien-son', 'dai-son')
                    key = key.replace('vinh-khuong', 'vinh-an')
                if current_district == 'yen-dung':
                    key = key.replace('neo', 'nham-bien')
                if current_district == 'ba-be':
                    key = key.replace('ba-be-cho-ra', 'cho-ra')
                if current_district == 'bach-thong':
                    key = key.replace('sy-binh', 'si-binh')
                if current_district == 'cho-don':
                    key = key.replace('dai-xao', 'dai-sao')
                if current_district == 'mo-cay-nam':
                    key = key.replace('mo-cay-nam', 'mo-cay')
                if current_district == 'ba-tri':
                    key = key.replace('an-nghai-trung', 'an-ngai-trung')
                if current_district == 'thanh-phu':
                    key = key.replace('an-qui', 'an-quy')
                if current_district == 'cao-bang':
                    key = key.replace('chu-chinh', 'chu-trinh')
                if current_district == 'bao-lam':
                    key = key.replace('pac-mieu', 'pac-miau')
                if current_district == 'ha-quang':
                    key = key.replace('hong-sy', 'hong-si')
                    key = key.replace('qui-quan', 'quy-quan')


                # Tây Ninh, Tinh Biên, nha-ban -> nha-bang
                if current_district == 'tinh-bien':
                    key = key.replace('nha-ban', 'nha-bang')

                value = item['WARDS_ID']
                ward_dict[key] = value

        district_ward_name = sheet.cell_value(
            i, 2) + ' - ' + sheet.cell_value(i, 4)
        district_ward_code = sheet.cell_value(i, 5)

        ward_slug = slugify(sheet.cell_value(i, 4))
        ward_slug = ward_slug.replace('phuong-05', '5')
        ward_slug = ward_slug.replace('phuong-01', '1')
        ward_slug = ward_slug.replace('phuong-02', '2')
        ward_slug = ward_slug.replace('phuong-03', '3')
        ward_slug = ward_slug.replace('phuong-04', '4')
        ward_slug = ward_slug.replace('phuong-06', '6')
        ward_slug = ward_slug.replace('phuong-07', '7')
        ward_slug = ward_slug.replace('phuong-08', '8')
        ward_slug = ward_slug.replace('phuong-09', '9')
       
        ward_slug = ward_slug.replace('phuong-hai-chau-ii', 'hai-chau-2')
        ward_slug = ward_slug.replace('phuong-hai-chau-i', 'hai-chau-1')
        ward_slug = ward_slug.replace('xa-', '')
        ward_slug = ward_slug.replace('phuong-', '')
        ward_slug = ward_slug.replace('thi-tran-', '')
        
        print('Phuong/xa %s' % ward_slug)
        if ward_slug in ward_dict:
            vtp_ward_code = ward_dict[ward_slug]
        else:
            print(ward_dict.keys())
            raise Exception(
                "Sorry, error ward", ward_slug)
            
        vn_data.append({
            'state_name': state_name,
            'state_code': 'P' + state_code,
            'vtp_state_code': vtp_state_code,
            'vtp_district_code': vtp_district_code,
            'district_ward_name': district_ward_name,
            'district_ward_code': 'W' + district_ward_code,
            'vtp_ward_code': vtp_ward_code,
        })

data = {}
data['VN'] = vn_data
with open('vn.json', 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=4)
