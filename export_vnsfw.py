# Reading an excel file using Python
import xlrd
import glob
import json
import requests
from slugify import slugify
import re

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
        key = re.sub('^quan-', '', key)
        key = re.sub('^huyen-', '', key)
        key = re.sub('^thi-xa-', '', key)
        key = re.sub('^thanh-pho-', '', key)

        key = key.replace('dao-cat-hai', 'cat-hai')
        value = item['DISTRICT_ID']
        district_dict[key] = value
    
    print(state_name_slug)

    current_district = None

    for i in range(1, sheet.nrows):
        
        district_slug = slugify(sheet.cell_value(i, 2))
        district_slug = re.sub('^quan-', '', district_slug)
        district_slug = re.sub('^huyen-', '', district_slug)
        district_slug = re.sub('^thi-xa-', '', district_slug)
        district_slug = re.sub('^thanh-pho-', '', district_slug)

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
                
                
                key = key.replace('thi-tran-', '')
                key = key.replace('tt-', '')
                if current_district == 'huong-khe':
                    key = key.replace('phuong-my', 'dien-my')
                if current_district == 'thach-ha':
                    key = key.replace('thach-lam', 'tan-lam-huong')
                    key = key.replace('thach-tien', 'viet-tien')
                    key = key.replace('thach-vinh', 'luu-vinh-son')
                    key = key.replace('nam-huong', 'nam-dien')
                    key = key.replace('thach-dinh', 'dinh-ban')
                if current_district == 'cam-xuyen':
                    key = key.replace('cam-yen', 'yen-hoa')
                    key = key.replace('cam-thang', 'nam-phuc-thang')
                if current_district == 'ky-anh':
                    key = key.replace('ky-lam', 'lam-hop')
                    key = key.replace('ky-hung', 'hung-tri')
                if current_district == 'loc-ha':
                    key = key.replace('binh-loc', 'binh-an')
                if current_district == 'bao-loc':
                    key = key.replace('blao', 'b-lao')
                    key = key.replace('dambri', 'dam-bri')
                if current_district == 'dam-rong':
                    key = key.replace('damrong', 'da-m-rong')
                    key = key.replace('lieng-s-ronh', 'lieng-sronh')
                if current_district == 'lac-duong':
                    key = key.replace('dung-k-noh', 'dung-kno')
                if current_district == 'don-duong':
                    key = key.replace('dran', 'd-ran')
                    key = key.replace('k-don', 'ka-don')
                    key = key.replace('p-roh', 'pro')
                if current_district == 'duc-trong':
                    key = key.replace('n-thon-ha', 'n-thol-ha')
                if current_district == 'bao-lam':
                    key = key.replace('loc-tlam', 'loc-lam')
                if current_district == 'da-huoai':
                    key = key.replace('dam-ri', 'da-m-ri')
                    key = key.replace('madagoil', 'ma-da-guoi')
                    key = key.replace('dap-loa', 'da-ploa')
                if current_district == 'cat-tien':
                    key = key.replace('phuoc-cat-1', 'phuoc-cat')
                if current_district == 'kon-tum':
                    key = key.replace('ngoc-bay', 'ngok-bay')
                    key = key.replace('dak-ro-va', 'dak-ro-wa')
                if current_district == 'dak-glei':
                    key = key.replace('dak-plo', 'dak-blo')
                    key = key.replace('da-pek', 'dak-pek')
                if current_district == 'ngoc-hoi':
                    key = key.replace('po-y', 'bo-y')
                if current_district == 'kon-plong':
                    key = key.replace('dak-long', 'mang-den')
                if current_district == 'sa-thay':
                    key = key.replace('mo-ray', 'mo-rai')
                if current_district == 'tu-mo-rong':
                    key = key.replace('ngok-lay', 'ngoc-lay')
                    key = key.replace('ngok-yeu', 'ngoc-yeu')
                if current_district == 'ia-h-drai':
                    key = key.replace('ia-dai', 'ia-dal')
                if current_district == 'pleiku':
                    key = key.replace('iakring', 'ia-kring')
                    key = key.replace('chuhdrong', 'chu-hdrong')
                    key = key.replace('iakenh', 'ia-kenh')
                if current_district == 'ayun-pa':
                    key = key.replace('rbol', 'ia-rbol')
                    key = key.replace('iarto', 'ia-rto')
                    key = key.replace('chu-bap', 'chu-bah')
                    key = key.replace('iasao', 'ia-sao')
                if current_district == 'kbang':
                    key = key.replace('so-lang', 'son-lang')
                    key = key.replace('k-roong', 'krong')
                    key = key.replace('dakrong', 'dak-roong')
                    key = key.replace('dak-sa-mar', 'dak-smar')
                    key = key.replace('kon-long-khong', 'kong-long-khong')
                    key = key.replace('kon-pla', 'kong-pla')
                    key = key.replace('dak-h-lo', 'dak-hlo')
                if current_district == 'dak-doa':
                    key = key.replace('dak-so-mei', 'dak-somei')
                    key = key.replace('kdang', 'k-dang')
                    key = key.replace('adok', 'a-dok')
                    key = key.replace('h-nol', 'hnol')
                if current_district == 'chu-pah':    
                    key = key.replace('ialy', 'ia-ly')
                    key = key.replace('iakreng', 'ia-kreng')
                    key = key.replace('iaka', 'ia-ka')
                if current_district == 'ia-grai':    
                    key = key.replace('iakha', 'ia-kha')
                    key = key.replace('iachia', 'ia-chia')
                if current_district == 'mang-yang':    
                    key = key.replace('kong-dong', 'kon-dong')  
                    key = key.replace('dak-taley', 'dak-ta-ley')
                    key = key.replace('ha-ra', 'hra')
                    key = key.replace('dak-jrang', 'dak-djrang')
                if current_district == 'kong-chro':    
                    key = key.replace('kon-chro', 'kong-chro')
                    key = key.replace('chu-krei', 'chu-krey')
                    key = key.replace('yama', 'ya-ma')
                if current_district == 'duc-co':  
                    key = key.replace('iadin', 'ia-din')
                if current_district == 'chu-prong':  
                    key = key.replace('ia-gar', 'ia-ga') 
                    key = key.replace('ia-mor', 'ia-mo') 
                if current_district == 'chu-se':  
                    key = key.replace('chupong', 'chu-pong') 
                    key = key.replace('iapal', 'ia-pal')
                    key = key.replace('alba', 'al-ba')
                    key = key.replace('konghtok', 'kong-htok')
                    key = key.replace('iako', 'ia-ko')
                    key = key.replace('hbong', 'h-bong')
                if current_district == 'dak-po':  
                    key = key.replace('dakpo', 'dak-po')
                if current_district == 'ia-pa':  
                    key = key.replace('iapa', 'ia-pa')
                    key = key.replace('iamoron', 'ia-ma-ron')
                    key = key.replace('iabroai', 'ia-broai')
                if current_district == 'krong-pa':  
                    key = key.replace('ia-sai', 'ia-rsai')
                    key = key.replace('iarsuom', 'ia-rsuom')
                    key = key.replace('iar-mook', 'ia-rmok')
                    key = key.replace('chur-cam', 'chu-rcam')
                if current_district == 'phu-thien':  
                    key = key.replace('iake', 'ia-ake')
                    key = key.replace('iasol', 'ia-sol')
                if current_district == 'chu-puh':  
                    key = key.replace('nhon-hoa-h-chu-puh', 'nhon-hoa')
                    key = key.replace('iale', 'ia-le')
                    key = key.replace('iarong', 'ia-rong')
                    key = key.replace('iadreng', 'ia-dreng')
                    key = key.replace('chudon', 'chu-don')
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
                if current_district == 'cai-nuoc':
                    key = key.replace('luong-the-chan', 'luong-the-tran')
                if current_district == 'muong-te':
                    key = key.replace('bun-to', 'bum-to')
                    key = key.replace('bun-nua', 'bum-nua')
                if current_district == 'phong-tho':
                    key = key.replace('lan-nhi-thang', 'la-nhi-thang')
                    key = key.replace('ma-li-pho', 'ma-ly-pho')
                if current_district == 'than-uyen':
                    key = key.replace('phu-ma', 'pha-mu')
                if current_district == 'nam-nhun':
                    key = key.replace('hua-bum', 'hua-bun')
                if current_district == 'ha-tinh':
                    key = key.replace('thach-mon', 'dong-mon')
                if current_district == 'huong-son':
                    key = key.replace('son-an', 'an-hoa-thinh')
                    key = key.replace('son-tan', 'tan-my-ha')
                    key = key.replace('son-thuy', 'kim-hoa')
                    key = key.replace('son-diem', 'quang-diem')
                if current_district == 'duc-tho':
                    key = key.replace('duc-vinh', 'quang-vinh')
                    key = key.replace('duc-tung', 'tung-chau')
                    key = key.replace('duc-lam', 'bui-la-nhan')
                    key = key.replace('trung-le', 'lam-trung-thuy')
                    key = key.replace('duc-thanh', 'thanh-binh-thinh')
                    key = key.replace('duc-dung', 'an-dung')
                    key = key.replace('duc-hoa', 'hoa-lac')
                    key = key.replace('duc-long', 'tan-dan')
                
                if current_district == 'vu-quang':
                    key = key.replace('huong-dien', 'tho-dien')
                    key = key.replace('huong-quang', 'quang-tho')

                if current_district == 'nghi-xuan':
                    key = key.replace('xuan-truong', 'dan-truong')
                if current_district == 'can-loc':
                    key = key.replace('kim-loc', 'kim-song-truong')
                    key = key.replace('khanh-loc', 'khanh-vinh-yen')
                if current_district == 'muong-lat':
                    key = key.replace('tam-trung', 'tam-chung')
                    key = key.replace('phu-nhi', 'pu-nhi')
                if current_district == 'quan-hoa':
                    key = key.replace('hien-trung', 'hien-chung')
                if current_district == 'lang-chanh':
                    key = key.replace('chi-nang', 'tri-nang')
                if current_district == 'cam-thuy':
                    key = key.replace('cam-thuy', 'phong-son')
                if current_district == 'thach-thanh':
                    key = key.replace('thanh-dinh', 'thach-dinh')
                if current_district == 'ha-trung':
                    key = key.replace('ha-thanh', 'hoat-giang')
                    key = key.replace('ha-yen', 'yen-duong')
                    key = key.replace('ha-ninh', 'yen-son')
                    key = key.replace('ha-toai', 'linh-toai')
                if current_district == 'vinh-loc':
                    key = key.replace('vinh-ninh', 'ninh-khang')
                    key = key.replace('vinh-minh', 'minh-tan')
                if current_district == 'yen-dinh':
                    key = key.replace('nong-truong-thong-nhat', 'thong-nhat')
                    key = key.replace('quy-loc', 'qui-loc')
                if current_district == 'tho-xuan':
                    key = key.replace('xuan-khanh', 'xuan-hong')
                    key = key.replace('xuan-son', 'xuan-sinh')
                    key = key.replace('xuan-tan', 'truong-xuan')
                    key = key.replace('xuan-yen', 'phu-xuan')
                    key = key.replace('tho-minh', 'thuan-minh')
                if current_district == 'trieu-son':
                    key = key.replace('tan-ninh', 'nua')
                if current_district == 'thieu-hoa':
                    key = key.replace('thieu-minh', 'minh-tam')
                    key = key.replace('thieu-tan', 'tan-chau')
                if current_district == 'nga-son':
                    key = key.replace('nga-linh', 'nga-phuong')
                if current_district == 'quang-xuong':
                    key = key.replace('quang-tan', 'tan-phong')
                    key = key.replace('quang-loi', 'tien-trang')
                if current_district == 'tinh-gia':
                    key = key.replace('nimh-hai', 'ninh-hai')
                if current_district == 'bat-xat':
                    key = key.replace('bat-sat', 'bat-xat')
                if current_district == 'muong-khuong':
                    key = key.replace('la-van-tan', 'la-pan-tan')
                if current_district == 'si-ma-cai':
                    key = key.replace('simacai', 'si-ma-cai')
                    key = key.replace('nan-sin', 'nan-xin')
                if current_district == 'bac-ha':
                    key = key.replace('lung-phin', 'lung-phinh')
                if current_district == 'bao-thang':
                    key = key.replace('phong-hai', 'n-t-phong-hai')
                    key = key.replace('tang-long', 'tang-loong')
                if current_district == 'sa-pa':
                    key = key.replace('trung-trai', 'trung-chai')
                    key = key.replace('ban-phung', 'thanh-binh')
                if current_district == 'quynh-nhai':
                    key = key.replace('pha-khinh', 'pa-ma-pha-khinh')
                    # key = key.replace('tho-binh', 'binh-son')

                if ('xa-xa-' in key):
                    key = key.replace('xa-xa-', 'xa-')
                elif ('phuong-phuong-' in key):
                    key = key.replace('phuong-phuong-', 'phuong-')                
                elif ('xa-phuong-' in key):
                    key = key.replace('xa-', '')
                elif ('phuong-xa-' in key):
                    key = key.replace('phuong-', '')
                elif ('phuong-' in key):
                    key = key.replace('phuong-', '')
                else:
                    key = key.replace('xa-', '')

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
        if ('xa-xa-' in ward_slug):
            ward_slug = ward_slug.replace('xa-xa-', 'xa-')
        elif ('phuong-phuong-' in ward_slug):
            ward_slug = ward_slug.replace('phuong-phuong-', 'phuong-')                
        elif ('xa-phuong-' in ward_slug):
            ward_slug = ward_slug.replace('xa-', '')
        elif ('phuong-xa-' in ward_slug):
            ward_slug = ward_slug.replace('phuong-', '')
        elif ('phuong-' in ward_slug):
            ward_slug = ward_slug.replace('phuong-', '')
        else:
            ward_slug = ward_slug.replace('xa-', '')
        ward_slug = ward_slug.replace('thi-tran-', '')
        ward_slug = ward_slug.replace('tt-', '')
        
        print('Phuong/xa %s' % ward_slug)
        if ward_slug in ward_dict:            
            vtp_ward_code = ward_dict[ward_slug]
        else:            
            if current_district == 'thanh-hoa':
                if ward_slug == 'thieu-van':
                    vtp_ward_code = 6697
            elif current_district == 'sam-son':
                if ward_slug == 'quang-minh':
                    vtp_ward_code = 6206
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
