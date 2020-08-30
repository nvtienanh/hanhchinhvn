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
        if (key == 'thi-xa-long-my') | (key == 'thi-xa-cai-lay') | (key == 'thi-xa-hong-ngu'):
            pass
        else:
            key = re.sub('^thi-xa-', '', key)
        if (key == 'thanh-pho-cao-lanh'):
            pass
        else:
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
        
        if (district_slug == 'thi-xa-long-my') | (district_slug == 'thi-xa-cai-lay') | (district_slug == 'thi-xa-hong-ngu'):
            pass
        else:
            district_slug = re.sub('^thi-xa-', '', district_slug)
        if (district_slug == 'thanh-pho-cao-lanh'):
            pass
        else:
            district_slug = re.sub('^thanh-pho-', '', district_slug)        

        if (district_slug != current_district):
            if district_slug in district_dict:
                vtp_district_code = district_dict[district_slug]
            elif district_slug == 'dak-r-lap':
                vtp_district_code = district_dict['dak-rlap']
            elif district_slug == 'krong-pac':
                vtp_district_code = district_dict['krong-pak']
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
                if current_district == 'thuan-chau':
                    key = key.replace('tong-lenh', 'tong-lanh')
                    key = key.replace('pung-cha', 'pung-tra')
                if current_district == 'bac-yen':
                    key = key.replace('ta-khoa', 'hua-nha')
                if current_district == 'moc-chau':
                    key = key.replace('nong-truong', 'nt-moc-chau')
                if current_district == 'yen-chau':
                    key = key.replace('vieng-lang', 'vieng-lan')
                if current_district == 'mai-son':
                    key = key.replace('na-bo', 'na-po')
                if current_district == 'song-ma':
                    key = key.replace('duc-mon', 'dua-mon')
                if current_district == 'sop-cop':
                    key = key.replace('xam-kha', 'sam-kha')
                if current_district == 'meo-vac':
                    key = key.replace('la-lung', 'ta-lung')
                    key = key.replace('nam-pan', 'nam-ban')
                if current_district == 'quan-ba':
                    key = key.replace('quang-ba', 'quan-ba')
                if current_district == 'xin-man':
                    key = key.replace('tran-tran-coc-pai', 'coc-pai')
                    key = key.replace('ban-riu', 'ban-diu')
                    key = key.replace('na-tri', 'na-chi')
                if current_district == 'tay-giang':
                    key = key.replace('g-ry', 'ga-ri')
                    key = key.replace('bhalle', 'bha-le')
                if current_district == 'dong-giang':
                    key = key.replace('prao', 'p-rao')
                if current_district == 'nam-giang':
                    key = key.replace('la-ee', 'laee')
                    key = key.replace('zuoih', 'zuoich')
                    key = key.replace('cha-vai', 'cha-val')
                    key = key.replace('dac-prinh', 'dac-pring')
                if current_district == 'nui-thanh':
                    key = key.replace('tam-xuan-2', 'tam-xuan-ii')
                    key = key.replace('tam-xuan-1', 'tam-xuan-i')
                if current_district == 'nho-quan':
                    key = key.replace('lac-phong', 'lang-phong')
                if current_district == 'hai-duong':
                    key = key.replace('an-chau', 'an-thuong')
                if current_district == 'kinh-mon':
                    key = key.replace('hoang-son', 'hoanh-son')
                    key = key.replace('pham-menh', 'pham-thai')
                    key = key.replace('quang-trung', 'quang-thanh')
                if current_district == 'kim-thanh':
                    key = key.replace('tuan-hung', 'tuan-viet')
                    key = key.replace('dong-gia', 'dong-cam')
                    key = key.replace('kim-khe', 'kim-lien')
                if current_district == 'thanh-ha':
                    key = key.replace('an-luong', 'an-phuong')
                    key = key.replace('hop-duc', 'thanh-quang')
                if current_district == 'cam-giang':
                    key = key.replace('cam-dinh', 'dinh-son')
                if current_district == 'binh-giang':
                    key = key.replace('vinh-tuy', 'vinh-hung')
                if current_district == 'tu-ky':
                    key = key.replace('ky-son', 'dai-son')
                    key = key.replace('dong-ky', 'chi-minh')
                if current_district == 'thanh-mien':
                    key = key.replace('tien-phong', 'hong-phong')
                if current_district == 'an-bien':
                    key = key.replace('thu-3', 'thu-ba')
                if current_district == 'an-minh':
                    key = key.replace('thu-11', 'thu-muoi-mot')
                if current_district == 'kien-hai':
                    key = key.replace('dao-nam-du', 'nam-du')
                if current_district == 'thai-binh':
                    key = key.replace('kcn-nguyen-duc-canh', 'nguyen-duc-canh')
                    key = key.replace('kcn-phu-khanh', 'phu-khanh')
                    key = key.replace('kcn-phong-phu', 'phong-phu')
                if current_district == 'quynh-phu':
                    key = key.replace('an-quy', 'an-qui')
                if current_district == 'hung-ha':
                    key = key.replace('kim-trung', 'kim-chung')
                if current_district == 'dong-hung':
                    key = key.replace('trong-quang', 'trong-quan')
                if current_district == 'tien-hai':
                    key = key.replace('dong-quy', 'dong-qui')
                if current_district == 'kien-xuong':
                    key = key.replace('vu-quy', 'vu-qui')
                if current_district == 'vinh-linh':
                    key = key.replace('vinh-trung', 'trung-nam')
                    key = key.replace('vinh-thach', 'kim-thach')
                    key = key.replace('pa-tang', 'ba-tang')
                    key = key.replace('vinh-hien', 'hien-thanh')
                if current_district == 'huong-hoa':
                    key = key.replace('pa-tang', 'ba-tang')
                    key = key.replace('a-xing', 'lia')
                if current_district == 'gio-linh':
                    key = key.replace('gio-phong', 'phong-binh')
                    key = key.replace('vinh-truong', 'linh-truong')
                if current_district == 'cam-lo':
                    key = key.replace('cam-an', 'thanh-an')
                if current_district == 'hai-lang':
                    key = key.replace('hai-tho', 'dien-sanh')
                    key = key.replace('hai-vinh', 'hai-hung')
                    key = key.replace('hai-thanh', 'hai-dinh')
                    key = key.replace('hai-hoa', 'hai-phong')
                if current_district == 'dak-glong':
                    key = key.replace('dak-rmang', 'dak-r-mang')
                if current_district == 'cu-jut':
                    key = key.replace('eatling', 'ea-t-ling')
                    key = key.replace('eapo', 'ea-po')
                if current_district == 'dak-mil':
                    key = key.replace('dak-rla', 'dak-r-la')
                    key = key.replace('dak-ndrot', 'dak-n-drot')
                if current_district == 'krong-no':
                    key = key.replace('buon-choach', 'buon-choah')
                    key = key.replace('dak-ro', 'dak-dro')
                    key = key.replace('nam-ndir', 'nam-n-dir')
                if current_district == 'dak-song':
                    key = key.replace('dak-moi', 'dak-mol')
                    key = key.replace('dak-ndrung', 'dak-n-dung')
                    key = key.replace('dak-rlap', 'dak-r-lap')
                if current_district == 'tuy-duc':
                    key = key.replace('dak-rtih', 'dak-r-tih')
                if current_district == 'phu-ly':
                    key = key.replace('p-thanh-tuyen', 'thanh-tuyen')
                if current_district == 'duy-tien':
                    key = key.replace('tien-phong', 'tien-son')
                if current_district == 'thanh-liem':
                    key = key.replace('thanh-binh', 'tan-thanh')
                if current_district == 'ly-nhan':
                    key = key.replace('nhan-dao', 'tran-hung-dao')
                if current_district == 'cam-pha':
                    key = key.replace('p-mong-duong', 'mong-duong')
                    key = key.replace('p-quang-hanh', 'quang-hanh')
                if current_district == 'thu-thua':
                    key = key.replace('tan-lap', 'tan-long')
                if current_district == 'tan-tru':
                    key = key.replace('my-binh', 'tan-binh')
                if current_district == 'can-giuoc':
                    key = key.replace('can-guoc', 'can-giuoc')
                    key = key.replace('tan-lap', 'tan-tap')
                if current_district == 'que-phong':
                    key = key.replace('muong-nooc', 'muong-noc')
                if current_district == 'quy-chau':
                    key = key.replace('quy-chau', 'tan-lac')
                if current_district == 'nghia-dan':
                    key = key.replace('nghia-lien', 'nghia-thanh')
                if current_district == 'dien-chau':
                    key = key.replace('dien-minh', 'minh-chau')
                if current_district == 'thanh-chuong':
                    key = key.replace('thanh-tuong', 'dai-dong')
                if current_district == 'nghi-loc':
                    key = key.replace('nghi-hop', 'khanh-hop')
                if current_district == 'nam-dan':
                    key = key.replace('nam-loc', 'thuong-tan-loc')
                    key = key.replace('nam-cuong', 'trung-phuc-cuong')
                if current_district == 'hung-nguyen':
                    key = key.replace('hung-thang', 'hung-nghia')
                    key = key.replace('hung-long', 'long-xa')
                    key = key.replace('hung-xuan', 'xuan-lam')
                    key = key.replace('hung-khanh', 'hung-thanh')
                    key = key.replace('hung-chau', 'chau-nhan')
                if current_district == 'vi-thanh':
                    key = key.replace('1', 'i')
                    key = key.replace('3', 'iii')
                    key = key.replace('4', 'iv')
                    key = key.replace('5', 'v')
                    key = key.replace('7', 'vii')
                if current_district == 'buon-ma-thuot':
                    key = key.replace('eatam', 'ea-tam')
                    key = key.replace('eatu', 'ea-tu')
                    key = key.replace('cuebuor', 'cu-ebur')
                    key = key.replace('eakao', 'ea-kao')
                if current_district == 'ea-sup':
                    key = key.replace('cu-kblang', 'cu-kbang')
                if current_district == 'cu-m-gar':
                    key = key.replace('ea-h-ding', 'ea-h-dinh')
                    key = key.replace('ea-m-dro-h', 'ea-m-droh')
                    key = key.replace('cu-m-ga', 'cu-m-gar')
                if current_district == 'krong-buk':
                    key = key.replace('pongdrang', 'pong-drang')
                if current_district == 'ea-kar':
                    key = key.replace('ea-tyh', 'ea-tih')
                    key = key.replace('cu-nie', 'cu-ni')
                    key = key.replace('ea-lang', 'cu-elang')
                if current_district == 'm-drak':
                    key = key.replace('em-m-doal', 'ea-m-doal')
                    key = key.replace('cu-kroa', 'cu-k-roa')
                if current_district == 'krong-bong':
                    key = key.replace('krong-pak', 'krong-pac')
                if current_district == 'krong-pac':
                    key = key.replace('ea-kmec', 'ea-knuec')
                    key = key.replace('ea-knang', 'ea-kuang')
                if current_district == 'hoa-binh':
                    key = key.replace('sui-ngoi', 'su-ngoi')
                    key = key.replace('hop-thinh', 'thinh-minh')
                    key = key.replace('phuc-tien', 'quang-tien')
                if current_district == 'da-bac':
                    key = key.replace('dong-nghe', 'nanh-nghe')
                if current_district == 'luong-son':
                    key = key.replace('cao-ram', 'cao-son')
                    key = key.replace('hop-chau', 'cao-duong')
                    key = key.replace('hop-thanh', 'thanh-son')
                    key = key.replace('cao-thang', 'thanh-cao')
                if current_district == 'kim-boi':
                    key = key.replace('hung-tien', 'hung-son')
                    key = key.replace('kim-son', 'kim-lap')
                    key = key.replace('thuong-bi', 'xuan-thuy')
                    key = key.replace('hop-dong', 'hop-tien')
                if current_district == 'cao-phong':
                    key = key.replace('dong-phong', 'hop-phong')
                    key = key.replace('yen-lap', 'thach-yen')
                if current_district == 'tan-lac':
                    key = key.replace('trung-hoa', 'suoi-hoa')
                    key = key.replace('lung-van', 'van-son')
                    key = key.replace('do-nhan', 'nhan-my')
                if current_district == 'mai-chau':
                    key = key.replace('tan-dan', 'tan-thanh')
                    key = key.replace('tan-mai', 'son-thuy')
                    key = key.replace('thung-khe', 'thanh-son')
                    key = key.replace('tan-son', 'dong-tan')
                    key = key.replace('xam-khoe', 'sam-khoe')
                if current_district == 'lac-son':
                    key = key.replace('phuc-tuy', 'quyet-thang')
                    key = key.replace('binh-chan', 'vu-binh')
                if current_district == 'yen-thuy':
                    key = key.replace('lac-huong', 'lac-luong')
                if current_district == 'lac-thuy':
                    key = key.replace('thanh-nong', 'ba-hang-doi')
                    key = key.replace('lien-hoa', 'thong-nhat')
                    key = key.replace('co-nghia', 'phu-nghia')
                    key = key.replace('chi-le', 'chi-ne')
                if current_district == 'trang-bom':
                    key = key.replace('bau-ham-1', 'bau-ham')
                if current_district == 'thong-nhat':
                    key = key.replace('gia-kiem-gia-kiem', 'gia-kiem')
                if current_district == 'yen-son':
                    key = key.replace('quy-quan', 'qui-quan')
                if current_district == 'son-duong':
                    key = key.replace('tuan-lo', 'tan-thanh')
                    key = key.replace('sam-duong', 'truong-sinh')
                if current_district == 'van-lang':
                    key = key.replace('tan-viet', 'bac-viet')
                    key = key.replace('an-hung', 'bac-hung')
                if current_district == 'van-quan':
                    key = key.replace('van-mong', 'lien-hoi')
                    key = key.replace('dai-an', 'an-son')
                    key = key.replace('van-an', 'diem-he')
                    key = key.replace('chi-le', 'tri-le')
                if current_district == 'bac-son':
                    key = key.replace('quynh-son', 'bac-quynh')
                if current_district == 'huu-lung':
                    key = key.replace('tan-lap', 'thien-tan')
                if current_district == 'loc-binh':
                    key = key.replace('xuan-man', 'khanh-xuan')
                    key = key.replace('dong-duc', 'dong-buc')
                    key = key.replace('minh-phat', 'minh-hiep')
                    key = key.replace('nhuong-ban', 'thong-nhat')
                    key = key.replace('huu-luan', 'huu-lan')
                if current_district == 'dinh-lap':
                    key = key.replace('che-nong-truong-thai-binh', 'nt-thai-binh')
                if current_district == 'trung-khanh':
                    key = key.replace('thong-hue', 'thong-hoe')
                if current_district == 'ha-lang':
                    key = key.replace('vinh-qui', 'vinh-quy')
                if current_district == 'hoa-an':
                    key = key.replace('trung-vuong', 'truong-vuong')
                if current_district == 'dong-hoi':
                    key = key.replace('dong-son-tieu-khu-8-11', 'dong-son')
                if current_district == 'bo-trach':
                    key = key.replace('nong-truong-viet-trung', 'nt-viet-trung')
                if current_district == 'le-thuy':
                    key = key.replace('nong-truong-le-ninh', 'nt-le-ninh')
                if current_district == 'dong-xuan':
                    key = key.replace('xuan-long-1', 'xuan-long')
                if current_district == 'tuy-an':
                    key = key.replace('an-hai', 'an-hoa-hai')
                if current_district == 'son-hoa':
                    key = key.replace('eacharang', 'eacha-rang')
                if current_district == 'song-hinh':
                    key = key.replace('eaba', 'ea-ba')
                    key = key.replace('ea-bar', 'eabar')
                if current_district == 'tay-hoa':
                    key = key.replace('hoa-binh-2', 'phu-thu')
                if current_district == 'song-cong':
                    key = key.replace('vinh-son', 'chau-son')
                if current_district == 'vo-nhai':
                    key = key.replace('nghin-tuong', 'nghinh-tuong')
                if current_district == 'cang-long':
                    key = key.replace('phuong-thanh', 'fuong-thanh')
                if current_district == 'muong-nhe':
                    key = key.replace('huoi-lech', 'huoi-lenh')
                if current_district == 'tua-chua':
                    key = key.replace('sin-chai', 'xin-chai')
                    key = key.replace('ta-sinh-thang', 'ta-sin-thang')
                if current_district == 'dien-bien-dong':
                    key = key.replace('noong-u', 'nong-u')
                if current_district == 'ba-to':
                    key = key.replace('xaba-giang', 'ba-giang')
                if current_district == 'viet-tri':
                    key = key.replace('kcn-thuy-van', 'thuy-van')
                if current_district == 'doan-hung':
                    key = key.replace('hung-quan', 'hung-xuyen')
                    key = key.replace('huu-do', 'hop-nhat')
                    key = key.replace('phong-phu', 'phu-lam')
                if current_district == 'ha-hoa':
                    key = key.replace('phu-khanh', 'tu-hiep')
                if current_district == 'thanh-ba':
                    key = key.replace('nang-yen', 'quang-yen')
                if current_district == 'phu-ninh':
                    key = key.replace('vinh-phu', 'binh-phu')
                if current_district == 'cam-khe':
                    key = key.replace('song-thao', 'cam-khe')
                    key = key.replace('phuong-xa', 'minh-tan')
                    key = key.replace('cat-tru', 'hung-viet')
                    key = key.replace('truong-xa', 'chuong-xa')
                if current_district == 'tam-nong':
                    key = key.replace('huong-nha', 'bac-son')
                    key = key.replace('hong-da', 'dan-quyen')
                    key = key.replace('co-tiet', 'van-xuan')
                    key = key.replace('hung-do', 'lam-son')
                if current_district == 'lam-thao':
                    key = key.replace('son-duong', 'phung-nguyen')
                if current_district == 'thanh-thuy':
                    key = key.replace('trung-thinh', 'dong-trung')
                if current_district == 'hue':
                    key = key.replace('vy-da', 'vi-da')
                    key = key.replace('phuong-duc', 'fuong-duc')
                if current_district == 'phu-vang':
                    key = key.replace('vinh-phu', 'phu-gia')
                if current_district == 'huong-tra':
                    key = key.replace('binh-dien', 'binh-tien')
                if current_district == 'a-luoi':
                    key = key.replace('hong-trung', 'trung-son')
                    key = key.replace('hong-quang', 'quang-nham')
                    key = key.replace('huong-lam', 'lam-dot')
                if current_district == 'phu-loc':
                    key = key.replace('vinh-giang', 'giang-hai')
                if current_district == 'nam-dong':
                    key = key.replace('huong-hoa', 'huong-xuan')
                if current_district == 'tan-hong':
                    key = key.replace('tan-hoi-co', 'tan-ho-co')
                if current_district == 'lap-vo':
                    key = key.replace('an-dong', 'hoi-an-dong')
                if current_district == 'chau-thanh':
                    key = key.replace('phu-trung', 'tan-phu-trung')
                    key = key.replace('phu-thuan', 'an-phu-thuan')
                if current_district == 'cu-lao-dung':
                    key = key.replace('an-thanh-iii', 'an-thanh-3')
                    key = key.replace('an-thanh-ii', 'an-thanh-2')
                    key = key.replace('an-thanh-i', 'an-thanh-1')
                    key = key.replace('dai-an-ii', 'dai-an-2')
                    key = key.replace('dai-an-i', 'dai-an-1')
                if current_district == 'my-xuyen':                    
                    key = key.replace('hoa-tu-i', 'hoa-tu-1')
                    key = key.replace('hoa-tu-1i', 'hoa-tu-ii')
                    key = key.replace('gia-hoa-ii', 'gia-hoa-2')
                    key = key.replace('gia-hoa-i', 'gia-hoa-1')

                if current_district == 'nga-nam':                    
                    key = key.replace('nga-nam', '1')
                    key = key.replace('long-tan', '2')
                    key = key.replace('vinh-bien', '3')
                if current_district == 'vinh-chau':                    
                    key = key.replace('vinh-chau', '1')












                if ('xa-xa-' in key):
                    print(key)
                    key = key.replace('xa-xa-', 'xa-')
                elif ('phuong-phuong-' in key):
                    key = key.replace('phuong-phuong-', 'phuong-')                
                elif ('xa-phuong-' in key):
                    key = key.replace('xa-', '')
                elif ('phuong-xa-' in key):
                    key = key.replace('phuong-', '')
                elif ('phuong-' in key):
                    key = key.replace('phuong-', '')
                elif key =='xa-ho':
                    pass
                else:
                    key = key.replace('xa-', '')

                # Tây Ninh, Tinh Biên, nha-ban -> nha-bang
                if current_district == 'tinh-bien':
                    key = key.replace('nha-ban', 'nha-bang')
                if current_district == 'cang-long':
                    key = key.replace('fuong-thanh', 'phuong-thanh')
                if current_district == 'hue':
                    key = key.replace('fuong-duc', 'phuong-duc')
                

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
            elif current_district == 'hai-duong':
                if ward_slug == 'quyet-thang':
                    vtp_ward_code = 2047
            elif current_district == 'hung-yen':
                if ward_slug == 'hung-cuong':
                    vtp_ward_code = 2357
                if ward_slug == 'phu-cuong':
                    vtp_ward_code = 2358
                if ward_slug == 'hoang-hanh':
                    vtp_ward_code = 2373
                if ward_slug == 'phuong-chieu':
                    vtp_ward_code = 2370
                if ward_slug == 'tan-hung':
                    vtp_ward_code = 2371
            elif current_district == 'thai-binh':
                if ward_slug == 'dong-tho':
                    vtp_ward_code = 2991
                if ward_slug == 'dong-my':
                    vtp_ward_code = 2986
            elif current_district == 'tuyen-quang':
                if ward_slug == 'kim-phu':
                    vtp_ward_code = 4235
            elif current_district == 'quang-ngai':
                if (ward_slug == 'tinh-hoa'):
                    vtp_ward_code = ward_dict['tran-phu']
            elif current_district == 'vinh-chau':
                if (ward_slug == '2'):
                    vtp_ward_code = ward_dict['1']
            elif current_district == 'tuy-hoa':
                if ward_slug == 'an-phu':
                    vtp_ward_code = 8596
            elif current_district == 'thai-nguyen':
                if ward_slug == 'son-cam':
                    vtp_ward_code = 4631
            elif current_district == 'bac-yen':
                if ward_slug == 'hang-dong':
                    vtp_ward_code = ward_dict['ta-xua']
            elif current_district == 'dien-bien-phu':
                if (ward_slug == 'na-tau') | (ward_slug == 'muong-phang') | (ward_slug == 'pa-khoang'):
                    vtp_ward_code = ward_dict['him-lam']
            elif current_district == 'dong-hung':
                if ward_slug == 'dong-quan':
                    vtp_ward_code = ward_dict['dong-huy']
            elif current_district == 'chau-thanh':
                if ward_slug == 'an-ninh':
                    vtp_ward_code = ward_dict['phu-tam']
            elif current_district == 'dien-khanh':
                if ward_slug == 'suoi-tien':
                    vtp_ward_code = ward_dict['suoi-hiep']
            elif current_district == 'ha-long':
                if ward_slug == 'hoanh-bo':
                    vtp_ward_code = ward_dict['hung-thang']
            elif current_district == 'duyen-hai':
                if (ward_slug == 'don-xuan') | (ward_slug == 'don-chau'):
                    vtp_ward_code = ward_dict['hiep-thanh']
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
    # data = {}
    # data['VN'] = vn_data

    # file_name = 'json/' + state_name_slug + '.json'
    # with open(file_name, 'w', encoding='utf-8') as f:
    #     json.dump(vn_data, f, ensure_ascii=False, indent=4)

data = {}
data['VN'] = vn_data
with open('vn.json', 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=4)
