from pypdf import PdfWriter #pip install pypdf #pdf병합기능 쓰기 위해.
from datetime import datetime #병합된 pdf이름에 오늘 날짜 쓰기 위해.
import pandas as pd #pip install pandas openpyxl #엑셀의 데이터를 읽어오기 위해.
import pyautogui #pip install pyautogui

####엑셀 파일 읽어오기
# 전채널 주문리스트 파일을 읽어오기
order_list = pd.read_csv('data.csv')

# 상품코드 열의 데이터를 문자열로 변환하고 NaN 값을 빈 문자열로 대체
codes = order_list['상품코드'].astype(str).fillna("").tolist()

#카페24상품코드와 네이버 상품코드를 매핑한 엑셀파일 읽어오기
product_code_mapping = pd.read_excel("product_code_mapping.xlsx", engine='openpyxl')

product_code_mapping['naver_code'] = product_code_mapping['naver_code'].astype(str).str.strip().str.replace('-', '')
product_code_mapping['kakao_code'] = product_code_mapping['kakao_code'].astype(str).str.strip().str.replace('-', '')
product_code_mapping['cafe24_code'] = product_code_mapping['cafe24_code'].astype(str).str.strip()


####상품코드를 카페24의 코드로 통일
def convert_to_cafe24(code, column):
    # 'column' 열에 'code'가 있는 행 ex) naver_code열에서 9708250509가 있는 행
    #result = product_code_mapping.query(f"{column} == {code}")
    result = product_code_mapping.query(f"{column} == @code")
    return result['cafe24_code'].iat[0]

converted_codes = []

for code in codes:
    if code.startswith("P00") : #카페24
        converted_codes.append(code)
    elif code.startswith("9") or code.startswith("1") : #네이버
        #상품코드 맵핑된 엑셀파일에서 네이버 상품코드에 해당하는 카페24상품코드 가져오기
        result = convert_to_cafe24(code, "naver_code")
        converted_codes.append(result)
    elif code.startswith("3") : #카카오
        result = convert_to_cafe24(code, "kakao_code")
        converted_codes.append(result)

#### PDF 파일 병합
merge_pdf = PdfWriter()
not_founded_files = {}

for converted_code in converted_codes:
    try:
        merge_pdf.append(f"sheets\\{converted_code}.pdf")
    except FileNotFoundError:
        not_founded_files[converted_code] = ''

####설명지없이 출고되는 상품인지 확인
#일단 보류

# 현재 날짜 가져오기
now = datetime.now().strftime("%m.%d.%a") #월.일.요일

merge_pdf.write(f"{now}_product_sheet.pdf")
merge_pdf.close()

####설명지 없는 상품코드와 상품명 알려주기
#설명지 없는 상품의 코드를 전채널주문리스트에서 찾고, 해당 상품의 이름을 가져와서 딕셔너리의 값으로 넣기
#매크로 돌리기 전의 열이름은 '상품명(한국어 쇼핑몰)' 돌린 후는 '상품명'이라서 상품명이 포함된 열을 지정
product_name_col = [col for col in order_list.columns if "상품명" in col]
for key in not_founded_files:
    key_in_order_list = order_list.query(f"상품코드 == @key")
    not_founded_files[key] = key_in_order_list[product_name_col[0]].iat[0]
    

#csv파일로 저장
converted_codes_df = pd.DataFrame(list(not_founded_files.items()), columns=['상품코드', '상품명'])
converted_codes_df.to_csv(f"{now}_not_founded_files.csv", index=False, encoding='utf-8-sig')

####pyautogui로 프로그램 실행 결과 알려주기
if len(not_founded_files) == 0:
    alert_msg = "모든 상품의 설명지를 찾았습니다!:)"
else:
    alert_msg=f"{len(not_founded_files)}개의 설명지를 찾지 못했습니다:(\nnot_founded_files을 확인해주세요"
pyautogui.alert(text=alert_msg, title='실행 결과!', button='OK')
