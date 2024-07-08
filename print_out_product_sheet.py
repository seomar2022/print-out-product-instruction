#각 상품 코드와 관리용 상품명 연결
#설명지 한장씩 저장. 이름을 상품 코드로 지정
#전채널 주문리스트를 입력. 그 중 상품코드 열만 가지고 오기.
#필요한 설명지를 순서대로 pdf하나로 병합
#인쇄 기능 넣어야하나?

#https://chuun92.tistory.com/10
#https://pypdf.readthedocs.io/en/stable/user/merging-pdfs.html
#https://wikidocs.net/226862
#https://wikidocs.net/153818

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
        print("kakao")

#### PDF 파일 병합
merge_pdf = PdfWriter()
file_not_found = []

for converted_code in converted_codes:
    try:
        merge_pdf.append(f"sheets\\{converted_code}.pdf")
    except FileNotFoundError:
        file_not_found.append(converted_code)

####설명지없이 출고되는 상품인지 확인
#일단 보류

# 현재 날짜 가져오기
now = datetime.now().strftime("%m.%d.%a") #월.일.요일

merge_pdf.write(f"{now}_product_sheet.pdf")
merge_pdf.close()

####설명지 없는 파일 알려주기
pyautogui.alert(file_not_found)

####pywin32로 프린트
