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

# 전채널 주문리스트 파일을 읽어오기
order_list = pd.read_csv('data.csv', usecols=['상품코드'])
#상품코드 열의 데이터들을 list에 넣기
codes = order_list['상품코드'].tolist()

#카페24의 상품코드만 가져오기위해서 P00로 시작하는 문자열 코드만 남겨둔다.
#걸러지는 데이터) 스마트스토어 상품 코드(9, 10으로 시작하는 숫자), 톡스토어 상품코드(3으로 시작하는 숫자), 비어있는 셀(nan)
#isinstance(item, str) ->item이 str이면 true를 반환
#cafe24_codes = [item for item in codes if isinstance(item, str) and item.startswith("P000")]


#
product_code_mapping = pd.read_csv("product_code_mapping.xslx", engine='openpyxl')

# PDF 파일 병합
merge_pdf = PdfWriter()
file_not_found = []


for code in codes:
    if code.startswith("P00") : #카페24
        try:
            merge_pdf.append(f"sheets\\{code}.pdf")
        except FileNotFoundError:
            file_not_found.append(code)
    elif code.startswith("9") or code.startswith("1") : #네이버
        #상품코드 맵핑된 엑셀파일에서 네이버 상품코드에 해당하는 카페24상품코드 가져오기

        print("네이버")
    elif code.startswith("3") : #카카오
        print("카카오")


# 현재 날짜 가져오기
now = datetime.now().strftime("%m.%d.%a") #월.일.요일

merge_pdf.write(f"{now}_product_sheet.pdf")
merge_pdf.close()

pyautogui.alert(file_not_found)

#pywin32로 프린트

#네이버의 상품코드(상품번호) 입력 시, 카페24 상품코드 출력. 
