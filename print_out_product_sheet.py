#각 상품 코드와 관리용 상품명 연결
#설명지 한장씩 저장. 이름을 상품 코드로 지정
#전채널 주문리스트를 입력. 그 중 상품코드 열만 가지고 오기.
#상품코드의 순서대로 병합 
#https://chuun92.tistory.com/10
#https://pypdf.readthedocs.io/en/stable/user/merging-pdfs.html
#https://wikidocs.net/226862
#https://wikidocs.net/153818

from pypdf import PdfWriter #pip install pypdf
from datetime import datetime
import pandas as pd #pip install pandas openpyxl

# 엑셀 파일을 읽어오기
df = pd.read_csv('data.csv', usecols=['상품코드'])

codes = df['상품코드'].tolist()
#print(name_array)

# PDF 파일 병합
merger = PdfWriter()

for code in codes:
    merger.append(f"{code}.pdf")


# 현재 날짜 가져오기
now = datetime.now().strftime("%m.%d.%a") #월.일.요일

#merger.write(f"{now}_product_sheet.pdf")
merger.close()

#pypdf가 나을까 pywin32가 나을까?
# 사용 예시
#print_pdf("C:\\Users\\User\\Desktop\\print-out-product-sheet\\1.pdf")


#필요한 설명지를 순서대로 pdf하나로 병합해서 프린트