from bs4 import BeautifulSoup
import requests
import time
import win32com.client

# 종목코드엑셀 주소 http://marketdata.krx.co.kr/mdi#document=040601
# 기존 SRIMevaluation원본_211121에서는 ROE컨센이 없을 때 전년도 ROE를 이용했지만, 여기에서는 컨센이 없으면 건너뛰었다. 속도향상 위해
# 달라진 곳 110~112 추가, 141~147 제거

# required information
#-----------------------------------------------------------
# 목록 파일
STOCK_LIST_FILE = 'KOSDAQ220224.xlsx'
# 결과 입력 받을 파일
RESULT_FILE = 'KOSDAQ230116_SRIMevaluation.xlsx'
# stock 개수 - KOSIP:837, KOSDAQ:1426
NUMBER_Of_STOCKS = 837
# 할인률
INTEREST_RATE = 11
# 가격 낮추기(price * PRICE_DISCOUNT)
PRICE_DISCOUNT = 1
#-----------------------------------------------------------


def get_url(company_code):
    url = "http://comp.fnguide.com/SVO2/ASP/SVD_main.asp?pGB=1&gicode=A" + company_code + "&cID=&MenuYn=Y&ReportGB=D&NewMenuID=Y&stkGb=701"
    result = requests.get(url)
    bs_obj = BeautifulSoup(result.content, "html.parser", from_encoding='cp949')
    return bs_obj


def get_url2(company_code):
    url = "https://finance.naver.com/item/main.nhn?code=" + company_code
    result = requests.get(url)
    bs_obj = BeautifulSoup(result.content, "html.parser", from_encoding='cp949')
    return bs_obj


def get_num(company_code):
    bs_obj = get_url(company_code)

    table1 = bs_obj.find("div", {"class": "ul_col2_r"})
    tr_myso = table1.find_all('tr')[5]
    td_myso = tr_myso.find_all('td')[1]

    table2 = bs_obj.find("div", {"id": "div15"})
    tr_toi = table2.find_all('tr')[11]
    td_toi = tr_toi.find_all('td')[2]  # 2:2019년, 3:2020년

    # 1년후 예상 ROE - ROE 2022/12(E)
    tr_roe = table2.find_all('tr')[19]
    td_roe = tr_roe.find_all('td')[3]  # 2:2019년, 3:2020년

    # 2년후 예상 ROE - 연결/연간 ROE 2023/12(E)
    #tr_roe = table2.find_all('tr')[46]
    #td_roe = tr_roe.find_all('td')[6]

    tr_so = table2.find_all('tr')[25]
    td_so = tr_so.find_all('td')[2]  # 2:2019년, 3:2020년

    return (td_myso.text, td_toi.text, td_roe.text, td_so.text)


def get_roe2(company_code):
    bs_obj = get_url(company_code)

    table2 = bs_obj.find("div", {"id": "div15"})
    tr_roe = table2.find_all('tr')[19]
    td_roe = tr_roe.find_all('td')[2]  # 2:2019년, 3:2020년

    return (td_roe.text)


def get_price(company_code):
    bs_obj = get_url2(company_code)
    no_today = bs_obj.find("p", {"class": "no_today"})
    blind = no_today.find("span", {"class": "blind"})
    now_price = blind.text
    return now_price


def get_name(company_code):
    bs_obj = get_url2(company_code)
    no_today = bs_obj.find("div", {"class": "wrap_company"})
    blind = no_today.find("a")
    now_price = blind.text
    return now_price


code_num_list = []
#company_detail_list = []

excel = win32com.client.Dispatch("Excel.Application")  # 코스피 코스닥 정보
wb1 = excel.Workbooks.Open('C:\\CodingProgram\\pycharm\\SRIMresult\\' + STOCK_LIST_FILE)  # KOSPI:KOSPI201010.xlsx

ws1 = wb1.ActiveSheet  # KOSDAQ:KOSDAQ201010.xlsx

i = 2
while i < NUMBER_Of_STOCKS:
    code_num_list.append(ws1.Cells(i, 2).Value)
    #company_detail_list.append(ws1.Cells(i, 5).Value)
    i = i + 1

excel.Visible = True  # 정보 저장 위치
wb2 = excel.Workbooks.Open('C:\\CodingProgram\\pycharm\\SRIMresult\\' + RESULT_FILE)
ws2 = wb2.ActiveSheet  # KOSPI:KOSPI200806_SRIMevaluation.xlsx
# KOSDAQ:KOSDAQ200806_SRIMevaluation.xlsx

# 계산 정보 출력하기
ws2.Cells(3, 7).Value = "적용 할인율: " + str(INTEREST_RATE) + "%"
ws2.Cells(4, 7).Value = "가격 낮추기: 가격 x " + str(PRICE_DISCOUNT)

i = 0
k = 11
load = 1

while i < NUMBER_Of_STOCKS - 2:  # KOSIP:836, KOSDAQ:1424

    if load % 10 == 0:
        print(load)
    load = load + 1
    roeCheck = ""

    try:
        (myso, toi, roe, so) = get_num(code_num_list[i])
        time.sleep(0.1)

        if len(roe) == 1: # ROE컨센이 없으면 건너뛰기 -> 바뀐점!!!!
            i = i + 1
            continue

    except:
        print(i + 1, "get_num")  # 정보없음
        i = i + 1
        continue

    try:
        price = get_price(code_num_list[i])
        time.sleep(0.1)
    except:
        print(i + 1, "price")  # 정보없음
        i = i + 1
        continue

    try:
        name = get_name(code_num_list[i])
    except:
        print(i + 1, "name")  # 정보없음
        i = i + 1
        continue

    try:

        if len(myso) == 1:
            myso = "0"
        if len(toi) == 1:
            toi = "0"

        # 2020년 ROE없으면 2019년 ROE사용!!!
        #if len(roe) == 1:
        #    roe = get_roe2(code_num_list[i]) #2019 ROE 가져오기
        #    roeCheck = "19"
        #    time.sleep(0.2)
        #    if len(roe) == 1:
        #        roe = "0"

        if len(so) == 1:
            so = "0"
        if roe == '완전잠식':
            roe = "0"
    except:
        print(i + 1, "error")
        i = i + 1
        continue

    try:
        myso_num = float(myso.replace(',', ''))
        toi_num = float(toi.replace(',', ''))
        roe_num = float(roe.replace(',', ''))
        so_num = float(so.replace(',', ''))
        price_num = float(price.replace(',', ''))
    except:
        print(i + 1, "error")  # 정보없음
        i = i + 1
        continue


    try:
        # time.sleep(0.3)
        if roe_num > INTEREST_RATE and (
                toi_num * 100000000 + toi_num * 100000000 * (roe_num - INTEREST_RATE) / 100 * 0.8 / (
                1 + (INTEREST_RATE / 100) - 0.8)) / (so_num * 1000 - myso_num) > price_num * PRICE_DISCOUNT:
            ws2.Cells(k, 6).Value = name
            #ws2.Cells(k, 7).Value = company_detail_list[i]
            ws2.Cells(k, 8).Value = code_num_list[i]
            ws2.Cells(k, 9).Value = price_num
            ws2.Cells(k, 10).Value = toi
            ws2.Cells(k, 11).Value = roe
            ws2.Cells(k, 12).Value = so
            ws2.Cells(k, 13).Value = myso
            #ws2.Cells(k, 14).Value = roeCheck
            # print(toi + so + myso + roe)
            # print((toi_num*100000000+toi_num*100000000*(roe_num-INTEREST_RATE)/100*0.8/(1+(INTEREST_RATE/100)-0.8))/(so_num*1000-myso_num))
            k = k + 1
    except:
        print(i + 1, "error")
        i = i + 1
        continue


    i = i + 1

wb2.Save()
excel.Quit()
