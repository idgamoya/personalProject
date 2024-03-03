# 프로젝트에 필요한 패키지 불러오기
from bs4 import BeautifulSoup as bs
import requests
import pandas as pd
from datetime import datetime

def checkIntegrity(arr, idx):
    if str(type(arr[idx]))!="<class 'int'>" or str(type(arr[idx+1]))!="<class 'int'>" or str(type(arr[idx+2])) != "<class 'int'>" :
        return False
    return True

# 엑셀파일 불러오기
df = pd.read_excel('coList.xlsx', sheet_name='상장법인목록', dtype=object)
#df = pd.read_excel('coList.xlsx', sheet_name='테스트', dtype=object)
wdf = pd.DataFrame(columns=['회사명', '매출[Y-2]', '매출[Y-1]', '매출[Y]',
                            '영업이익[Y-2]', '영업이익[Y-1]', '영업이익[Y]',
                            '당기순이익[Y-2]', '당기순이익[Y-1]', '당기순이익[Y]',
                            '최근분기영업이익률', 'PER', '배당수익률'])

idx = 0
debug = 0
# 행 단위로 처리
for name, code in zip(df['회사명'], df['종목코드']):
    # 입력받은 query가 포함된 url 주소 저장
    url = 'https://finance.naver.com/item/main.naver?code='+ code

    # requests 패키지를 이용해 'url'의 html 문서 가져오기
    response = requests.get(url)

    # beautifulsoup 패키지로 파싱 후, 'soup' 변수에 저장
    soup = bs(response.text, 'html.parser')

    print(name)
    # 기업실적 테이블 추출
    errFlag = False
    salesList = []
    quaterList = []
    salesRow = soup.select("#content > div.section.cop_analysis > div.sub_section > table > tbody > tr:nth-child(1)")
    if not(salesRow):
        continue
    # 4개 연도 매출액을 추출한다.
    for i in range(4):
        try:
            strX = salesRow[0].select('td')[i].text.strip("\n\t\xa0").replace(",","")
            salesList.append(int(strX)) if strX != "" and strX != '-' else salesList.append(strX)
        except:
            errFlag = True
            continue
    if errFlag:
        continue

    # 4번째 연도 매출액이 없으면
    lastValueList = []
    lastValue = 0
    if str(type(salesList[3]))=="<class 'str'>":
        for i in range(4, 10):
            strX = salesRow[0].select('td')[i].text.strip("\n\t\xa0").replace(",", "")
            if strX != "" and strX != '-':
                lastValueList.append(int(strX))
        #print(lastValueList)
        sInx = len(lastValueList) - 4
        if sInx >= 0:
            for i in range(4):
                lastValue += lastValueList[sInx+i]
            #print("Last Value = ", lastValue)
        salesList[3] = lastValue
    #print(salesList)

    # 4개 연도 영업이익을 추출한다.
    profitList = []
    profitRow = soup.select("#content > div.section.cop_analysis > div.sub_section > table > tbody > tr:nth-child(2)")
    for i in range(4):
        strX = profitRow[0].select('td')[i].text.strip("\n\t\xa0").replace(",","")
        profitList.append(int(strX)) if strX != "" and strX != '-' else profitList.append(strX)

    # 4번째 연도 영업이익이 없으면
    lastValueList = []
    lastValue = 0
    if str(type(profitList[3]))=="<class 'str'>":
        # 4개 분기 영엽이익 추출
        for i in range(4, 10):
            strX = profitRow[0].select('td')[i].text.strip("\n\t\xa0").replace(",", "")
            if strX != "" and strX != '-':
                lastValueList.append(int(strX))
        if debug: print("분기 영업이익 리스트 =", lastValueList)
        sInx = len(lastValueList) - 4
        if sInx >= 0:
            for i in range(4):
                lastValue += lastValueList[sInx+i]
            if debug: print("최근 4개분기 영업이익 =", lastValue)
        profitList[3] = lastValue
    if debug: print("연간 영업이익 리스트 =", profitList)

    # 4개 연도 당기순이익을 추출한다.
    incomeList = []
    incomeRow = soup.select("#content > div.section.cop_analysis > div.sub_section > table > tbody > tr:nth-child(3)")
    for i in range(4):
        strX = incomeRow[0].select('td')[i].text.strip("\n\t\xa0").replace(",","")
        incomeList.append(int(strX)) if strX != "" and strX != '-' else incomeList.append(strX)
    # for i in range(4):
    #     print(incomeList[i])

    # 4번째 연도 당기순이익이 없으면
    lastValueList = []
    lastValue = 0
    if str(type(incomeList[3]))=="<class 'str'>":
        # 4개 분기 당기 순이익 추출
        for i in range(4, 10):
            strX = incomeRow[0].select('td')[i].text.strip("\n\t\xa0").replace(",", "")
            if strX != "" and strX != '-':
                lastValueList.append(int(strX))
        if debug: print("분기 당기순이익 리스트 =", lastValueList)
        sInx = len(lastValueList) - 4
        if sInx >= 0:
            for i in range(4):
                lastValue += lastValueList[sInx+i]
            if debug: print("최근 4개분기 당기순이익 =", lastValue)
        incomeList[3] = lastValue
        if debug: print("연간 당기순이익 리스트 =", incomeList)

    # 4개 연도 영엽이익률을 추출한다.
    marginList = []
    marginRow = soup.select("#content > div.section.cop_analysis > div.sub_section > table > tbody > tr:nth-child(4)")
    for i in range(4):
        strX = marginRow[0].select('td')[i].text.strip("\n\t\xa0").replace(",","")
        if strX == '-': strX = "0"
        marginList.append(float(strX)) if strX != "" else marginList.append(strX)

    # 4번째 연도 영업이익률이 없으면
    lastValueList = []
    lastValue = 0
    if str(type(marginList[3]))=="<class 'str'>":
        for i in range(4, 10):
            strX = marginRow[0].select('td')[i].text.strip("\n\t\xa0").replace(",", "")
            if strX != "" and strX != '-':
                lastValueList.append(float(strX))
        if debug: print("분기 영업이이귤 =", lastValueList)
        sInx = len(lastValueList) - 4
        if sInx >= 0:
            for i in range(4):
                lastValue += lastValueList[sInx+i]
            if debug: print("최근 4개 분기 영업이익률 = ", lastValue/4)
        marginList[3] = lastValue/4
        if debug: print("연간 영업이익률 리스트 =", marginList)

    # PER 추출
    per_tag = soup.select_one("#_per")
    per = -999
    if per_tag:
        per = float(per_tag.get_text().replace(",",""))
    # strX = per_tag[0].select('td')[0].text.strip("\n\t\xa0").replace(",", "")
    # print("PER = ", strX)

    # 추정 PER 추출
    cns_per_tag = soup.select_one("#_cns_per")
    if cns_per_tag:     # 추정 PER이 있다면 PER 대체
        per = float(cns_per_tag.get_text().replace(",", ""))

    # 배당 수익률 추출
    allocRateTag = soup.select_one("#_dvr")
    allocRate = -99.9
    if allocRateTag:
        allocRate = float(allocRateTag.get_text().replace(",", ""))

    newRow = []
    #print(type(salesList[3]))
    # 4개 연도 중 추출한 3개 연도 시작점 결정
    sInx = 0 if str(type(salesList[3]))=="<class 'str'>" else 1
    #print(sInx)
    if not(checkIntegrity(salesList, sInx)):
        continue
    # 매출액이 최근 3개년도 동안 계속 상승하고
    if salesList[sInx] <= salesList[sInx+1] and salesList[sInx+1] <= salesList[sInx+2]:
        # 영업이익이 최근 3개년도 동안 계속 상승하면
        if profitList[sInx] <= profitList[sInx+1] and profitList[sInx+1] <= profitList[sInx+2]:
            #newRow.append(code)
            newRow.append(name)
            newRow[1:4] = salesList[sInx:sInx+3]
            newRow[4:7] = profitList[sInx:sInx+3]
            newRow[7:10] = incomeList[sInx:sInx+3]

            # 최근 분기 영업이익률 추가
            newRow.append(marginList[sInx+2])
            # PER 추가
            newRow.append(per)
            if debug: print(newRow)

            # 배당수익률 추가
            newRow.append(allocRate)

            # 영업이익률이 0보다 커야 기록
            if marginList[sInx+2] > 0 and per > 0 and allocRate > 0:
                wdf.loc[idx] = newRow
                idx = idx + 1




wdf['이익률순위'] = wdf['최근분기영업이익률'].rank(method='dense', ascending=False)
wdf['PER순위'] = wdf['PER'].rank(method='dense', ascending=True)
wdf['배당수익률순위'] = wdf['배당수익률'].rank(method='dense', ascending=False)
wdf['순위합산값'] = wdf['이익률순위'] + wdf['PER순위'] + wdf['배당수익률순위']
wdf['합산순위'] = wdf['순위합산값'].rank(method='dense', ascending=True)
sorted_df = wdf.sort_values(by='순위합산값', ascending=True)

now = datetime.now()
wfName = 'result' + now.strftime('%Y%m%d%H%M%S') + '.xlsx'  # 쓰기 파일 이름에 오늘날짜시간 붙이기
sorted_df.to_excel(wfName)



