# 프로젝트에 필요한 패키지 불러오기
from bs4 import BeautifulSoup as bs
import requests
import pandas as pd

def checkIntegrity(arr, idx):
    if str(type(arr[idx]))!="<class 'int'>" or str(type(arr[idx+1]))!="<class 'int'>" or str(type(arr[idx+2])) != "<class 'int'>" :
        return False
    return True

# 엑셀파일 불러오기
df = pd.read_excel('copList.xlsx', sheet_name='copList', dtype=object)
wdf = pd.DataFrame(columns=['회사명', 'Y-3', 'Y-2', 'Y-1', 'Y', 'MAvg'])

idx = 0
# 행 단위로 처리
for name, code in zip(df['회사명'], df['종목코드']):
    # 입력받은 query가 포함된 url 주소 저장
    url = 'https://finance.naver.com/item/main.naver?code='+ code

    # requests 패키지를 이용해 'url'의 html 문서 가져오기
    response = requests.get(url)
    html_text = response.text

    # beautifulsoup 패키지로 파싱 후, 'soup' 변수에 저장
    soup = bs(response.text, 'html.parser')

    print(name)
    # 기업실적 테이블 추출
    errFlag = False
    salesList = []
    salesRow = soup.select("#content > div.section.cop_analysis > div.sub_section > table > tbody > tr:nth-child(1)")
    if not(salesRow):
        continue
    for i in range(4):
        try:
            strX = salesRow[0].select('td')[i].text.strip("\n\t\xa0").replace(",","")
            salesList.append(int(strX)) if strX != "" and strX != '-' else salesList.append(strX)
        except:
            errFlag = True
            continue
    if errFlag:
        continue

    # for i in range(4):
    #     print(salesList[i])

    profitList = []
    profitRow = soup.select("#content > div.section.cop_analysis > div.sub_section > table > tbody > tr:nth-child(2)")
    for i in range(4):
        strX = profitRow[0].select('td')[i].text.strip("\n\t\xa0").replace(",","")
        profitList.append(int(strX)) if strX != "" and strX != '-' else profitList.append(strX)

    # for i in range(4):
    #     print(profitList[i])

    incomeList = []
    incomeRow = soup.select("#content > div.section.cop_analysis > div.sub_section > table > tbody > tr:nth-child(3)")
    for i in range(4):
        strX = incomeRow[0].select('td')[i].text.strip("\n\t\xa0").replace(",","")
        incomeList.append(int(strX)) if strX != "" and strX != '-' else incomeList.append(strX)
    # for i in range(4):
    #     print(incomeList[i])

    marginList = []
    marginRow = soup.select("#content > div.section.cop_analysis > div.sub_section > table > tbody > tr:nth-child(4)")
    for i in range(4):
        strX = marginRow[0].select('td')[i].text.strip("\n\t\xa0").replace(",","")
        marginList.append(float(strX)) if len(strX) > 2 else marginList.append(strX)
    # for i in range(4):
    #     print(marginList[i])

    newRow = []
    #print(type(salesList[3]))
    sInx = 0 if str(type(salesList[3]))=="<class 'str'>" else 1
    #print(sInx)
    if not(checkIntegrity(salesList, sInx)):
        continue
    if salesList[sInx] < salesList[sInx+1] and salesList[sInx+1] < salesList[sInx+2]:
        if profitList[sInx] < profitList[sInx+1] and profitList[sInx+1] < profitList[sInx+2]:
            newRow.append(name)
            newRow[1:] = salesList

            # sInx = 0 if str(type(marginList[3])) == "<class 'str'>" else 1
            # if not(checkIntegrity(marginList, sInx)):
            #     continue
            marginAvg = (marginList[sInx] + marginList[sInx+1] + marginList[sInx+2])/3
            if marginAvg < 10:
                continue
            newRow.append(marginAvg)

            wdf.loc[idx] = newRow
            idx = idx + 1

wdf.to_excel('result.xlsx')
