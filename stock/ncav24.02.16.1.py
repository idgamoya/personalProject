# 프로젝트에 필요한 패키지 불러오기
from bs4 import BeautifulSoup as bs
import requests
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import re

driver = webdriver.Chrome()

def checkIntegrity(arr, idx):
    if str(type(arr[idx]))!="<class 'int'>" or str(type(arr[idx+1]))!="<class 'int'>" or str(type(arr[idx+2])) != "<class 'int'>" :
        return False
    return True

# 엑셀파일 불러오기
df = pd.read_excel('coList.xlsx', sheet_name='상장법인목록', dtype=object)
#df = pd.read_excel('coList.xlsx', sheet_name='테스트', dtype=object)
wdf = pd.DataFrame(columns=['회사명', '유동자산', '부채총액', '시가총액','최근분기순이익'])

idx = 1
# 행 단위로 처리
for name, code in zip(df['회사명'], df['종목코드']):
    try:
        # 입력받은 query가 포함된 url 주소 저장
        url = 'https://navercomp.wisereport.co.kr/v2/company/c1030001.aspx?cmp_cd='+ code + '&cn='
        print(url)

        driver.get(url)
        driver.find_element(By.XPATH, '//*[@id="rpt_tab2"]').click()
        time.sleep(1)

        html_text = driver.page_source

        # beautifulsoup 패키지로 파싱 후, 'soup' 변수에 저장
        soup = bs(html_text, 'html.parser')
        print(name)
        # 유동자산 추출
        ext4List = soup.select('td.num.ext4')
        assets = ext4List[1].attrs['title'].replace(",", "")
        totalDebt = ext4List[126].attrs['title'].replace(",", "")
        print("유동자산 = " + assets)
        print("부채총계 = " + totalDebt)

        url = 'https://finance.naver.com/item/main.naver?code=' + code

        # requests 패키지를 이용해 'url'의 html 문서 가져오기
        response = requests.get(url)
        html_text = response.text

        # beautifulsoup 패키지로 파싱 후, 'soup' 변수에 저장
        soup = bs(response.text, 'html.parser')
        marketCapList = soup.select("#_market_sum")
        marketCap = marketCapList[0].get_text().replace(",","")
        marketCap = re.sub('[\t\n]', '', marketCap)
        marketCap = marketCap.strip()
        if "조" in marketCap:
            marketCap = marketCap.replace("조", "")
        print("시가총액 = " + marketCap)

        netProfitList = soup.select('#content div.section.cop_analysis div.sub_section table tbody tr:nth-child(3) td.last.cell_strong')
        netProfit = netProfitList[0].get_text().replace(",","")
        netProfit = re.sub('[\t\n]', '', netProfit)
        netProfit = netProfit.strip()
        #print("1: ",netProfit, len(netProfit))
        if len(netProfit) == 0:
            netProfitList = soup.select(
                        '#content div.section.cop_analysis div.sub_section table tbody tr:nth-child(3) td:nth-child(10)')
            netProfit = netProfitList[0].get_text().replace(",", "")
            netProfit = re.sub('[\t\n]', '', netProfit)
            netProfit = netProfit.strip()
            #print("2: ", netProfit)
        # try:
        #     netProfit = float(netProfit)
        # except ValueError:
        #     netProfitList = soup.select(
        #         '#content div.section.cop_analysis div.sub_section table tbody tr:nth-child(3) td:nth-child(10) em')
        #     netProfit = netProfitList[0].get_text().replace(",", "")
        #     netProfit = re.sub('[\t\n]', '', netProfit)
        #     print("2nd: " + netProfit)
        #     try:
        #         netProfit = float(netProfit)
        #         print(netProfit)
        #     except ValueError:
        #         print("error")
        #         continue

        if len(netProfit) > 0:
            print("최근분기실적 =", netProfit)
        else:
            print("최근분기실적: 없음")

        newRow = []
        if assets=="" or totalDebt=="" or netProfit=="":
            print("데이터 누락")
            continue
        try:
            if (float(assets)-float(totalDebt) > (float(marketCap)*1.5)) and (int(netProfit)>0):
                newRow.append(name)
                newRow.append(assets)
                newRow.append(totalDebt)
                newRow.append(marketCap)
                newRow.append(netProfit)
                wdf.loc[idx] = newRow
                idx = idx + 1
                print("Hit!")
            else:
                print("Not")
        except ValueError:
            print("데이터 변환 에러")
            continue
    except:
        continue

# tab_con1 > div.first > table > tbody > tr.strong > td
    # if not(assetsRow):
    #     continue
    # for i in range(5):
    #     try:
    #         strX = assetsRow[0].select('td')[i].text.strip("\n\t\xa0").replace(",","")
    #         assestsList.append(int(strX)) if strX != "" and strX != '-' else assestsList.append(strX)
    #     except:
    #         errFlag = True
    #         continue
    # if errFlag:
    #     continue
    #
    # for i in range(5):
    #     print(assestsList[i])
#
#     profitList = []
#     profitRow = soup.select("#content > div.section.cop_analysis > div.sub_section > table > tbody > tr:nth-child(2)")
#     for i in range(4):
#         strX = profitRow[0].select('td')[i].text.strip("\n\t\xa0").replace(",","")
#         profitList.append(int(strX)) if strX != "" and strX != '-' else profitList.append(strX)
#
#     # for i in range(4):
#     #     print(profitList[i])
#
#     incomeList = []
#     incomeRow = soup.select("#content > div.section.cop_analysis > div.sub_section > table > tbody > tr:nth-child(3)")
#     for i in range(4):
#         strX = incomeRow[0].select('td')[i].text.strip("\n\t\xa0").replace(",","")
#         incomeList.append(int(strX)) if strX != "" and strX != '-' else incomeList.append(strX)
#     # for i in range(4):
#     #     print(incomeList[i])
#
#     marginList = []
#     marginRow = soup.select("#content > div.section.cop_analysis > div.sub_section > table > tbody > tr:nth-child(4)")
#     for i in range(4):
#         strX = marginRow[0].select('td')[i].text.strip("\n\t\xa0").replace(",","")
#         marginList.append(float(strX)) if len(strX) > 2 else marginList.append(strX)
#     # for i in range(4):
#     #     print(marginList[i])
#
#     newRow = []
#     #print(type(salesList[3]))
#     sInx = 0 if str(type(salesList[3]))=="<class 'str'>" else 1
#     #print(sInx)
#     if not(checkIntegrity(salesList, sInx)):
#         continue
#     if salesList[sInx] < salesList[sInx+1] and salesList[sInx+1] < salesList[sInx+2]:
#         if profitList[sInx] < profitList[sInx+1] and profitList[sInx+1] < profitList[sInx+2]:
#             #newRow.append(code)
#             newRow.append(name)
#             newRow[1:] = salesList
#
#             # sInx = 0 if str(type(marginList[3])) == "<class 'str'>" else 1
#             # if not(checkIntegrity(marginList, sInx)):
#             #     continue
#             marginAvg = (marginList[sInx] + marginList[sInx+1] + marginList[sInx+2])/3
#             if marginAvg < 10:
#                 continue
#             newRow.append(marginAvg)
#
#             wdf.loc[idx] = newRow
#             idx = idx + 1
#
now = datetime.now()
wfName = 'ncav' + now.strftime('%Y%m%d%H%M%S') + '.xlsx'  # 쓰기 파일 이름에 오늘날짜시간 붙이기
wdf.to_excel(wfName)
print("Writing Done")

