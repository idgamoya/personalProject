import pandas as pd

df = pd.read_html('http://kind.krx.co.kr/corpgeneral/corpList.do?method=download&searchType=13',header=0)[0]
df.code = df.code.map('{:06d}'.format) # 6자라로 맞추기
df = df[['회사명', '종목코드', '상장일', '업종', '주요제품']]
df.to_excel('codes.xlsx')