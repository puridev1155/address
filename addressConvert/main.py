import time
import warnings

import pandas as pd
import requests as req
from pandas import DataFrame

# Excel Reading
xlsx = pd.read_excel('./document.xlsx', usecols=[7]) 

# Test Excel Read
print(xlsx['Address'])

# API KEY (dev)
# API KEY 받기 (https://www.juso.go.kr/addrlink/devAddrLinkRequestWrite.do?returnFn=write&cntcMenu=URL)
#http://www.juso.go.kr/addrlink/addrLinkApi.do?currentPage="+currentPage+"&countPerPage="+countPerPage+"&keyword="+URLEncoder.encode(keyword,"UTF-8")+"&confmKey="+confmKey; 
#https://www.juso.go.kr/addrlink/addrLinkApi.do?currentPage=1&countPerPage=10&keyword=&confmKey=devU01TX0FVVEgyMDIxMTAyODE4MjczMDExMTgxNjc=
KEY = 'devU01TX0FVVEgyMDIyMTAxMjE4MjgyNjExMzA0OTE='

# Request URL
URL = 'https://www.juso.go.kr/addrlink/addrLinkApi.do?confmKey=' + KEY + '&currentPage=' + '1' + '&countPerPage=' + \
      '10'+ '&addInfoYn=' + 'Y' + '&resultType=' + 'json' + '&keyword=' 

# 시도
city = []
# 시군구
gun = []
# 읍면동
dong = []

# API Called
for idx, keyword in enumerate(xlsx['Address']):
    try:
        searchUrl = URL + keyword
        # Test searchUrl
        print(searchUrl)
        response = req.get(searchUrl)
        print(response.json()["results"]['juso'][0]['hemdNm'])
        juso = response.json()["results"]['juso'][0]['hemdNm']
        juso = juso.split()
        if len(juso) >= 4:
            print(juso)
            juso[1] = juso[1] + " " + juso[2]
            juso[2] = juso[3]

        if juso[0] == '세종특별자치시':
            juso.append('')
            juso[2] = juso[1]
            juso[1] = '세종시'

        if juso[1] == '성북구' or juso[1] == '은평구':
            juso[2] = juso[2].replace('제', '')


        # 승현아 여기부터 안된다 왜그럴까?
        if juso[2] in '.':
            juso[2] = juso[2].replace('.', '·')

        print(juso[0], juso[1], juso[2])
        city.append(juso[0])
        gun.append(juso[1])
        dong.append(juso[2])
    except:
        # 주소 잘못된 경우 or API Response Time out
        print(keyword, '변환 에러')
        city.append('변환 에러')
        gun.append('변환 에러')
        dong.append('변환 에러')
    # delay 0.5s
    time.sleep(.5)


# DataFrame
df = DataFrame({"시도": city, "시군구": gun, "읍면동": dong})

# XlsxWriter 엔진으로 Pandas writer 객체 만들기
writer = pd.ExcelWriter('result.xlsx', engine='xlsxwriter') # pylint: disable=abstract-class-instantiated

## DataFrame을 xlsx에 쓰기
df.to_excel(writer, sheet_name='Sheet1')

## Pandas writer 객체 닫기
writer.close()
