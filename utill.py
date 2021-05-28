#-*- encoding: utf-8 -*-
from pytrends.request import TrendReq
import pandas as pd
import csv
import seaborn as sns
import matplotlib.pyplot as plt
from os import path
import json

pytrends = TrendReq(hl = "ko", tz = 540)

# 파일을 읽어오기
filePath = "raw.xlsx"
if(not path.exists(filePath)):
    print("raw 파일을 확인해주세요")

try:
    rawData = pd.read_excel(filePath, sheet_name='Sheet', usecols = ["정치", "연예계","시작","종료"])
    rawDataJson = json.loads(rawData.to_json(orient='records', force_ascii=False))
except Exception as e:
    print('RAW 파일에서 읽기에서 에러가 발생했습니다. Sheet 이름이 sheet 이며, 칼럼 이름이 "정치", "연예계" 인지 확인해주세요')
    print(e)

keywords = []
data = []

for i in rawDataJson:
    keywords.append([ i["정치"], i["연예계"] ])
    data.append( [ i["시작"], i["종료"] ] )

for idx, key in enumerate(keywords):
    period = str(data[idx][0]) + ' ' + str(data[idx][1])
    pytrends.build_payload(key, cat = 0, timeframe = period, geo = 'KR', gprop = '')
    getDataInfo = pytrends.interest_over_time()

    del getDataInfo['isPartial']

    print(getDataInfo.keys())
    # key값으로 키워드만 있음. 날짜가 안나옴

    getDataInfo.to_csv(key[0] + ".txt" ,encoding = 'utf-8-sig')

    getDataInfoCsv = pd.read_csv(key[0] + ".txt")

    # 키워드의 최고 vol날짜를 찾는 것
    for i in key:
        maxVal = max(getDataInfoCsv[i])
        for idx, valVol in enumerate(getDataInfoCsv[i]):
            if(maxVal == valVol):
                print(maxVal)
                print(getDataInfoCsv.iloc[idx]["date"])


    # for i in key:
    #     maxVal = max(getDataInfo[i])
    #     print(maxVal)
        # maxVal = max(getDataInfo[i])
        # for idx,data in enumerate(getDataInfo[i]):
        #     if(data == maxVal):
        #         print(i, "  idx :: ", getDataInfo["idx"], "  data :: ", data)
            
    plt.figure()
    sns.set(style="darkgrid", palette = 'rocket')
    ax = sns.lineplot(data=getDataInfo)
    ax.set_title('THAAD & Hong,Kim Love', fontsize=20)

    # 윈도우에서 한글 글꼴 설정하면 굳이 안써도 돼
    plt.legend('TH', ncol=4, loc='upper left')

    plt.xlabel('date')
    plt.ylabel('Vol')

    plt.show()

# pytrends.build_payload(keywords2, cat = 0, timeframe = 'today 5-y', geo = 'KR', gprop = '')
# getDataInfo2 = pytrends.interest_over_time()

# del getDataInfo1['isPartial']
# del getDataInfo2['isPartial']

# print(getDataInfo1[keywords1[0]])

# 관심도가 90점 이상인것은 버렸어
# 극단치를 버리는 로직을 사용해서 90 말고 다른 값을 쓰는 것도 좋을 둣.(시간나면 ㄱ)

# plt.rcParams['axes.unicode_minus'] = False



# getDataInfo1.to_csv("사드, 홍상수.csv", encoding = 'utf-8-sig')
# getDataInfo2.to_csv("기타.csv", encoding = 'utf-8-sig')
