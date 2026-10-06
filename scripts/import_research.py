"""Normalize bounded research tables. Formula evaluation is deliberately limited."""
import re

MARKET = '시장조사.xlsx'
STRATEGY = '게임전략_MVP_마케팅.xlsx'

def normalize_research(market, strategy, identity):
    def source(file, sheet, row, cells, kind, text, date=None):
        return dict(file=file, sheet=sheet, row=row, range=cells, kind=kind,
                    originalText=text, date=date, urls=re.findall(r'https?://[^\s)]+', str(text)))
    def ss(sheet, row, cells, kind, text, date=None):
        return source(STRATEGY, sheet, row, cells, kind, text, date)
    def ms(sheet, row, cells, kind, text, date=None):
        return source(MARKET, sheet, row, cells, kind, text, date)
    channels_sheet = '3-2 광고 채널 분석'
    rows = strategy[channels_sheet]
    channels = []
    for row in range(17,24):
        score, public = rows[row-1], rows[row-12]
        channels.append(dict(name=score[0], scores=score[1:6], score=score[6], rank=score[7], reason=score[8],
                             public=dict(name=public[0], buying=public[1], reach=public[2], grade=public[3], performance=public[4], genre=public[5], cost=public[6], text=public[7], source=ss(channels_sheet,row-11,f'A{row-11}:H{row-11}','secondary',public[7])),
                             source=ss(channels_sheet,row,f'A{row}:I{row}','analystReading',score[8])))
    ad_sheet = '3-3 후킹 영상 분석'
    rows = strategy[ad_sheet]
    names = rows[10][2:5]
    stats = [dict(gameId=identity(r[0]),name=r[0],active=r[2],tracked=r[3],classified=r[4],longestDays=r[5],spendEstimate=r[6],note=r[7],source=ss(ad_sheet,i,f'A{i}:H{i}','estimate',r[7],'2026-10-05')) for i,r in enumerate(rows[5:8],6)]
    samples = []
    for table, start, end in [('hook',12,15),('format',20,25)]:
        for row in range(start,end+1):
            r = rows[row-1]
            for col,name in enumerate(names,2):
                samples.append(dict(gameId=identity(name),name=name,table=table,label=r[0],description=r[1],count=r[col],sampleSize=25,source=ss(ad_sheet,row,f'A{row}:G{row}','aiClassification',r[1],'2026-10-05')))
    creatives = []
    readings = {str(r[1]): (i,r[:6]) for i,r in enumerate(rows[62:92],63)}
    for row,r in enumerate(rows[29:59],30):
        reading_row, reading = readings[str(r[1])]
        creatives.append(dict(gameId=identity(r[0]),name=r[0],adId=str(r[1]),days=r[2],hook=r[3],format=r[4],tone=r[5],offer=r[6],copy=r[7],
                             actualType=reading[2],requestedType=reading[3],description=reading[4],videoGroup=reading[5] or str(r[1]),
                             source=ss(ad_sheet,row,f'A{row}:H{row}','aiClassification',r[7],'2026-10-05'),
                             readingSource=ss(ad_sheet,reading_row,f'A{reading_row}:F{reading_row}','analystReading',reading[4],'2026-10-05')))
    benchmarks = []
    for row,r in enumerate(market['4-2 지역 벤치마크'],1):
        if len(r)>6 and isinstance(r[2],(int,float)):
            for col,p in enumerate(['P50','P90','P99'],2):
                benchmarks.append(dict(group='region',label=r[0],metric=r[1],value=r[col],unit=r[5],percentile=p,publicationYear=2026,period='2025년 주간 데이터의 연평균',source=ms('4-2 지역 벤치마크',row,f'A{row}:G{row}','secondary',r[6])))
    for row,r in enumerate(market['2-7 동남아 분석'][66:72],67):
        for col,label in enumerate(['동남아','아시아','북미'],1):
            benchmarks.append(dict(group='sea',label=label,metric=r[0],value=r[col],unit=r[4],percentile='P50',publicationYear=2025,period='2024년 지역별 중앙값',source=ms('2-7 동남아 분석',row,f'A{row}:F{row}','secondary',r[5])))
    for start,metric,unit in [(7,'D1 리텐션','%'),(26,'D7 리텐션','%'),(45,'D28 리텐션','%'),(64,'하루 플레이 시간','분'),(83,'세션 길이','분'),(102,'하루 세션 수','회')]:
        for row in range(start,start+16):
            r=market['4-3 장르 차트 판독값'][row-1]
            benchmarks.append(dict(group='genre',label=r[1],metric=metric,value=r[14],unit=unit,percentile='P50',publicationYear=2025,period='2024년 월별 중앙값의 산술평균',source=ms('4-3 장르 차트 판독값',row,f'A{row}:S{row}','analystReading',market['4-3 장르 차트 판독값'][2][0])))
    plan_sheet='3-4 마케팅 실행안'
    rows=strategy[plan_sheet]
    cpi_rows={r[0]:(i,r) for i,r in enumerate(strategy['3-1 광고 대비 유입'],1) if isinstance(r[0],str) and re.fullmatch(r'C\d\d',r[0])}
    phases=[]
    for row,r in enumerate(rows[7:12],8):
        retention=next((b for b in benchmarks if b['group']=='region' and b['label']==r[10] and b['metric']=='D7 리텐션' and b['percentile']=='P50'),None)
        cpi=cpi_rows.get(r[7])
        phases.append(dict(name=r[0],timing=r[1],region=r[2],os=r[3],channel=r[4],channelRank=r[5],budget=r[6],cpiId=r[7],cpi=r[8],installs=r[9],retentionRegion=r[10],d7Rate=retention['value'] if retention else None,d7=r[11],criterion=r[12],reason=r[13],
                           source=ss(plan_sheet,row,f'A{row}:N{row}','proposal',r[13]),retentionSource=retention['source'] if retention else None,
                           cpiSource=ss('3-1 광고 대비 유입',cpi[0],f'A{cpi[0]}:I{cpi[0]}','secondary',cpi[1][8],'2026-10-05') if cpi else None,cpiRegion=cpi[1][1] if cpi else None,cpiPeriod=cpi[1][5] if cpi else None))
    plan=dict(candidate=rows[3][1],summary=rows[3][2],source=ss(plan_sheet,4,'A4:D4','proposal',rows[3][3]),phases=phases,
              materials=[dict(name=r[0],opening=r[1],reference=r[2],difficulty=r[3],source=ss(plan_sheet,i,f'A{i}:D{i}','proposal',r[2])) for i,r in enumerate(rows[16:21],17)],
              metrics=[dict(name=r[0],meaning=r[1],criterion=r[2],reference=r[3],source=ss(plan_sheet,i,f'A{i}:D{i}','proposal',r[3])) for i,r in enumerate(rows[25:31],26)])
    os_sheet='2-6 대륙별 OS 분포'
    os_rows=market[os_sheet]
    os=[dict(group=r[0],region=r[1],android=r[2],ios=r[3],other=r[4],month=r[6],note=r[8],source=ms(os_sheet,i,f'A{i}:I{i}','observed',r[7],r[6])) for i,r in enumerate(os_rows[2:21],3)]
    priorities=[dict(region=r[0],priority=r[2],reason=r[3],source=ms(os_sheet,i,f'A{i}:D{i}','proposal',r[3])) for i,r in enumerate(os_rows[45:62],46)]
    mvp_sheet='2-4 2인 개발 MVP 판단'
    rows=strategy[mvp_sheet]
    assessments=[dict(name=r[0],gameId=identity(r[0]),genre=r[1],subgenre=r[2],grossingRank=r[3],freeRank=r[4],classification=r[5],candidate=r[6],score=r[7],decision=r[8],source=ss(mvp_sheet,i,f'A{i}:I{i}','analystReading',r[5]),scoreState='intentionalBlank' if r[7] is None and r[6]=='후보 아님' else 'analysisMissing' if r[7] is None else 'available') for i,r in enumerate(rows[78:250],79)]
    recommended=[dict(name=r[0],gameId=identity(r[0]),subgenre=r[1],score=r[4],decision=r[6],roles=r[7],duration=r[8],scope=r[9],caution=r[10],source=ss(mvp_sheet,i,f'A{i}:K{i}','proposal',r[10])) for i,r in enumerate(rows[57:64],58)]
    criteria=[dict(genre=r[0],subgenre=r[1],scores=r[2:7],score=r[7],decision=r[8],reason=r[9],source=ss(mvp_sheet,i,f'A{i}:L{i}','analystReading',r[9])) for i,r in enumerate(rows[17:54],18)]
    targets=[dict(name=r[0],gameId=identity(r[0]),region=r[1],genre=r[2],d1=dict(criterion=r[3],target=r[4],median=r[5]),d7=dict(criterion=r[6],target=r[7],median=r[8]),d30=dict(criterion=r[9],target=r[10],median=r[11]),iapArpu=r[12],adArpu=r[13],iapArppu=r[14],source=ss('2-2 MVP 검증 목표',i,f'A{i}:P{i}','proposal','지역·장르·합격선 선택은 분석자 제안. 수익 참고는 AppsFlyer D90 글로벌 평균.')) for i,r in enumerate(strategy['2-2 MVP 검증 목표'][2:14],3)]
    specs=[dict(name=r[0],values=r,source=ss('2-1 MVP 명세',i,f'A{i}:H{i}','proposal',r)) for i,r in enumerate(strategy['2-1 MVP 명세'][1:],2) if r[0]]
    sea=[dict(text=r[0],evidence=r[1],kind=r[2],source=ms('2-7 동남아 분석',i,f'A{i}:C{i}','proposal' if i==63 else 'analystReading',r[1],'2026-10-05')) for i,r in enumerate(market['2-7 동남아 분석'][55:63],56)]
    return dict(advertisingChannels=channels,channelAxes=strategy[channels_sheet][14][1:6],channelWeights=strategy[channels_sheet][15][1:6],adBrandStats=stats,adHookSamples=samples,adCreatives=creatives,marketingPlan=plan,regionalBenchmarks=benchmarks,osDistribution=os,osPriorities=priorities,
                osGameMarket=dict(revenue=os_rows[34][:6],downloads=os_rows[40][:6],revenueSource=ms(os_sheet,35,'A35:F35','secondary',os_rows[34][5]),downloadSource=ms(os_sheet,41,'A41:F41','secondary',os_rows[40][5])),
                mvpAssessments=assessments,mvpRecommended=recommended,mvpCriteria=criteria,mvpTargets=targets,mvpSpecs=specs,seaAnalysis=sea,
                documentReferences=[dict(file=STRATEGY,sheet=s,originalFile=MARKET,originalSheet=t) for s,t in [('4-1 지역 벤치마크(참조)','4-2 지역 벤치마크'),('4-2 장르 차트 판독값(참조)','4-3 장르 차트 판독값'),('4-3 장르 판독 요약(참조)','4-4 장르 판독 요약'),('4-4 Google Play Top100(참조)','1-2 Google Play Top100'),('4-5 매출 Top100(참조)','3-2 매출 Top100'),('4-6 매출-다운로드 비교(참조)','3-3 매출-다운로드 비교'),('4-7 대륙별 OS 분포(참조)',os_sheet)]])
