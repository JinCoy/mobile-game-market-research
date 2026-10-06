"""Independently reconcile generated JSON with source cells (read only)."""
import hashlib
import json
import math
import unicodedata
from pathlib import Path
import openpyxl

ROOT=Path(__file__).resolve().parents[1]
data=json.loads((ROOT/'data/market.json').read_text())
config=json.loads((ROOT/'data/import-config.json').read_text())
supplemental=[r for path in config.get('snapshotFiles',[]) for r in json.loads((ROOT/path).read_text())]
supplemental_ids={f"{r['country']}-{r['store']}-{r['chart']}-{r['date']}-{r['rank']}":r for r in supplemental}
books={unicodedata.normalize('NFC',p.name):openpyxl.load_workbook(p,data_only=True) for p in ROOT.glob('*.xlsx')}
hashes={unicodedata.normalize('NFC',p.name):hashlib.sha256(p.read_bytes()).hexdigest() for p in ROOT.glob('*') if p.suffix=='.xlsx' or ('생성코드' in unicodedata.normalize('NFC',p.name) and p.suffix=='.py')}
def eq(a,b):
    assert math.isclose(a,b,rel_tol=1e-12,abs_tol=1e-12) if isinstance(a,(int,float)) and isinstance(b,(int,float)) else a==b,(a,b)
def row(source): return [c.value for c in books[source['file']][source['sheet']][source['row']]]
assert len(data['documents'])==sum(len(w.sheetnames) for w in books.values())
market=books['시장조사.xlsx']
def valid_rank(value): return type(value) is int and 1<=value<=100
expected=sum(valid_rank(r[0].value) for sheet in ['1-1 App Store Top100','1-2 Google Play Top100','3-2 매출 Top100'] for r in market[sheet])
expected+=sum(valid_rank(r[3].value) for sheet in ['2-2 국가별 Top20','2-7 동남아 분석'] for r in market[sheet] if len(r)>=8 and isinstance(r[2].value,str) and len(r[2].value)==10)
assert len(data['observations'])==expected+len(supplemental)
assert len(data['games'])==len({o['gameId'] for o in data['observations']})
for o in data['observations']:
    if o['id'] in supplemental_ids:
        original=supplemental_ids[o['id']]
        for field in ['name','publisher','genre','country','store','date','rank','chart']:
            eq(o[field],original[field])
        eq(o['source']['file'],unicodedata.normalize('NFC',original['source']['file']))
        eq(o['source']['sheet'],original['source']['sheet']);eq(o['source']['row'],original['source']['row'])
        row(o['source'])
        continue
    r=row(o['source'])
    if o['source']['sheet'] in ['2-2 국가별 Top20','2-7 동남아 분석']:
        eq(o['name'],r[4]);eq(o['rank'],r[3]);eq(o['date'],r[2])
    else:
        eq(o['name'],r[1]);eq(o['rank'],r[0])
assert len(data['advertisingChannels'])==7
for c in data['advertisingChannels']:
    r=row(c['source']);eq(c['scores'],r[1:6]);eq(c['score'],r[6]);eq(c['rank'],r[7])
    eq(sum(s*w/5*100 for s,w in zip(c['scores'],data['channelWeights'])),c['score'])
    eq(c['public']['name'],row(c['public']['source'])[0])
for table in ['hook','format']:
    for game in data['adBrandStats']:
        samples=[s for s in data['adHookSamples'] if s['gameId']==game['gameId'] and s['table']==table]
        assert sum(s['count'] for s in samples)==25
        for s in samples:
            r=row(s['source']);idx=['Royal Match','MONOPOLY GO!','Magic Sort!'].index(s['name'])+2
            eq(s['count'],r[idx])
assert len(data['adCreatives'])==30
assert len({(c['gameId'],c['videoGroup']) for c in data['adCreatives']})==26
for c in data['adCreatives']:
    r=row(c['source']);reading=row(c['readingSource'])
    eq(c['adId'],r[1]);eq(c['adId'],reading[1]);eq(c['actualType'],reading[2]);eq(c['requestedType'],reading[3]);eq(c['description'],reading[4]);eq(c['videoGroup'],reading[5] or r[1])
for game in data['adBrandStats']:
    assert len([c for c in data['adCreatives'] if c['gameId']==game['gameId']])==10
for p in data['marketingPlan']['phases']:
    r=row(p['source']);eq(p['budget'],r[6]);eq(p['cpi'],r[8])
    if p['cpi'] is not None:
        eq(p['cpi'],row(p['cpiSource'])[4]);eq(p['d7Rate'],row(p['retentionSource'])[2])
        eq(p['budget']/p['cpi'],p['installs']);eq(p['installs']*p['d7Rate'],p['d7'])
totals={'budget':sum(p['budget'] for p in data['marketingPlan']['phases']), 'installs':sum(p['installs'] or 0 for p in data['marketingPlan']['phases']), 'd7':sum(p['d7'] or 0 for p in data['marketingPlan']['phases'])}
cached=books['게임전략_MVP_마케팅.xlsx']['3-4 마케팅 실행안']
eq(totals['budget'],cached['G13'].value);eq(totals['installs'],cached['J13'].value);eq(totals['d7'],cached['L13'].value)
for o in data['osDistribution']:
    r=row(o['source']);eq(o['android'],r[2]);eq(o['ios'],r[3]);eq(o['other'],r[4]);eq(o['android']+o['ios']+o['other'],1)
    assert o['source']['file']=='시장조사.xlsx'
for b in data['regionalBenchmarks']:
    r=row(b['source'])
    idx=14 if b['group']=='genre' else ['동남아','아시아','북미'].index(b['label'])+1 if b['group']=='sea' else ['P50','P90','P99'].index(b['percentile'])+2
    eq(b['value'],r[idx]);assert b['source']['file']=='시장조사.xlsx'
for a in data['mvpAssessments']:
    r=row(a['source']);eq(a['name'],r[0]);eq(a['score'],r[7]);eq(a['decision'],r[8])
blank=next(d for d in data['documents'] if d['sheet']=='2-4 2인 개발 MVP 판단')
assert blank['blankStates']['H39']==blank['blankStates']['H51']=='intentionalBlank'
assert all(c.data_type!='e' for w in books.values() for s in w for r in s for c in r)
output=ROOT/'docs/validation'
output.mkdir(exist_ok=True)
(output/'source-sha256.json').write_text(json.dumps(hashes,ensure_ascii=False,indent=2)+'\n')
summary=dict(sheets=len(data['documents']),observations=len(data['observations']),knownGames=sum(not g['missing'] for g in data['games']),missingSlots=sum(g['missing'] for g in data['games']),channels=len(data['advertisingChannels']),aiSamplesPerTable=75,creatives=30,uniqueVideos=26,mvpAssessments=len(data['mvpAssessments']),benchmarks=len(data['regionalBenchmarks']),osRecords=len(data['osDistribution']),planTotals=totals)
(output/'data-checks.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2)+'\n')
print(json.dumps(summary,ensure_ascii=False,indent=2))
