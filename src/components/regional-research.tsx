'use client';
import {useResearchState} from '@/lib/use-research-state';
import type {Dataset} from '@/lib/schema';
import {BenchmarkExplorer,Evidence,type OpenSource} from './analytics';

export default function RegionalResearch({data,open}:{data:Dataset;open:OpenSource}) {const [group,setGroup]=useResearchState(`${data.importedAt}/osGroup`,'country');const rows=data.osDistribution.filter(r=>group==='all'||(group==='country'?r.group.startsWith('국가'):r.group==='대륙'));
  return <div className="regional-research"><h3>동남아 조사 범위와 지역 특징</h3><p className="data-note">순위 수집은 인도네시아 Google Play 무료 Top 20입니다. 태국·베트남·필리핀·말레이시아의 순위는 미수집이며 OS 점유율 자료만 있습니다. 기존 시트의 해설에는 인도네시아 Top 100을 참고한 사례도 있지만 정규화 순위 집계에는 Top 20만 포함합니다.</p><div className="sea-notes">{data.seaAnalysis.map(r=><details key={r.text}><summary>{r.text} · {r.kind}</summary><p>{r.evidence}</p><Evidence source={r.source} open={open}/></details>)}</div>
    <BenchmarkExplorer data={data} open={open} initialGroup="sea"/>
    <h3>모바일 OS 사용 점유율</h3><p className="data-note">StatCounter 웹 페이지뷰 기준. Android·iOS·기타 합계 100%. 기기 대수, 게임 이용자 비율, 게임 매출 비율과 다른 측정값입니다. {osMonthNote(data)}</p><label>표시 범위<select aria-label="OS 표시 범위" value={group} onChange={e=>setGroup(e.target.value)}><option value="country">국가</option><option value="continent">대륙</option><option value="all">전체</option></select></label><div className="os-chart">{rows.map(r=><article key={r.region}><div><strong>{r.region}</strong><span>{r.month}</span></div><div className="sample-stack" role="img" aria-label={`${r.region}: Android ${(r.android*100).toFixed(2)}%, iOS ${(r.ios*100).toFixed(2)}%, 기타 ${(r.other*100).toFixed(2)}%`}><span className="os-android" style={{width:`${r.android*100}%`}}>Android {(r.android*100).toFixed(1)}%</span><span className="os-ios" style={{width:`${r.ios*100}%`}}>iOS {(r.ios*100).toFixed(1)}%</span><span className="os-other" style={{width:`${r.other*100}%`}}/></div><small>Android {(r.android*100).toFixed(2)}% · iOS {(r.ios*100).toFixed(2)}% · 기타 {(r.other*100).toFixed(2)}%</small><details><summary>OS 출시 우선순위·근거</summary><p>{data.osPriorities.find(p=>p.region===r.region)?.priority||'제안 미작성'} · iOS 50% 이상 기준의 분석자 제안</p><p>{r.note}</p><Evidence source={r.source} open={open}/>{data.osPriorities.filter(p=>p.region===r.region).map(p=><Evidence key={p.region} source={p.source} open={open}/>)}</details></article>)}</div>
    <details><summary>2025년 세계 게임 매출·다운로드의 스토어 구성</summary><p>2026년 보고서의 2025년 세계 데이터 · 2차 인용. App Store 매출 ${data.osGameMarket.revenue[1]}B ({(Number(data.osGameMarket.revenue[3])*100).toFixed(1)}%), Google Play ${data.osGameMarket.revenue[2]}B ({(Number(data.osGameMarket.revenue[4])*100).toFixed(1)}%).</p><p>다운로드: App Store {data.osGameMarket.downloads[1]}억 건, Google Play {data.osGameMarket.downloads[2]}억 건. Google Play에는 중국의 다른 Android 마켓이 포함되지 않습니다. 국가별 매출·다운로드 OS 구성은 미수집입니다.</p><Evidence source={data.osGameMarket.revenueSource} open={open}/><Evidence source={data.osGameMarket.downloadSource} open={open}/></details>
  </div>;
}

/** e.g. "2026-09 관측, 오세아니아는 2026-08." — the usual month plus the regions observed in another month. */
function osMonthNote(data:Dataset) {
  const months=data.osDistribution.map(o=>o.month);
  const usual=[...new Set(months)].sort((a,b)=>months.filter(m=>m===b).length-months.filter(m=>m===a).length)[0];
  const others=data.osDistribution.filter(o=>o.month!==usual).map(o=>`${o.region}${(o.region.charCodeAt(o.region.length-1)-0xac00)%28?'은':'는'} ${o.month}`);
  return `${usual} 관측${others.length?`, ${others.join(', ')}`:''}.`;
}
