import type {Observation, PlanPhase} from './schema';

export function latestRecords(records: Observation[]) {
  const dates = new Map<string,string>();
  records.forEach(o => {const key = `${o.country}/${o.store}/${o.chart}`; if(o.date > (dates.get(key) || '')) dates.set(key,o.date);});
  return records.filter(o => o.date === dates.get(`${o.country}/${o.store}/${o.chart}`));
}
export function channelScore(scores:number[], weights:number[]) {
  if(weights.some(w => !Number.isFinite(w) || w < 0 || w > 100) || Math.abs(weights.reduce((a,b)=>a+b,0)-100)>1e-8) return null;
  return scores.reduce((sum,score,i)=>sum+score*weights[i]/5,0);
}
export function planResult(phase: PlanPhase, budget: string, cpi: string) {
  if(!budget.trim() || !Number.isFinite(Number(budget)) || Number(budget)<0) return null;
  if(phase.cpi === null) return Number(budget)===0 ? {installs:null,d7:null} : null;
  if(!cpi.trim() || !Number.isFinite(Number(cpi)) || Number(cpi)<=0 || phase.d7Rate===null) return null;
  const installs = Number(budget)/Number(cpi);
  return {installs,d7:installs*phase.d7Rate};
}
export const evidenceLabels:Record<string,string> = {observed:'관측',secondary:'공개 자료·2차 인용',estimate:'추정',aiClassification:'AI 분류',analystReading:'분석자 판독·평가',proposal:'분석자 제안'};
export const genrePalette:Record<string,string> = {'퍼즐':'#cdd1d6','하이퍼·하이브리드 캐주얼':'#8da2b8','카드·보드·카지노':'#c5af80','시뮬레이션':'#94b7a5','전략':'#aba1ba','액션·슈팅':'#ae8f91','스포츠':'#89aeb1','RPG':'#9d9ebc','파티·음악':'#b9a58c','기타(UGC·리워드)':'#88988f','확인 필요':'#626b78','확인 불가':'#424954'};
