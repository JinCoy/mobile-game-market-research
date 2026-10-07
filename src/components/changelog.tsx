import type {Dataset} from '@/lib/schema';
import {hashFor} from '@/lib/routes';
import ResearchText from './research-text';

/** Update log read from 0-1 개요 stage rows (출처·한계) and 0 목차 notes. Newest stage first. */
export default function Changelog({data}: {data: Dataset}) {
  return <section className="changelog">
    <p className="data-note">엑셀 두 파일의 ‘0-1 개요’ 단계별 출처·한계 줄과 ‘0 목차’ 비고 칸에서 읽습니다. 날짜는 원자료 조회일이며 사이트를 고친 날짜가 아닙니다.</p>
    <ol>{data.changelog.map(entry => <li key={entry.stage}>
      <h2>{entry.stage}{entry.date && <time dateTime={entry.date}>{entry.date.replaceAll('-', '.')} 조회</time>}</h2>
      {entry.sheets.length > 0 && <div className="changelog-sheets"><strong>만들거나 바꾼 시트</strong><ul>{entry.sheets.map(s => <li key={s.file + s.sheet}><a href={hashFor('sources', s)}>{s.sheet}</a><small>{s.file} · {s.note}</small></li>)}</ul></div>}
      {entry.sources && <div><strong>출처</strong><p><ResearchText text={entry.sources}/></p></div>}
      {entry.limits && <div><strong>한계</strong><p><ResearchText text={entry.limits}/></p></div>}
      {entry.row && <a className="source-link" href={hashFor('sources', {file: '시장조사.xlsx', sheet: '0-1 개요'})}>원자료 0-1 개요 · {entry.row}행</a>}
    </li>)}</ol>
  </section>;
}
