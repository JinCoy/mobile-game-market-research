'use client';

import { useEffect, useMemo, useRef, useState } from 'react';
import { ArrowDownRight, ArrowLeft, ArrowRight, ArrowUpRight, BookmarkSimple, ChartBar, Check, DownloadSimple, FunnelSimple, GlobeHemisphereWest, MagnifyingGlass, Plus, SquaresFour, Stack, X } from '@phosphor-icons/react';
import type { Dataset, Game, Observation, ResearchDocument } from '@/lib/schema';
import ResearchText from './research-text';

const sections = [
  ['overview', '시장 개요'], ['rankings', '게임 순위'], ['comparison', '스토어 비교'],
  ['gameplay', '게임플레이'], ['monetization', '수익 모델'], ['marketing', '마케팅'],
  ['art', '아트 디렉션'], ['regions', '지역 비교'], ['opportunity', '개발 기회'], ['sources', '자료·출처'],
] as const;
type Section = typeof sections[number][0];
const storeName = (store: string) => store === 'ios' ? 'App Store' : 'Google Play';
const colors = ['#de633d', '#d4a15c', '#74938c', '#7e8eac', '#b18b92', '#9c9b76', '#82939c', '#b9aa94'];
const percent = (n: number, d: number) => d ? Math.round(n / d * 100) : 0;
const compact = (n: number) => Intl.NumberFormat('en', {notation: 'compact', maximumFractionDigits: 1}).format(n);
function dateLabel(date: string) { return date.replaceAll('-', '.'); }

export default function Dashboard({data}: {data: Dataset}) {
  const [section, setSection] = useState<Section>('overview');
  const [country, setCountry] = useState('미국');
  const year = String(Math.max(...data.observations.map(o => o.year)));
  const [snapshot, setSnapshot] = useState('latest');
  const [store, setStore] = useState('all');
  const [top, setTop] = useState(100);
  const [genre, setGenre] = useState('all');
  const [chart, setChart] = useState('free');
  const [query, setQuery] = useState('');
  const [sort, setSort] = useState('rank');
  const [selected, setSelected] = useState<Game | null>(null);
  const [compare, setCompare] = useState<string[]>([]);
  const [compareOpen, setCompareOpen] = useState(false);
  const [saved, setSaved] = useState<string[]>([]);
  const [savedOnly, setSavedOnly] = useState(false);
  const [notice, setNotice] = useState('');
  const [documentId, setDocumentId] = useState(data.documents[1].id);
  const [documentQuery, setDocumentQuery] = useState('');
  const [documentReturn, setDocumentReturn] = useState<{section: Section; scroll: number; game: Game | null; compareOpen: boolean} | null>(null);
  const returnScroll = useRef<number | null>(null);
  const dialog = useRef<HTMLDialogElement>(null);
  const gameMap = useMemo(() => new Map(data.games.map(game => [game.id, game])), [data.games]);
  useEffect(() => {
    try { const value = JSON.parse(localStorage.getItem('playfield-saved') || '[]'); if (Array.isArray(value)) setSaved(value.filter(id => typeof id === 'string')); } catch { /* A corrupt local preference can be discarded. */ }
  }, []);
  useEffect(() => {
    if (returnScroll.current !== null) {
      const scroll = returnScroll.current;
      returnScroll.current = null;
      requestAnimationFrame(() => window.scrollTo({top: scroll, behavior: 'instant'}));
    }
  }, [section]);
  useEffect(() => {
    if (selected || compareOpen) dialog.current?.showModal(); else dialog.current?.close();
  }, [selected, compareOpen]);
  const countries = [...new Set(data.observations.map(o => o.country))];
  const countryLabel = country === 'all' ? '전체 시장' : country;
  const available = data.observations.filter(o => (country === 'all' || o.country === country) && String(o.year) === year && o.chart === chart);
  const dates = [...new Set(available.map(o => o.date))].sort().reverse();
  const availableStores = [...new Set(available.map(o => o.store))];
  // Keep each country's own latest snapshot; one country's date must not discard another country's records.
  const latestDates = new Map<string, string>();
  available.forEach(o => {
    const key = `${o.country}-${o.store}`;
    if (o.date > (latestDates.get(key) || '')) latestDates.set(key, o.date);
  });
  const base = available.filter(o => snapshot === 'latest'
    ? o.date === latestDates.get(`${o.country}-${o.store}`)
    : o.date === snapshot).filter(o => o.rank <= top);
  const scope = base.filter(o => store === 'all' || o.store === store);
  const filtered = scope.filter(o => (genre === 'all' || o.genre === genre)
    && (!savedOnly || saved.includes(o.gameId)) && `${o.name} ${o.publisher} ${o.genre} ${o.subgenre || ''}`.toLowerCase().includes(query.toLowerCase()));
  const genreNames = [...new Set(base.map(o => o.genre))].filter(g => g !== '확인 불가').sort();
  const grouped = useMemo(() => {
    const result = new Map<string, Observation[]>();
    filtered.forEach(o => result.set(o.gameId, [...(result.get(o.gameId) || []), o]));
    return [...result.values()].sort((a,b) => sort === 'name' ? a[0].name.localeCompare(b[0].name) : Math.min(...a.map(o => o.rank)) - Math.min(...b.map(o => o.rank)));
  }, [filtered, sort]);
  const ios = scope.filter(o => o.store === 'ios');
  const android = scope.filter(o => o.store === 'android');
  const known = scope.filter(o => !gameMap.get(o.gameId)?.missing);
  const unique = new Set(known.map(o => o.gameId)).size;
  const overlap = new Set(ios.filter(o => android.some(p => p.gameId === o.gameId) && !gameMap.get(o.gameId)?.missing).map(o => o.gameId)).size;
  const genreCounts = genreNames.map(name => ({name, ios: ios.filter(o => o.genre === name).length, android: android.filter(o => o.genre === name).length})).filter(g => g.ios + g.android > 0).sort((a,b) => b.ios + b.android - a.ios - a.android);
  const storeRows = known;
  const leaders = [...new Map(storeRows.map(o => [o.genre, 0])).keys()].map(name => ({name, count: storeRows.filter(o => o.genre === name).length})).sort((a,b) => b.count - a.count);
  const leader = leaders[0];
  const leaderCount = leader?.count || 0;
  const selectedDates = [...new Set(base.map(o => o.date))].sort();
  const tableRows = grouped.slice(0, top);
  const researchGames = grouped.map(group => gameMap.get(group[0].gameId)!).filter(game => game && !game.missing);
  function navigate(next: Section) {
    if (next === 'sources' && section !== 'sources') setDocumentReturn({section, scroll: window.scrollY, game: selected, compareOpen});
    else if (next !== 'sources') setDocumentReturn(null);
    setSection(next); setNotice('');
  }
  function returnFromDocument() {
    if (!documentReturn) return;
    returnScroll.current = documentReturn.scroll;
    setSection(documentReturn.section); setSelected(documentReturn.game); setCompareOpen(documentReturn.compareOpen);
    setDocumentReturn(null); setNotice('');
  }
  function bookmark(id: string) {
    const next = saved.includes(id) ? saved.filter(item => item !== id) : [...saved, id];
    setSaved(next);
    try {localStorage.setItem('playfield-saved', JSON.stringify(next));} catch {setNotice('브라우저 저장 공간을 사용할 수 없습니다. 현재 화면에서는 관심 게임이 유지됩니다.');}
  }
  function toggleCompare(id: string) {
    if (compare.includes(id)) setCompare(compare.filter(item => item !== id));
    else if (compare.length < 6) setCompare([...compare, id]);
    else setNotice('최대 6개 게임을 비교할 수 있습니다. 기존 선택을 해제한 뒤 추가하세요.');
  }
  function exportCsv() {
    const rows = [['기준일','국가','스토어','차트','순위','게임','퍼블리셔','장르','세부 장르','메모','파일','시트','행','출처'],
      ...filtered.map(o => [o.date,o.country,storeName(o.store),o.chart,o.rank,o.name,o.publisher,o.genre,o.subgenre,o.notes,o.source.file,o.source.sheet,o.source.row,o.source.url])];
    const csv = '\uFEFF' + rows.map(row => row.map(value => { const s = String(value ?? ''); return '"' + (/^[=+@-]/.test(s) ? "'" + s : s).replaceAll('"','""') + '"'; }).join(',')).join('\r\n');
    const url = URL.createObjectURL(new Blob([csv], {type: 'text/csv;charset=utf-8;'}));
    const a = document.createElement('a'); a.href = url; a.download = `playfield-${countryLabel}-${chart}-top${top}.csv`; a.click(); URL.revokeObjectURL(url);
    setNotice(`${filtered.length}개 순위 기록을 CSV로 내보냈습니다.`);
  }
  function showDocument(file: string, sheet: string) {
    const doc = data.documents.find(d => d.file === file && d.sheet === sheet);
    if(doc) {setDocumentId(doc.id); setDocumentQuery(''); navigate('sources');setSelected(null);setCompareOpen(false);}
  }
  const rankOf = (id: string, s: string) => {
    const ranks = scope.filter(o => o.gameId === id && o.store === s).map(o => o.rank);
    return ranks.length ? Math.min(...ranks) : undefined;
  };
  const selectedDocument = data.documents.find(d => d.id === documentId)!;

  return <>
    <header className="masthead"><div className="header-inner">
      <button className="brand" onClick={() => navigate('overview')} aria-label="PLAYFIELD 시장 개요"><span className="brand-mark"><SquaresFour size={24} weight="fill"/></span>PLAYFIELD<span className="brand-caption">GAME MARKET RESEARCH</span></button>
      <div className="header-actions"><span className="internal-label">TEAM RESEARCH / VOL. 01</span><button className={savedOnly ? 'plain-button active' : 'plain-button'} onClick={() => {setSavedOnly(!savedOnly);navigate('rankings');}}><BookmarkSimple size={17} weight={savedOnly ? 'fill' : 'regular'}/>관심 게임 <span className="small-count">{saved.length}</span></button></div>
    </div></header>
    <nav className="navigation" aria-label="주요 분석"><div className="nav-inner">{sections.map(([id,label]) => <button key={id} onClick={() => navigate(id)} aria-current={section === id ? 'page' : undefined} className={section === id ? 'nav-item selected' : 'nav-item'}>{label}</button>)}</div></nav>
    <main className="container">
      <div className="report-meta"><span><span className="status-dot"/> RESEARCH EDITION {year}</span><span>워크북 기반 스냅샷<span className="meta-divider">/</span>{selectedDates.map(dateLabel).join(' – ') || '해당 조건의 자료 없음'}</span></div>
      <>{section === 'sources' && documentReturn && <button className="source-back" onClick={returnFromDocument}><ArrowLeft size={17}/>{sections.find(([id]) => id === documentReturn.section)?.[1]}로 돌아가기</button>}</><div className="title-row"><div><div className="eyebrow">MOBILE GAME INTELLIGENCE</div><h1>{section === 'overview' ? '모바일 게임 시장 분석' : sections.find(([id]) => id === section)?.[1]}</h1><>{section !== 'overview' && <p className="intro">{section === 'sources' ? '원자료의 기록과 분석 범위를 확인합니다.' : '같은 조사 범위 안에서 게임과 전략을 비교합니다.'}</p>}</></div><button className="export-button" onClick={exportCsv}><DownloadSimple size={18}/>순위 CSV 내보내기</button></div>

      {section !== 'sources' && <>
        <div className="filter-strip">
          <div className="filter-group"><GlobeHemisphereWest size={18}/><label><span>시장</span><select aria-label="시장" value={country} onChange={e => {setCountry(e.target.value);setStore('all');setSnapshot('latest');setGenre('all');setTop(e.target.value === '미국' || e.target.value === 'all' ? 100 : 20);setChart('free');}}><option value="all">전체</option>{countries.map(c => <option key={c}>{c}</option>)}</select></label></div>
          <div className="filter-group"><label><span>연도</span><span className="year-value" aria-label="연도">{year}</span></label></div>
          <div className="filter-group date-filter"><label><span>기준일</span><select aria-label="기준일" value={snapshot} onChange={e => setSnapshot(e.target.value)}><option value="latest">{country === 'all' ? '국가·스토어별 최신' : '스토어별 최신'}</option>{dates.map(d => <option key={d}>{d}</option>)}</select></label></div>
          <div className="filter-group"><label><span>스토어</span><select aria-label="스토어" value={store} onChange={e => setStore(e.target.value)}><option value="all">전체 스토어</option>{availableStores.map(s => <option key={s} value={s}>{storeName(s)}</option>)}</select></label></div>
          <div className="filter-group"><label><span>순위 유형</span><select aria-label="순위 유형" value={chart} onChange={e => {setChart(e.target.value);setStore('all');setSnapshot('latest');setGenre('all');}}><option value="free">무료 순위</option>{(country === '미국' || country === 'all') && <option value="grossing">매출 순위</option>}</select></label></div>
          <div className="top-switch" role="group" aria-label="분석 범위">{[10,20,100].map(n => <button key={n} className={top === n ? 'active' : ''} onClick={() => setTop(n)}>Top {n}</button>)}</div>
        </div>
        <div className="scope-note"><span>{countryLabel} / {chart === 'free' ? 'Top Free' : 'Top Grossing'} / Top {top}</span><span>{availableStores.length === 1 ? '이 시장은 Google Play 자료만 있습니다.' : selectedDates.length > 1 ? '스토어별 조사일이 다릅니다.' : '선택한 기준일의 자료입니다.'}{country !== '미국' && chart === 'free' ? ' 미국 외 시장의 무료 순위는 Top 20까지만 수집했습니다.' : ''}{country === 'all' && chart === 'grossing' ? ' 매출 순위는 미국 자료만 수집했습니다.' : ''}</span></div>
      </>}

      {section === 'overview' && <>
        <section className="stats" aria-label="조사 범위 요약">
          <div><span className="stat-label">조사 게임 <span>중복 제외</span></span><strong>{unique.toLocaleString()}<small>개</small></strong><span className="stat-foot">확인된 {known.length}개 스토어 기록</span></div>
          <div><span className="stat-label">최다 장르</span><strong className="genre-stat">{leader?.name || '자료 없음'}</strong><span className="stat-foot">확인된 기록의 {percent(leaderCount, known.length)}%</span></div>
          <div><span className="stat-label">양쪽 스토어 진입</span><strong>{ios.length && android.length ? overlap : '—'}<small>{ios.length && android.length ? '개' : ''}</small></strong><span className="stat-foot">동일 게임명 매핑 기준</span></div>
          <div><span className="stat-label">수집 범위</span><strong>{scope.length}<small>건</small></strong><span className="stat-foot">누락 {scope.length - known.length}건 포함 / 실제 다운로드 수 아님</span></div>
        </section>
        <section className="editorial"><div className="editorial-heading"><span className="section-index">01 / MARKET SIGNAL</span><ChartBar size={23}/></div><div className="editorial-body"><h2>{leader ? `${leader.name}, ${chart === 'free' ? '무료' : '매출'} 시장의 중심을 차지하다.` : '선택한 범위의 자료가 없습니다.'}</h2><p><ResearchText text={leader ? `${countryLabel} ${chart === 'free' ? '무료' : '매출'} Top ${top}의 확인된 순위 기록 ${known.length}건 중 ${leaderCount}건이 ${leader.name}입니다. 장르 비중과 개별 게임의 수익 모델을 함께 살펴보세요.` : '시장이나 기준일을 변경해 자료를 확인하세요.'}/></p><button onClick={() => {setGenre(leader?.name || 'all');navigate('rankings');}}>해당 장르 게임 보기 <ArrowUpRight size={18}/></button></div><div className="editorial-number"><strong>{percent(leaderCount,known.length)}<span>%</span></strong><span>확인된 스토어 기록 기준</span></div></section>
        <GenreChart counts={genreCounts} ios={ios.length} android={android.length} top={top} onGenre={g => {setGenre(g);navigate('rankings');}}/>
      </>}

      {(section === 'rankings' || section === 'comparison' || section === 'overview') && <section className="ranking-section">
        <div className="section-title"><div><span className="section-index">{section === 'overview' ? '03 / GAME INDEX' : 'GAME INDEX'}</span><h2>{section === 'comparison' ? '같은 게임, 다른 스토어 순위' : section === 'overview' ? '상위 게임 살펴보기' : '게임 리서치 목록'} <span className="title-count">{tableRows.length}</span></h2></div>{section === 'overview' && <button className="text-button" onClick={() => navigate('rankings')}>전체 순위 보기 <ArrowRight size={17}/></button>}</div>
        <div className="table-tools"><label className="search-field"><MagnifyingGlass size={19}/><input aria-label="게임 검색" placeholder="게임, 퍼블리셔, 장르 검색" value={query} onChange={e => setQuery(e.target.value)}/>{query && <button aria-label="검색 지우기" onClick={() => setQuery('')}><X size={15}/></button>}</label><label className="genre-filter"><FunnelSimple size={17}/><select aria-label="장르" value={genre} onChange={e => setGenre(e.target.value)}><option value="all">전체 장르</option>{genreNames.map(g => <option key={g}>{g}</option>)}</select></label><select aria-label="정렬" value={sort} onChange={e => setSort(e.target.value)}><option value="rank">최고 순위순</option><option value="name">게임 이름순</option></select>{savedOnly && <button className="filter-chip" onClick={() => setSavedOnly(false)}>관심 게임만 <X size={13}/></button>}</div>
        <p className="scroll-hint">표를 좌우로 스크롤하면 스토어별 순위와 조사 상태를 볼 수 있습니다.</p>
        <div className="table-scroll" role="region" aria-label="게임 순위 표" tabIndex={0}><table className="game-table"><thead><tr><th className="compare-col">비교</th><th>게임 / 퍼블리셔</th><th>장르 / 세부 장르</th><th className="rank-column"><span className="legend-dot ios"/>App Store</th><th className="rank-column"><span className="legend-dot android"/>Google Play</th><th>조사 상태</th><th><span className="sr-only">관심 게임</span></th></tr></thead><tbody>{tableRows.map(group => {
          const o = group[0]; const game = gameMap.get(o.gameId)!;
          return <tr key={o.gameId}><td><button className={compare.includes(game.id) ? 'compare-check checked' : 'compare-check'} disabled={game.missing} aria-label={`${o.name} 비교 선택`} aria-pressed={compare.includes(game.id)} onClick={() => toggleCompare(game.id)}>{compare.includes(game.id) ? <Check size={12}/> : <Plus size={12}/>}</button></td><td><button className="game-name" onClick={() => setSelected(game)}><span aria-hidden="true" className={`game-initial tone-${Math.min(7, Math.max(0,genreNames.indexOf(o.genre)))}`}>{game.missing ? '?' : o.name.slice(0,1).toUpperCase()}</span><span><strong>{o.name}</strong><small>{o.publisher}</small>{country === 'all' && <small className="market-coverage" title={[...new Set(group.map(item => item.country))].join(', ')}>{[...new Set(group.map(item => item.country))].join(' · ')}</small>}</span><ArrowUpRight className="row-arrow" size={15}/></button></td><td><span className="genre-label">{o.genre}</span><small className="subgenre">{o.subgenre || '미분류'}</small></td><td className="rank-value">{rankOf(o.gameId,'ios') ?? '—'}</td><td className="rank-value">{rankOf(o.gameId,'android') ?? '—'}</td><td><span className={game.missing || group.some(item => item.notes?.includes('확인')) ? 'research-status pending' : 'research-status'}>{game.missing ? '순위 누락' : group.some(item => item.notes?.includes('확인')) ? '확인 필요' : game.gameplay ? '분석 있음' : '순위 수집'}</span></td><td><button className="icon-button" aria-label={`${o.name} 관심 게임 ${saved.includes(game.id) ? '해제' : '저장'}`} aria-pressed={saved.includes(game.id)} onClick={() => bookmark(game.id)}><BookmarkSimple size={18} weight={saved.includes(game.id) ? 'fill' : 'regular'}/></button></td></tr>;
        })}</tbody></table></div>
        {!tableRows.length && <Empty onReset={() => {setQuery('');setGenre('all');setSavedOnly(false);setStore('all');}}/>}
        <div className="table-bottom"><span><strong className="text-number">{tableRows.length}개 게임 표시</strong> · Top {top} 범위의 검색 결과 {grouped.length}개 중 최대 {top}개<span className="meta-divider">/</span>— 선택 범위에서 미수집</span></div><p className="ranking-scope-note">각 시장·스토어의 Top {top}에 진입한 게임을 중복 제거하고, 선택한 정렬 기준으로 최대 {top}개 표시합니다.{country === 'all' && ' 전체 시장의 순위는 국가별 기록 중 게임의 최고 순위이며, 세계 통합 순위가 아닙니다.'}</p>
        {section === 'comparison' && <p className="data-note">스토어별 날짜와 장르 분류는 개별 기록에 유지됩니다. 게임명을 정규화하고 검토된 이름 매핑을 적용한 비교이며, 미수집은 해당 스토어에 게임이 없다는 의미가 아닙니다.</p>}
      </section>}

      {['gameplay','monetization','marketing','art'].includes(section) && <section className="research-section"><div className="section-title"><div><span className="section-index">GAME STRATEGY / RESEARCH NOTES</span><h2>{section === 'gameplay' ? '성공한 게임의 핵심 루프' : section === 'monetization' ? '수익 모델과 순위의 관계' : section === 'marketing' ? '유입 전략과 후킹 요소' : '반복해서 등장하는 시각적 접근'}</h2></div></div>
        <p className="data-note">아래 내용은 제공된 전략 워크북의 분석입니다. 추정 또는 공개 정보 기반 판단은 각 기록의 근거 수준을 확인하세요.</p>
        {section === 'marketing' && <aside className="method-note"><strong>광고 효과의 인과관계는 아직 미확인</strong><p><ResearchText text={"무료 순위와 추정 누적 설치만으로 광고의 효과를 계산할 수 없습니다. 캠페인별 광고비·기간·소재와 유료/자연 유입, 전환 이벤트, 대조군 자료가 필요합니다. 현재 자료의 UA·타깃 해석은 분석자 추정입니다."}/></p></aside>}
        <label className="search-field research-search"><MagnifyingGlass size={19}/><input aria-label="게임 검색" placeholder="게임 또는 퍼블리셔 검색" value={query} onChange={e => setQuery(e.target.value)}/></label>
        {researchGames.filter(game => game[section === 'gameplay' ? 'gameplay' : section === 'monetization' ? 'monetization' : section === 'marketing' ? 'marketing' : 'artStyle']).map((game,index) => <article className="research-row" key={game.id}><span className="research-number">{String(index+1).padStart(2,'0')}</span><div className="research-identity"><button className="text-button" onClick={() => setSelected(game)}>{game.name}<ArrowUpRight size={16}/></button><small>{game.publisher}</small><span className="genre-label">{game.genre}</span></div><div className="research-content"><p><ResearchText text={game[section === 'gameplay' ? 'gameplay' : section === 'monetization' ? 'monetization' : section === 'marketing' ? 'marketing' : 'artStyle'] || '미수집'}/></p>{section === 'marketing' && game.marketingHook && <div className="hook-hypothesis"><span>소재 테스트 가설 · 미검증</span><p><ResearchText text={game.marketingHook}/></p></div>}<span className="evidence-label">근거: {game.evidence}</span><button className="source-link" onClick={() => game.analysis && showDocument(game.analysis.file,game.analysis.sheet)}>원자료 {game.analysis?.sheet} · {game.analysis?.row}행 <ArrowUpRight size={13}/></button></div></article>)}
        {!researchGames.some(game => game[section === 'gameplay' ? 'gameplay' : section === 'monetization' ? 'monetization' : section === 'marketing' ? 'marketing' : 'artStyle']) && <Empty onReset={() => {setQuery('');setGenre('all');setSavedOnly(false);}}/>}
        {section === 'monetization' && <button className="text-button end-link" onClick={() => showDocument('시장조사.xlsx','3-3 매출-다운로드 비교')}>매출·무료 순위와 추정 설치 원자료 <ArrowRight size={16}/></button>}
      </section>}

      {section === 'regions' && <section><div className="section-title"><div><span className="section-index">GEOGRAPHY / GOOGLE PLAY TOP 20</span><h2>국가별로 다른 장르 구성</h2></div></div><p className="data-note">각 국가의 최신 무료 순위 Top 20을 동일한 범위로 집계합니다. 조사일이 달라 시점에 따른 차이가 포함됩니다. Google Play 기준입니다.</p><div className="region-grid">{countries.map(c => {
        const records = data.observations.filter(o => o.country === c && o.store === 'android' && o.chart === 'free' && String(o.year) === year && o.rank <= 20);
        const date = records.map(o => o.date).sort().at(-1); const rows = records.filter(o => o.date === date);
        const counts = [...new Set(rows.map(o => o.genre))].map(g => ({name:g,count:rows.filter(o => o.genre === g).length})).sort((a,b) => b.count-a.count);
        return <article key={c} className="region-item"><div><h3>{c}</h3><small>{date ? dateLabel(date) : '자료 없음'}</small></div><div className="stack-bar" role="img" aria-label={counts.map(g => `${g.name} ${g.count}개`).join(', ')}>{counts.map(g => <span key={g.name} style={{width:`${percent(g.count,rows.length)}%`,background:colors[Math.max(0,genreNames.indexOf(g.name)) % colors.length]}}/>)}</div><p>{counts[0]?.name || '자료 없음'} <strong>{percent(counts[0]?.count || 0,rows.length)}%</strong></p><button className="text-button" onClick={() => {setCountry(c);setStore('android');setChart('free');setSnapshot('latest');setTop(20);setGenre('all');setQuery('');navigate('rankings');}}>Top 20 보기<ArrowUpRight size={16}/></button></article>;
      })}</div><button className="text-button end-link" onClick={() => showDocument('시장조사.xlsx','2-5 지역별 특징')}>국가별 특징과 추정 근거 <ArrowRight size={16}/></button></section>}

      {section === 'opportunity' && <section><div className="section-title"><div><span className="section-index">DEVELOPMENT / HYPOTHESES</span><h2>시장 패턴을 개발 가설로</h2></div></div><p className="data-note">시장 공백과 성공 가능성은 검증 전 가설입니다. 아래 워크북의 MVP 범위와 검증 지표를 함께 확인합니다.</p>{['1-7 성공 공식 TOP5','2-1 MVP 명세','2-2 MVP 검증 목표'].map(sheet => <article className="opportunity-row" key={sheet}><span className="section-index">{sheet.split(' ')[0]}</span><div><h3>{sheet.slice(4)}</h3><p>{sheet.includes('TOP5') ? '간단한 규칙, 반복 플레이, 메타 진행과 소셜 요소에 대한 원자료의 분석.' : sheet.includes('명세') ? '필수 기능과 후속 기능, 예상 인원·기간을 게임별로 비교합니다. 규모와 일정은 분석자 추정입니다.' : '리텐션·수익 참고값의 지역, 백분위, 조사 기간을 확인합니다. 벤치마크는 목표 설정의 참고 자료입니다.'}</p></div><button className="text-button" onClick={() => showDocument('게임전략_MVP_마케팅.xlsx',sheet)}>자료 보기<ArrowUpRight size={18}/></button></article>)}<aside className="method-note"><strong>현재 자료가 답하지 못하는 질문</strong><p><ResearchText text={"실제 캠페인 ROAS, 게임별 광고 매출, 정확한 유저 연령, 국가별 개발 경쟁 강도는 미수집입니다. 무료 순위만으로 수익성이나 미개척 시장을 판단하지 않습니다."}/></p></aside></section>}

      {section === 'sources' && <section className="source-section"><div className="source-summary"><Stack size={23}/><div><strong>{data.files.length}개 워크북 · {data.documents.length}개 시트</strong><p>가져온 시각 {new Date(data.importedAt).toLocaleString('ko-KR',{timeZone:'America/Los_Angeles'})} (미국 태평양 시간). 원자료의 조사일과는 별개입니다.</p></div></div><div className="source-controls"><label><span className="sr-only">원자료 시트</span><select aria-label="원자료 시트" value={documentId} onChange={e => {setDocumentId(e.target.value);setDocumentQuery('');}}>{data.files.map(file => <optgroup label={file.name} key={file.name}>{data.documents.filter(d => d.file === file.name).map(d => <option key={d.id} value={d.id}>{d.sheet}</option>)}</optgroup>)}</select></label><label className="search-field"><MagnifyingGlass size={18}/><input aria-label="원자료 검색" placeholder="시트 안에서 검색" value={documentQuery} onChange={e => setDocumentQuery(e.target.value)}/></label></div><div className="section-title"><div><span className="section-index">{selectedDocument.file}</span><h2>{selectedDocument.sheet}</h2></div></div><RawDocument document={selectedDocument} query={documentQuery}/><p className="data-note">원본 워크북의 저장된 셀 값을 표시합니다. 비어 있는 값은 미수집 또는 수식 캐시 미저장을 의미할 수 있습니다. 이 화면에서 원본 파일을 수정하지 않습니다. 보고서 발행 연도와 데이터 관측 연도는 개별 시트에 기재된 범위를 따릅니다.</p></section>}

      <section className="methodology"><span className="section-index">READING THE DATA</span><p><ResearchText text={"무료·매출 순위는 해당 시점의 상대적 순위입니다. 장르와 게임 전략은 분석자 분류이며, 추정 설치 수는 실제 다운로드 수와 다릅니다. 출처와 조사일은 각 게임의 상세 정보에서 확인할 수 있습니다."}/></p><button className="text-button" onClick={() => showDocument('시장조사.xlsx','0-1 개요')}>조사 방법과 범위<ArrowUpRight size={15}/></button></section>
      <footer><span className="footer-brand">PLAYFIELD</span><span>Mobile Game Market Research</span><span>팀 내부 리서치 자료</span></footer>
    </main>
    {notice && <div className="toast" role="status">{notice}<button aria-label="알림 닫기" onClick={() => setNotice('')}><X size={16}/></button></div>}
    {compare.length > 0 && <div className="compare-tray"><span><SquaresFour size={18}/><strong>{compare.length}/6</strong> 비교할 게임 선택</span><button onClick={() => {setSelected(null);setCompareOpen(true);}}>게임 비교<ArrowRight size={16}/></button><button className="tray-clear" aria-label="비교 선택 지우기" onClick={() => setCompare([])}><X size={18}/></button></div>}
    <dialog ref={dialog} className={compareOpen ? 'detail-dialog compare-dialog' : 'detail-dialog'} onCancel={() => {setSelected(null);setCompareOpen(false);}} onClick={e => {if (e.target === dialog.current) {setSelected(null);setCompareOpen(false);}}}>
      <div className="dialog-inner"><button className="dialog-close icon-button" aria-label="상세 정보 닫기" onClick={() => {setSelected(null);setCompareOpen(false);}}><X size={23}/></button>
        {selected && <><span className="eyebrow">GAME RESEARCH</span><h2>{selected.name}</h2><p className="detail-publisher">{selected.publisher} / {selected.genre}</p><div className="detail-ranks">{base.filter(o => o.gameId === selected.id).map(o => <div key={o.id}><span>{storeName(o.store)} {chart === 'free' ? '무료' : '매출'}</span><strong>#{o.rank}</strong><small>{o.country} · {dateLabel(o.date)}</small><small>{o.notes}</small><a href={o.source.url || undefined} target="_blank" rel="noreferrer">출처 열기 <ArrowUpRight size={13}/></a><button className="source-link" onClick={() => showDocument(o.source.file,o.source.sheet)}>{o.source.sheet} · {o.source.row}행</button></div>)}</div><dl className="detail-fields">{[['핵심 게임플레이',selected.gameplay],['Monetization / BM',selected.monetization],['마케팅',selected.marketing],['아트 스타일',selected.artStyle],['타깃 / 연령',selected.audience],['실제 다운로드 / 매출',null]].map(([label,value]) => <div key={label}><dt>{label}</dt><dd><ResearchText text={value || '미수집'}/></dd></div>)}{selected.downloads && <div><dt>추정 누적 설치</dt><dd><strong className="text-number">{compact(selected.downloads.value)}</strong> · {selected.downloads.kind} · Google Play<br/><small>{selected.downloads.date} / 국가별 수치 아님</small></dd></div>}</dl>{selected.evidence && <p className="data-note">근거 수준: {selected.evidence}</p>}<div className="dialog-actions"><button className="export-button" disabled={selected.missing} onClick={() => toggleCompare(selected.id)}>{compare.includes(selected.id) ? '비교에서 해제' : '비교에 추가'}<Plus size={16}/></button><button className="plain-button" onClick={() => bookmark(selected.id)}><BookmarkSimple size={18} weight={saved.includes(selected.id) ? 'fill' : 'regular'}/>{saved.includes(selected.id) ? '관심 게임 저장됨' : '관심 게임 저장'}</button></div></>}
        {compareOpen && <><span className="eyebrow">SIDE BY SIDE</span><h2>게임 비교</h2><p className="data-note">현재 시장·순위 범위의 기록을 비교합니다. 전략 분석의 근거 수준은 게임별로 다릅니다.</p><div className="table-scroll"><table className="comparison-table" style={{minWidth: Math.max(650, 140 + compare.length * 220)}}><thead><tr><th>항목</th>{compare.map(id => <th key={id}>{gameMap.get(id)?.name}</th>)}</tr></thead><tbody>{['publisher','genre','ios','android','gameplay','monetization','marketing','artStyle','evidence'].map(field => <tr key={field}><th>{({publisher:'퍼블리셔',genre:'장르',ios:'App Store 순위',android:'Google Play 순위',gameplay:'핵심 루프',monetization:'수익 모델',marketing:'마케팅',artStyle:'아트',evidence:'근거 수준'} as Record<string,string>)[field]}</th>{compare.map(id => {const game = gameMap.get(id)!;return <td key={id}>{field === 'ios' || field === 'android' ? rankOf(id,field) || '선택 범위에서 미수집' : <ResearchText text={String(game[field as keyof Game] || '미수집')}/>}</td>;})}</tr>)}</tbody></table></div></>}
      </div>
    </dialog>
  </>;
}

function GenreChart({counts,ios,android,top,onGenre}: {counts:{name:string;ios:number;android:number}[];ios:number;android:number;top:number;onGenre:(name:string)=>void}) {
  const scale = Math.max(1,...counts.flatMap(c => [c.ios,c.android]));
  return <section className="genre-section"><div className="section-title"><div><span className="section-index">02 / GENRE LANDSCAPE</span><h2>상위권은 어떤 장르로 이루어져 있나</h2></div><div className="chart-legend"><span><i className="legend-dot ios"/>App Store <small>{ios}건</small></span><span><i className="legend-dot android"/>Google Play <small>{android}건</small></span></div></div><div className="chart-head"><span>장르</span><span>Top {top} 내 게임 수 · 스토어 기록별 집계</span></div><div className="genre-chart">{counts.map(g => <div className="genre-chart-row" key={g.name}><button onClick={() => onGenre(g.name)}>{g.name}<ArrowDownRight size={14}/></button><div className="bar-pair"><div className="bar-line"><span className="chart-bar ios-bar" style={{width:`${g.ios/scale*85}%`}}/><span className="bar-value">{g.ios}</span></div><div className="bar-line"><span className="chart-bar android-bar" style={{width:`${g.android/scale*85}%`}}/><span className="bar-value">{g.android}</span></div></div><div className="chart-percent">{percent(g.ios,ios)}%<small>{percent(g.android,android)}%</small></div></div>)}</div><p className="chart-caption">분모는 각 스토어의 수집된 순위 기록입니다. 순위 누락은 분모에 포함하고 장르 막대에서 제외합니다. 장르를 누르면 해당 게임 목록으로 이동합니다.</p></section>;
}
function Empty({onReset}: {onReset:()=>void}) {return <div className="empty-state"><MagnifyingGlass size={27}/><h3>조건에 맞는 자료가 없습니다.</h3><p><ResearchText text={"검색어와 필터 또는 수집 범위를 확인하세요."}/></p><button className="text-button" onClick={onReset}>검색·장르·관심 필터 초기화<ArrowRight size={16}/></button></div>;}
function RawDocument({document,query}: {document:ResearchDocument;query:string}) {
  const rows = document.rows.map((values,index) => ({values,index})).filter(row => row.values.some(v => v !== null && v !== '') && row.values.some(v => String(v ?? '').toLowerCase().includes(query.toLowerCase())));
  const maxCols = Math.max(...document.rows.map(row => {let last = row.length;while(last > 0 && row[last-1] === null) last--;return last;}));
  return <div className="table-scroll raw-scroll"><table className="raw-table"><thead><tr><th>행</th>{Array.from({length:maxCols},(_,i) => <th key={i}>{i<26 ? String.fromCharCode(65+i) : `열 ${i+1}`}</th>)}</tr></thead><tbody>{rows.map(({values,index}) => <tr key={index}><th>{index+1}</th>{values.slice(0,maxCols).map((value,i) => <td key={i}>{typeof value === 'string' ? value.split(/(https?:\/\/[^\s)]+)/g).map((part,j) => /^https?:\/\//.test(part) ? <a key={j} href={part} target="_blank" rel="noreferrer">{part}</a> : <ResearchText key={j} text={part}/>) : value === null ? '—' : <strong className="text-number">{String(value)}</strong>}</td>)}</tr>)}</tbody></table>{!rows.length && <div className="empty-state">검색 결과가 없습니다.</div>}</div>;
}
