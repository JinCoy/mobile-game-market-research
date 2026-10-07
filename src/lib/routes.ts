import type {Dataset} from './schema';

/** Analysis views built into the site. Sheets come from each workbook's 0 목차 at import time. */
export const sections = [
  ['overview', '시장 개요'], ['rankings', '게임 순위'], ['comparison', '스토어 비교'],
  ['gameplay', '게임플레이'], ['monetization', '수익 모델'], ['marketing', '마케팅'],
  ['art', '아트 디렉션'], ['regions', '지역 비교'], ['opportunity', '개발 기회'], ['sources', '원자료 시트'],
  ['updates', '업데이트 기록'],
] as const;
export type Section = typeof sections[number][0];
export const sectionLabel = (id: Section) => sections.find(([s]) => s === id)![1];

export const workbooks = [
  {file: '시장조사.xlsx', slug: 'market', label: '시장조사'},
  {file: '게임전략_MVP_마케팅.xlsx', slug: 'strategy', label: '게임전략'},
] as const;
/** Which 0 목차 group (leading number) each analysis view belongs to. */
export const viewPlacement: Partial<Record<Section, {slug: string; group: string}>> = {
  rankings: {slug: 'market', group: '1'}, comparison: {slug: 'market', group: '1'}, regions: {slug: 'market', group: '2'},
  gameplay: {slug: 'strategy', group: '1'}, monetization: {slug: 'strategy', group: '1'}, art: {slug: 'strategy', group: '1'},
  opportunity: {slug: 'strategy', group: '2'}, marketing: {slug: 'strategy', group: '3'},
};
export const sheetSlug = (sheet: string) => sheet.split(' ')[0];
export const groupKey = (group: string) => group.split(' ')[0];
export const bookBySlug = (slug: string) => workbooks.find(b => b.slug === slug);
export const bookByFile = (file: string) => workbooks.find(b => b.file === file);

export type Route = {section: Section; documentId?: string};
export function parseHash(hash: string, data: Dataset): Route {
  const [slug, part] = hash.replace(/^#\/?/, '').split('/').map(decodeURIComponent);
  if (slug === 'updates') return {section: 'updates'};
  const book = bookBySlug(slug);
  if (!book) return {section: 'overview'};
  const view = sections.find(([id]) => id === part && viewPlacement[id]?.slug === slug);
  if (view) return {section: view[0]};
  const sheet = part ?? data.toc.find(t => t.file === book.file)?.items[0]?.sheet;
  const doc = data.documents.find(d => d.file === book.file && (d.sheet === sheet || sheetSlug(d.sheet) === sheet));
  return doc ? {section: 'sources', documentId: doc.id} : {section: 'overview'};
}
export function hashFor(section: Section, doc?: {file: string; sheet: string}) {
  if (section === 'overview') return '#/';
  if (section === 'updates') return '#/updates';
  if (section === 'sources' && doc) return `#/${bookByFile(doc.file)?.slug}/${sheetSlug(doc.sheet)}`;
  const place = viewPlacement[section];
  return place ? `#/${place.slug}/${section}` : '#/';
}
