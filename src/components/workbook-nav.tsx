'use client';
import {useEffect, useRef, useState} from 'react';
import {CaretDown, ChartLineUp, ListBullets} from '@phosphor-icons/react';
import type {Dataset} from '@/lib/schema';
import {groupKey, hashFor, sectionLabel, sections, sheetSlug, viewPlacement, workbooks, type Section} from '@/lib/routes';

/** Sidebar on desktop, a toggled drawer below 1024px. Groups and sheets come from the workbook's 0 목차. */
export default function WorkbookToc({data, slug, section, sheet}: {data: Dataset; slug: string; section: Section; sheet?: string}) {
  const book = workbooks.find(b => b.slug === slug)!;
  const items = data.toc.find(t => t.file === book.file)?.items || [];
  const groups = [...new Set(items.map(i => i.group))];
  const currentGroup = section === 'sources' ? items.find(i => i.sheet === sheet)?.group : groups.find(g => groupKey(g) === viewPlacement[section]?.group);
  const currentLabel = section === 'sources' ? sheet : sectionLabel(section);
  const [open, setOpen] = useState(false);
  const toggle = useRef<HTMLButtonElement>(null);
  useEffect(() => setOpen(false), [section, sheet]);
  return <nav className="workbook-toc" aria-label={`${book.label} 목차`} onKeyDown={e => {if (e.key === 'Escape' && open) {setOpen(false); toggle.current?.focus();}}}>
    <button ref={toggle} className="toc-toggle" aria-expanded={open} aria-controls="toc-panel" onClick={() => setOpen(!open)}>
      <ListBullets size={18}/><span><small>{book.label} 목차</small>{currentLabel}</span><CaretDown size={16} className={open ? 'flip' : undefined}/>
    </button>
    <div id="toc-panel" className={open ? 'toc-panel open' : 'toc-panel'}>
      <p className="toc-title">{book.label} <small>{items.length}개 시트</small></p>
      {groups.map(group => {
        const views = sections.filter(([id]) => viewPlacement[id]?.slug === slug && viewPlacement[id]?.group === groupKey(group));
        const sheets = items.filter(i => i.group === group);
        return <details key={group} open={group === currentGroup || undefined}>
          <summary>{group}<small>{sheets.length}</small></summary>
          <ul>
            {views.map(([id, label]) => <li key={id}><a className="toc-view" href={hashFor(id)} aria-current={section === id ? 'page' : undefined}><ChartLineUp size={14}/>{label} 분석</a></li>)}
            {sheets.map(i => <li key={i.sheet}><a href={hashFor('sources', {file: book.file, sheet: i.sheet})} title={i.description || undefined} aria-current={section === 'sources' && sheet === i.sheet ? 'page' : undefined}><span className="toc-num">{sheetSlug(i.sheet)}</span>{i.sheet.slice(sheetSlug(i.sheet).length + 1)}</a></li>)}
          </ul>
        </details>;
      })}
    </div>
  </nav>;
}
