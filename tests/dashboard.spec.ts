import {test,expect} from '@playwright/test';
import fs from 'node:fs';
import data from '../data/market.json' with {type:'json'};
import {hashFor,sections} from '../src/lib/routes';
/** Navigate inside the open page (keeps local state) via the same hash links the menu uses. */
const go=(page:import('@playwright/test').Page,target:string)=>page.evaluate(h=>{location.hash=h},target.startsWith('#')?target:hashFor(sections.find(([,label])=>label===target)![0]));

test('rank graphs use the selected ordinal scale and distinguish missing observations', async ({page}) => {
  await page.goto('./');
  await expect(page.locator('.report-meta')).toContainText('v2.0.0');
  await page.getByRole('textbox',{name:'게임 검색'}).fill('Meowdoku');
  const rankCells=page.locator('.rank-value');
  await expect(rankCells).toHaveText(['1','2']);
  await expect(rankCells.nth(1).getByRole('img')).toHaveAttribute('aria-label','1위부터 100위 공통 순위 축에서 2위');
  await page.getByRole('button',{name:'Top 10',exact:true}).click();
  await expect(rankCells.nth(1).getByRole('img')).toHaveAttribute('aria-label','1위부터 10위 공통 순위 축에서 2위');
  await page.getByRole('button',{name:'Meowdoku! 비교 선택',exact:true}).click();
  await page.getByRole('textbox',{name:'게임 검색'}).fill('Block Out!');
  await page.getByRole('button',{name:'Block Out! - Color Sort Puzzle 비교 선택',exact:true}).click();
  await page.locator('.compare-tray').getByRole('button',{name:'게임 비교',exact:true}).click();
  const androidRow=page.locator('.comparison-table tbody tr').filter({hasText:'Google Play 순위'});
  // Block Out! has no Google Play observation in Top 10; missing is not zero.
  await expect(androidRow.locator('td').nth(1)).toHaveText('수집 Top 10에서 미관측');
  await expect(androidRow.locator('td').nth(1).locator('.chart-bar')).toHaveCount(0);
  await expect(androidRow.locator('td').first().getByRole('img')).toHaveAttribute('aria-label','1위부터 10위 공통 순위 축에서 2위');
});

test('all sections retain their layout and render the local design fonts', async ({page}) => {
  const errors:string[]=[];
  page.on('pageerror',error=>errors.push(error.message));
  page.on('console',message=>{if(message.type()==='error') errors.push(message.text());});
  await page.goto('./');
  await page.waitForLoadState('networkidle');
  await page.evaluate(()=>document.fonts.ready);
  expect(await page.evaluate(()=>document.fonts.check('500 14px "Inter Variable"','PLAYFIELD') && document.fonts.check('500 14px "IBM Plex Sans"','145') && document.fonts.check('400 14px "Noto Sans KR"','시장'))).toBe(true);
  expect(await page.locator('body').evaluate(el=>getComputedStyle(el).backgroundColor)).toBe('rgb(11, 14, 17)');
  const mainSections=await page.locator('.main-column>section').evaluateAll(els=>els.map(el=>el.className));
  expect(mainSections).toEqual(['stats','editorial','genre-section','ranking-section','methodology']);
  for(const width of [1440,390]) {
    await page.setViewportSize({width,height:1000});
    for(const section of ['시장 개요','게임 순위','스토어 비교','게임플레이','수익 모델','마케팅','아트 디렉션','지역 비교','개발 기회','#/market/3-4','#/updates']) {
      await go(page,section);
      await expect(page.getByRole('heading',{level:1})).toBeVisible();
      expect(await page.evaluate(()=>document.documentElement.scrollWidth<=window.innerWidth)).toBe(true);
      await page.screenshot({path:`test-results/section-${section}-${width}.png`,fullPage:true});
    }
  }
  expect(errors).toEqual([]);
});

test('data preserves all rank slots, dates, missing values and traceable origins', () => {
  for (const store of ['ios','android']) {
    const rows = data.observations.filter(o => o.country === '미국' && o.chart === 'free' && o.store === store);
    expect(rows.length).toBe(100);
    expect(new Set(rows.map(o=>o.rank)).size).toBe(100);
    expect(rows.filter(o=>o.genre === '퍼즐').length).toBe(store === 'ios' ? 48 : 53);
    expect(rows.every(o=>o.year === 2026 && o.source.row >= 2)).toBe(true);
  }
  // Google Play 무료 80위 + App Store 매출 89~100위 (AppBrain page lists 88 only).
  expect(data.games.filter(g=>g.missing).length).toBe(13);
  expect(data.games.filter(g=>g.missing).every(g=>g.name==='확인 불가')).toBe(true);
  const asGrossing=data.observations.filter(o=>o.store==='ios'&&o.chart==='grossing');
  expect(asGrossing).toHaveLength(100);
  expect(asGrossing.every(o=>o.date==='2026-10-06')).toBe(true);
  expect(asGrossing.filter(o=>o.rank>88).every(o=>o.name==='확인 불가'&&o.notes?.includes('88위까지'))).toBe(true);
  expect(data.observations.find(o=>o.store==='android' && o.chart==='free' && o.country==='미국' && o.rank===80)?.notes).toContain('누락');
});

test('search, ranking bands, comparison, bookmarks and export work', async ({page}) => {
  const errors:string[]=[];page.on('pageerror',error=>errors.push(error.message));
  await page.goto('./');
  await expect(page.getByRole('heading',{level:1})).toHaveText('모바일 게임 시장 분석');
  await expect(page.locator('.intro')).toHaveCount(0);
  await expect(page.getByRole('combobox',{name:'연도',exact:true})).toHaveCount(0);
  await expect(page.locator('.year-value')).toHaveText('2026');
  await go(page,'게임 순위');
  await page.getByRole('textbox',{name:'게임 검색'}).fill('Meowdoku');
  await expect(page.locator('.game-table tbody tr')).toHaveCount(1);
  await expect(page.locator('.rank-value')).toHaveText(['1','2']);
  await page.getByRole('button',{name:'Meowdoku! 비교 선택',exact:true}).click();
  await page.getByRole('button',{name:'Meowdoku! 관심 게임 저장',exact:true}).click();
  await page.locator('.game-name').first().click();
  await expect(page.locator('dialog')).toBeVisible();
  await expect(page.locator('dialog')).toContainText('칸 고르기');
  await page.keyboard.press('Escape');
  await expect(page.locator('dialog')).not.toBeVisible();
  const downloadPromise=page.waitForEvent('download');
  await page.getByRole('button',{name:'순위 CSV 내보내기'}).click();
  const download=await downloadPromise; const path=await download.path();
  expect(fs.readFileSync(path!,'utf8')).toContain('Meowdoku');
  await page.getByRole('textbox',{name:'게임 검색'}).fill('Rotate Rings');
  await page.getByRole('button',{name:'Rotate Rings 비교 선택',exact:true}).click();
  await page.locator('.compare-tray').getByRole('button',{name:'게임 비교',exact:true}).click();
  await expect(page.locator('.comparison-table')).toContainText('Meowdoku');
  await expect(page.locator('.comparison-table')).toContainText('Rotate Rings');
  await page.keyboard.press('Escape');
  await page.getByRole('textbox',{name:'게임 검색'}).fill('');
  await page.getByRole('button',{name:'Top 10',exact:true}).click();
  await expect(page.locator('.game-table tbody tr').first()).toBeVisible();
  await expect(page.locator('.game-table tbody tr')).toHaveCount(10);
  await page.getByRole('button',{name:'Top 100',exact:true}).click();
  await expect(page.locator('.game-table tbody tr')).toHaveCount(100);
  await page.getByRole('combobox',{name:'시장',exact:true}).selectOption('캐나다');
  await expect(page.locator('.scope-note')).toContainText('Google Play 자료만');
  await expect(page.locator('.game-table tbody tr')).toHaveCount(20);
  await page.reload();
  await expect(page.locator('.small-count')).toHaveText('1');
  expect(errors).toEqual([]);
});

test('research, region navigation and source sheets are usable', async ({page}) => {
  await page.goto('./');
  await go(page,'마케팅');
  await expect(page.locator('.method-note')).toContainText('인과관계는 아직 미확인');
  await expect(page.locator('.research-row')).not.toHaveCount(0);
  await page.locator('.research-row').first().getByRole('button',{name:/원자료/}).click();
  await expect(page.getByRole('heading',{level:1})).toHaveText(/^\d+-\d+ /);
  await expect(page.locator('.raw-table')).toContainText('핵심 재미 루프');
  await page.getByRole('textbox',{name:'원자료 검색'}).fill('Meowdoku');
  await expect(page.locator('.raw-table tbody tr')).toHaveCount(1);
  await go(page,'지역 비교');
  await expect(page.locator('.region-item')).toHaveCount(7);
  await page.locator('.region-item').filter({has:page.getByRole('heading',{name:'브라질',exact:true})}).getByRole('button').click();
  await expect(page.getByRole('combobox',{name:'시장',exact:true})).toHaveValue('브라질');
  await page.getByRole('textbox',{name:'게임 검색'}).fill('does not exist');
  await expect(page.locator('.empty-state')).toBeVisible();
});

test('Top controls change the actual overview and ranking row counts', async ({page}) => {
  await page.goto('./');
  for (const section of ['시장 개요','게임 순위','스토어 비교']) {
    await go(page,section);
    for (const count of [10,20,100]) {
      await page.getByRole('button',{name:`Top ${count}`,exact:true}).click();
      await expect(page.locator('.game-table tbody tr')).toHaveCount(count);
      await expect(page.locator('.table-bottom')).toContainText(`${count}개 게임 표시`);
    }
  }
});

test('all markets retain each country snapshot and display highest observed ranks', async ({page}) => {
  await page.goto('./');
  await page.getByRole('combobox',{name:'시장',exact:true}).selectOption('all');
  await expect(page.locator('.stats>div').last()).toContainText('320');
  await expect(page.locator('.game-table tbody tr')).toHaveCount(100);
  await page.getByRole('textbox',{name:'게임 검색'}).fill('Meowdoku');
  await expect(page.locator('.game-table tbody tr')).toHaveCount(1);
  await expect(page.locator('.rank-value')).toHaveText(['1','1']);
  await expect(page.locator('.market-coverage')).toContainText('캐나다');
  await expect(page.locator('.market-coverage')).toContainText('일본');
  await page.getByRole('textbox',{name:'게임 검색'}).fill('Block Blast!');
  await page.locator('.game-name').click();
  await expect(page.locator('.detail-ranks')).toContainText('2026.09.25');
  await expect(page.locator('.detail-ranks')).toContainText('2026.10.05');
  await page.keyboard.press('Escape');
  await page.getByRole('textbox',{name:'게임 검색'}).fill('');
  const downloadPromise=page.waitForEvent('download');
  await page.getByRole('button',{name:'순위 CSV 내보내기'}).click();
  const downloaded=await downloadPromise; const path=await downloaded.path();
  const csv=fs.readFileSync(path!,'utf8');
  for(const country of ['미국','캐나다','브라질','멕시코','일본','한국','인도네시아']) expect(csv).toContain(country);
});

test('source back restores the research section, filters, scroll and game detail', async ({page}) => {
  await page.goto('./');
  await go(page,'수익 모델');
  await page.getByRole('button',{name:'Top 20',exact:true}).click();
  await page.getByRole('textbox',{name:'게임 검색'}).fill('Meowdoku');
  const description=page.locator('.research-content>p').first();
  const originalDescription=data.games.find(g=>g.name==='Meowdoku!')!.monetization;
  await expect(description).toHaveText(originalDescription!);
  await expect(description.locator('.text-number')).toContainText(['6개국','100']);
  await expect(description.locator('.text-keyword')).toContainText(['광고 중심']);
  await page.locator('.research-row').first().getByRole('button',{name:/원자료/}).click();
  await expect(page.getByRole('heading',{level:1})).toHaveText(/^\d+-\d+ /);
  await page.getByRole('button',{name:'수익 모델로 돌아가기',exact:true}).click();
  await expect(page.getByRole('heading',{level:1})).toHaveText('수익 모델');
  await expect(page.getByRole('textbox',{name:'게임 검색'})).toHaveValue('Meowdoku');
  await expect(page.getByRole('button',{name:'Top 20',exact:true})).toHaveClass('active');
  await page.locator('.research-identity').getByRole('button').click();
  await expect(page.locator('dialog')).toBeVisible();
  await page.locator('.detail-ranks').getByRole('button').first().click();
  await expect(page.locator('dialog')).not.toBeVisible();
  await page.getByRole('button',{name:'수익 모델로 돌아가기',exact:true}).click();
  await expect(page.locator('dialog')).toBeVisible();
  await expect(page.locator('dialog h2')).toHaveText('Meowdoku!');
});

test('six comparison columns remain readable and a seventh selection is rejected', async ({page}) => {
  await page.goto('./');
  const buttons=page.locator('.compare-check:not([disabled])');
  for(let i=0;i<6;i++) await buttons.nth(i).click();
  await expect(page.locator('.compare-tray')).toContainText('6/6');
  await buttons.nth(6).click();
  await expect(page.locator('.toast')).toContainText('최대 6개');
  await page.locator('.compare-tray').getByRole('button',{name:'게임 비교',exact:true}).click();
  await expect(page.locator('.comparison-table thead th')).toHaveCount(7);
  expect(await page.locator('.comparison-table').evaluate(el=>el.getBoundingClientRect().width)).toBeGreaterThanOrEqual(1460);
  await page.setViewportSize({width:390,height:844});
  expect(await page.evaluate(()=>document.documentElement.scrollWidth <= window.innerWidth)).toBe(true);
  expect(await page.locator('.compare-dialog .table-scroll').evaluate(el=>el.scrollWidth>el.clientWidth)).toBe(true);
  await page.keyboard.press('Escape');
  await buttons.nth(0).click();
  await buttons.nth(6).click();
  await expect(page.locator('.compare-tray')).toContainText('6/6');
});

for (const width of [1920,1440,834,390]) {
  test(`responsive layout and console at ${width}px`, async ({page}) => {
    const errors:string[]=[];
    page.on('pageerror',error=>errors.push(error.message));
    page.on('console',message=>{if(message.type()==='error') errors.push(`${message.text()} ${message.location().url}`);});
    await page.setViewportSize({width,height:1000});
    await page.goto('./');
    await page.waitForLoadState('networkidle');
    await expect(page.locator('h1')).toBeVisible();
    expect(await page.evaluate(()=>document.documentElement.scrollWidth <= window.innerWidth)).toBe(true);
    await page.screenshot({path:`test-results/overview-${width}.png`,fullPage:true});
    await page.screenshot({path:`test-results/overview-viewport-${width}.png`});
    await go(page,'게임 순위');
    await page.getByRole('combobox',{name:'스토어',exact:true}).selectOption('ios');
    await page.getByRole('combobox',{name:'장르',exact:true}).selectOption('퍼즐');
    await expect(page.locator('.game-table tbody tr')).toHaveCount(48);
    expect(await page.evaluate(()=>document.documentElement.scrollWidth <= window.innerWidth)).toBe(true);
    if(width===390) expect(await page.getByRole('region',{name:'게임 순위 표'}).evaluate(el=>el.scrollWidth>el.clientWidth)).toBe(true);
    expect(errors).toEqual([]);
  });
}
