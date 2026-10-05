import {test,expect} from '@playwright/test';
import fs from 'node:fs';
import data from '../data/market.json' with {type:'json'};

test('data preserves all rank slots, dates, missing values and traceable origins', () => {
  for (const store of ['ios','android']) {
    const rows = data.observations.filter(o => o.country === '미국' && o.chart === 'free' && o.store === store);
    expect(rows.length).toBe(100);
    expect(new Set(rows.map(o=>o.rank)).size).toBe(100);
    expect(rows.filter(o=>o.genre === '퍼즐').length).toBe(store === 'ios' ? 48 : 53);
    expect(rows.every(o=>o.year === 2026 && o.source.row >= 2)).toBe(true);
  }
  expect(data.games.filter(g=>g.missing).length).toBe(1);
  expect(data.games.find(g=>g.missing)?.name).toBe('확인 불가');
  expect(data.observations.find(o=>o.store==='android' && o.chart==='free' && o.country==='미국' && o.rank===80)?.notes).toContain('누락');
});

test('search, ranking bands, comparison, bookmarks and export work', async ({page}) => {
  const errors:string[]=[];page.on('pageerror',error=>errors.push(error.message));
  await page.goto('/');
  await expect(page.getByRole('heading',{level:1})).toHaveText('모바일 게임 시장, 한눈에.');
  await page.getByRole('button',{name:'게임 순위',exact:true}).click();
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
  expect(await page.locator('.game-table tbody tr').count()).toBeLessThanOrEqual(20);
  await page.getByRole('button',{name:'Top 100',exact:true}).click();
  await page.getByRole('button',{name:'다음',exact:true}).click();
  await expect(page.locator('.pagination')).toContainText('2 /');
  await page.getByRole('combobox',{name:'시장',exact:true}).selectOption('캐나다');
  await expect(page.locator('.scope-note')).toContainText('Google Play 자료만');
  await expect(page.locator('.pagination')).toContainText('1 / 1');
  await page.reload();
  await expect(page.locator('.small-count')).toHaveText('1');
  expect(errors).toEqual([]);
});

test('research, region navigation and source sheets are usable', async ({page}) => {
  await page.goto('/');
  await page.getByRole('button',{name:'마케팅',exact:true}).click();
  await expect(page.locator('.method-note')).toContainText('인과관계는 아직 미확인');
  await expect(page.locator('.research-row')).not.toHaveCount(0);
  await page.locator('.research-row').first().getByRole('button',{name:/원자료/}).click();
  await expect(page.getByRole('heading',{level:1})).toHaveText('자료·출처');
  await expect(page.locator('.raw-table')).toContainText('핵심 재미 루프');
  await page.getByRole('textbox',{name:'원자료 검색'}).fill('Meowdoku');
  await expect(page.locator('.raw-table tbody tr')).toHaveCount(1);
  await page.getByRole('button',{name:'지역 비교',exact:true}).click();
  await expect(page.locator('.region-item')).toHaveCount(7);
  await page.locator('.region-item').filter({has:page.getByRole('heading',{name:'브라질',exact:true})}).getByRole('button').click();
  await expect(page.getByRole('combobox',{name:'시장',exact:true})).toHaveValue('브라질');
  await page.getByRole('textbox',{name:'게임 검색'}).fill('does not exist');
  await expect(page.locator('.empty-state')).toBeVisible();
});

for (const width of [1920,1440,834,390]) {
  test(`responsive layout and console at ${width}px`, async ({page}) => {
    const errors:string[]=[];
    page.on('pageerror',error=>errors.push(error.message));
    page.on('console',message=>{if(message.type()==='error') errors.push(message.text());});
    await page.setViewportSize({width,height:1000});
    await page.goto('/');
    await page.waitForLoadState('networkidle');
    await expect(page.locator('h1')).toBeVisible();
    expect(await page.evaluate(()=>document.documentElement.scrollWidth <= window.innerWidth)).toBe(true);
    await page.screenshot({path:`test-results/overview-${width}.png`,fullPage:true});
    await page.getByRole('button',{name:'게임 순위',exact:true}).click();
    await page.getByRole('combobox',{name:'스토어',exact:true}).selectOption('ios');
    await page.getByRole('combobox',{name:'장르',exact:true}).selectOption('퍼즐');
    await expect(page.locator('.game-table tbody tr')).toHaveCount(20);
    expect(await page.evaluate(()=>document.documentElement.scrollWidth <= window.innerWidth)).toBe(true);
    if(width===390) expect(await page.locator('.table-scroll').first().evaluate(el=>el.scrollWidth>el.clientWidth)).toBe(true);
    expect(errors).toEqual([]);
  });
}
