// タスクボード「かんたんガイド」PDF を生成するスクリプト。
//
// 使い方（リポジトリのルートで）:
//   npm install
//   npx playwright install chromium
//   node apps/task-board/guide/build_guide.mjs
//
// 出力: apps/task-board/タスクボード_かんたんガイド.pdf
//   1) アプリ(../index.html)にサンプルデータを入れて画面を board.png に撮影
//   2) guide.html を A4 で PDF 化（board.png を埋め込み）
import { chromium } from 'playwright';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const appIndex = resolve(here, '..', 'index.html');
const guideHtml = join(here, 'guide.html');
const boardPng = join(here, 'board.png');
const outPdf = resolve(here, '..', 'タスクボード_かんたんガイド.pdf');

const browser = await chromium.launch();

// 1) 画面イメージ（NEON）を撮影
const ctx = await browser.newContext({ viewport: { width: 430, height: 820 }, deviceScaleFactor: 2 });
const page = await ctx.newPage();
await page.goto(pathToFileURL(appIndex).href, { waitUntil: 'load' });
await page.evaluate(() => {
  const pad = (n) => (n < 10 ? '0' : '') + n;
  const ymd = (d) => d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
  const add = (n) => { const d = new Date(); d.setDate(d.getDate() + n); return ymd(d); };
  const mondayOf = (s) => { const q = s.split('-'); const d = new Date(+q[0], +q[1] - 1, +q[2]); const w = d.getDay(); d.setDate(d.getDate() + (w === 0 ? -6 : 1 - w)); return ymd(d); };
  const today = ymd(new Date()), thisMon = mondayOf(today), nextMon = mondayOf(add(7));
  const T = (title, o) => Object.assign({ id: Math.random().toString(36).slice(2), title, deadline: '', workDate: '', workWeek: '', prio: false, done: false, t: Math.random() }, o);
  const st = {
    tasks: [
      T('進度予定表を提出', { workDate: today, prio: true, recurringId: 'r1', recurKey: 'r1|' + today }),
      T('保護者へ折り返し電話', { workDate: today }),
      T('中間テスト作問', { workWeek: nextMon, prio: true, deadline: add(9) }),
      T('遺伝の授業プリント修正', { workDate: add(3) }),
      T('席替えの希望を集計', { workWeek: thisMon }),
    ],
    showDone: { today: false, week: false },
    recurring: [{ id: 'r1', title: '進度予定表を提出', prio: true, freq: 'weekly', weekdays: [5], monthday: 1, holidayShift: 'before', t: 1 }],
    recurredKeys: ['r1|' + today], recurBreaks: [],
  };
  localStorage.setItem('taskboard.v2', JSON.stringify(st));
});
await page.reload({ waitUntil: 'load' });
await page.waitForTimeout(500);
await page.locator('#app').screenshot({ path: boardPng });
await ctx.close();

// 2) guide.html → PDF (A4)
const p2 = await browser.newPage();
await p2.goto(pathToFileURL(guideHtml).href, { waitUntil: 'load' });
await p2.waitForTimeout(400);
await p2.pdf({ path: outPdf, format: 'A4', printBackground: true, preferCSSPageSize: true });
await browser.close();
console.log('生成しました:', outPdf);
