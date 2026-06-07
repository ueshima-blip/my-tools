// apps/ 配下の各 index.html をヘッドレス Chromium で開き、
// 読み込み時に未捕捉の JavaScript 例外（pageerror）が出ないかを確認するスモークテスト。
//
// 方針:
//  - 失敗とみなすのは「未捕捉の JS 例外」のみ。
//  - CDN（フォント・SheetJS・html2canvas 等）の読み込み失敗はネットワーク次第で
//    変動するため、CI を不安定にしないよう失敗扱いにしない。
//  - これにより「JS の構文ミス」「未定義関数の呼び出し」「onload で落ちる」など、
//    “うっかり壊した” を低ノイズで検知できる。
import { chromium } from 'playwright';
import { readdirSync, statSync, existsSync } from 'node:fs';
import { join, resolve, dirname } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const appsDir = join(root, 'apps');

const targets = [];
for (const name of readdirSync(appsDir).sort()) {
  const dir = join(appsDir, name);
  if (!statSync(dir).isDirectory()) continue;
  // 各アプリフォルダ内の .html をすべて対象にする（配付用オフライン版なども含む）
  for (const f of readdirSync(dir).sort()) {
    if (f.toLowerCase().endsWith('.html')) targets.push({ name: `${name}/${f}`, idx: join(dir, f) });
  }
}

if (targets.length === 0) {
  console.error('apps/*/index.html が見つかりません');
  process.exit(1);
}

const browser = await chromium.launch();
let failed = 0;

for (const { name, idx } of targets) {
  const page = await browser.newPage();
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  try {
    await page.goto(pathToFileURL(idx).href, { waitUntil: 'load', timeout: 30000 });
    await page.waitForTimeout(1500); // onload / 遅延初期化が走る余地を与える
  } catch (e) {
    errors.push('navigation: ' + e.message);
  }
  await page.close();

  if (errors.length) {
    failed++;
    console.log(`✗ ${name}`);
    for (const e of errors) console.log(`    ${e}`);
  } else {
    console.log(`✓ ${name}`);
  }
}

await browser.close();
console.log(`\n${targets.length - failed}/${targets.length} アプリが読み込みエラーなし`);
process.exit(failed ? 1 : 0);
