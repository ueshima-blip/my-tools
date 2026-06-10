// OMRコアの自動テスト: 合成スキャン(BMP)を読み取り、正しい
// 出席番号とマークが得られることを検証する。
//   python3 make_synthetic.py && node test_omr.js
"use strict";
const fs = require("fs");
const path = require("path");
const OMR = require("../omr.js");

function readBmp(file) {
  const b = fs.readFileSync(file);
  if (b.toString("ascii", 0, 2) !== "BM") throw new Error("not BMP");
  const off = b.readUInt32LE(10);
  const w = b.readInt32LE(18);
  const h = b.readInt32LE(22);
  const bpp = b.readUInt16LE(28);
  if (bpp !== 24) throw new Error("need 24bit BMP");
  const row = (w * 3 + 3) & ~3;
  const data = new Uint8ClampedArray(w * h * 4);
  for (let y = 0; y < h; y++) {
    const src = off + (h - 1 - y) * row;
    for (let x = 0; x < w; x++) {
      const p = (y * w + x) * 4;
      data[p] = b[src + x * 3 + 2];
      data[p + 1] = b[src + x * 3 + 1];
      data[p + 2] = b[src + x * 3];
      data[p + 3] = 255;
    }
  }
  return { width: w, height: h, data };
}

const dir = __dirname;
const def = OMR.parseDefinition(fs.readFileSync(path.join(dir, "test_def.csv"), "utf-8"));
console.log(`def: ${def.days}日 x ${def.slots}コマ, 生徒 ${def.students.length}名, ラベル ${def.labels.length}`);

const manifest = fs.readFileSync(path.join(dir, "manifest.txt"), "utf-8").trim().split("\n");
let failures = 0;
const pages = [];

for (const entry of manifest) {
  const [file, idStr, marksStr] = entry.split(";");
  const expectId = parseInt(idStr, 10);
  const expectMarks = new Set(marksStr ? marksStr.split("|") : []);
  const img = readBmp(path.join(dir, file));
  const res = OMR.decodePage(img, def);

  if (!res.ok) {
    console.log(`✗ ${file}: 解析失敗 (${res.reason})`);
    failures++;
    continue;
  }
  const gotMarks = new Set(res.marks.filter((m) => m.state === 1).map((m) => `${m.d}-${m.k}`));
  const ambiguous = res.marks.filter((m) => m.state === 2).length;

  const idPass = res.idOk && res.id === expectId;
  const marksPass =
    gotMarks.size === expectMarks.size && [...expectMarks].every((m) => gotMarks.has(m));

  console.log(
    `${idPass && marksPass ? "✓" : "✗"} ${file}: id=${res.id}(${res.idOk ? "ok" : "NG"}) ` +
    `marks=[${[...gotMarks].sort().join(" ")}] expected=[${[...expectMarks].sort().join(" ")}] ` +
    `要確認=${ambiguous}`
  );
  if (!idPass || !marksPass) failures++;
  pages.push({ id: res.id, marks: res.marks });
}

// CSV出力の形が取り込みパーサ(VBA)の期待と一致するか
const csv = OMR.buildCsv(def, pages);
const lines = csv.trim().split("\r\n");
if (!lines[0].startsWith("番号,氏名,6/20 13:00")) {
  console.log("✗ CSVヘッダ不正:", lines[0].slice(0, 60));
  failures++;
} else {
  console.log("✓ CSV: " + lines.length + " 行, ヘッダOK");
}

if (failures) {
  console.log(`\n${failures} failure(s)`);
  process.exit(1);
}
console.log("\nALL PASS");
