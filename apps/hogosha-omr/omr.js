/* 保護者会 調査票OMR — 画像処理コア（DOM非依存・Nodeでテスト可能）
 *
 * 入力: { width, height, data } (RGBA / ImageData互換) と 調査票定義
 * 出力: マーカー検出 → 射影補正なしの双一次マッピング → ID/マーク読み取り
 *
 * 座標系: 定義ファイルの値は「左上マーカー中心が原点のポイント(pt)」。
 * 4隅マーカーの実測位置に (x/W, y/H) を双一次補間して画像座標へ写す。
 */
(function (global) {
  "use strict";

  // ---------- 定義ファイル ----------
  function parseDefinition(text) {
    const def = { students: [], labels: [], idMods: [] };
    const lines = text.replace(/^﻿/, "").split(/\r?\n/);
    for (const line of lines) {
      if (!line.trim()) continue;
      const f = line.split(",");
      switch (f[0]) {
        case "PTMOMR":
          def.version = parseInt(f[1], 10);
          break;
        case "size":
          def.W = parseFloat(f[1]);
          def.H = parseFloat(f[2]);
          break;
        case "days":
          def.days = parseInt(f[1], 10);
          def.slots = parseInt(f[3], 10);
          break;
        case "grid":
          def.grid = {
            x0: parseFloat(f[1]), y0: parseFloat(f[2]),
            pitchX: parseFloat(f[3]), pitchY: parseFloat(f[4]),
            cellW: parseFloat(f[5]), cellH: parseFloat(f[6]),
          };
          break;
        case "idstrip": {
          def.idX = parseFloat(f[1]);
          def.idW = parseFloat(f[2]);
          for (let i = 3; i + 1 < f.length; i += 2) {
            def.idMods.push({ y: parseFloat(f[i]), h: parseFloat(f[i + 1]) });
          }
          break;
        }
        case "labels":
          def.labels = f.slice(1);
          break;
        case "student":
          def.students.push({ num: parseInt(f[1], 10), name: f.slice(2).join(",") });
          break;
      }
    }
    if (!def.W || !def.grid || !def.days || !def.slots) {
      throw new Error("定義ファイルの形式が正しくありません");
    }
    return def;
  }

  // ---------- 基本画像処理 ----------
  function toGray(img) {
    const { width: w, height: h, data } = img;
    const g = new Uint8Array(w * h);
    for (let i = 0, p = 0; i < g.length; i++, p += 4) {
      g[i] = (data[p] * 299 + data[p + 1] * 587 + data[p + 2] * 114) / 1000;
    }
    return { w, h, g };
  }

  function otsu(gray) {
    const hist = new Array(256).fill(0);
    const { g } = gray;
    for (let i = 0; i < g.length; i++) hist[g[i]]++;
    const total = g.length;
    let sum = 0;
    for (let i = 0; i < 256; i++) sum += i * hist[i];
    let sumB = 0, wB = 0, best = 0, thr = 127;
    for (let t = 0; t < 256; t++) {
      wB += hist[t];
      if (!wB) continue;
      const wF = total - wB;
      if (!wF) break;
      sumB += t * hist[t];
      const mB = sumB / wB, mF = (sum - sumB) / wF;
      const between = wB * wF * (mB - mF) * (mB - mF);
      if (between > best) { best = between; thr = t; }
    }
    return thr;
  }

  // ---------- マーカー検出 ----------
  // 窓内の連結成分から「マーカーらしい」塊を探す
  //   scoreMode "corner": 窓の隅に近い + 大きいものを優先
  //   scoreMode "predict": 予測位置 (towardX,towardY) に最も近いものを優先
  function findMarkerIn(gray, thr, x0, y0, x1, y1, expArea, towardX, towardY, scoreMode) {
    const { w, g } = gray;
    x0 = Math.max(0, x0 | 0); y0 = Math.max(0, y0 | 0);
    x1 = Math.min(gray.w, x1 | 0); y1 = Math.min(gray.h, y1 | 0);
    const ww = x1 - x0, wh = y1 - y0;
    if (ww <= 2 || wh <= 2) return null;
    const seen = new Uint8Array(ww * wh);
    let best = null;

    for (let yy = 0; yy < wh; yy++) {
      for (let xx = 0; xx < ww; xx++) {
        const li = yy * ww + xx;
        if (seen[li]) continue;
        if (g[(y0 + yy) * w + (x0 + xx)] >= thr) { seen[li] = 1; continue; }
        // flood fill
        let area = 0, sx = 0, sy = 0;
        let minX = xx, maxX = xx, minY = yy, maxY = yy;
        const stack = [li];
        seen[li] = 1;
        while (stack.length) {
          const cur = stack.pop();
          const cy = (cur / ww) | 0, cx = cur % ww;
          area++; sx += cx; sy += cy;
          if (cx < minX) minX = cx;
          if (cx > maxX) maxX = cx;
          if (cy < minY) minY = cy;
          if (cy > maxY) maxY = cy;
          const nb = [];
          if (cx > 0) nb.push(cur - 1);
          if (cx < ww - 1) nb.push(cur + 1);
          if (cy > 0) nb.push(cur - ww);
          if (cy < wh - 1) nb.push(cur + ww);
          for (const n of nb) {
            if (!seen[n]) {
              const ny = (n / ww) | 0, nx = n % ww;
              if (g[(y0 + ny) * w + (x0 + nx)] < thr) {
                seen[n] = 1;
                stack.push(n);
              } else {
                seen[n] = 1;
              }
            }
          }
        }
        if (area < expArea * 0.2 || area > expArea * 5) continue;
        const bw = maxX - minX + 1, bh = maxY - minY + 1;
        const aspect = bw / bh;
        if (aspect < 0.4 || aspect > 2.5) continue;
        const fill = area / (bw * bh);
        if (fill < 0.5) continue;
        const cx = x0 + sx / area, cy = y0 + sy / area;
        const dx = cx - towardX, dy = cy - towardY;
        const dist = dx * dx + dy * dy;
        const score = scoreMode === "predict" ? -dist : -dist + area;
        if (!best || score > best.score) best = { x: cx, y: cy, area, score };
      }
    }
    return best;
  }

  function findMarkers(gray, thr, def) {
    const { w, h } = gray;
    const sc = w / 595.0;                 // A4幅(595pt)基準のスケール推定
    const expArea = (18 * sc) * (16 * sc);

    // 1. 上の2マーカーを画像の上部帯から探す
    //    （フォームの高さはコマ数で変わるので、下は後で「予測」して探す）
    const wx = Math.round(w * 0.32), wy = Math.round(h * 0.35);
    const tl = findMarkerIn(gray, thr, 0, 0, wx, wy, expArea, 0, 0, "corner");
    const tr = findMarkerIn(gray, thr, w - wx, 0, w, wy, expArea, w, 0, "corner");
    if (global.OMR_DEBUG) console.log("top:", [tl, tr].map((m) => (m ? `${Math.round(m.x)},${Math.round(m.y)} a=${m.area}` : "MISS")));
    if (!tl || !tr) return null;

    // 2. 上辺ベクトルから下マーカーの位置を予測して、その近傍だけを探す
    const ux = (tr.x - tl.x) / def.W, uy = (tr.y - tl.y) / def.W;   // px / pt
    const vx = -uy, vy = ux;                                        // 紙面の下方向
    const predBl = { x: tl.x + vx * def.H, y: tl.y + vy * def.H };
    const predBr = { x: tr.x + vx * def.H, y: tr.y + vy * def.H };
    const win = Math.max(50, Math.round(0.05 * h));
    const bl = findMarkerIn(gray, thr, predBl.x - win, predBl.y - win,
      predBl.x + win, predBl.y + win, expArea, predBl.x, predBl.y, "predict");
    const br = findMarkerIn(gray, thr, predBr.x - win, predBr.y - win,
      predBr.x + win, predBr.y + win, expArea, predBr.x, predBr.y, "predict");
    if (global.OMR_DEBUG) console.log("bottom:", [bl, br].map((m) => (m ? `${Math.round(m.x)},${Math.round(m.y)} a=${m.area}` : "MISS")), "pred:", Math.round(predBl.x), Math.round(predBl.y));
    if (!bl && !br) return null;

    const res = {
      tl, tr,
      bl: bl || { x: tl.x + (br.x - tr.x), y: tl.y + (br.y - tr.y) },
      br: br || { x: tr.x + (bl.x - tl.x), y: tr.y + (bl.y - tl.y) },
    };

    // 妥当性: 縦横比が定義と合うか
    const dW = Math.hypot(res.tr.x - res.tl.x, res.tr.y - res.tl.y);
    const dH = Math.hypot(res.bl.x - res.tl.x, res.bl.y - res.tl.y);
    const ratio = (dW / dH) / (def.W / def.H);
    if (ratio < 0.9 || ratio > 1.1) return null;
    return res;
  }

  // 双一次写像: 定義pt座標 → 画像px
  function makeMapper(mk, def) {
    const { tl, tr, bl, br } = mk;
    return function (fx, fy) {
      const u = fx / def.W, v = fy / def.H;
      const x = (1 - u) * (1 - v) * tl.x + u * (1 - v) * tr.x + (1 - u) * v * bl.x + u * v * br.x;
      const y = (1 - u) * (1 - v) * tl.y + u * (1 - v) * tr.y + (1 - u) * v * bl.y + u * v * br.y;
      return [x, y];
    };
  }

  // 矩形の平均グレー値（中心部 frac のみ）
  function sampleMean(gray, mapper, fx, fy, fw, fh, frac) {
    const { w, h, g } = gray;
    const [cx, cy] = mapper(fx, fy);
    // 1ptあたりの画素数（局所スケール）
    const [x2] = mapper(fx + 1, fy);
    const [, y2] = mapper(fx, fy + 1);
    const sx = Math.abs(x2 - cx), sy = Math.abs(y2 - cy);
    const rw = Math.max(2, fw * frac * sx / 2), rh = Math.max(2, fh * frac * sy / 2);
    let sum = 0, n = 0;
    const px0 = Math.max(0, Math.round(cx - rw)), px1 = Math.min(w - 1, Math.round(cx + rw));
    const py0 = Math.max(0, Math.round(cy - rh)), py1 = Math.min(h - 1, Math.round(cy + rh));
    for (let y = py0; y <= py1; y++) {
      for (let x = px0; x <= px1; x++) {
        sum += g[y * w + x];
        n++;
      }
    }
    return n ? sum / n : 255;
  }

  // ---------- ページ読み取り ----------
  function decodePage(img, def) {
    const gray = toGray(img);
    const thr = otsu(gray);
    const mk = findMarkers(gray, thr, def);
    if (!mk) return { ok: false, reason: "マーカー（四隅の黒四角）を検出できませんでした" };
    const mapper = makeMapper(mk, def);

    // ID帯（8モジュール: b6..b0 + 奇数パリティ）
    let bits = [];
    for (const mod of def.idMods) {
      const mean = sampleMean(gray, mapper, def.idX, mod.y, def.idW, mod.h, 0.5);
      bits.push(mean < thr ? 1 : 0);
    }
    let id = 0, ones = 0;
    for (let i = 0; i < 7; i++) { id = (id << 1) | bits[i]; ones += bits[i]; }
    ones += bits[7];
    const idOk = bits.length === 8 && (ones % 2) === 1 && id >= 1 && id <= 99;

    // セル読み取り（グレー平均 vs 紙の白レベル）
    const cells = [];
    const means = [];
    for (let d = 0; d < def.days; d++) {
      for (let k = 0; k < def.slots; k++) {
        const fx = def.grid.x0 + d * def.grid.pitchX;
        const fy = def.grid.y0 + k * def.grid.pitchY;
        const mean = sampleMean(gray, mapper, fx, fy, def.grid.cellW, def.grid.cellH, 0.55);
        means.push(mean);
        cells.push({ d, k, mean });
      }
    }
    // 紙の地レベル: 平均値の中央値（大半のセルは空欄想定）
    const sorted = means.slice().sort((a, b) => a - b);
    const paper = sorted[Math.floor(sorted.length * 0.5)];
    const marks = [];
    for (const c of cells) {
      const ratio = c.mean / Math.max(1, paper);
      let state = 0;                       // 0=空欄
      if (ratio < 0.82) state = 1;         // 1=マーク
      else if (ratio < 0.92) state = 2;    // 2=要確認
      marks.push({ d: c.d, k: c.k, state, ratio: Math.round(ratio * 100) / 100 });
    }

    return { ok: true, markers: mk, id, idOk, bits, marks, paper, thr };
  }

  // ---------- CSV出力 ----------
  function buildCsv(def, pages) {
    // pages: [{id, marks(state==1のみ採用)}] → マトリクス形式
    const head = ["番号", "氏名"].concat(def.labels);
    const lines = [head.join(",")];
    for (const pg of pages) {
      const st = def.students.find((s) => s.num === pg.id);
      const row = [String(pg.id), st ? st.name : ""];
      const byIdx = new Array(def.labels.length).fill("");
      for (const m of pg.marks) {
        if (m.state === 1) byIdx[m.d * def.slots + m.k] = "○";
      }
      lines.push(row.concat(byIdx).join(","));
    }
    return lines.join("\r\n") + "\r\n";
  }

  const api = { parseDefinition, toGray, otsu, findMarkers, makeMapper, decodePage, buildCsv };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else global.OMR = api;
})(typeof window !== "undefined" ? window : globalThis);
