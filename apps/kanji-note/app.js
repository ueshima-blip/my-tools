/* app.js — かんじノート（iPad ノートアプリ／漢字練習）
 *
 * 機能概要:
 *  - Apple Pencil 対応（Pointer Events: pen / touch / mouse、筆圧で線幅変化）
 *  - パームリジェクション（"Pencilのみ" トグルで touch を無視）
 *  - ペン／マーカー／けしゴム、6色、太さ、Undo/Redo、複数ページ、保存/読込
 *  - 漢字練習モード: お題を入れるとマス目が並び、書いたあと自動で
 *    まちがった漢字を見つけ、ただしい字を表示（recognizer.js）
 */

'use strict';

// =================================================================
// 状態
// =================================================================

const State = {
  mode: 'note',          // 'note' | 'practice'
  tool: 'pen',           // 'pen' | 'highlighter' | 'eraser'
  color: '#1a1a1a',
  thickness: 3,
  penOnly: false,

  note: {
    pages: [[]],         // 各ページ = ストローク配列
    pageIndex: 0,
    undo: [],
    redo: [],
  },

  practice: {
    target: '学校',
    cells: [],           // [{char, bbox, strokes, result}]
    loose: [],           // マスの外に書いたストローク
    autoJudge: true,
    showGhost: true,
    undo: [],
    redo: [],
  },
};

// =================================================================
// DOM / Canvas セットアップ
// =================================================================

const $ = (id) => document.getElementById(id);
const canvas = $('board');
const ctx = canvas.getContext('2d');
let DPR = 1;

function cssSize() {
  const r = canvas.getBoundingClientRect();
  return { w: r.width, h: r.height };
}

function fitCanvas() {
  DPR = window.devicePixelRatio || 1;
  const { w, h } = cssSize();
  canvas.width = Math.max(1, Math.round(w * DPR));
  canvas.height = Math.max(1, Math.round(h * DPR));
  ctx.setTransform(DPR, 0, 0, DPR, 0, 0);
  ctx.lineCap = 'round';
  ctx.lineJoin = 'round';
  if (State.mode === 'practice') layoutCells();
  redraw();
}

window.addEventListener('resize', fitCanvas);
window.addEventListener('orientationchange', () => setTimeout(fitCanvas, 80));

// =================================================================
// ストローク描画
// =================================================================

function drawStroke(stroke) {
  const pts = stroke.points;
  if (!pts || pts.length === 0) return;
  ctx.save();
  ctx.lineCap = 'round';
  ctx.lineJoin = 'round';
  ctx.strokeStyle = stroke.color;
  if (stroke.tool === 'highlighter') {
    ctx.globalAlpha = 0.32;
    ctx.lineWidth = stroke.thickness * 4;
    ctx.beginPath();
    ctx.moveTo(pts[0].x, pts[0].y);
    for (let i = 1; i < pts.length; i++) ctx.lineTo(pts[i].x, pts[i].y);
    ctx.stroke();
  } else {
    if (pts.length === 1) {
      ctx.fillStyle = stroke.color;
      ctx.beginPath();
      ctx.arc(pts[0].x, pts[0].y, Math.max(1, stroke.thickness * 0.6), 0, Math.PI * 2);
      ctx.fill();
    } else {
      // 筆圧で太さを変えるため、セグメントごとに lineWidth を変えて描く
      for (let i = 1; i < pts.length; i++) {
        const a = pts[i - 1], b = pts[i];
        const w = stroke.thickness * (0.55 + 0.9 * ((a.p + b.p) / 2));
        ctx.lineWidth = w;
        ctx.beginPath();
        ctx.moveTo(a.x, a.y);
        ctx.lineTo(b.x, b.y);
        ctx.stroke();
      }
    }
  }
  ctx.restore();
}

function drawSegmentImmediate(stroke) {
  // pointermove 中の最後のセグメントだけ描く（低レイテンシ用）
  const pts = stroke.points;
  if (pts.length < 1) return;
  ctx.save();
  ctx.lineCap = 'round';
  ctx.lineJoin = 'round';
  ctx.strokeStyle = stroke.color;
  if (stroke.tool === 'highlighter') {
    ctx.globalAlpha = 0.32;
    ctx.lineWidth = stroke.thickness * 4;
  }
  if (pts.length === 1) {
    ctx.fillStyle = stroke.color;
    ctx.beginPath();
    ctx.arc(pts[0].x, pts[0].y, Math.max(1, stroke.thickness * 0.6), 0, Math.PI * 2);
    ctx.fill();
  } else {
    const a = pts[pts.length - 2], b = pts[pts.length - 1];
    if (stroke.tool !== 'highlighter') {
      ctx.lineWidth = stroke.thickness * (0.55 + 0.9 * ((a.p + b.p) / 2));
    }
    ctx.beginPath();
    ctx.moveTo(a.x, a.y);
    ctx.lineTo(b.x, b.y);
    ctx.stroke();
  }
  ctx.restore();
}

// =================================================================
// 背景・マス・結果オーバーレイ
// =================================================================

function drawBackground() {
  const { w, h } = cssSize();
  ctx.save();
  ctx.fillStyle = '#fefdf7';
  ctx.fillRect(0, 0, w, h);
  if (State.mode === 'note') {
    // 罫線
    ctx.strokeStyle = '#dde6f3';
    ctx.lineWidth = 1;
    for (let y = 56; y < h; y += 36) {
      ctx.beginPath(); ctx.moveTo(0, y); ctx.lineTo(w, y); ctx.stroke();
    }
    // 左マージン
    ctx.strokeStyle = '#f3c8c8';
    ctx.beginPath(); ctx.moveTo(56, 0); ctx.lineTo(56, h); ctx.stroke();
  }
  ctx.restore();
}

function drawCells() {
  for (const cell of State.practice.cells) {
    const { x, y, w: cw, h: ch } = cell.bbox;
    ctx.save();
    ctx.fillStyle = '#fffdf3';
    ctx.fillRect(x, y, cw, ch);
    ctx.strokeStyle = '#cfc6a8';
    ctx.lineWidth = 2;
    ctx.strokeRect(x, y, cw, ch);
    // 十字ガイド
    ctx.setLineDash([4, 6]);
    ctx.strokeStyle = '#ddd29a';
    ctx.beginPath();
    ctx.moveTo(x + cw / 2, y); ctx.lineTo(x + cw / 2, y + ch);
    ctx.moveTo(x, y + ch / 2); ctx.lineTo(x + cw, y + ch / 2);
    ctx.stroke();
    ctx.setLineDash([]);
    // お手本（うすく）
    if (State.practice.showGhost && cell.char) {
      ctx.fillStyle = 'rgba(60,40,20,0.10)';
      const fs = Math.floor(cw * 0.78);
      ctx.font = `${fs}px "Hiragino Mincho ProN","YuMincho","Yu Mincho","Noto Serif JP","Hiragino Sans",serif`;
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.fillText(cell.char, x + cw / 2, y + ch / 2 + cw * 0.04);
    }
    ctx.restore();
  }
}

function drawCellResults() {
  for (const cell of State.practice.cells) {
    if (!cell.result) continue;
    const { x, y, w: cw, h: ch } = cell.bbox;
    const r = cell.result;
    ctx.save();
    if (r.kind === 'ok') {
      ctx.strokeStyle = '#43a047';
      ctx.lineWidth = 4;
      ctx.strokeRect(x - 3, y - 3, cw + 6, ch + 6);
      ctx.fillStyle = '#43a047';
      ctx.font = 'bold 28px sans-serif';
      ctx.textAlign = 'right'; ctx.textBaseline = 'top';
      ctx.fillText('○', x + cw - 6, y + 4);
    } else if (r.kind === 'ng') {
      ctx.strokeStyle = '#e53935';
      ctx.lineWidth = 4;
      ctx.strokeRect(x - 3, y - 3, cw + 6, ch + 6);
      ctx.fillStyle = '#e53935';
      ctx.font = 'bold 28px sans-serif';
      ctx.textAlign = 'right'; ctx.textBaseline = 'top';
      ctx.fillText('×', x + cw - 6, y + 4);
      // 「正しくは：X」バッジをマスの下に
      ctx.fillStyle = '#e53935';
      ctx.fillRect(x, y + ch + 4, cw, 28);
      ctx.fillStyle = '#fff';
      ctx.font = 'bold 16px sans-serif';
      ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
      ctx.fillText(`正しくは「${cell.char}」`, x + cw / 2, y + ch + 4 + 14);
    } else if (r.kind === 'unsure') {
      ctx.strokeStyle = '#fb8c00';
      ctx.lineWidth = 3;
      ctx.setLineDash([6, 4]);
      ctx.strokeRect(x - 3, y - 3, cw + 6, ch + 6);
      ctx.setLineDash([]);
    }
    ctx.restore();
  }
}

function redraw() {
  const { w, h } = cssSize();
  ctx.clearRect(0, 0, w, h);
  drawBackground();
  if (State.mode === 'note') {
    for (const s of State.note.pages[State.note.pageIndex]) drawStroke(s);
  } else {
    drawCells();
    for (const cell of State.practice.cells) for (const s of cell.strokes) drawStroke(s);
    for (const s of State.practice.loose) drawStroke(s);
    drawCellResults();
  }
}

// =================================================================
// 練習マスのレイアウト
// =================================================================

function layoutCells() {
  const { w, h } = cssSize();
  const target = State.practice.target;
  const chars = [...target];
  const n = chars.length;
  if (n === 0) { State.practice.cells = []; return; }

  const padX = 24, padY = 56, gap = 16;
  let perRow = n;
  let cellSize = Math.min(220, Math.floor((w - padX * 2 - gap * (n - 1)) / Math.max(1, n)));
  if (cellSize < 100) {
    perRow = Math.max(1, Math.floor((w - padX * 2 + gap) / (140 + gap)));
    cellSize = Math.floor((w - padX * 2 - gap * (perRow - 1)) / Math.max(1, perRow));
  }
  cellSize = Math.max(80, Math.min(220, cellSize));

  const rows = Math.ceil(n / perRow);
  const totalH = rows * cellSize + (rows - 1) * gap + 36; // 余白（下のラベル用）
  const startY = Math.max(padY, Math.floor((h - totalH) / 2));

  const oldCells = State.practice.cells || [];
  const cells = [];
  for (let i = 0; i < n; i++) {
    const row = Math.floor(i / perRow);
    const col = i % perRow;
    const inRowCount = (row === rows - 1) ? (n - row * perRow) : perRow;
    const rowWidth = inRowCount * cellSize + (inRowCount - 1) * gap;
    const startX = Math.floor((w - rowWidth) / 2);
    const x = startX + col * (cellSize + gap);
    const y = startY + row * (cellSize + gap);
    const newC = {
      char: chars[i],
      bbox: { x, y, w: cellSize, h: cellSize },
      strokes: [],
      result: null,
    };
    // 同じ位置の旧セルが同じ文字なら、ストロークを新 bbox に変換して引き継ぐ
    const oldC = oldCells[i];
    if (oldC && oldC.char === chars[i] && oldC.strokes.length > 0) {
      const sx = newC.bbox.w / oldC.bbox.w;
      const sy = newC.bbox.h / oldC.bbox.h;
      newC.strokes = oldC.strokes.map(s => ({
        ...s,
        points: s.points.map(p => ({
          ...p,
          x: newC.bbox.x + (p.x - oldC.bbox.x) * sx,
          y: newC.bbox.y + (p.y - oldC.bbox.y) * sy,
        })),
      }));
      newC.result = oldC.result;
    }
    cells.push(newC);
  }
  State.practice.cells = cells;
}

function cellAt(p) {
  for (let i = 0; i < State.practice.cells.length; i++) {
    const { x, y, w, h } = State.practice.cells[i].bbox;
    if (p.x >= x && p.x <= x + w && p.y >= y && p.y <= y + h) return i;
  }
  return -1;
}

// =================================================================
// Undo / Redo / 追加 / 削除
// =================================================================

function activeUndo() { return State.mode === 'note' ? State.note.undo : State.practice.undo; }
function activeRedo() { return State.mode === 'note' ? State.note.redo : State.practice.redo; }

function pushUndoEntry(entry) {
  activeUndo().push(entry);
  if (State.mode === 'note') State.note.redo = [];
  else State.practice.redo = [];
}

function addStrokeToActive(stroke, cellIndex) {
  let container;
  if (State.mode === 'note') {
    container = State.note.pages[State.note.pageIndex];
  } else if (cellIndex != null && cellIndex >= 0) {
    container = State.practice.cells[cellIndex].strokes;
    State.practice.cells[cellIndex].result = null; // 再判定対象に
  } else {
    container = State.practice.loose;
  }
  container.push(stroke);
  pushUndoEntry({ type: 'add', container, stroke, cellIndex: cellIndex ?? null });
}

function strokeNearPoint(stroke, p, R) {
  const r2 = R * R;
  for (const sp of stroke.points) {
    const dx = sp.x - p.x, dy = sp.y - p.y;
    if (dx * dx + dy * dy < r2) return true;
  }
  return false;
}

function activeContainersForErase() {
  if (State.mode === 'note') return [{ container: State.note.pages[State.note.pageIndex], cellIndex: null }];
  const list = [];
  for (let i = 0; i < State.practice.cells.length; i++) {
    list.push({ container: State.practice.cells[i].strokes, cellIndex: i });
  }
  list.push({ container: State.practice.loose, cellIndex: null });
  return list;
}

let pendingEraseEntry = null;

function eraserBegin() { pendingEraseEntry = { type: 'erase', removed: [] }; }
function eraserCommit() {
  if (pendingEraseEntry && pendingEraseEntry.removed.length > 0) {
    activeUndo().push(pendingEraseEntry);
    if (State.mode === 'note') State.note.redo = [];
    else State.practice.redo = [];
  }
  pendingEraseEntry = null;
}

function eraseAt(p) {
  const R = 16;
  let removedAny = false;
  for (const { container, cellIndex } of activeContainersForErase()) {
    for (let i = container.length - 1; i >= 0; i--) {
      const s = container[i];
      if (strokeNearPoint(s, p, R)) {
        container.splice(i, 1);
        if (pendingEraseEntry) pendingEraseEntry.removed.push({ container, index: i, stroke: s, cellIndex });
        removedAny = true;
        if (State.mode === 'practice' && cellIndex != null) {
          State.practice.cells[cellIndex].result = null;
        }
      }
    }
  }
  if (removedAny) redraw();
}

function undo() {
  const u = activeUndo(), r = activeRedo();
  const e = u.pop();
  if (!e) return;
  if (e.type === 'add') {
    const idx = e.container.indexOf(e.stroke);
    if (idx >= 0) e.container.splice(idx, 1);
    if (e.cellIndex != null && State.mode === 'practice') State.practice.cells[e.cellIndex].result = null;
  } else if (e.type === 'erase') {
    // 元のインデックスに昇順で挿入して復元
    const sorted = [...e.removed].sort((a, b) => a.index - b.index);
    for (const item of sorted) item.container.splice(item.index, 0, item.stroke);
  } else if (e.type === 'clear-page') {
    State.note.pages[e.pageIndex] = e.snapshot.slice();
  } else if (e.type === 'practice-clear') {
    for (let i = 0; i < State.practice.cells.length; i++) {
      State.practice.cells[i].strokes = (e.cellSnapshots[i] || []).slice();
    }
    State.practice.loose = (e.looseSnapshot || []).slice();
  }
  r.push(e);
  if (State.mode === 'practice' && State.practice.autoJudge) judgeAll();
  else { redraw(); renderResultsPanel(); }
}

function redo() {
  const u = activeUndo(), r = activeRedo();
  const e = r.pop();
  if (!e) return;
  if (e.type === 'add') {
    e.container.push(e.stroke);
    if (e.cellIndex != null && State.mode === 'practice') State.practice.cells[e.cellIndex].result = null;
  } else if (e.type === 'erase') {
    for (const item of e.removed) {
      const idx = item.container.indexOf(item.stroke);
      if (idx >= 0) item.container.splice(idx, 1);
    }
  } else if (e.type === 'clear-page') {
    State.note.pages[e.pageIndex] = [];
  } else if (e.type === 'practice-clear') {
    for (const c of State.practice.cells) { c.strokes = []; c.result = null; }
    State.practice.loose = [];
  }
  u.push(e);
  if (State.mode === 'practice' && State.practice.autoJudge) judgeAll();
  else { redraw(); renderResultsPanel(); }
}

// =================================================================
// ポインタ入力（Apple Pencil / タッチ / マウス）
// =================================================================

let drawing = false;
let activePointerId = null;
let currentStroke = null;
let currentCellIndex = null;

function pointFromEvent(e) {
  const rect = canvas.getBoundingClientRect();
  return {
    x: e.clientX - rect.left,
    y: e.clientY - rect.top,
    p: e.pressure > 0 ? e.pressure : (e.pointerType === 'pen' ? 0.5 : 0.5),
    t: performance.now(),
  };
}

function shouldIgnore(e) {
  return State.penOnly && e.pointerType === 'touch';
}

canvas.addEventListener('pointerdown', (e) => {
  if (shouldIgnore(e)) return;
  if (drawing) return;
  e.preventDefault();
  drawing = true;
  activePointerId = e.pointerId;
  try { canvas.setPointerCapture(e.pointerId); } catch {}
  const p = pointFromEvent(e);
  currentStroke = {
    tool: State.tool,
    color: State.color,
    thickness: State.thickness,
    points: [p],
  };
  if (State.tool === 'eraser') {
    eraserBegin();
    eraseAt(p);
    return;
  }
  currentCellIndex = (State.mode === 'practice') ? cellAt(p) : null;
  drawSegmentImmediate(currentStroke);
});

canvas.addEventListener('pointermove', (e) => {
  if (!drawing || e.pointerId !== activePointerId) return;
  if (shouldIgnore(e)) return;
  e.preventDefault();
  // CoalescedEvents で滑らかに（非対応なら通常のイベントで）
  const evs = e.getCoalescedEvents ? e.getCoalescedEvents() : null;
  const points = (evs && evs.length > 0) ? evs.map(pointFromEvent) : [pointFromEvent(e)];
  for (const p of points) {
    currentStroke.points.push(p);
    if (State.tool === 'eraser') {
      eraseAt(p);
    } else {
      drawSegmentImmediate(currentStroke);
    }
  }
});

function endPointer(e) {
  if (!drawing) return;
  if (e && e.pointerId !== activePointerId) return;
  drawing = false;
  try { canvas.releasePointerCapture(activePointerId); } catch {}
  activePointerId = null;

  if (currentStroke && currentStroke.points.length > 0) {
    if (State.tool === 'eraser') {
      eraserCommit();
    } else {
      addStrokeToActive(currentStroke, currentCellIndex);
      if (State.mode === 'practice') {
        if (State.practice.autoJudge) scheduleAutoJudge();
        else renderResultsPanel();
      }
    }
  }
  currentStroke = null;
  currentCellIndex = null;
}
canvas.addEventListener('pointerup', endPointer);
canvas.addEventListener('pointercancel', endPointer);
canvas.addEventListener('pointerleave', (e) => { if (e.pointerId === activePointerId) endPointer(e); });

// 二本指ジェスチャ等によるスクロール抑止（保険）
canvas.addEventListener('touchstart', e => e.preventDefault(), { passive: false });
canvas.addEventListener('touchmove',  e => e.preventDefault(), { passive: false });

// =================================================================
// 漢字判定
// =================================================================

let autoJudgeTimer = null;
function scheduleAutoJudge() {
  clearTimeout(autoJudgeTimer);
  autoJudgeTimer = setTimeout(() => judgeAll(), 700);
}

function judgeCell(cell) {
  if (!cell.strokes || cell.strokes.length === 0) {
    cell.result = null;
    return null;
  }
  const result = KanjiRecognizer.recognize(cell.strokes, { target: cell.char });
  if (!result) { cell.result = null; return null; }
  const kind = KanjiRecognizer.classify(result, cell.char);
  cell.result = { kind, ...result };
  return cell.result;
}

function judgeAll() {
  for (const cell of State.practice.cells) judgeCell(cell);
  redraw();
  renderResultsPanel();
}

function renderResultsPanel() {
  const root = $('results');
  if (!root) return;
  root.innerHTML = '';
  for (const cell of State.practice.cells) {
    const card = document.createElement('div');
    card.className = 'result-card';
    const badge = document.createElement('div'); badge.className = 'badge';
    const target = document.createElement('div'); target.className = 'target';
    const text = document.createElement('div'); text.className = 'text';
    target.textContent = cell.char;

    if (!cell.result) {
      badge.textContent = '…';
      card.classList.add('unsure');
      text.innerHTML = `まだ書かれていません`;
    } else if (cell.result.kind === 'ok') {
      badge.textContent = '○';
      card.classList.add('ok');
      text.innerHTML = `<b>せいかい！</b> よく書けたね。`;
    } else if (cell.result.kind === 'ng') {
      badge.textContent = '×';
      card.classList.add('ng');
      const wrong = cell.result.best;
      text.innerHTML = `これは <b>「${wrong}」</b> に近いよ。<br>正しくは <b style="color:#1b5e20">「${cell.char}」</b>`;
    } else {
      badge.textContent = '？';
      card.classList.add('unsure');
      text.innerHTML = `もう少していねいに書いてみよう。`;
    }
    card.append(badge, target, text);
    root.appendChild(card);
  }

  // くわしい結果（直近で判定されたセルの上位候補）
  const adv = $('advanced-results');
  if (adv) {
    adv.innerHTML = '';
    const last = [...State.practice.cells].reverse().find(c => c.result && c.result.top);
    if (last) {
      for (const [c, s] of last.result.top) {
        const row = document.createElement('div');
        row.className = 'ar-row';
        row.innerHTML = `<span class="c">${c}</span><span class="s">${(s * 100).toFixed(0)}</span>`;
        adv.appendChild(row);
      }
    }
  }
}

// =================================================================
// UI 配線
// =================================================================

document.querySelectorAll('.mode-btn').forEach(b => {
  b.addEventListener('click', () => setMode(b.dataset.mode));
});

function setMode(mode) {
  State.mode = mode;
  for (const b of document.querySelectorAll('.mode-btn')) b.classList.toggle('active', b.dataset.mode === mode);
  $('practice-panel').hidden = (mode !== 'practice');
  $('pages-group').style.display = (mode === 'note') ? '' : 'none';
  if (mode === 'practice') {
    layoutCells();
    renderResultsPanel();
  }
  redraw();
}

document.querySelectorAll('.tool-btn').forEach(b => {
  b.addEventListener('click', () => {
    State.tool = b.dataset.tool;
    for (const tb of document.querySelectorAll('.tool-btn')) tb.classList.toggle('active', tb === b);
    canvas.style.cursor = (State.tool === 'eraser') ? 'cell' : 'crosshair';
  });
});

document.querySelectorAll('.swatch').forEach(b => {
  b.addEventListener('click', () => {
    State.color = b.dataset.color;
    for (const sw of document.querySelectorAll('.swatch')) sw.classList.toggle('active', sw === b);
  });
});

$('thickness').addEventListener('input', e => {
  State.thickness = parseInt(e.target.value, 10) || 3;
});

$('pen-only').addEventListener('change', e => {
  State.penOnly = e.target.checked;
});

$('undo').addEventListener('click', undo);
$('redo').addEventListener('click', redo);

$('clear').addEventListener('click', () => {
  if (!confirm('このページを消します。よろしいですか？')) return;
  if (State.mode === 'note') {
    const snap = State.note.pages[State.note.pageIndex].slice();
    State.note.pages[State.note.pageIndex] = [];
    State.note.undo.push({ type: 'clear-page', pageIndex: State.note.pageIndex, snapshot: snap });
    State.note.redo = [];
  } else {
    const cellSnapshots = State.practice.cells.map(c => c.strokes.slice());
    const looseSnapshot = State.practice.loose.slice();
    for (const c of State.practice.cells) { c.strokes = []; c.result = null; }
    State.practice.loose = [];
    State.practice.undo.push({ type: 'practice-clear', cellSnapshots, looseSnapshot });
    State.practice.redo = [];
    renderResultsPanel();
  }
  redraw();
});

$('prev-page').addEventListener('click', () => {
  if (State.note.pageIndex > 0) { State.note.pageIndex--; updatePageIndicator(); redraw(); }
});
$('next-page').addEventListener('click', () => {
  if (State.note.pageIndex < State.note.pages.length - 1) { State.note.pageIndex++; updatePageIndicator(); redraw(); }
});
$('add-page').addEventListener('click', () => {
  State.note.pages.push([]);
  State.note.pageIndex = State.note.pages.length - 1;
  updatePageIndicator();
  redraw();
});

function updatePageIndicator() {
  $('page-indicator').textContent = `${State.note.pageIndex + 1} / ${State.note.pages.length}`;
}

$('save').addEventListener('click', () => {
  try {
    const payload = {
      v: 1,
      note: {
        pageIndex: State.note.pageIndex,
        pages: State.note.pages,
      },
      practice: { target: State.practice.target },
    };
    localStorage.setItem('kanji-note', JSON.stringify(payload));
    toast('保存しました');
  } catch (e) {
    toast('保存できません: ' + e.message);
  }
});

$('load').addEventListener('click', () => {
  try {
    const raw = localStorage.getItem('kanji-note');
    if (!raw) { toast('保存データがありません'); return; }
    const data = JSON.parse(raw);
    if (data.note && Array.isArray(data.note.pages)) {
      State.note.pages = data.note.pages.length ? data.note.pages : [[]];
      State.note.pageIndex = Math.min(data.note.pageIndex || 0, State.note.pages.length - 1);
      State.note.undo = []; State.note.redo = [];
    }
    if (data.practice && data.practice.target) {
      State.practice.target = data.practice.target;
      $('target-word').value = data.practice.target;
      State.practice.cells = []; // 既存セルを破棄して新規生成
      layoutCells();
    }
    updatePageIndicator();
    redraw();
    renderResultsPanel();
    toast('読み込みました');
  } catch (e) {
    toast('読み込み失敗: ' + e.message);
  }
});

// 漢字練習パネル
$('set-target').addEventListener('click', applyTarget);
$('target-word').addEventListener('keydown', e => { if (e.key === 'Enter') applyTarget(); });
$('check-all').addEventListener('click', () => judgeAll());
$('reset-cells').addEventListener('click', () => {
  const cellSnapshots = State.practice.cells.map(c => c.strokes.slice());
  const looseSnapshot = State.practice.loose.slice();
  for (const c of State.practice.cells) { c.strokes = []; c.result = null; }
  State.practice.loose = [];
  State.practice.undo.push({ type: 'practice-clear', cellSnapshots, looseSnapshot });
  State.practice.redo = [];
  redraw();
  renderResultsPanel();
});
$('auto-judge').addEventListener('change', e => {
  State.practice.autoJudge = e.target.checked;
  if (e.target.checked) judgeAll();
});
$('show-ghost').addEventListener('change', e => {
  State.practice.showGhost = e.target.checked;
  redraw();
});

function applyTarget() {
  const v = $('target-word').value.trim();
  if (!v) return;
  State.practice.target = v;
  State.practice.cells = []; // 強制再構築
  State.practice.loose = [];
  State.practice.undo = [];
  State.practice.redo = [];
  layoutCells();
  redraw();
  renderResultsPanel();
}

// トースト
let toastTimer;
function toast(msg) {
  const t = $('toast');
  t.textContent = msg;
  t.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { t.hidden = true; }, 1800);
}

// ショートカット（PCで開発しやすいように）
document.addEventListener('keydown', (e) => {
  if (e.target.matches('input,textarea')) return;
  const ctrlOrMeta = e.ctrlKey || e.metaKey;
  if (ctrlOrMeta && e.key === 'z' && !e.shiftKey) { e.preventDefault(); undo(); }
  else if (ctrlOrMeta && (e.key === 'y' || (e.key === 'z' && e.shiftKey))) { e.preventDefault(); redo(); }
  else if (!ctrlOrMeta) {
    if (e.key === '1') document.querySelector('.tool-btn[data-tool="pen"]').click();
    else if (e.key === '2') document.querySelector('.tool-btn[data-tool="highlighter"]').click();
    else if (e.key === '3') document.querySelector('.tool-btn[data-tool="eraser"]').click();
  }
});

// 初期化
fitCanvas();
updatePageIndicator();
