/* 保護者会 調査票OMR — UI（omr.js のコアを使う） */
(function () {
  "use strict";

  if (window.pdfjsLib) {
    pdfjsLib.GlobalWorkerOptions.workerSrc = "lib/pdf.worker.min.js";
  }

  const state = { def: null, pages: [] };

  const $ = (id) => document.getElementById(id);
  const defStatus = $("defStatus");
  const scanStatus = $("scanStatus");
  const pagesDiv = $("pages");
  const exportBtn = $("exportBtn");

  // ---------- 1. 定義ファイル ----------
  $("defFile").addEventListener("change", (ev) => {
    const f = ev.target.files[0];
    if (!f) return;
    const fr = new FileReader();
    fr.onload = () => {
      try {
        state.def = OMR.parseDefinition(fr.result);
        defStatus.textContent = `OK: ${state.def.days}日 × ${state.def.slots}コマ / 名簿 ${state.def.students.length} 名`;
        defStatus.className = "status ok";
        $("scanFiles").disabled = false;
      } catch (e) {
        defStatus.textContent = "読み込み失敗: " + e.message;
        defStatus.className = "status bad";
      }
    };
    fr.readAsText(f, "utf-8");
  });

  // ---------- 2. スキャン ----------
  $("scanFiles").addEventListener("change", async (ev) => {
    const files = [...ev.target.files];
    if (!files.length || !state.def) return;
    scanStatus.textContent = "処理中…";
    scanStatus.className = "status";
    let nPages = 0;
    try {
      for (const f of files) {
        if (/\.pdf$/i.test(f.name)) {
          nPages += await loadPdf(f);
        } else {
          nPages += await loadImage(f);
        }
      }
      scanStatus.textContent = `${nPages} ページを読み取りました`;
      scanStatus.className = "status ok";
    } catch (e) {
      console.error(e);
      scanStatus.textContent = "エラー: " + e.message;
      scanStatus.className = "status bad";
    }
    render();
    ev.target.value = "";
  });

  async function loadPdf(file) {
    if (!window.pdfjsLib) throw new Error("pdf.js が読み込めません（lib/ フォルダごと配置してください）");
    const buf = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
    for (let p = 1; p <= pdf.numPages; p++) {
      const page = await pdf.getPage(p);
      let vp = page.getViewport({ scale: 1 });
      const scale = 1400 / vp.width;
      vp = page.getViewport({ scale });
      const cv = document.createElement("canvas");
      cv.width = Math.round(vp.width);
      cv.height = Math.round(vp.height);
      await page.render({ canvasContext: cv.getContext("2d"), viewport: vp }).promise;
      addPage(cv, `${file.name} p.${p}`);
    }
    return pdf.numPages;
  }

  function loadImage(file) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      const url = URL.createObjectURL(file);
      img.onload = () => {
        const scale = Math.min(1, 1600 / img.width);
        const cv = document.createElement("canvas");
        cv.width = Math.round(img.width * scale);
        cv.height = Math.round(img.height * scale);
        cv.getContext("2d").drawImage(img, 0, 0, cv.width, cv.height);
        URL.revokeObjectURL(url);
        addPage(cv, file.name);
        resolve(1);
      };
      img.onerror = () => reject(new Error("画像を開けません: " + file.name));
      img.src = url;
    });
  }

  function addPage(canvas, label) {
    const ctx = canvas.getContext("2d");
    const imgData = ctx.getImageData(0, 0, canvas.width, canvas.height);
    const res = OMR.decodePage(imgData, state.def);
    state.pages.push({
      label,
      canvas,
      res,
      id: res.ok && res.idOk ? res.id : 0,
      marks: res.ok ? res.marks.map((m) => ({ ...m })) : [],
      excluded: !res.ok,
    });
  }

  // ---------- 3. 表示・修正 ----------
  function render() {
    pagesDiv.innerHTML = "";
    const counts = {};
    for (const pg of state.pages) {
      if (!pg.excluded && pg.id) counts[pg.id] = (counts[pg.id] || 0) + 1;
    }

    state.pages.forEach((pg, idx) => {
      const div = document.createElement("div");
      div.className = "page";

      const cv = document.createElement("canvas");
      drawOverlay(cv, pg);
      cv.addEventListener("click", (ev) => onCanvasClick(ev, cv, pg));
      div.appendChild(cv);

      const info = document.createElement("div");
      info.className = "info";
      const st = state.def.students.find((s) => s.num === pg.id);
      let tag;
      if (!pg.res.ok) tag = '<span class="tag bad">読取失敗</span>';
      else if (!pg.res.idOk) tag = '<span class="tag warn">番号要確認</span>';
      else if (counts[pg.id] > 1) tag = '<span class="tag bad">番号重複</span>';
      else tag = '<span class="tag ok">OK</span>';

      const nAmb = pg.marks.filter((m) => m.state === 2).length;
      const ambTag = nAmb ? `<span class="tag warn">要確認 ${nAmb}</span>` : "";

      info.innerHTML = `<h3>${escapeHtml(pg.label)} ${tag}${ambTag}</h3>`;

      if (pg.res.ok) {
        const sel = document.createElement("select");
        sel.innerHTML =
          '<option value="0">— 番号を選択 —</option>' +
          state.def.students
            .map((s) => `<option value="${s.num}" ${s.num === pg.id ? "selected" : ""}>${s.num}番 ${escapeHtml(s.name)}</option>`)
            .join("");
        sel.addEventListener("change", () => { pg.id = parseInt(sel.value, 10); render(); });
        const selWrap = document.createElement("div");
        selWrap.appendChild(sel);
        if (counts[pg.id] > 1) {
          const d = document.createElement("span");
          d.className = "dup";
          d.textContent = " ← 同じ番号のページが複数あります";
          selWrap.appendChild(d);
        }
        info.appendChild(selWrap);

        const ml = document.createElement("div");
        ml.className = "marklist";
        ml.innerHTML = pg.marks
          .filter((m) => m.state >= 1)
          .map((m) => `<span class="chip ${m.state === 2 ? "warn" : ""}">${escapeHtml(state.def.labels[m.d * state.def.slots + m.k] || `${m.d + 1}-${m.k + 1}`)}</span>`)
          .join("") || "<i>マークなし</i>";
        info.appendChild(ml);
      } else {
        info.innerHTML += `<p>${escapeHtml(pg.res.reason || "")}<br>このページは書き出しから除外されます。スキャンし直すか、Excelのクイック入力をご利用ください。</p>`;
      }

      const excl = document.createElement("label");
      excl.className = "excl";
      excl.innerHTML = `<input type="checkbox" ${pg.excluded ? "checked" : ""}> このページを除外する`;
      excl.querySelector("input").addEventListener("change", (e) => {
        pg.excluded = e.target.checked;
        render();
      });
      info.appendChild(excl);

      div.appendChild(info);
      pagesDiv.appendChild(div);
    });

    const valid = state.pages.filter((p) => !p.excluded && p.id >= 1);
    exportBtn.disabled = valid.length === 0;
    $("summary").innerHTML = state.pages.length
      ? `読み取り済み: ${state.pages.length} ページ / 書き出し対象: <b>${valid.length} 名</b>`
      : "";
  }

  function drawOverlay(cv, pg) {
    const W = 420;
    const sc = W / pg.canvas.width;
    cv.width = W;
    cv.height = Math.round(pg.canvas.height * sc);
    const ctx = cv.getContext("2d");
    ctx.drawImage(pg.canvas, 0, 0, cv.width, cv.height);
    if (!pg.res.ok) {
      ctx.strokeStyle = "#d04545";
      ctx.lineWidth = 4;
      ctx.strokeRect(2, 2, cv.width - 4, cv.height - 4);
      return;
    }
    const mapper = OMR.makeMapper(pg.res.markers, state.def);
    // マーカー
    ctx.strokeStyle = "#2563b0";
    ctx.lineWidth = 2;
    for (const key of ["tl", "tr", "bl", "br"]) {
      const m = pg.res.markers[key];
      ctx.strokeRect(m.x * sc - 8, m.y * sc - 8, 16, 16);
    }
    // セル
    const g = state.def.grid;
    for (const m of pg.marks) {
      const [x, y] = mapper(g.x0 + m.d * g.pitchX, g.y0 + m.k * g.pitchY);
      const [x2] = mapper(g.x0 + m.d * g.pitchX + g.cellW / 2, g.y0 + m.k * g.pitchY);
      const rw = Math.abs(x2 - x) * sc;
      const [, y2] = mapper(g.x0 + m.d * g.pitchX, g.y0 + m.k * g.pitchY + g.cellH / 2);
      const rh = Math.abs(y2 - y) * sc;
      if (m.state === 1) {
        ctx.strokeStyle = "#2e9e44";
        ctx.lineWidth = 2.5;
      } else if (m.state === 2) {
        ctx.strokeStyle = "#e6a700";
        ctx.lineWidth = 2.5;
      } else {
        ctx.strokeStyle = "rgba(120,140,180,0.35)";
        ctx.lineWidth = 1;
      }
      ctx.strokeRect(x * sc - rw, y * sc - rh, rw * 2, rh * 2);
    }
  }

  function onCanvasClick(ev, cv, pg) {
    if (!pg.res.ok) return;
    const rect = cv.getBoundingClientRect();
    const px = ((ev.clientX - rect.left) / rect.width) * cv.width;
    const py = ((ev.clientY - rect.top) / rect.height) * cv.height;
    const sc = cv.width / pg.canvas.width;
    const mapper = OMR.makeMapper(pg.res.markers, state.def);
    const g = state.def.grid;
    let bestM = null, bestD = Infinity;
    for (const m of pg.marks) {
      const [x, y] = mapper(g.x0 + m.d * g.pitchX, g.y0 + m.k * g.pitchY);
      const d = (x * sc - px) ** 2 + (y * sc - py) ** 2;
      if (d < bestD) { bestD = d; bestM = m; }
    }
    if (!bestM) return;
    // クリックで 要確認→○ / ○→なし / なし→○
    if (bestM.state === 2) bestM.state = 1;
    else if (bestM.state === 1) bestM.state = 0;
    else bestM.state = 1;
    render();
  }

  // ---------- 4. CSV ----------
  exportBtn.addEventListener("click", () => {
    const valid = state.pages.filter((p) => !p.excluded && p.id >= 1);
    // 同じ番号が複数 → 最後のページを採用（後から読んだ方が新しい想定）
    const byId = new Map();
    for (const p of valid) byId.set(p.id, p);
    const csv = OMR.buildCsv(state.def, [...byId.values()].map((p) => ({ id: p.id, marks: p.marks })));
    const blob = new Blob(["﻿" + csv], { type: "text/csv" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "希望取込.csv";
    a.click();
    URL.revokeObjectURL(a.href);
  });

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }
})();
