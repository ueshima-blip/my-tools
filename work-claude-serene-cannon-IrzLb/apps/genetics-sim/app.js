/* 遺伝のモデル実験 — 共有フォルダ同期版（File System Access API） */
'use strict';

const POLL_INTERVAL_MS = 1500;
const HEARTBEAT_MS = 5000;
const STALE_AFTER_MS = 30000;
const REQUEST_TTL_MS = 60000;
const TEACHER_PASSWORD = 'sensei';
const HISTORY_LIMIT = 200;

// ----- DOM ヘルパ -----
const $ = (id) => document.getElementById(id);
const els = {};
[
  'folderScreen', 'loginScreen', 'mainScreen',
  'chooseFolderBtn', 'folderError', 'browserNote',
  'folderPath', 'nameInput', 'teacherCheck',
  'passwordRow', 'passwordInput', 'loginBtn', 'loginError',
  'changeFolderBtn',
  'myBadge', 'myName', 'myCards', 'syncDot', 'logoutBtn',
  'bigCards', 'myPhenotype', 'myGenotype',
  'pairTarget', 'pairBtn', 'giftTarget', 'giftBtn',
  'roster', 'rosterCount',
  'cAA', 'cAa', 'caa', 'phenotypeRatio', 'genotypeRatio', 'history',
  'teacherPanel', 'initMode', 'customRatioRow',
  'ratioAA', 'ratioAa', 'ratioaa', 'initPreview',
  'resetBtn', 'cleanupBtn',
  'toast',
  'pairDialog', 'pairDialogText', 'pairAccept', 'pairDecline',
  'giftDialog', 'giftDialogText', 'giftAccept', 'giftDecline',
  'resultDialog', 'resultTitle', 'resultBody', 'resultClose',
].forEach((id) => (els[id] = $(id)));

// ----- 状態 -----
let rootHandle = null;     // 共有フォルダのハンドル
let dataHandle = null;     // data/ のハンドル
let me = null;             // { id, name, alleles, isTeacher, lastSeen }
let config = { initialGenotype: 'Aa', classCode: '3-1' };
let knownStudents = new Map();
let seenEventIds = new Set();
let allEvents = [];
let pendingPair = null;
let pendingGift = null;
let pollBusy = false;
let pollTimer = null;

// ----- ユーティリティ -----
function showToast(msg, ms = 2400) {
  els.toast.textContent = msg;
  els.toast.classList.remove('hidden');
  clearTimeout(showToast._t);
  showToast._t = setTimeout(() => els.toast.classList.add('hidden'), ms);
}

function escapeHTML(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
  }[c]));
}

function newId() {
  return Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 8);
}

function genotypeOf(alleles) {
  const sorted = [...alleles].sort((a, b) => {
    if (a === b) return 0;
    return a === 'A' ? -1 : 1;
  });
  return sorted.join('');
}

function phenotypeOf(genotype) {
  return genotype.includes('A') ? '丸' : 'しわ';
}

function defaultAlleles(g) {
  if (g === 'AA') return ['A', 'A'];
  if (g === 'aa') return ['a', 'a'];
  return ['A', 'a'];
}

function pickRandom(arr) {
  return arr[Math.floor(Math.random() * arr.length)];
}

// ----- File System Access ヘルパ -----
async function getOrCreateDir(parent, name) {
  return parent.getDirectoryHandle(name, { create: true });
}

async function writeJSON(dirHandle, fileName, obj) {
  const fh = await dirHandle.getFileHandle(fileName, { create: true });
  const w = await fh.createWritable();
  await w.write(JSON.stringify(obj));
  await w.close();
}

async function readJSON(dirHandle, fileName) {
  try {
    const fh = await dirHandle.getFileHandle(fileName);
    const file = await fh.getFile();
    const text = await file.text();
    return JSON.parse(text);
  } catch (e) {
    if (e.name === 'NotFoundError') return null;
    throw e;
  }
}

async function listFiles(dirHandle) {
  const out = [];
  for await (const entry of dirHandle.values()) {
    if (entry.kind === 'file') out.push(entry);
  }
  return out;
}

async function tryRemove(dirHandle, name) {
  try {
    await dirHandle.removeEntry(name);
  } catch (e) { /* ignore */ }
}

// ----- 起動・初期化 -----
function checkBrowser() {
  if (!('showDirectoryPicker' in window)) {
    els.chooseFolderBtn.disabled = true;
    els.browserNote.innerHTML =
      '⚠️ このブラウザは対応していません。<strong>Google Chrome</strong> で開いてください。';
    return false;
  }
  // 共有フォルダから直接開いていると失敗することがあるため事前警告
  const isUNC = location.href.startsWith('file:////') || location.href.startsWith('file://///');
  if (isUNC) {
    els.browserNote.innerHTML =
      '⚠ このページは共有フォルダから<strong>直接</strong>開かれています。<br>'
      + 'ブラウザの仕様で動作しないことがあります。動かない場合は、共有フォルダにある '
      + '<strong>「GeneticsApp起動.bat」</strong> をダブルクリックして開き直してください。';
  }
  return true;
}

async function withTimeout(promise, ms, message) {
  let timer;
  const timeoutPromise = new Promise((_, reject) => {
    timer = setTimeout(() => reject(new Error(message)), ms);
  });
  try {
    return await Promise.race([promise, timeoutPromise]);
  } finally {
    clearTimeout(timer);
  }
}

function setFolderStatus(text) {
  if (text) {
    els.folderError.style.color = '#4f6056';
    els.folderError.textContent = text;
  } else {
    els.folderError.style.color = '';
    els.folderError.textContent = '';
  }
}

function setFolderError(html) {
  els.folderError.style.color = '#c0392b';
  els.folderError.innerHTML = html;
}

async function chooseFolder() {
  setFolderStatus('');
  els.chooseFolderBtn.disabled = true;
  const originalBtnText = els.chooseFolderBtn.textContent;
  try {
    setFolderStatus('① フォルダ選択ダイアログを開いています…');
    rootHandle = await window.showDirectoryPicker({ mode: 'readwrite' });

    setFolderStatus('② 書き込み許可を確認しています…');
    let perm = await rootHandle.queryPermission({ mode: 'readwrite' });
    if (perm !== 'granted') {
      perm = await rootHandle.requestPermission({ mode: 'readwrite' });
    }
    if (perm !== 'granted') {
      throw new Error('フォルダへの書き込みが許可されませんでした。もう一度試して「編集を許可」を選んでください。');
    }

    setFolderStatus('③ data フォルダを準備しています…');
    dataHandle = await withTimeout(
      getOrCreateDir(rootHandle, 'data'),
      10000,
      'data フォルダの作成が10秒以内に完了しませんでした。共有フォルダへの書き込みが制限されている可能性があります。',
    );
    await withTimeout(getOrCreateDir(dataHandle, 'students'), 10000, 'students フォルダの作成がタイムアウトしました。');
    await withTimeout(getOrCreateDir(dataHandle, 'events'), 10000, 'events フォルダの作成がタイムアウトしました。');
    await withTimeout(getOrCreateDir(dataHandle, 'requests'), 10000, 'requests フォルダの作成がタイムアウトしました。');

    setFolderStatus('④ 設定ファイルを読み込んでいます…');
    const cfg = await withTimeout(readJSON(dataHandle, 'config.json'), 10000, 'config.json の読み込みがタイムアウトしました。');
    if (cfg) config = { ...config, ...cfg };
    else await withTimeout(writeJSON(dataHandle, 'config.json', config), 10000, 'config.json の書き込みがタイムアウトしました。');

    setFolderStatus('');
    els.folderScreen.classList.add('hidden');
    els.loginScreen.classList.remove('hidden');
    els.folderPath.textContent = rootHandle.name + ' / data /';
  } catch (e) {
    if (e.name === 'AbortError') {
      setFolderStatus('');
      return;
    }
    console.error('chooseFolder error', e);
    const isFileUrl = location.protocol === 'file:';
    const isUNC = location.href.startsWith('file:////') || location.href.startsWith('file://///');
    const hint = (isFileUrl && isUNC)
      ? '<br><br><strong>⚠ このHTMLは共有フォルダから直接開かれています。</strong><br>'
        + 'ブラウザのセキュリティ制限により、この形式では動作しないことがあります。<br>'
        + '<strong>対処：</strong>共有フォルダ内の <strong>「GeneticsApp起動.bat」</strong> をダブルクリックして開き直してください。'
        + '（裏でローカルにコピーしてからEdgeを起動する仕組みです）'
      : '<br><br>共有フォルダへの書き込み権限を確認するか、もう一度やり直してください。';
    setFolderError(
      'フォルダを開けませんでした：<br><strong>' + escapeHTML(e.message || String(e)) + '</strong>' + hint,
    );
  } finally {
    els.chooseFolderBtn.disabled = false;
    els.chooseFolderBtn.textContent = originalBtnText;
  }
}

// ----- 参加処理 -----
async function login() {
  const name = els.nameInput.value.trim();
  if (!name) {
    els.loginError.textContent = '名前を入力してください';
    return;
  }
  const isTeacher = els.teacherCheck.checked;
  if (isTeacher && els.passwordInput.value !== TEACHER_PASSWORD) {
    els.loginError.textContent = '先生用パスワードが違います';
    return;
  }

  // 重複チェック
  await loadStudents();
  for (const s of knownStudents.values()) {
    if (!s.isTeacher && s.name === name && (Date.now() - s.lastSeen) < STALE_AFTER_MS) {
      els.loginError.textContent = 'その名前はすでに使われています';
      return;
    }
  }

  me = {
    id: newId(),
    name,
    alleles: isTeacher ? ['—', '—'] : defaultAlleles(config.initialGenotype),
    isTeacher,
    lastSeen: Date.now(),
    joinedAt: Date.now(),
  };

  try {
    const studentsDir = await dataHandle.getDirectoryHandle('students');
    await writeJSON(studentsDir, me.id + '.json', me);
    await writeEvent({ type: 'join', actorId: me.id, message: `${name} さんが${isTeacher ? '（先生として）' : ''}参加しました` });
  } catch (e) {
    els.loginError.textContent = '書き込み失敗：' + e.message;
    return;
  }

  els.loginScreen.classList.add('hidden');
  els.mainScreen.classList.remove('hidden');
  els.myBadge.classList.remove('hidden');
  if (isTeacher) els.teacherPanel.classList.remove('hidden');

  try { localStorage.setItem('genName', name); } catch (e) {}

  startPolling();
  renderMe();
}

// ----- イベント書き込み -----
async function writeEvent(e) {
  const eventsDir = await dataHandle.getDirectoryHandle('events');
  const id = String(Date.now()).padStart(15, '0') + '-' + Math.random().toString(36).slice(2, 6);
  const obj = { id, at: Date.now(), ...e };
  await writeJSON(eventsDir, id + '.json', obj);
  return obj;
}

async function writeRequest(r) {
  const reqDir = await dataHandle.getDirectoryHandle('requests');
  const id = String(Date.now()).padStart(15, '0') + '-' + Math.random().toString(36).slice(2, 6);
  const obj = { id, at: Date.now(), ...r };
  await writeJSON(reqDir, id + '.json', obj);
  return obj;
}

// ----- ポーリング -----
function startPolling() {
  if (pollTimer) clearInterval(pollTimer);
  poll();
  pollTimer = setInterval(poll, POLL_INTERVAL_MS);
}

async function poll() {
  if (pollBusy || !dataHandle || !me) return;
  pollBusy = true;
  els.syncDot.classList.add('active');
  try {
    const studentsDir = await dataHandle.getDirectoryHandle('students');

    // 先生が初期化等で自分のファイルを書き換えている可能性があるので、まず再読込
    if (!me.isTeacher) {
      const fresh = await readJSON(studentsDir, me.id + '.json');
      if (fresh && Array.isArray(fresh.alleles)) {
        me.alleles = fresh.alleles;
      }
    }

    // 心拍を更新
    if (Date.now() - me.lastSeen > HEARTBEAT_MS - 200) {
      me.lastSeen = Date.now();
      await writeJSON(studentsDir, me.id + '.json', me);
    }

    // 設定を再読込
    const cfg = await readJSON(dataHandle, 'config.json');
    if (cfg) config = { ...config, ...cfg };

    await loadStudents();
    await loadEvents();
    await processRequests();

    renderMe();
    renderRoster();
    renderSummary();
  } catch (e) {
    console.error('poll error', e);
    showToast('同期エラー：' + e.message, 4000);
  } finally {
    pollBusy = false;
    setTimeout(() => els.syncDot.classList.remove('active'), 200);
  }
}

async function loadStudents() {
  const studentsDir = await dataHandle.getDirectoryHandle('students');
  const files = await listFiles(studentsDir);
  const next = new Map();
  for (const f of files) {
    if (!f.name.endsWith('.json')) continue;
    try {
      const file = await f.getFile();
      const s = JSON.parse(await file.text());
      if (s && s.id) next.set(s.id, s);
    } catch (e) { /* ignore */ }
  }
  knownStudents = next;
}

async function loadEvents() {
  const eventsDir = await dataHandle.getDirectoryHandle('events');
  const files = await listFiles(eventsDir);
  const newOnes = [];
  for (const f of files) {
    if (!f.name.endsWith('.json')) continue;
    const id = f.name.replace(/\.json$/, '');
    if (seenEventIds.has(id)) continue;
    try {
      const file = await f.getFile();
      const e = JSON.parse(await file.text());
      seenEventIds.add(id);
      newOnes.push(e);
    } catch (err) { /* ignore */ }
  }
  newOnes.sort((a, b) => (a.at || 0) - (b.at || 0));
  for (const e of newOnes) {
    allEvents.push(e);
    handleNewEvent(e);
  }
  if (allEvents.length > HISTORY_LIMIT * 2) {
    allEvents = allEvents.slice(-HISTORY_LIMIT);
  }
  if (newOnes.length > 0) renderHistory();
}

async function processRequests() {
  const reqDir = await dataHandle.getDirectoryHandle('requests');
  const files = await listFiles(reqDir);
  for (const f of files) {
    if (!f.name.endsWith('.json')) continue;
    let req;
    try {
      const file = await f.getFile();
      req = JSON.parse(await file.text());
    } catch (e) { continue; }
    if (!req || !req.to) continue;

    // 古い依頼を掃除
    if (Date.now() - (req.at || 0) > REQUEST_TTL_MS) {
      await tryRemove(reqDir, f.name);
      continue;
    }

    // 自分宛て？
    if (req.to !== me.id) continue;
    if (req._handled) continue;

    if (req.type === 'pair' && !pendingPair && els.pairDialog.classList.contains('hidden')) {
      pendingPair = { req, fileName: f.name };
      const partner = knownStudents.get(req.from);
      if (!partner) {
        await tryRemove(reqDir, f.name);
        pendingPair = null;
        continue;
      }
      els.pairDialogText.textContent =
        `${partner.name} さん（${genotypeOf(partner.alleles)}）から交配の申し込みが届きました。受けますか？`;
      els.pairDialog.classList.remove('hidden');
      break;
    }
    if (req.type === 'gift' && !pendingGift && els.giftDialog.classList.contains('hidden')) {
      pendingGift = { req, fileName: f.name };
      const partner = knownStudents.get(req.from);
      if (!partner) {
        await tryRemove(reqDir, f.name);
        pendingGift = null;
        continue;
      }
      els.giftDialogText.textContent =
        `${partner.name} さんから「${req.allele}」を受け取りますか？（自分の遺伝子1つと交換になります）`;
      els.giftDialog.classList.remove('hidden');
      break;
    }
  }
}

function handleNewEvent(e) {
  if (!e || !e.type) return;
  if (e.type === 'cross') {
    if (e.participantIds && e.participantIds.includes(me.id)) {
      showResultDialog({
        title: '🌱 子の遺伝子ができました！',
        body: `
          <p><strong>${escapeHTML(e.data.parentA.name)}</strong>（${e.data.parentA.genotype}）から <strong>${e.data.pickedFromA}</strong>、
          <strong>${escapeHTML(e.data.parentB.name)}</strong>（${e.data.parentB.genotype}）から <strong>${e.data.pickedFromB}</strong> をランダムに選びました。</p>
          <div class="child-card">${e.data.childGenotype}</div>
          <p>形質：<strong>${e.data.childPhenotype}</strong>　／　遺伝子型：<strong>${e.data.childGenotype}</strong></p>
          <p class="hint">※ この結果はクラス全体の集計には加わりません。記録用紙にメモしましょう。</p>
        `,
      });
    }
  } else if (e.type === 'pairDeclined') {
    if (e.toId === me.id) {
      showToast(`${e.byName} さんに断られました`);
    }
  } else if (e.type === 'gift') {
    // 申込側の自分の遺伝子を更新
    if (e.from === me.id && !e._appliedHere) {
      e._appliedHere = true;
      me.alleles = [...me.alleles];
      me.alleles[e.sourceIndex] = e.returnedAllele;
      me.lastSeen = Date.now();
      dataHandle.getDirectoryHandle('students').then((d) => writeJSON(d, me.id + '.json', me));
    }
    if (e.participantIds && e.participantIds.includes(me.id)) {
      const fromName = e.fromName, toName = e.toName;
      showResultDialog({
        title: '🔄 遺伝子を交換しました',
        body: `
          <p><strong>${escapeHTML(fromName)}</strong> さんが「${e.sentAllele}」を渡し、
          <strong>${escapeHTML(toName)}</strong> さんが「${e.returnedAllele}」を返しました。</p>
        `,
      });
    }
  } else if (e.type === 'giftDeclined') {
    if (e.toId === me.id) showToast(`${e.byName} さんに交換を断られました`);
  }
  // reset / teacherSet は先生が各生徒ファイルを直接書き換えるため、
  // クライアント側での自動上書きは行わない（次のポーリングで再読込される）。
}

// ----- 描画 -----
function renderMe() {
  if (!me) return;
  els.myName.textContent = me.name + (me.isTeacher ? '（先生）' : '');
  if (me.isTeacher) {
    els.myCards.textContent = '👩‍🏫';
    els.bigCards.innerHTML = '<p class="muted">先生は遺伝子を持ちません</p>';
    els.myGenotype.textContent = '—';
    els.myPhenotype.textContent = '—';
  } else {
    const g = genotypeOf(me.alleles);
    els.myCards.textContent = g;
    els.bigCards.innerHTML = genotypeToCards(g);
    els.myGenotype.textContent = g;
    els.myPhenotype.textContent = phenotypeOf(g);
  }
}

function genotypeToCards(g) {
  return [...g].map((c) => `<div class="gene-card ${c === c.toLowerCase() ? 'lower' : ''}">${c}</div>`).join('');
}

function activeStudents() {
  const now = Date.now();
  return [...knownStudents.values()].filter((s) => now - (s.lastSeen || 0) < STALE_AFTER_MS);
}

function renderRoster() {
  const all = activeStudents().sort((a, b) => (a.joinedAt || 0) - (b.joinedAt || 0));
  const others = all.filter((s) => !s.isTeacher && s.id !== me.id);
  els.rosterCount.textContent = `（生徒 ${all.filter((s) => !s.isTeacher).length} 人）`;

  const fillSelect = (sel) => {
    const prev = sel.value;
    sel.innerHTML = '<option value="">- 相手を選んでください -</option>' +
      others.map((s) => `<option value="${s.id}">${escapeHTML(s.name)} (${genotypeOf(s.alleles)})</option>`).join('');
    if (others.find((s) => s.id === prev)) sel.value = prev;
  };
  fillSelect(els.pairTarget);
  fillSelect(els.giftTarget);

  els.roster.innerHTML = all.map((s) => {
    const isMe = s.id === me.id;
    const cls = ['student'];
    if (isMe) cls.push('me');
    if (s.isTeacher) cls.push('teacher');
    const g = s.isTeacher ? '—' : genotypeOf(s.alleles);
    const teacherCtl = me.isTeacher && !s.isTeacher
      ? `<button data-set="${s.id}" type="button">設定</button>`
      : '';
    return `<div class="${cls.join(' ')}">
      <div class="name">${escapeHTML(s.name)}${s.isTeacher ? ' 👩‍🏫' : ''}${isMe ? ' (あなた)' : ''}</div>
      <div class="mini-cards">${g}</div>
      <div class="pheno">${s.isTeacher ? '' : '形質：' + phenotypeOf(g)}</div>
      ${teacherCtl}
    </div>`;
  }).join('');

  els.roster.querySelectorAll('button[data-set]').forEach((b) => {
    b.addEventListener('click', () => teacherSetStudent(b.dataset.set));
  });

  if (me.isTeacher) refreshInitPreview();
}

function renderSummary() {
  const counts = { AA: 0, Aa: 0, aa: 0 };
  for (const s of activeStudents()) {
    if (s.isTeacher) continue;
    const g = genotypeOf(s.alleles);
    if (counts[g] != null) counts[g] += 1;
  }
  els.cAA.textContent = counts.AA;
  els.cAa.textContent = counts.Aa;
  els.caa.textContent = counts.aa;
  const round = counts.AA + counts.Aa;
  const wrinkle = counts.aa;
  els.phenotypeRatio.textContent = `${round} : ${wrinkle}`;
  const total = counts.AA + counts.Aa + counts.aa;
  els.genotypeRatio.textContent = total > 0 ? `${counts.AA} : ${counts.Aa} : ${counts.aa}` : '-';
}

function renderHistory() {
  const recent = allEvents.slice(-HISTORY_LIMIT).reverse();
  els.history.innerHTML = recent.map((e) => {
    const t = new Date(e.at);
    const ts = `${String(t.getHours()).padStart(2, '0')}:${String(t.getMinutes()).padStart(2, '0')}:${String(t.getSeconds()).padStart(2, '0')}`;
    const msg = e.message || formatEventMessage(e);
    return `<div class="entry ${e.type || ''}">[${ts}] ${escapeHTML(msg)}</div>`;
  }).join('');
}

function formatEventMessage(e) {
  if (e.type === 'cross' && e.data) {
    return `${e.data.parentA.name}（${e.data.parentA.genotype}）× ${e.data.parentB.name}（${e.data.parentB.genotype}）→ 子: ${e.data.childGenotype}（${e.data.childPhenotype}）`;
  }
  if (e.type === 'gift') {
    return `${e.fromName} と ${e.toName} が遺伝子を交換しました（${e.sentAllele} ⇄ ${e.returnedAllele}）`;
  }
  if (e.type === 'pairDeclined') return `${e.byName} さんが交配を断りました`;
  if (e.type === 'giftDeclined') return `${e.byName} さんが交換を断りました`;
  if (e.type === 'reset') {
    const c = e.counts;
    if (c) return `先生が遺伝子を初期化しました（AA ${c.AA} ／ Aa ${c.Aa} ／ aa ${c.aa}）`;
    return `先生が遺伝子を初期化しました（初期: ${e.initialGenotype}）`;
  }
  if (e.type === 'teacherSet') return `先生が ${e.targetName} さんを ${e.genotype} に設定しました`;
  return e.type;
}

function showResultDialog({ title, body }) {
  els.resultTitle.textContent = title;
  els.resultBody.innerHTML = body;
  els.resultDialog.classList.remove('hidden');
}

// ----- 行動 -----
async function requestPair() {
  const targetId = els.pairTarget.value;
  if (!targetId) return showToast('相手を選んでください');
  const target = knownStudents.get(targetId);
  if (!target) return showToast('相手が見つかりません');
  els.pairBtn.disabled = true;
  try {
    await writeRequest({ type: 'pair', from: me.id, fromName: me.name, to: targetId, toName: target.name });
    showToast('交配を申し込みました。相手の返事を待っています…');
  } catch (e) {
    showToast('送信失敗：' + e.message);
  }
  els.pairBtn.disabled = false;
}

async function respondPair(accept) {
  if (!pendingPair) return;
  const { req, fileName } = pendingPair;
  pendingPair = null;
  els.pairDialog.classList.add('hidden');

  const reqDir = await dataHandle.getDirectoryHandle('requests');
  await tryRemove(reqDir, fileName);

  const partner = knownStudents.get(req.from);
  if (!partner) return;

  if (!accept) {
    await writeEvent({
      type: 'pairDeclined',
      toId: req.from,
      byName: me.name,
    });
    return;
  }

  // 子をつくる
  const childAlleles = [pickRandom(partner.alleles), pickRandom(me.alleles)];
  const childGenotype = genotypeOf(childAlleles);
  const childPhenotype = phenotypeOf(childGenotype);

  await writeEvent({
    type: 'cross',
    participantIds: [partner.id, me.id],
    data: {
      parentA: { id: partner.id, name: partner.name, genotype: genotypeOf(partner.alleles) },
      parentB: { id: me.id, name: me.name, genotype: genotypeOf(me.alleles) },
      pickedFromA: childAlleles[0],
      pickedFromB: childAlleles[1],
      childGenotype,
      childPhenotype,
    },
  });
}

async function requestGift() {
  const targetId = els.giftTarget.value;
  if (!targetId) return showToast('相手を選んでください');
  const target = knownStudents.get(targetId);
  if (!target) return showToast('相手が見つかりません');
  const idx = parseInt(document.querySelector('input[name="giveAllele"]:checked').value, 10);
  const allele = me.alleles[idx];
  els.giftBtn.disabled = true;
  try {
    await writeRequest({
      type: 'gift', from: me.id, fromName: me.name,
      to: targetId, toName: target.name,
      allele, sourceIndex: idx,
    });
    showToast(`「${allele}」を渡そうとしています。相手の返事を待っています…`);
  } catch (e) {
    showToast('送信失敗：' + e.message);
  }
  els.giftBtn.disabled = false;
}

async function respondGift(accept) {
  if (!pendingGift) return;
  const { req, fileName } = pendingGift;
  pendingGift = null;
  els.giftDialog.classList.add('hidden');

  const reqDir = await dataHandle.getDirectoryHandle('requests');
  await tryRemove(reqDir, fileName);

  if (!accept) {
    await writeEvent({ type: 'giftDeclined', toId: req.from, byName: me.name });
    return;
  }

  const idx = parseInt(document.querySelector('input[name="replaceIdx"]:checked').value, 10);
  const replaced = me.alleles[idx];
  me.alleles = [...me.alleles];
  me.alleles[idx] = req.allele;
  me.lastSeen = Date.now();
  const studentsDir = await dataHandle.getDirectoryHandle('students');
  await writeJSON(studentsDir, me.id + '.json', me);

  await writeEvent({
    type: 'gift',
    from: req.from, fromName: req.fromName,
    to: me.id, toName: me.name,
    sentAllele: req.allele,
    returnedAllele: replaced,
    sourceIndex: req.sourceIndex,
    targetIndex: idx,
    participantIds: [req.from, me.id],
  });
  renderMe();
}

// ----- 先生メニュー -----
function shuffleInPlace(arr) {
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
  return arr;
}

function readCustomRatio() {
  const rAA = Math.max(0, parseInt(els.ratioAA.value, 10) || 0);
  const rAa = Math.max(0, parseInt(els.ratioAa.value, 10) || 0);
  const raa = Math.max(0, parseInt(els.ratioaa.value, 10) || 0);
  return { rAA, rAa, raa, total: rAA + rAa + raa };
}

// 配分を指定された比率に従って n 人分の遺伝子型配列を返す
function distributeByRatio(n, rAA, rAa, raa) {
  const total = rAA + rAa + raa;
  if (total === 0 || n === 0) return [];
  const nAA = Math.round((n * rAA) / total);
  const naa = Math.round((n * raa) / total);
  const nAa = Math.max(0, n - nAA - naa);
  return [
    ...Array(nAA).fill('AA'),
    ...Array(nAa).fill('Aa'),
    ...Array(naa).fill('aa'),
  ];
}

function buildAssignments(mode, students) {
  const n = students.length;
  if (mode === 'all-AA') return Array(n).fill('AA');
  if (mode === 'all-Aa') return Array(n).fill('Aa');
  if (mode === 'all-aa') return Array(n).fill('aa');
  if (mode === 'half') {
    // 半数を AA、残りを aa（奇数なら AA を1人多く）
    const half = Math.ceil(n / 2);
    return [...Array(half).fill('AA'), ...Array(n - half).fill('aa')];
  }
  if (mode === 'custom') {
    const { rAA, rAa, raa } = readCustomRatio();
    return distributeByRatio(n, rAA, rAa, raa);
  }
  return Array(n).fill('Aa');
}

// 「初期化したらこうなる」プレビューを更新
function refreshInitPreview() {
  if (!me || !me.isTeacher) return;
  const mode = els.initMode.value;
  els.customRatioRow.classList.toggle('hidden', mode !== 'custom');

  if (mode === 'custom') {
    const { total } = readCustomRatio();
    if (total === 0) {
      els.initPreview.textContent = '⚠ 比率を1つ以上指定してください';
      return;
    }
  }
  const students = activeStudents().filter((s) => !s.isTeacher);
  const assignments = buildAssignments(mode, students);
  const counts = { AA: 0, Aa: 0, aa: 0 };
  for (const g of assignments) counts[g] = (counts[g] || 0) + 1;
  els.initPreview.textContent =
    `初期化すると：AA ${counts.AA}人 ／ Aa ${counts.Aa}人 ／ aa ${counts.aa}人 （合計 ${assignments.length}人）`;
}

async function teacherReset() {
  if (!me.isTeacher) return;
  const mode = els.initMode.value;

  if (mode === 'custom') {
    const { total } = readCustomRatio();
    if (total === 0) { showToast('比率を1つ以上指定してください'); return; }
  }

  if (!confirm('全員の遺伝子を初期化します。よろしいですか？')) return;

  // 新規参加者用のデフォルトを保存（custom/half の場合は Aa を採用）
  const defaultForNew = mode.startsWith('all-') ? mode.slice(4) : 'Aa';
  config.initialGenotype = defaultForNew;
  config.lastResetMode = mode;
  if (mode === 'custom') config.lastRatio = readCustomRatio();
  await writeJSON(dataHandle, 'config.json', config);

  // 生徒一覧を取り出してランダム順に並べてから割り当て
  const students = activeStudents().filter((s) => !s.isTeacher);
  shuffleInPlace(students);
  const assignments = buildAssignments(mode, students);
  // 念のため、人数が足りない場合は末尾を Aa で埋める（custom の丸め誤差対策）
  while (assignments.length < students.length) assignments.push('Aa');

  const studentsDir = await dataHandle.getDirectoryHandle('students');
  const counts = { AA: 0, Aa: 0, aa: 0 };
  for (let i = 0; i < students.length; i++) {
    const s = students[i];
    const g = assignments[i] || 'Aa';
    s.alleles = defaultAlleles(g);
    s.lastSeen = Date.now();
    await writeJSON(studentsDir, s.id + '.json', s);
    counts[g] = (counts[g] || 0) + 1;
  }

  await writeEvent({
    type: 'reset',
    mode,
    counts,
    initialGenotype: defaultForNew,
  });
  showToast(`初期化しました（AA ${counts.AA} ／ Aa ${counts.Aa} ／ aa ${counts.aa}）`, 4000);
  refreshInitPreview();
}

async function teacherSetStudent(studentId) {
  if (!me.isTeacher) return;
  const g = prompt('遺伝子型を入力してください（AA / Aa / aa）', 'Aa');
  if (!g || !['AA', 'Aa', 'aa'].includes(g)) return showToast('AA / Aa / aa のいずれかを入力してください');
  const studentsDir = await dataHandle.getDirectoryHandle('students');
  const s = knownStudents.get(studentId);
  if (!s) return;
  s.alleles = defaultAlleles(g);
  s.lastSeen = Date.now();
  await writeJSON(studentsDir, s.id + '.json', s);
  await writeEvent({ type: 'teacherSet', targetId: s.id, targetName: s.name, genotype: g });
}

async function teacherCleanup() {
  if (!me.isTeacher) return;
  if (!confirm('古いイベント・依頼ファイルと退出した生徒のファイルを削除します。よろしいですか？')) return;
  let removed = 0;
  const eventsDir = await dataHandle.getDirectoryHandle('events');
  for (const f of await listFiles(eventsDir)) {
    await tryRemove(eventsDir, f.name); removed++;
  }
  const reqDir = await dataHandle.getDirectoryHandle('requests');
  for (const f of await listFiles(reqDir)) {
    await tryRemove(reqDir, f.name); removed++;
  }
  const studentsDir = await dataHandle.getDirectoryHandle('students');
  for (const s of [...knownStudents.values()]) {
    if (s.id === me.id) continue;
    if (Date.now() - (s.lastSeen || 0) > STALE_AFTER_MS * 2) {
      await tryRemove(studentsDir, s.id + '.json');
      removed++;
    }
  }
  seenEventIds = new Set();
  allEvents = [];
  renderHistory();
  showToast(`${removed} 個のファイルを削除しました`);
}

// ----- 退出 -----
async function logout() {
  if (!confirm('退出しますか？')) return;
  if (pollTimer) clearInterval(pollTimer);
  try {
    const studentsDir = await dataHandle.getDirectoryHandle('students');
    await tryRemove(studentsDir, me.id + '.json');
    await writeEvent({ type: 'leave', actorId: me.id, message: `${me.name} さんが退出しました` });
  } catch (e) {}
  location.reload();
}

window.addEventListener('beforeunload', () => {
  if (!dataHandle || !me) return;
  // 退出時にファイルを削除（同期APIではなくfire-and-forget）
  dataHandle.getDirectoryHandle('students')
    .then((d) => d.removeEntry(me.id + '.json'))
    .catch(() => {});
});

// ----- イベント結線 -----
function setup() {
  if (!checkBrowser()) return;

  els.chooseFolderBtn.addEventListener('click', chooseFolder);
  els.changeFolderBtn.addEventListener('click', () => location.reload());

  els.teacherCheck.addEventListener('change', () => {
    els.passwordRow.classList.toggle('hidden', !els.teacherCheck.checked);
  });

  try {
    const saved = localStorage.getItem('genName');
    if (saved) els.nameInput.value = saved;
  } catch (e) {}

  els.loginBtn.addEventListener('click', login);
  els.nameInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') login();
  });

  els.logoutBtn.addEventListener('click', logout);

  document.querySelectorAll('.mode-tabs .tab').forEach((tab) => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.mode-tabs .tab').forEach((t) => t.classList.remove('active'));
      tab.classList.add('active');
      const mode = tab.dataset.mode;
      $('pairMode').classList.toggle('hidden', mode !== 'pair');
      $('giftMode').classList.toggle('hidden', mode !== 'gift');
    });
  });

  els.pairBtn.addEventListener('click', requestPair);
  els.giftBtn.addEventListener('click', requestGift);
  els.pairAccept.addEventListener('click', () => respondPair(true));
  els.pairDecline.addEventListener('click', () => respondPair(false));
  els.giftAccept.addEventListener('click', () => respondGift(true));
  els.giftDecline.addEventListener('click', () => respondGift(false));
  els.resultClose.addEventListener('click', () => els.resultDialog.classList.add('hidden'));

  els.resetBtn.addEventListener('click', teacherReset);
  els.cleanupBtn.addEventListener('click', teacherCleanup);

  els.initMode.addEventListener('change', refreshInitPreview);
  ['ratioAA', 'ratioAa', 'ratioaa'].forEach((id) => {
    els[id].addEventListener('input', refreshInitPreview);
  });
}

setup();
