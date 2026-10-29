/* Tide term audit pane.
 *
 * Runs inside the WPS task pane. Termbase and settings stay local. The change
 * archive is AES-GCM encrypted at rest and only the approved termbase can be
 * pushed, through the local helper, with its own password.
 */
(function () {
  'use strict';

  var K = {
    settings: 'tide.settings',
    terms: 'tide.terms',
    archive: 'tide.archive',
    snapshot: 'tide.snapshot',
    lastTab: 'tide.lastTab',
    repo: 'tide.syncRepo',
    branch: 'tide.syncBranch'
  };
  var HISTORY_LIMIT = 3000;
  var HELPER = 'http://127.0.0.1:8765';

  var state = {
    terms: [],
    settings: { sourceCol: 'C', targetCol: 'F', uidCol: 'A' },
    scanned: [],
    archive: { history: [], snapshot: {} },
    passphrase: null,
    unlocked: false,
    lastTab: 'matches',
    helper: { ok: false, token: false },
    bundledTerms: [],
    matchTerms: [],
    bundledMeta: null,
    termsBlob: null,
    entries: [],
    cursor: 0,
    lastWrite: null
  };

  function $(id) { return document.getElementById(id); }

  function readJSON(key, fallback) {
    try {
      var raw = window.localStorage.getItem(key);
      return raw ? JSON.parse(raw) : fallback;
    } catch (err) {
      return fallback;
    }
  }

  function writeJSON(key, value) {
    try {
      window.localStorage.setItem(key, JSON.stringify(value));
      return true;
    } catch (err) {
      return false;
    }
  }

  function sheetApp() {
    return (typeof window.Application !== 'undefined' && window.Application) ? window.Application : null;
  }

  function currentSheet() {
    var app = sheetApp();
    return app && app.ActiveSheet ? app.ActiveSheet : null;
  }

  function colToNumber(letter) {
    var text = String(letter || '').toUpperCase().replace(/[^A-Z]/g, '');
    var value = 0;
    for (var i = 0; i < text.length; i += 1) value = value * 26 + (text.charCodeAt(i) - 64);
    return value;
  }

  // WPS returns Value2 as a 1-based 2D array; tolerate 0-based arrays too.
  function toGrid(values) {
    if (values === null || values === undefined) return [];
    if (!Array.isArray(values)) return [[values]];
    var oneBased = values[0] === undefined && values[1] !== undefined;
    var rows = oneBased ? values.slice(1) : values;
    return rows.map(function (row) {
      if (!Array.isArray(row)) return [row];
      return oneBased ? row.slice(1) : row;
    });
  }

  function cellText(grid, r, c) {
    var row = grid[r];
    if (!row || c < 0) return '';
    var value = row[c];
    return value === null || value === undefined ? '' : String(value);
  }

  function buildRows() {
    var sheet = currentSheet();
    if (!sheet) return { error: 'not-in-wps' };
    var used = sheet.UsedRange;
    if (!used) return { rows: [], sheetName: sheet.Name };
    var grid = toGrid(used.Value2);
    var startRow = Number(used.Row) || 1;
    var startCol = Number(used.Column) || 1;
    var srcIdx = colToNumber(state.settings.sourceCol) - startCol;
    var tgtIdx = colToNumber(state.settings.targetCol) - startCol;
    var uidIdx = colToNumber(state.settings.uidCol) - startCol;
    var rows = [];
    for (var r = 1; r < grid.length; r += 1) {
      var source = cellText(grid, r, srcIdx);
      if (!source.trim()) continue;
      rows.push({
        row: startRow + r,
        uid: uidIdx >= 0 ? cellText(grid, r, uidIdx) : '',
        source: source,
        target: tgtIdx >= 0 ? cellText(grid, r, tgtIdx) : ''
      });
    }
    return { rows: rows, sheetName: sheet.Name };
  }

  // ---------- encrypted archive ----------

  async function saveArchive() {
    if (!state.unlocked || !state.passphrase) return false;
    try {
      var blob = await window.TideCrypto.encrypt(state.passphrase, JSON.stringify(state.archive));
      window.localStorage.setItem(K.archive, JSON.stringify(blob));
      return true;
    } catch (err) {
      $('lockStatus').textContent = '归档写入失败：' + err.message;
      return false;
    }
  }

  async function unlockArchive(password) {
    var raw = window.localStorage.getItem(K.archive);
    if (!raw) throw new Error('尚未设置归档口令');
    var text = await window.TideCrypto.decrypt(password, JSON.parse(raw));
    state.archive = JSON.parse(text);
    state.passphrase = password;
    state.unlocked = true;
  }

  async function createArchive(password) {
    state.archive = { history: [], snapshot: {} };
    state.passphrase = password;
    state.unlocked = true;
    await saveArchive();
  }

  function lockArchive() {
    state.passphrase = null;
    state.unlocked = false;
    state.archive = { history: [], snapshot: {} };
    renderLock();
    renderHistory();
  }

  function archiveChanges(sheetName, result) {
    if (!state.unlocked) return 0;
    var snapshot = state.archive.snapshot;
    var history = state.archive.history;
    var changed = 0;
    var stamp = new Date().toISOString();
    result.rows.forEach(function (row) {
      [[state.settings.sourceCol, row.source], [state.settings.targetCol, row.target]].forEach(function (pair) {
        var key = sheetName + '!' + pair[0] + row.row;
        var before = Object.prototype.hasOwnProperty.call(snapshot, key) ? snapshot[key] : null;
        if (before !== null && before !== pair[1]) {
          history.push({ ts: stamp, sheet: sheetName, cell: pair[0] + row.row, before: before, after: pair[1] });
          changed += 1;
        }
        snapshot[key] = pair[1];
      });
    });
    if (history.length > HISTORY_LIMIT) history.splice(0, history.length - HISTORY_LIMIT);
    return changed;
  }

  // ---------- rendering ----------

  function textNode(tag, text, cls) {
    var el = document.createElement(tag);
    if (cls) el.className = cls;
    el.textContent = text;
    return el;
  }

  function highlighted(source, matches) {
    var wrap = document.createElement('span');
    wrap.className = 'zh';
    var cursor = 0;
    matches.forEach(function (hit) {
      var start = source.indexOf(hit.matched, cursor);
      if (start < 0) return;
      if (start > cursor) wrap.appendChild(document.createTextNode(source.slice(cursor, start)));
      var mark = document.createElement('mark');
      mark.textContent = hit.matched;
      wrap.appendChild(mark);
      cursor = start + hit.matched.length;
    });
    if (cursor < source.length) wrap.appendChild(document.createTextNode(source.slice(cursor)));
    return wrap;
  }

  function renderMatches() {
    var table = $('matchTable');
    var body = table.tBodies[0];
    body.textContent = '';
    $('matchCount').textContent = state.scanned.length ? state.scanned.length + ' 行命中' : '无命中';
    $('matchEmpty').hidden = state.scanned.length > 0;
    table.hidden = state.scanned.length === 0;
    state.scanned.forEach(function (row) {
      var tr = document.createElement('tr');
      tr.appendChild(textNode('td', String(row.row), 'num'));
      var tdSource = document.createElement('td');
      tdSource.appendChild(highlighted(row.source, row.matches));
      tr.appendChild(tdSource);
      var tdPair = document.createElement('td');
      row.matches.forEach(function (match) {
        tdPair.appendChild(textNode('div', (match.zh || '') + ' → ' + (match.en || '—'), 'pair'));
      });
      tr.appendChild(tdPair);
      tr.appendChild(textNode('td', row.target || '—', 'hint'));
      body.appendChild(tr);
    });
  }

  function renderTerms() {
    var grouped = window.TideEngine.groupByTerm(state.scanned);
    var table = $('termTable');
    var body = table.tBodies[0];
    body.textContent = '';
    $('termCount').textContent = state.terms.length ? state.terms.length + ' 条术语' : '术语表未导入';
    $('termEmpty').hidden = state.terms.length > 0;
    table.hidden = !state.terms.length;
    grouped.forEach(function (group) {
      var tr = document.createElement('tr');
      tr.appendChild(textNode('td', group.zh));
      tr.appendChild(textNode('td', group.en || '—'));
      tr.appendChild(textNode('td', String(group.rows.length), 'num'));
      body.appendChild(tr);
    });
  }

  // Entries flagged needs_review carry notes or several variants in the English
  // column; they would produce false matches, so they stay visible but unused.
  function setTerms(list) {
    state.terms = list || [];
    state.matchTerms = state.terms.filter(function (term) {
      return term.status !== 'needs_review';
    });
  }

  function renderTermSource() {
    var el = $('termSource');
    if (!el) return;
    if (state.bundledMeta) {
      var meta = state.bundledMeta;
      var held = state.terms.length - state.matchTerms.length;
      el.textContent = '内置术语表 ' + meta.count + ' 条 · ' + meta.source + ' · 构建 ' +
        String(meta.built_at || '').slice(0, 10) +
        (held ? ' · ' + held + ' 条待确认，不参与匹配' : '');
    } else if (state.termsBlob) {
      el.textContent = '术语表已加密随包分发；到下方输入口令解锁。';
    } else {
      el.textContent = '内置术语表未加载；可用「导入术语表」或「恢复内置术语表」。';
    }
  }

  function usePayload(data) {
    state.bundledTerms = window.TideEngine.parseTermbase(data.terms);
    state.bundledMeta = data.meta || { count: state.bundledTerms.length, source: 'bundled' };
    if (!state.terms.length) setTerms(state.bundledTerms.slice());
    renderTermSource();
    renderTerms();
    return state.bundledTerms;
  }

  // Portable builds ship the termbase encrypted (data/terms.enc.json); a local
  // build has the plaintext file, which needs no passphrase.
  function loadBundled() {
    return fetch('../data/terms.json').then(function (response) {
      return response.ok ? response.json() : null;
    }).then(function (data) {
      if (data && Array.isArray(data.terms)) return usePayload(data);
      return fetch('../data/terms.enc.json').then(function (response) {
        return response.ok ? response.json() : null;
      }).then(function (blob) {
        if (!blob) return null;
        state.termsBlob = blob;
        var saved = window.localStorage.getItem('tide.termsPass');
        if (saved) return decryptTerms(saved).catch(function () { return null; });
        $('termsUnlockBlock').hidden = false;
        renderTermSource();
        return null;
      });
    }).catch(function () {
      return null;
    });
  }

  function decryptTerms(passphrase) {
    return window.TideCrypto.decrypt(passphrase, state.termsBlob).then(function (text) {
      window.localStorage.setItem('tide.termsPass', passphrase);
      $('termsUnlockBlock').hidden = true;
      return usePayload(JSON.parse(text));
    });
  }

  // ---------- translation workspace ----------

  function loadEntries() {
    var result = buildRows();
    if (result.error) {
      state.entries = [];
      return [];
    }
    state.entries = result.rows;
    var firstEmpty = 0;
    for (var i = 0; i < state.entries.length; i += 1) {
      if (!String(state.entries[i].target || '').trim()) { firstEmpty = i; break; }
    }
    state.cursor = firstEmpty;
    return state.entries;
  }

  function currentEntry() {
    return state.entries[state.cursor] || null;
  }

  function entryAddress(row) {
    return state.settings.targetCol + row;
  }

  function writeCell(row, value) {
    var sheet = currentSheet();
    if (!sheet) return false;
    var cell = sheet.Range(entryAddress(row));
    var before = cell.Value2;
    cell.Value2 = value;
    state.lastWrite = { row: row, before: before, after: value, address: entryAddress(row) };
    return true;
  }

  function recordWrite(entry, before, after) {
    if (!state.unlocked) return;
    state.archive.history.push({
      ts: new Date().toISOString(),
      sheet: currentSheet() ? currentSheet().Name : '',
      cell: entryAddress(entry.row),
      before: before == null ? '' : String(before),
      after: String(after)
    });
    saveArchive();
  }

  function renderEntry() {
    var entry = currentEntry();
    $('entryPos').textContent = state.entries.length
      ? (state.cursor + 1) + ' / ' + state.entries.length
      : '0 / 0';
    if (!entry) {
      $('entrySource').textContent = '当前表没有可翻译的条目。';
      $('entryDraft').value = '';
      $('entryCheck').textContent = '';
      return;
    }
    $('entrySource').textContent = entry.source;
    $('entryDraft').value = entry.target || '';
    $('entryCheck').textContent = describeEntry(entry, $('entryDraft').value);
  }

  function describeEntry(entry, draft) {
    var head = '第 ' + entry.row + ' 行 · ' + entryAddress(entry.row);
    if (!String(draft || '').trim()) return head + ' · 尚未翻译';
    var check = window.TideEngine.checkTokens(entry.source, draft);
    if (check.ok) return head + ' · 占位符一致';
    return head + ' · 占位符不一致：缺 ' + (check.missing.join(' ') || '无') +
      ' / 多 ' + (check.extra.join(' ') || '无');
  }

  function move(delta) {
    if (!state.entries.length) return;
    state.cursor = Math.min(Math.max(state.cursor + delta, 0), state.entries.length - 1);
    renderEntry();
  }

  function writeAndNext() {
    var entry = currentEntry();
    if (!entry) return;
    var sheet = currentSheet();
    if (!sheet) { $('entryCheck').textContent = '未连接 WPS，无法写入。'; return; }
    var value = $('entryDraft').value;
    if (!value.trim()) { $('entryCheck').textContent = '草稿为空，未写入。'; return; }
    var before = entry.target;
    if (!writeCell(entry.row, value)) return;
    var address = entryAddress(entry.row);
    entry.target = value;
    recordWrite(entry, before, value);
    move(1);
    $('entryCheck').textContent = '已写入 ' + address + ' · ' + $('entryCheck').textContent;
  }

  function undoSelf() {
    if (!state.lastWrite) { $('entryCheck').textContent = '没有可撤销的写入。'; return; }
    var sheet = currentSheet();
    if (!sheet) return;
    sheet.Range(state.lastWrite.address).Value2 = state.lastWrite.before;
    var entry = null;
    for (var i = 0; i < state.entries.length; i += 1) {
      if (state.entries[i].row === state.lastWrite.row) { entry = state.entries[i]; break; }
    }
    if (entry) entry.target = state.lastWrite.before == null ? '' : String(state.lastWrite.before);
    $('entryCheck').textContent = '已撤销 ' + state.lastWrite.address;
    state.lastWrite = null;
    renderEntry();
  }

  function locateTarget() {
    var entry = currentEntry();
    var sheet = currentSheet();
    if (!entry || !sheet) return;
    sheet.Range(entryAddress(entry.row)).Select();
  }

  function deliveryCheck() {
    loadEntries();
    var issues = [];
    state.entries.forEach(function (entry) {
      var target = String(entry.target || '');
      if (!target.trim()) {
        issues.push({ row: entry.row, problem: '空译文', source: entry.source });
        return;
      }
      var check = window.TideEngine.checkTokens(entry.source, target);
      if (!check.ok) {
        issues.push({ row: entry.row, problem: '占位符不匹配', source: entry.source });
        return;
      }
      if (target === entry.source && window.TideEngine.hasCJK(target)) {
        issues.push({ row: entry.row, problem: '疑似未翻译', source: entry.source });
      }
    });
    $('deliveryCount').textContent = issues.length + ' 个问题 / ' + state.entries.length + ' 条';
    var table = $('deliveryTable');
    var body = table.tBodies[0];
    body.textContent = '';
    table.hidden = issues.length === 0;
    issues.slice(0, 100).forEach(function (issue) {
      var tr = document.createElement('tr');
      tr.appendChild(textNode('td', String(issue.row), 'num'));
      tr.appendChild(textNode('td', issue.problem));
      tr.appendChild(textNode('td', issue.source.slice(0, 40), 'hint'));
      body.appendChild(tr);
    });
    renderEntry();
  }

  function renderLock() {
    var exists = !!window.localStorage.getItem(K.archive);
    $('btnUnlock').hidden = !exists || state.unlocked;
    $('btnSetPass').hidden = exists || state.unlocked;
    $('passInput').hidden = state.unlocked;
    $('lockBar').hidden = state.unlocked && !!state.passphrase;
    $('lockStatus').textContent = state.unlocked
      ? '归档已解锁，本次修改会加密记录。'
      : (exists ? '归档已锁定。解锁后才会记录改动。' : '尚未设置归档口令。设置后修改才会被记录。');
  }

  function renderHistory() {
    var history = state.archive.history || [];
    var table = $('historyTable');
    var body = table.tBodies[0];
    body.textContent = '';
    $('historyCount').textContent = state.unlocked ? history.length + ' 条' : '已锁定';
    $('historyEmpty').hidden = history.length > 0 || !state.unlocked;
    table.hidden = history.length === 0 || !state.unlocked;
    history.slice(-200).reverse().forEach(function (entry) {
      var tr = document.createElement('tr');
      tr.appendChild(textNode('td', entry.ts.slice(11, 19), 'num'));
      tr.appendChild(textNode('td', entry.sheet, 'num'));
      tr.appendChild(textNode('td', entry.cell, 'num'));
      tr.appendChild(textNode('td', entry.before + ' → ' + entry.after, 'hint'));
      body.appendChild(tr);
    });
  }

  async function scan() {
    var result = buildRows();
    if (result.error) {
      $('matchCount').textContent = '未在 WPS 中运行';
      return;
    }
    state.scanned = window.TideEngine.scanRows(result.rows, state.matchTerms);
    renderMatches();
    if (state.terms.length) {
      var changed = archiveChanges(result.sheetName, result);
      if (state.unlocked) await saveArchive();
      $('matchCount').textContent = state.scanned.length + ' 行命中' + (changed ? ' · 记录 ' + changed + ' 处改动' : '');
      renderHistory();
    }
    $('docName').textContent = result.sheetName || '当前工作簿';
  }

  function showTab(name) {
    var buttons = $('tabs').querySelectorAll('button');
    for (var i = 0; i < buttons.length; i += 1) buttons[i].className = buttons[i].dataset.tab === name ? 'on' : '';
    ['matches', 'translate', 'terms', 'history', 'sync', 'settings'].forEach(function (tab) {
      $('panel-' + tab).hidden = tab !== name;
    });
    state.lastTab = name;
    writeJSON(K.lastTab, name);
    if (name === 'translate') { loadEntries(); renderEntry(); }
    if (name === 'history') { renderLock(); renderHistory(); }
    if (name === 'terms') renderTerms();
    if (name === 'sync') refreshHelper();
  }

  function refreshHelper() {
    fetch(HELPER + '/status').then(function (response) {
      return response.json();
    }).then(function (data) {
      state.helper = data;
      $('syncStatus').textContent = '助手在线 · 口令' + (data.password_set ? '已设' : '未设') +
        ' · 令牌' + (data.token_set ? '已备' : '缺失');
      $('btnPush').disabled = !data.password_set;
    }).catch(function () {
      state.helper = { ok: false };
      $('syncStatus').textContent = '助手未运行。在 sync/ 目录执行 python3 helper.py serve。';
      $('btnPush').disabled = true;
    });
  }

  async function pushTerms() {
    var password = $('syncPass').value;
    if (!password) { $('syncStatus').textContent = '请先输入口令。'; return; }
    if (!state.terms.length) { $('syncStatus').textContent = '没有可上传的术语表。'; return; }
    $('syncStatus').textContent = '上传中…';
    try {
      var response = await fetch(HELPER + '/push', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'X-Tide-Password': password },
        body: JSON.stringify({
          terms: state.terms,
          repo: $('syncRepo').value.trim(),
          branch: $('syncBranch').value.trim() || 'main'
        })
      });
      var data = await response.json();
      $('syncStatus').textContent = response.ok
        ? '已上传 ' + data.count + ' 条术语，提交 ' + String(data.commit || '').slice(0, 8)
        : '失败：' + (data.error || response.status);
    } catch (err) {
      $('syncStatus').textContent = '无法连接本机助手：' + err.message;
    }
  }

  function bind() {
    $('tabs').addEventListener('click', function (event) {
      if (event.target && event.target.dataset && event.target.dataset.tab) showTab(event.target.dataset.tab);
    });
    $('btnScan').addEventListener('click', scan);
    $('btnSave').addEventListener('click', function () {
      state.settings = {
        sourceCol: $('colSource').value.toUpperCase() || 'C',
        targetCol: $('colTarget').value.toUpperCase() || 'F',
        uidCol: $('colUid').value.toUpperCase() || 'A'
      };
      writeJSON(K.settings, state.settings);
      writeJSON(K.repo, $('syncRepo').value);
      writeJSON(K.branch, $('syncBranch').value);
      $('settingsStatus').textContent = '已保存到本机。';
    });
    $('termFile').addEventListener('change', function (event) {
      var file = event.target.files && event.target.files[0];
      if (!file) return;
      var reader = new FileReader();
      reader.onload = function () {
        try {
          setTerms(window.TideEngine.parseTermbase(JSON.parse(String(reader.result))));
          writeJSON(K.terms, state.terms);
          $('settingsStatus').textContent = '已导入 ' + state.terms.length + ' 条术语。';
          renderTerms();
        } catch (err) {
          $('settingsStatus').textContent = '术语表解析失败：' + err.message;
        }
      };
      reader.readAsText(file, 'utf-8');
    });
    $('btnUnlock').addEventListener('click', async function () {
      try {
        await unlockArchive($('passInput').value);
        $('passInput').value = '';
        renderLock(); renderHistory();
      } catch (err) {
        $('lockStatus').textContent = '解锁失败：' + err.message;
      }
    });
    $('btnSetPass').addEventListener('click', async function () {
      var value = $('passInput').value;
      if (value.length < 8) { $('lockStatus').textContent = '口令至少 8 位。'; return; }
      await createArchive(value);
      $('passInput').value = '';
      renderLock(); renderHistory();
    });
    $('btnExport').addEventListener('click', function () {
      var raw = window.localStorage.getItem(K.archive);
      if (!raw) { $('syncStatus').textContent = '没有可导出的归档。'; return; }
      var blob = new Blob([raw], { type: 'application/json' });
      var link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = 'tide-archive.enc.json';
      link.click();
      URL.revokeObjectURL(link.href);
    });
    $('btnClear').addEventListener('click', async function () {
      if (!state.unlocked) { $('lockStatus').textContent = '请先解锁再清空。'; return; }
      if (!window.confirm('清空本机加密归档？此操作不可撤销。')) return;
      state.archive = { history: [], snapshot: {} };
      await saveArchive();
      renderHistory();
    });
    $('btnResetTerms').addEventListener('click', function () {
      if (!state.bundledTerms.length) {
        $('settingsStatus').textContent = '没有可恢复的内置术语表。';
        return;
      }
      setTerms(state.bundledTerms.slice());
      window.localStorage.removeItem(K.terms);
      $('settingsStatus').textContent = '已恢复内置术语表：' + state.terms.length + ' 条。';
      renderTermSource();
      renderTerms();
    });
    $('btnExportTerms').addEventListener('click', function () {
      if (!state.terms.length) {
        $('settingsStatus').textContent = '当前没有可导出的术语。';
        return;
      }
      var blob = new Blob([JSON.stringify({ meta: state.bundledMeta, terms: state.terms }, null, 1)], { type: 'application/json' });
      var link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = 'tide-terms.json';
      link.click();
      URL.revokeObjectURL(link.href);
    });
    $('btnPrevEntry').addEventListener('click', function () { move(-1); });
    $('btnNextEntry').addEventListener('click', function () { move(1); });
    $('btnWriteNext').addEventListener('click', writeAndNext);
    $('btnLocate').addEventListener('click', locateTarget);
    $('btnUndoWrite').addEventListener('click', undoSelf);
    $('btnDelivery').addEventListener('click', deliveryCheck);
    $('entryDraft').addEventListener('input', function () {
      var entry = currentEntry();
      if (!entry) return;
      $('entryCheck').textContent = describeEntry(entry, $('entryDraft').value);
    });
    $('btnUnlockTerms').addEventListener('click', function () {
      var pass = $('termsPass').value;
      if (!pass) { $('settingsStatus').textContent = '请输入术语表口令。'; return; }
      decryptTerms(pass).then(function () {
        $('settingsStatus').textContent = '术语表已解锁：' + state.terms.length + ' 条。';
      }).catch(function () {
        $('settingsStatus').textContent = '口令错误，或加密包已损坏。';
      });
    });
    $('btnPush').addEventListener('click', pushTerms);
    $('btnPull').addEventListener('click', refreshHelper);
  }

  function loadSettings() {
    var saved = readJSON(K.settings, null);
    if (saved) state.settings = saved;
    setTerms(readJSON(K.terms, []));
    $('colSource').value = state.settings.sourceCol;
    $('colTarget').value = state.settings.targetCol;
    $('colUid').value = state.settings.uidCol;
    $('syncRepo').value = readJSON(K.repo, '') || '';
    $('syncBranch').value = readJSON(K.branch, 'main') || 'main';
    $('envHint').textContent = sheetApp() ? '已连接 WPS 表格。' : '当前不在 WPS 中，仅能查看设置与记录。';
    $('docName').textContent = sheetApp() && currentSheet() ? currentSheet().Name : '未连接 WPS';
  }

  function watchRibbon() {
    var app = sheetApp();
    if (!app || !app.PluginStorage) return;
    var last = app.PluginStorage.getItem('tide_scan_request');
    window.setInterval(function () {
      var current = app.PluginStorage.getItem('tide_scan_request');
      if (current && current !== last) {
        last = current;
        showTab('matches');
        scan();
      }
    }, 1200);
  }

  function boot() {
    bind();
    loadSettings();
    var app = sheetApp();
    var requested = app && app.PluginStorage ? app.PluginStorage.getItem('tide_open_tab') : null;
    showTab(requested || readJSON(K.lastTab, 'matches') || 'matches');
    renderLock();
    renderTermSource();
    loadBundled().then(function () { renderTermSource(); });
    renderHistory();
    watchRibbon();
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
})();
