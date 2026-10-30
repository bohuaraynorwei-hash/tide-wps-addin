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
    termStore: 'tide.termbase.v1',
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
    settings: { sourceCol: 'C', targetCol: 'F', oldCol: 'E', uidCol: 'A' },
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
    termsMeta: null,
    entries: [],
    findResults: [],
    tm: [],
    qaFindings: [],
    gridRows: [],
    projectScan: [],
    project: null,
    projectPath: '',
    selectedSegment: '',
    batchPreview: [],
    lastBatch: [],
    sourceDirty: false,
    memory: [],
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

  // The pane knows the build it was opened with (?b=…); build.json on the server
  // says what is deployed. A mismatch is offered as a one-click reload.
  var SELF_BUILD = (function () {
    var m = /[?&]b=([^&]+)/.exec(window.location.search);
    return m ? decodeURIComponent(m[1]) : '';
  })();

  function checkForUpdate() {
    if (window.location.protocol === 'file:') return;
    fetch('../build.json', { cache: 'no-store' }).then(function (response) {
      return response.ok ? response.json() : null;
    }).then(function (data) {
      // No stamp (pane opened by an older ribbon) is treated as stale, so the
      // update button still appears at least once.
      if (!data || !data.build || data.build === SELF_BUILD) return;
      var bar = $('updateBar');
      bar.hidden = false;
      bar.onclick = function () {
        var url = new URL(window.location.href);
        url.searchParams.set('b', data.build);
        window.location.replace(url.toString());
      };
    }).catch(function () { /* offline: keep the current pane */ });
  }

  function sheetApp() {
    return (typeof window.Application !== 'undefined' && window.Application) ? window.Application : null;
  }

  function currentSheet() {
    var app = sheetApp();
    return app && app.ActiveSheet ? app.ActiveSheet : null;
  }

  function sheetByName(name) {
    var app = sheetApp();
    if (!app || !name) return currentSheet();
    var book = app.ActiveWorkbook;
    var sheets = book && book.Worksheets;
    if (!sheets) return currentSheet();
    try {
      var direct = sheets.Item(name);
      if (direct) return direct;
    } catch (err) { /* some hosts only accept an index */ }
    try {
      var count = sheets.Count;
      for (var i = 1; i <= count; i += 1) {
        var candidate = sheets.Item(i);
        if (candidate && candidate.Name === name) return candidate;
      }
    } catch (err2) { /* fall through */ }
    return currentSheet();
  }

  // Reveal a cell reliably: Application.Goto handles activation + scroll in one
  // call; otherwise activate the sheet, select, then scroll the window.
  function revealCell(sheet, address) {
    if (!sheet) return false;
    var app = sheetApp();
    try {
      var cell = sheet.Range(address);
      if (app && typeof app.Goto === 'function') {
        app.Goto(cell, true);
        return true;
      }
    } catch (err) { /* fall back below */ }
    try {
      sheet.Activate();
    } catch (err) { /* already active */ }
    try {
      if (app && app.ActiveWindow && typeof app.ActiveWindow.Activate === 'function') {
        app.ActiveWindow.Activate();
      }
    } catch (err) { /* window may already be frontmost */ }
    try {
      if (app) app.ActiveSheet = sheet;
    } catch (err) { /* read-only in some hosts */ }
    try {
      var target = sheet.Range(address);
      // Activate sets the active cell (the visible cursor); Select sets the
      // highlighted block. Some hosts need Activate before Select takes effect.
      if (typeof target.Activate === 'function') target.Activate();
      target.Select();
      if (app && app.ActiveWindow && typeof app.ActiveWindow.ScrollIntoView === 'function') {
        app.ActiveWindow.ScrollIntoView(target, true);
      }
      return true;
    } catch (err) {
      return false;
    }
  }

  function activateSheet(sheet) {
    try { if (sheet && sheet.Activate) sheet.Activate(); } catch (err) { /* already active */ }
    return sheet;
  }

  function entrySheet(entry) {
    return sheetByName(entry && entry.sheet ? entry.sheet : state.sheetName);
  }

  function numberToCol(value) {
    var n = Number(value) || 0;
    if (n <= 0) return '';
    var text = '';
    while (n > 0) {
      var rem = (n - 1) % 26;
      text = String.fromCharCode(65 + rem) + text;
      n = Math.floor((n - 1) / 26);
    }
    return text;
  }

  // Column letter of whatever cell is selected in WPS right now.
  function selectedColumnLetter() {
    var app = sheetApp();
    if (!app || !app.Selection) return '';
    try {
      var column = Number(app.Selection.Column);
      if (column > 0) return numberToCol(column);
      var address = String(app.Selection.Address || '');
      var match = /\$?([A-Z]{1,3})\$?\d+/.exec(address);
      if (match) return match[1];
    } catch (err) { /* no selection */ }
    return '';
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
    // The used range may start on a blank row; the header is the first row with
    // a non-empty source cell, and data starts after it.
    var headerAt = 0;
    for (var h = 0; h < grid.length; h += 1) {
      if (cellText(grid, h, srcIdx).trim()) { headerAt = h; break; }
    }
    var rows = [];
    for (var r = headerAt + 1; r < grid.length; r += 1) {
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
      tr.dataset.row = String(row.row);
      tr.appendChild(textNode('td', String(row.row), 'num'));
      var tdSource = document.createElement('td');
      tdSource.appendChild(highlighted(row.source, row.matches));
      tr.appendChild(tdSource);
      var tdPair = document.createElement('td');
      row.matches.forEach(function (match) {
        tdPair.appendChild(textNode('div', (match.zh || '') + ' → ' + (match.en || '—'), 'pair'));
      });
      tr.appendChild(tdPair);
      var tdCopy = document.createElement('td');
      var copyBtn = document.createElement('button');
      copyBtn.className = 'act ghost';
      copyBtn.textContent = '复制';
      copyBtn.addEventListener('click', function () {
        var pairs = row.matches.map(function (m) { return m.zh + ' → ' + m.en; }).join('\n');
        copyText(pairs).then(function (ok) {
          copyBtn.textContent = ok ? '已复制' : '失败';
          window.setTimeout(function () { copyBtn.textContent = '复制'; }, 1500);
        });
      });
      tdCopy.appendChild(copyBtn);
      tr.appendChild(tdCopy);
      tr.appendChild(textNode('td', row.target || '—', 'hint'));
      body.appendChild(tr);
    });
  }

  function currentTermQuery() {
    var el = $('termSearch');
    return el ? String(el.value || '').trim() : '';
  }

  // Substring first; when that finds nothing, fall back to similarity so a
  // partial or mistyped word still surfaces something.
  function filterTerms(query) {
    var list = state.terms || [];
    var q = query.toLowerCase();
    if (!q) return { rows: list, fuzzy: false };
    var direct = list.filter(function (t) {
      return String(t.zh || '').toLowerCase().indexOf(q) > -1 ||
             String(t.en || '').toLowerCase().indexOf(q) > -1;
    });
    if (direct.length) return { rows: direct, fuzzy: false };
    var scored = [];
    list.forEach(function (t) {
      var score = Math.max(
        window.TideEngine.similarity(query, t.zh || ''),
        window.TideEngine.similarity(query, t.en || '')
      );
      if (score >= 0.34) scored.push({ term: t, score: score });
    });
    scored.sort(function (a, b) { return b.score - a.score; });
    return { rows: scored.slice(0, 60).map(function (x) { return x.term; }), fuzzy: true };
  }

  // Re-read the sheet with whatever column letters are in the inputs right now,
  // so a mapping change is visible without visiting 设置 → 保存 first.
  function syncTermColumns() {
    if ($('termColSource')) $('termColSource').value = state.settings.sourceCol || 'C';
    if ($('termColTarget')) $('termColTarget').value = state.settings.targetCol || 'F';
    if ($('termColUid')) $('termColUid').value = state.settings.uidCol || 'A';
  }

  function readTermColumns() {
    var src = (($('termColSource') && $('termColSource').value) || state.settings.sourceCol || 'C').toUpperCase();
    var tgt = (($('termColTarget') && $('termColTarget').value) || state.settings.targetCol || 'F').toUpperCase();
    var uid = (($('termColUid') && $('termColUid').value) || state.settings.uidCol || 'A').toUpperCase();
    state.settings.sourceCol = src;
    state.settings.targetCol = tgt;
    state.settings.uidCol = uid;
    writeJSON(K.settings, state.settings);
    if ($('colSource')) $('colSource').value = src;
    if ($('colTarget')) $('colTarget').value = tgt;
    if ($('colUid')) $('colUid').value = uid;
    return { src: src, tgt: tgt };
  }

  function refreshTermsFromSheet() {
    var cols = readTermColumns();
    var src = cols.src;
    var tgt = cols.tgt;

    var result = buildRows();
    if (result.error) {
      $('termRefreshStatus').textContent = '未连接 WPS，无法刷新。';
      renderTerms();
      return;
    }
    state.scanned = window.TideEngine.scanRows(result.rows, state.matchTerms);
    var hitRows = state.scanned.length;
    var hitTerms = 0;
    state.scanned.forEach(function (row) { hitTerms += row.matches.length; });
    renderTerms();
    $('termRefreshStatus').textContent = '已刷新「' + (result.sheetName || '当前表') + '」：' +
      result.rows.length + ' 条原文，' + hitRows + ' 行命中（' + src + '→' + tgt + '），共 ' + hitTerms + ' 个术语命中。';
  }

  function renderTerms() {
    var table = $('termTable');
    if (!table) return;
    var body = table.tBodies[0];
    body.textContent = '';
    var total = (state.terms || []).length;
    if (!total) {
      $('termCount').textContent = '术语表未导入';
      $('termEmpty').hidden = false;
      table.hidden = true;
      return;
    }
    var query = currentTermQuery();
    var result = filterTerms(query);
    var counts = {};
    (state.scanned || []).forEach(function (row) {
      row.matches.forEach(function (m) {
        var key = m.zh + '\u0000' + m.en;
        counts[key] = (counts[key] || 0) + 1;
      });
    });
    var rows = result.rows.slice(0, 300);
    $('termCount').textContent = query
      ? rows.length + ' / ' + total + ' 条' + (result.fuzzy ? ' · 近似匹配' : '')
      : total + ' 条术语';
    table.hidden = rows.length === 0;
    $('termEmpty').hidden = rows.length > 0;
    $('termEmpty').textContent = query ? '没有匹配的术语，换个词试试。' : '术语表为空。';
    rows.forEach(function (term) {
      var tr = document.createElement('tr');
      tr.appendChild(textNode('td', term.zh || ''));
      tr.appendChild(textNode('td', term.en || '—'));
      var hits = counts[(term.zh || '') + '\u0000' + (term.en || '')] || 0;
      var hitCell = textNode('td', hits ? String(hits) : '—', 'num');
      if (term.status === 'needs_review') {
        hitCell = textNode('td', '待确认', 'num');
        hitCell.style.color = 'var(--orange)';
      }
      tr.appendChild(hitCell);
      var td = document.createElement('td');
      var use = document.createElement('button');
      use.className = 'act ghost';
      use.textContent = '采用';
      use.addEventListener('click', function () {
        $('entryDraft').value = term.en || '';
        $('entryDraft').dispatchEvent(new Event('input'));
        showTab('translate');
        $('entryCheck').textContent = '已采用术语：' + term.zh + ' → ' + (term.en || '');
      });
      td.appendChild(use);
      tr.appendChild(td);
      body.appendChild(tr);
    });
  }

  function persistTerms(meta) {
    var payload = { meta: meta || null, terms: state.terms, savedAt: new Date().toISOString() };
    try {
      window.localStorage.setItem(K.termStore, JSON.stringify(payload));
      return true;
    } catch (err) {
      return false;
    }
  }

  function loadSavedTerms() {
    try {
      var raw = window.localStorage.getItem(K.termStore);
      if (raw) {
        var saved = JSON.parse(raw);
        if (saved && Array.isArray(saved.terms) && saved.terms.length) {
          setTerms(window.TideEngine.parseTermbase(saved.terms), saved.meta || { count: saved.terms.length, source: '本机保存' });
          return true;
        }
      }
      // migrate a termbase imported by an older build
      var legacy = window.localStorage.getItem(K.terms);
      if (legacy) {
        var list = window.TideEngine.parseTermbase(JSON.parse(legacy));
        if (list.length) {
          setTerms(list, { count: list.length, source: '旧版本导入' });
          persistTerms(state.termsMeta);
          window.localStorage.removeItem(K.terms);
          return true;
        }
      }
    } catch (err) { /* corrupted store: fall back to the bundled termbase */ }
    return false;
  }

  // Entries flagged needs_review carry notes or several variants in the English
  // column; they would produce false matches, so they stay visible but unused.
  function setTerms(list, meta) {
    state.terms = list || [];
    state.termsMeta = meta || null;
    state.matchTerms = state.terms.filter(function (term) {
      return term.status !== 'needs_review';
    });
  }

  function localStamp(iso) {
    var d = new Date(iso);
    if (isNaN(d.getTime())) return '';
    function two(n) { return (n < 10 ? '0' : '') + n; }
    return d.getFullYear() + '-' + two(d.getMonth() + 1) + '-' + two(d.getDate()) +
      ' ' + two(d.getHours()) + ':' + two(d.getMinutes());
  }

  function renderTermSource() {
    var el = $('termSource');
    if (!el) return;
    if (state.termsMeta) {
      var meta = state.termsMeta;
      var held = state.terms.length - state.matchTerms.length;
      el.textContent = '已保存导入：' + meta.count + ' 条 · ' + (meta.source || '本机') +
        (meta.importedAt ? ' · ' + localStamp(meta.importedAt) : '') +
        ' · 仅存本机' + (held ? ' · ' + held + ' 条待确认' : '');
      return;
    }
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
    // file:// preview (design lint, local double-click) cannot fetch; stay quiet.
    if (window.location.protocol === 'file:') return Promise.resolve(null);
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

  function activeSheetName() {
    var sheet = currentSheet();
    return sheet && sheet.Name ? String(sheet.Name) : '';
  }

  // WPS switched the visible worksheet: the batch must follow it, otherwise
  // every locate and write points at the previous sheet.
  function syncActiveSheet() {
    var name = activeSheetName();
    if (!name || name === state.sheetName) return false;
    var draft = $('entryDraft') ? $('entryDraft').value : '';
    var entry = currentEntry();
    var dirty = state.sourceDirty || (entry && draft !== (entry.target || ''));
    if (dirty) {
      // Do not overwrite the result line the user is reading; flag it in the
      // position line instead and wait for an explicit refresh.
      state.sheetMismatch = name;
      renderPosition();
      return false;
    }
    state.sheetMismatch = '';
    loadEntries();
    renderEntry();
    $('entryCheck').textContent = '工作表已切换到 ' + name + '，条目已更新';
    return true;
  }

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
    state.sheetName = result.sheetName;
    state.entries.forEach(function (entry) { entry.sheet = result.sheetName; });
    state.cursor = firstEmpty;
    refreshMemory();
    return state.entries;
  }

  function currentEntry() {
    return state.entries[state.cursor] || null;
  }

  function entryAddress(row) {
    return state.settings.targetCol + row;
  }

  // EN (old) holds exactly the previous translation; the current EN is the
  // only official text. Keeping the archive in the sheet lets anyone roll back.
  function archivePrevious(sheet, row, before) {
    var column = state.settings.oldCol;
    if (!column || !sheet) return;
    var previous = String(before == null ? '' : before);
    if (!previous.trim()) return;
    try { sheet.Range(column + row).Value2 = previous; } catch (err) { /* read-only cell */ }
  }

  function rollbackEntry() {
    var entry = currentEntry();
    if (!entry) return;
    var sheet = entrySheet(entry);
    var oldAddress = state.settings.oldCol + entry.row;
    var archived = '';
    try { archived = String(sheet.Range(oldAddress).Value2 == null ? '' : sheet.Range(oldAddress).Value2); }
    catch (err) { archived = ''; }
    if (!archived.trim()) { $('entryCheck').textContent = '旧译文列是空的，没有可回退的版本。'; return; }
    var current = entry.target;
    try {
      sheet.Range(entryAddress(entry.row)).Value2 = archived;
      sheet.Range(oldAddress).Value2 = current;   // swap, so pressing again undoes
    } catch (err) {
      $('entryCheck').textContent = '回退失败：' + err.message;
      return;
    }
    recordRevision(entry.sheet, entry.row, entry.source, current, archived);
    entry.target = archived;
    renderEntry();
    $('entryCheck').textContent = '已回退到旧译文（' + oldAddress + '），再次点击可换回。';
  }

  function writeCell(entry, value) {
    var sheet = entrySheet(entry);
    if (!sheet) return false;
    var cell = sheet.Range(entryAddress(entry.row));
    var before = cell.Value2;
    cell.Value2 = value;
    state.lastWrite = {
      row: entry.row, sheet: entry.sheet, before: before, after: value, address: entryAddress(entry.row)
    };
    return true;
  }

  function recordWrite(entry, before, after) {
    if (!state.unlocked) return;
    state.archive.history.push({
      ts: new Date().toISOString(),
      sheet: entry.sheet || (currentSheet() ? currentSheet().Name : ''),
      cell: entryAddress(entry.row),
      before: before == null ? '' : String(before),
      after: String(after)
    });
    saveArchive();
  }

  function refreshMemory() {
    state.memory = window.TideEngine.buildMemory(state.entries);
  }

  function applyMemory(target) {
    $('entryDraft').value = target;
    $('entryDraft').dispatchEvent(new Event('input'));
    $('entryDraft').focus();
  }

  function renderMemory(entry, sourceText) {
    var table = $('memoryTable');
    var body = table.tBodies[0];
    body.textContent = '';
    var count = $('memoryCount');
    if (!entry || !state.memory.length) {
      table.hidden = true;
      $('memoryEmpty').hidden = false;
      $('memoryEmpty').textContent = state.memory.length ? '当前条目没有匹配。' : '当前表还没有可用的已确认译文。';
      count.textContent = state.memory.length ? state.memory.length + ' 条记忆' : '';
      return;
    }
    var result = window.TideEngine.findMemory(sourceText || entry.source, state.memory, 0.5, 5);
    var rows = result.exact.map(function (item) { return { item: item, exact: true, score: 1 }; })
      .concat(result.fuzzy.map(function (item) { return { item: item, exact: false, score: item.score }; }));
    count.textContent = rows.length ? rows.length + ' 条匹配 / ' + state.memory.length + ' 条记忆' : '';
    table.hidden = rows.length === 0;
    $('memoryEmpty').hidden = rows.length > 0;
    $('memoryEmpty').textContent = '当前条目没有匹配。';
    rows.forEach(function (row) {
      var tr = document.createElement('tr');
      tr.appendChild(textNode('td', row.exact ? '精确' : Math.round(row.score * 100) + '%', 'num'));
      tr.appendChild(textNode('td', row.item.source.slice(0, 40), 'hint'));
      tr.appendChild(textNode('td', row.item.target.slice(0, 40)));
      var tdUse = document.createElement('td');
      var use = document.createElement('button');
      use.className = 'act ghost';
      use.textContent = '采用';
      use.addEventListener('click', function () { applyMemory(row.item.target); });
      tdUse.appendChild(use);
      tr.appendChild(tdUse);
      body.appendChild(tr);
    });
  }

  // ---------- workbook-wide substring search ----------

  function scanSheetFor(sheet, needle) {
    var hits = [];
    var used;
    try { used = sheet.UsedRange; } catch (err) { return hits; }
    if (!used) return hits;
    var grid;
    try { grid = toGrid(used.Value2); } catch (err) { return hits; }
    var startRow = Number(used.Row) || 1;
    var startCol = Number(used.Column) || 1;
    var srcIdx = colToNumber(state.settings.sourceCol) - startCol;
    var tgtIdx = colToNumber(state.settings.targetCol) - startCol;
    for (var i = 0; i < grid.length; i += 1) {
      var source = srcIdx >= 0 ? cellText(grid, i, srcIdx) : '';
      if (source && source.toLowerCase().indexOf(needle) > -1) {
        hits.push({ sheet: sheet.Name, row: startRow + i, column: 'source', text: source });
      }
      var target = tgtIdx >= 0 ? cellText(grid, i, tgtIdx) : '';
      if (target && target.toLowerCase().indexOf(needle) > -1) {
        hits.push({ sheet: sheet.Name, row: startRow + i, column: 'target', text: target });
      }
    }
    return hits;
  }

  function renderFind() {
    var table = $('findTable');
    var body = table.tBodies[0];
    body.textContent = '';
    var results = state.findResults || [];
    $('findCount').textContent = results.length ? results.length + ' 处命中' : '';
    table.hidden = results.length === 0;
    $('findEmpty').hidden = results.length > 0;
    results.forEach(function (hit) {
      var tr = document.createElement('tr');
      tr.appendChild(textNode('td', hit.sheet, 'num'));
      tr.appendChild(textNode('td', (hit.column === 'target' ? state.settings.targetCol : state.settings.sourceCol) + hit.row, 'num'));
      tr.appendChild(textNode('td', hit.column === 'target' ? '译文' : '原文', 'num'));
      tr.appendChild(textNode('td', hit.text.slice(0, 60), 'hint'));
      tr.addEventListener('click', function () {
        var sheet = sheetByName(hit.sheet);
        var column = hit.column === 'target' ? state.settings.targetCol : state.settings.sourceCol;
        var ok = revealCell(sheet, column + hit.row);
        $('findStatus').textContent = ok
          ? '已跳转 ' + hit.sheet + ' ' + column + hit.row
          : '跳转失败：' + hit.sheet + ' ' + column + hit.row;
      });
      body.appendChild(tr);
    });
  }

  // Scans one sheet per tick so a 93-sheet workbook does not freeze the pane.
  function runFind() {
    var term = $('findInput').value.trim();
    if (!term) { $('findStatus').textContent = '先输入要搜索的片段。'; return; }
    var app = sheetApp();
    var book = app && app.ActiveWorkbook;
    if (!book || !book.Worksheets) { $('findStatus').textContent = '未连接工作簿。'; return; }
    var sheets = book.Worksheets;
    var count = 0;
    try { count = sheets.Count; } catch (err) { count = 0; }
    var needle = term.toLowerCase();
    var results = [];
    var index = 1;
    state.findResults = [];
    $('findStatus').textContent = '搜索中…';
    $('btnFindAll').disabled = true;

    function step() {
      if (index > count || results.length >= 200) {
        state.findResults = results;
        renderFind();
        $('btnFindAll').disabled = false;
        $('findStatus').textContent = results.length >= 200
          ? '已到 200 条上限，请把片段写长一点。'
          : '搜索完成：' + results.length + ' 处。';
        return;
      }
      var sheet;
      try { sheet = sheets.Item(index); } catch (err) { sheet = null; }
      index += 1;
      if (sheet) {
        try { results = results.concat(scanSheetFor(sheet, needle)); } catch (err) { /* skip sheet */ }
      }
      window.setTimeout(step, 0);
    }
    step();
  }

  // ---------- batch find & replace ----------

  function replaceAllIn(text, find, repl, caseSensitive) {
    if (!find) return text;
    if (caseSensitive) return text.split(find).join(repl);
    var escaped = find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    return text.replace(new RegExp(escaped, 'gi'), repl);
  }

  function batchColumns() {
    var mode = $('repColumn').value;
    var cols = [];
    if (mode === 'source' || mode === 'both') cols.push({ kind: 'source', letter: state.settings.sourceCol });
    if (mode === 'target' || mode === 'both') cols.push({ kind: 'target', letter: state.settings.targetCol });
    return cols;
  }

  function scanSheetForReplace(sheet, find, repl, caseSensitive, cols) {
    var hits = [];
    var used;
    try { used = sheet.UsedRange; } catch (err) { return hits; }
    if (!used) return hits;
    var grid;
    try { grid = toGrid(used.Value2); } catch (err) { return hits; }
    var startRow = Number(used.Row) || 1;
    var startCol = Number(used.Column) || 1;
    for (var c = 0; c < cols.length; c += 1) {
      var idx = colToNumber(cols[c].letter) - startCol;
      if (idx < 0) continue;
      for (var i = 0; i < grid.length; i += 1) {
        var before = cellText(grid, i, idx);
        if (!before) continue;
        var after = replaceAllIn(before, find, repl, caseSensitive);
        if (after === before) continue;
        hits.push({
          sheet: sheet.Name, row: startRow + i, column: cols[c].kind,
          letter: cols[c].letter, before: before, after: after
        });
      }
    }
    return hits;
  }

  function runBatchScan(find, repl, caseSensitive, scope, onDone) {
    var app = sheetApp();
    var book = app && app.ActiveWorkbook;
    if (!book || !book.Worksheets) { onDone([]); return; }
    var sheets = book.Worksheets;
    var count = 0;
    try { count = sheets.Count; } catch (err) { count = 0; }
    var current = currentSheet();
    var cols = batchColumns();
    var hits = [];
    var index = 1;

    function step() {
      if (index > count || hits.length >= 500) { onDone(hits); return; }
      var sheet = null;
      try { sheet = sheets.Item(index); } catch (err) { sheet = null; }
      index += 1;
      var skip = scope === 'sheet' && current && (!sheet || sheet.Name !== current.Name);
      if (sheet && !skip) {
        try { hits = hits.concat(scanSheetForReplace(sheet, find, repl, caseSensitive, cols)); } catch (err) { /* skip */ }
      }
      if (index % 5 === 0) $('repStatus').textContent = '扫描中 ' + index + '/' + count;
      window.setTimeout(step, 0);
    }
    step();
  }

  function renderBatch() {
    var table = $('repTable');
    var body = table.tBodies[0];
    body.textContent = '';
    var hits = state.batchPreview || [];
    $('repCount').textContent = hits.length ? hits.length + ' 处待替换' : '';
    table.hidden = hits.length === 0;
    hits.slice(0, 200).forEach(function (hit) {
      var tr = document.createElement('tr');
      tr.appendChild(textNode('td', hit.sheet, 'num'));
      tr.appendChild(textNode('td', hit.letter + hit.row, 'num'));
      tr.appendChild(textNode('td', hit.before.slice(0, 34), 'hint'));
      tr.appendChild(textNode('td', hit.after.slice(0, 34)));
      body.appendChild(tr);
    });
    $('btnRepApply').disabled = hits.length === 0;
  }

  function previewBatch() {
    var find = $('repFind').value;
    if (!find) { $('repStatus').textContent = '先填写查找内容。'; return; }
    var repl = $('repWith').value;
    $('repStatus').textContent = '扫描中…';
    $('btnRepPreview').disabled = true;
    runBatchScan(find, repl, $('repCase').checked, $('repScope').value, function (hits) {
      state.batchPreview = hits;
      renderBatch();
      $('btnRepPreview').disabled = false;
      $('repStatus').textContent = hits.length
        ? '预览完成：' + hits.length + ' 处。确认后点「执行替换」。'
        : '没有匹配的单元格。';
    });
  }

  function applyBatch() {
    var hits = state.batchPreview || [];
    if (!hits.length) return;
    if (!window.confirm('将修改 ' + hits.length + ' 个单元格，确定执行？')) return;
    var sheetCache = {};
    var applied = [];
    var stamp = new Date().toISOString();
    hits.forEach(function (hit) {
      var sheet = sheetCache[hit.sheet] || (sheetCache[hit.sheet] = sheetByName(hit.sheet));
      if (!sheet) return;
      try {
        var cell = sheet.Range(hit.letter + hit.row);
        cell.Value2 = hit.after;
        applied.push(hit);
      } catch (err) { /* skip unreadable cell */ }
    });
    state.lastBatch = applied;
    if (state.unlocked) {
      applied.forEach(function (hit) {
        state.archive.history.push({
          ts: stamp, sheet: hit.sheet, cell: hit.letter + hit.row,
          before: hit.before, after: hit.after
        });
      });
      saveArchive();
    }
    state.batchPreview = [];
    renderBatch();
    $('btnRepUndo').disabled = applied.length === 0;
    $('repStatus').textContent = '已替换 ' + applied.length + ' 个单元格' + (state.unlocked ? '，已记入归档。' : '（归档未解锁，未记录）。');
    loadEntries();
    renderEntry();
  }

  function undoBatch() {
    var batch = state.lastBatch || [];
    if (!batch.length) return;
    if (!window.confirm('撤销上次批量替换（' + batch.length + ' 个单元格）？')) return;
    var sheetCache = {};
    var restored = 0;
    batch.forEach(function (hit) {
      var sheet = sheetCache[hit.sheet] || (sheetCache[hit.sheet] = sheetByName(hit.sheet));
      if (!sheet) return;
      try { sheet.Range(hit.letter + hit.row).Value2 = hit.before; restored += 1; } catch (err) { /* skip */ }
    });
    state.lastBatch = [];
    $('btnRepUndo').disabled = true;
    $('repStatus').textContent = '已撤销 ' + restored + ' 个单元格。';
    loadEntries();
    renderEntry();
  }

  // ---------- review layer: revisions + comments ----------

  function authorName() {
    return window.localStorage.getItem('tide.author') || '本机';
  }

  function segKey(sheet, row) { return String(sheet) + '!' + row; }

  function ensureProject() {
    if (!state.project) state.project = { version: 1, segments: {} };
    if (!state.project.segments) state.project.segments = {};
    return state.project;
  }

  function segmentRecord(sheet, row, sourceText, targetText) {
    var project = ensureProject();
    var key = segKey(sheet, row);
    var record = project.segments[key];
    if (!record) {
      record = project.segments[key] = { source: sourceText || '', target: targetText || '', revisions: [], comments: [], status: '' };
    } else {
      if (sourceText) record.source = sourceText;
      if (targetText !== undefined && targetText !== null) record.target = targetText;
    }
    return record;
  }

  // Keys are "sheet!row", which collide across workbooks; scope storage by
  // workbook name so one project's revisions never appear in another.
  function reviewStoreKey() {
    var book = sheetApp() && sheetApp().ActiveWorkbook;
    var name = book && (book.FullName || book.Name) ? String(book.FullName || book.Name) : 'default';
    return 'tide.review:' + name;
  }

  function saveProject() {
    var project = ensureProject();
    if (window.TideProject) { window.TideProject.save(project); return; }
    try { window.localStorage.setItem(reviewStoreKey(), JSON.stringify(project)); } catch (err) { /* quota */ }
  }

  function recordRevision(sheet, row, sourceText, before, after) {
    if (String(before == null ? '' : before) === String(after == null ? '' : after)) return;
    var record = segmentRecord(sheet, row, sourceText, after);
    // A new edit replaces any earlier pending revision of the same segment;
    // keeping both is what made the review list look like stacked versions.
    (record.revisions || []).forEach(function (rev) {
      if (rev.state === 'pending') rev.state = 'superseded';
    });
    record.revisions.push({
      id: 'r' + Date.now() + Math.floor(Math.random() * 1000),
      ts: new Date().toISOString(),
      author: authorName(),
      before: before == null ? '' : String(before),
      after: String(after),
      state: 'pending'
    });
    record.target = String(after);
    saveProject();
  }

  // Common prefix/suffix trimming: cheaper and clearer than a full diff for
  // short UI strings, and it reads correctly for both Chinese and English.
  function diffParts(before, after) {
    var a = String(before), b = String(after);
    var start = 0;
    while (start < a.length && start < b.length && a[start] === b[start]) start += 1;
    var endA = a.length, endB = b.length;
    while (endA > start && endB > start && a[endA - 1] === b[endB - 1]) { endA -= 1; endB -= 1; }
    return {
      prefix: a.slice(0, start),
      removed: a.slice(start, endA),
      added: b.slice(start, endB),
      suffix: a.slice(endA)
    };
  }

  function diffNode(before, after) {
    var wrap = document.createElement('span');
    var parts = diffParts(before, after);
    if (parts.prefix) wrap.appendChild(document.createTextNode(parts.prefix));
    if (parts.removed) {
      var del = document.createElement('del');
      del.textContent = parts.removed;
      wrap.appendChild(del);
    }
    if (parts.added) {
      var ins = document.createElement('ins');
      ins.textContent = parts.added;
      wrap.appendChild(ins);
    }
    if (parts.suffix) wrap.appendChild(document.createTextNode(parts.suffix));
    return wrap;
  }

  function pendingCount(record) {
    return (record.revisions || []).filter(function (r) { return r.state === 'pending'; }).length;
  }

  function openComments(record) {
    return (record.comments || []).filter(function (c) { return !c.resolved; }).length;
  }

  // Bring the stored revisions back in line with the sheet: anything whose
  // result is no longer in the cell is history, not a pending decision.
  function reconcileReview() {
    var project = ensureProject();
    var changed = false;
    Object.keys(project.segments).forEach(function (key) {
      var record = project.segments[key];
      var match = /^(.*)!(\d+)$/.exec(key);
      if (!match) return;
      var sheet = sheetByName(match[1]);
      var row = Number(match[2]);
      var current = '';
      try { current = String(sheet.Range(state.settings.targetCol + row).Value2 == null ? '' : sheet.Range(state.settings.targetCol + row).Value2); }
      catch (err) { return; }
      if (record.target !== current) { record.target = current; changed = true; }
      (record.revisions || []).forEach(function (rev) {
        if (rev.state === 'pending' && String(rev.after) !== current) {
          rev.state = 'superseded';
          changed = true;
        }
      });
    });
    if (changed) saveProject();
    return changed;
  }

  function renderReview() {
    reconcileReview();
    var project = ensureProject();
    var keys = Object.keys(project.segments);
    var table = $('reviewTable');
    var body = table.tBodies[0];
    body.textContent = '';
    var rows = keys.filter(function (key) {
      var record = project.segments[key];
      return pendingCount(record) > 0 || (record.comments || []).length > 0;
    });
    $('reviewCount').textContent = rows.length ? rows.length + ' 条有修订或评论 / ' + keys.length + ' 条记录' : '';
    table.hidden = rows.length === 0;
    $('reviewEmpty').hidden = rows.length > 0;
    rows.forEach(function (key) {
      var record = project.segments[key];
      var tr = document.createElement('tr');
      var rowNumber = key.split('!')[1];
      tr.appendChild(textNode('td', rowNumber, 'num'));
      var summary = document.createElement('td');
      summary.appendChild(textNode('div', record.source.slice(0, 30), 'zh'));
      summary.appendChild(textNode('div', (record.target || '—').slice(0, 40), 'hint'));
      tr.appendChild(summary);
      var status = document.createElement('td');
      if (pendingCount(record)) status.appendChild(textNode('span', pendingCount(record) + ' 处修订', 'badge'));
      if (openComments(record)) status.appendChild(textNode('span', openComments(record) + ' 条评论', 'badge'));
      tr.appendChild(status);
      tr.appendChild(textNode('td', key.split('!')[0], 'num'));
      tr.addEventListener('click', function () {
        state.selectedSegment = key;
        renderReviewDetail();
        // Jump to the target cell this revision belongs to.
        var parts = key.split('!');
        var targetSheet = sheetByName(parts[0]);
        var address = state.settings.targetCol + parts[1];
        var ok = revealCell(targetSheet, address);
        $('reviewStatus').textContent = ok
          ? '已定位 ' + address + '（' + parts[0] + '）'
          : '定位失败：请手动跳到 ' + address;
      });
      body.appendChild(tr);
    });
    if (state.selectedSegment && project.segments[state.selectedSegment]) renderReviewDetail();
    else $('reviewDetail').hidden = true;
  }

  function renderReviewDetail() {
    var project = ensureProject();
    var key = state.selectedSegment;
    var record = project.segments[key];
    var panel = $('reviewDetail');
    if (!record) { panel.hidden = true; return; }
    panel.hidden = false;
    $('reviewDetailTitle').textContent = key + ' · ' + (record.source || '').slice(0, 40);

    var revisions = $('revisionList');
    revisions.textContent = '';
    var list = (record.revisions || []).slice().reverse();
    if (!list.length) revisions.appendChild(textNode('div', '还没有修订。', 'hint'));
    list.forEach(function (rev) {
      var line = document.createElement('div');
      line.className = 'rev' + (rev.state === 'accepted' ? ' accepted' : '');
      line.appendChild(textNode('span', rev.author + ' · ' + rev.ts.slice(5, 16).replace('T', ' '), 'who'));
      line.appendChild(diffNode(rev.before, rev.after));
      if (rev.state === 'superseded') {
        line.appendChild(textNode('span', '已过期（单元格已是别的值）', 'badge'));
      } else if (rev.state === 'pending') {
        var accept = document.createElement('button');
        accept.className = 'act ghost';
        accept.textContent = '接受';
        accept.addEventListener('click', function () { resolveRevision(rev.id, 'accepted'); });
        var reject = document.createElement('button');
        reject.className = 'act ghost';
        reject.textContent = '拒绝';
        reject.addEventListener('click', function () { resolveRevision(rev.id, 'rejected'); });
        line.appendChild(accept);
        line.appendChild(reject);
      } else {
        line.appendChild(textNode('span', rev.state === 'accepted' ? '已接受' : '已拒绝', 'badge ' + (rev.state === 'accepted' ? 'ok' : '')));
      }
      revisions.appendChild(line);
    });

    var comments = $('commentList');
    comments.textContent = '';
    if (!(record.comments || []).length) comments.appendChild(textNode('div', '还没有评论。', 'hint'));
    (record.comments || []).forEach(function (comment) {
      var box = document.createElement('div');
      box.className = 'comment' + (comment.resolved ? ' resolved' : '');
      var meta = document.createElement('div');
      meta.className = 'meta';
      meta.appendChild(textNode('span', comment.author + ' · ' + comment.ts.slice(5, 16).replace('T', ' ')));
      box.appendChild(meta);
      box.appendChild(textNode('div', comment.body));
      if (!comment.resolved) {
        var done = document.createElement('button');
        done.className = 'act ghost';
        done.textContent = '标记解决';
        done.style.marginTop = '6px';
        done.addEventListener('click', function () {
          comment.resolved = true;
          saveProject();
          renderReviewDetail();
          renderReview();
        });
        box.appendChild(done);
      }
      comments.appendChild(box);
    });
  }

  function resolveRevision(id, next) {
    var project = ensureProject();
    var key = state.selectedSegment;
    var record = project.segments[key];
    if (!record) return;
    var rev = (record.revisions || []).find(function (r) { return r.id === id; });
    if (!rev) return;
    rev.state = next;
    if (next === 'rejected') {
      var match = /^(.*)!(\d+)$/.exec(key);
      if (match) {
        var sheet = sheetByName(match[1]);
        var row = Number(match[2]);
        try { sheet.Range(state.settings.targetCol + row).Value2 = rev.before; } catch (err) { /* cell gone */ }
      }
      record.target = rev.before;
    }
    saveProject();
    renderReviewDetail();
    renderReview();
  }

  function addComment() {
    var body = ($('commentInput').value || '').trim();
    if (!body) { $('reviewStatus').textContent = '先写点内容。'; return; }
    var project = ensureProject();
    var record = project.segments[state.selectedSegment];
    if (!record) return;
    record.comments.push({
      id: 'c' + Date.now(),
      ts: new Date().toISOString(),
      author: authorName(),
      body: body,
      resolved: false
    });
    $('commentInput').value = '';
    saveProject();
    renderReviewDetail();
    renderReview();
    $('reviewStatus').textContent = '已添加评论。';
  }

  function loadProject() {
    if (!window.TideProject) {
      try {
        var raw = window.localStorage.getItem(reviewStoreKey());
        state.project = raw ? JSON.parse(raw) : { version: 1, segments: {} };
      } catch (err) {
        state.project = { version: 1, segments: {} };
      }
      if (!state.project.segments) state.project.segments = {};
      return Promise.resolve(state.project);
    }
    return window.TideProject.load().then(function (result) {
      state.projectPath = result && result.path ? result.path : '';
      state.project = (result && result.data) || { version: 1, segments: {} };
      if (!state.project.segments) state.project.segments = {};
      return state.project;
    }).catch(function () {
      state.project = { version: 1, segments: {} };
    });
  }

  // ---------- translation memory ----------

  function tmIndex() {
    return window.TideTM.buildIndex(state.tm || []);
  }

  function persistTM() {
    try { window.localStorage.setItem('tide.tm', JSON.stringify(state.tm)); } catch (err) { /* quota */ }
  }

  function loadTM() {
    try {
      var raw = window.localStorage.getItem('tide.tm');
      state.tm = raw ? JSON.parse(raw) : [];
    } catch (err) { state.tm = []; }
    if (!Array.isArray(state.tm)) state.tm = [];
  }

  function mergeTM(units) {
    var index = tmIndex();
    var added = 0;
    (units || []).forEach(function (unit) {
      var source = String(unit.source || '').trim();
      var target = String(unit.target || '').trim();
      if (!source || !target || index.bySource[source]) return;
      state.tm.push({ source: source, target: target });
      index.bySource[source] = { source: source, target: target };
      added += 1;
    });
    persistTM();
    return added;
  }

  function renderTM() {
    var table = $('tmTable');
    var body = table.tBodies[0];
    body.textContent = '';
    $('tmCount').textContent = state.tm.length ? state.tm.length + ' 条记忆' : '记忆库为空';
    var query = ($('tmQuery').value || '').trim();
    if (!query) {
      table.hidden = true;
      $('tmEmpty').hidden = state.tm.length > 0;
      $('tmEmpty').textContent = state.tm.length ? '输入片段开始检索。' : '导入 TMX，或把当前工作表的已有译文合并进来。';
      return;
    }
    var hits = window.TideTM.concordance(tmIndex(), query, { limit: 50, threshold: 0.4 });
    table.hidden = hits.length === 0;
    $('tmEmpty').hidden = hits.length > 0;
    $('tmEmpty').textContent = '记忆库里没有匹配。';
    hits.forEach(function (hit) {
      var tr = document.createElement('tr');
      tr.appendChild(textNode('td', Math.round(hit.score * 100) + '%', 'num'));
      tr.appendChild(textNode('td', hit.entry.source.slice(0, 30), 'hint'));
      tr.appendChild(textNode('td', hit.entry.target.slice(0, 34)));
      var td = document.createElement('td');
      var use = document.createElement('button');
      use.className = 'act ghost';
      use.textContent = '采用';
      use.addEventListener('click', function () {
        $('entryDraft').value = hit.entry.target;
        $('entryDraft').dispatchEvent(new Event('input'));
        showTab('translate');
      });
      td.appendChild(use);
      tr.appendChild(td);
      body.appendChild(tr);
    });
  }

  function tmAnalyze() {
    var result = buildRows();
    if (result.error) { $('tmStats').textContent = '未连接表格。'; return; }
    var sources = result.rows.map(function (row) { return row.source; });
    var stats = window.TideTM.analyze(sources, tmIndex());
    $('tmStats').textContent = '「' + result.sheetName + '」' + stats.total + ' 条源文：'
      + '精确 ' + stats.exact + '（' + stats.exactPct + '%） · '
      + '高相似 ' + stats.high + '（' + stats.highPct + '%） · '
      + '低相似 ' + stats.low + '（' + stats.lowPct + '%） · '
      + '新增 ' + stats.new + '（' + stats.newPct + '%） · '
      + '复用率 ' + stats.leveragePct + '%';
  }

  function tmCollectFromSheet() {
    var result = buildRows();
    if (result.error) { $('tmStats').textContent = '未连接表格。'; return; }
    var pairs = result.rows.filter(function (row) {
      return String(row.source || '').trim() && String(row.target || '').trim();
    });
    var added = mergeTM(pairs);
    renderTM();
    $('tmStats').textContent = '从「' + result.sheetName + '」合并 ' + added + ' 条新记忆。';
  }

  function tmImport(file) {
    var isSheet = /\.(xlsx|xls)$/i.test(file.name);
    var reader = new FileReader();
    reader.onload = function () {
      try {
        var units;
        if (isSheet) {
          var terms = parseWorkbook(reader.result);
          units = terms.map(function (term) { return { source: term.zh, target: term.en }; });
        } else {
          units = window.TideTM.parseTmx(String(reader.result));
        }
        var added = mergeTM(units);
        renderTM();
        $('tmStats').textContent = (isSheet ? 'Excel ' : 'TMX ') + '导入 ' + units.length +
          ' 条，新增 ' + added + ' 条。';
      } catch (err) {
        $('tmStats').textContent = '导入失败：' + err.message;
      }
    };
    if (isSheet) reader.readAsArrayBuffer(file);
    else reader.readAsText(file, 'utf-8');
  }

  function tmExport() {
    if (!state.tm.length) { $('tmStats').textContent = '记忆库为空，无法导出。'; return; }
    var xml = window.TideTM.buildTmx(state.tm, 'zh-CN', 'en-US');
    var blob = new Blob([xml], { type: 'application/xml' });
    var link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'tide-memory.tmx';
    link.click();
    URL.revokeObjectURL(link.href);
    $('tmStats').textContent = '已导出 ' + state.tm.length + ' 条记忆。';
  }

  // ---------- QA ----------

  function severityLabel(sev) {
    return sev === 'high' ? '高' : (sev === 'medium' ? '中' : '低');
  }

  function runQA() {
    var result = buildRows();
    if (result.error) { $('qaSummary').textContent = '未连接表格。'; return; }
    var findings = window.TideQA.runQA(result.rows, state.matchTerms, {});
    state.qaFindings = findings;
    var summary = window.TideQA.summarize(findings);
    var parts = Object.keys(summary.byCategory).sort().map(function (key) {
      return key + ' ' + summary.byCategory[key];
    });
    $('qaCount').textContent = findings.length ? findings.length + ' 个问题' : '无问题';
    $('qaSummary').textContent = findings.length
      ? '「' + result.sheetName + '」' + result.rows.length + ' 条：高 ' + summary.bySeverity.high +
        ' · 中 ' + summary.bySeverity.medium + ' · 低 ' + summary.bySeverity.low +
        '（' + parts.join('，') + '）'
      : '「' + result.sheetName + '」' + result.rows.length + ' 条全部通过。';
    var table = $('qaTable');
    var body = table.tBodies[0];
    body.textContent = '';
    table.hidden = findings.length === 0;
    findings.slice(0, 300).forEach(function (finding) {
      var tr = document.createElement('tr');
      tr.appendChild(textNode('td', String(finding.row), 'num'));
      var sev = textNode('td', severityLabel(finding.severity), 'num');
      sev.style.color = finding.severity === 'high' ? 'var(--red)'
        : (finding.severity === 'medium' ? 'var(--orange)' : 'var(--label-3)');
      tr.appendChild(sev);
      tr.appendChild(textNode('td', finding.category, 'num'));
      tr.appendChild(textNode('td', finding.message, 'hint'));
      tr.addEventListener('click', function () {
        $('qaSummary').textContent = '第 ' + finding.row + ' 行 · 原文：' + finding.source.slice(0, 40) +
          ' · 译文：' + (finding.target || '（空）').slice(0, 40);
      });
      body.appendChild(tr);
    });
  }

  function exportQA() {
    var findings = state.qaFindings || [];
    if (!findings.length) { $('qaSummary').textContent = '先运行一次质检。'; return; }
    var header = ['row', 'severity', 'category', 'message', 'source', 'target'];
    var lines = [header.join(',')];
    findings.forEach(function (f) {
      var cells = [f.row, f.severity, f.category, f.message, f.source, f.target].map(function (value) {
        return '"' + String(value == null ? '' : value).replace(/"/g, '""') + '"';
      });
      lines.push(cells.join(','));
    });
    var blob = new Blob([lines.join('\n')], { type: 'text/csv;charset=utf-8' });
    var link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'tide-qa.csv';
    link.click();
    URL.revokeObjectURL(link.href);
  }

  // ---------- segment grid ----------

  function gridIssuesFor(row, terms) {
    try {
      return window.TideQA.checkRow({ row: row.row, sheet: row.sheet, source: row.source, target: row.target }, terms, {}).length;
    } catch (err) {
      return 0;
    }
  }

  function gridPendingRevision(row) {
    var project = state.project;
    if (!project || !project.segments) return 0;
    var record = project.segments[segKey(row.sheet, row.row)];
    return record ? pendingCount(record) : 0;
  }

  function writeGridCell(row, value) {
    var sheet = sheetByName(row.sheet);
    if (!sheet) { $('gridHint').textContent = '未连接表格。'; return false; }
    var before = row.target;
    archivePrevious(sheet, row.row, before);
    try {
      sheet.Range(state.settings.targetCol + row.row).Value2 = value;
    } catch (err) {
      $('gridHint').textContent = '写入失败：' + err.message;
      return false;
    }
    row.target = value;
    recordRevision(row.sheet, row.row, row.source, before, value);
    recordWrite({ sheet: row.sheet, row: row.row, source: row.source }, before, value);
    return true;
  }

  function selectGridRow(rowNumber) {
    var rows = document.querySelectorAll('#gridTable tbody tr');
    for (var i = 0; i < rows.length; i += 1) {
      rows[i].className = Number(rows[i].dataset.row) === rowNumber ? 'on' : '';
    }
  }

  function updateGridRow(tr, row) {
    var cells = tr.querySelectorAll('td');
    var qaCell = cells[cells.length - 1];
    if (!qaCell) return;
    var issues = gridIssuesFor(row, state.matchTerms);
    qaCell.textContent = issues ? String(issues) : '—';
    qaCell.className = issues ? 'num grid-qa' : 'num grid-qa ok';
  }

  function renderGrid() {
    var result = buildRows();
    if (result.error) { $('gridHint').textContent = '未连接表格。'; return; }
    state.gridRows = result.rows.map(function (row) {
      row.sheet = result.sheetName;
      return row;
    });
    var filter = $('gridFilter').value;
    var rows = state.gridRows.filter(function (row) {
      if (filter === 'empty') return !String(row.target || '').trim();
      if (filter === 'qa') return gridIssuesFor(row, state.matchTerms) > 0;
      if (filter === 'rev') return gridPendingRevision(row) > 0;
      return true;
    });
    var shown = rows.slice(0, 200);
    var table = $('gridTable');
    var body = table.tBodies[0];
    body.textContent = '';
    table.hidden = shown.length === 0;
    $('gridCount').textContent = rows.length
      ? shown.length + ' / ' + rows.length + ' 条' + (rows.length > shown.length ? '（仅显示前 200）' : '')
      : '没有符合条件的条目';
    shown.forEach(function (row) {
      var tr = document.createElement('tr');
      tr.dataset.row = String(row.row);
      tr.appendChild(textNode('td', String(row.row), 'num'));
      var tdSource = document.createElement('td');
      tdSource.appendChild(textNode('div', row.source.slice(0, 60), 'zh'));
      tr.appendChild(tdSource);
      var tdTarget = document.createElement('td');
      var input = document.createElement('input');
      input.className = 'grid-cell';
      input.value = row.target || '';
      input.addEventListener('change', function () {
        if (writeGridCell(row, input.value)) {
          tr.classList.add('dirty');
          // Update this row in place; rebuilding the grid would replace the
          // input the user is still in.
          updateGridRow(tr, row);
        }
      });
      tdTarget.appendChild(input);
      tr.appendChild(tdTarget);
      var issues = gridIssuesFor(row, state.matchTerms);
      var tdQA = textNode('td', issues ? String(issues) : '—', 'num');
      tdQA.className = issues ? 'num grid-qa' : 'num grid-qa ok';
      tr.appendChild(tdQA);
      tr.addEventListener('mousedown', function () { selectGridRow(row.row); });
      body.appendChild(tr);
    });
  }

  // ---------- project-wide analysis ----------

  function rowsForSheet(sheet) {
    var used;
    try { used = sheet.UsedRange; } catch (err) { return []; }
    if (!used) return [];
    var grid = toGrid(used.Value2);
    var startRow = Number(used.Row) || 1;
    var startCol = Number(used.Column) || 1;
    var srcIdx = colToNumber(state.settings.sourceCol) - startCol;
    var tgtIdx = colToNumber(state.settings.targetCol) - startCol;
    if (srcIdx < 0) return [];
    var headerAt = 0;
    for (var h = 0; h < grid.length; h += 1) {
      if (cellText(grid, h, srcIdx).trim()) { headerAt = h; break; }
    }
    var rows = [];
    for (var r = headerAt + 1; r < grid.length; r += 1) {
      var source = cellText(grid, r, srcIdx);
      if (!source.trim()) continue;
      rows.push({
        row: startRow + r,
        sheet: sheet.Name,
        source: source,
        target: tgtIdx >= 0 ? cellText(grid, r, tgtIdx) : ''
      });
    }
    return rows;
  }

  function analyzeWorkbook() {
    var app = sheetApp();
    var book = app && app.ActiveWorkbook;
    if (!book || !book.Worksheets) { $('projectSummary').textContent = '未连接工作簿。'; return; }
    var sheets = book.Worksheets;
    var count = 0;
    try { count = sheets.Count; } catch (err) { count = 0; }
    var results = [];
    var index = 1;
    var btn = $('btnProjectScan');
    btn.disabled = true;
    $('projectSummary').textContent = '分析中…';

    function step() {
      if (index > count) {
        state.projectScan = results;
        renderProject();
        btn.disabled = false;
        return;
      }
      var sheet = null;
      try { sheet = sheets.Item(index); } catch (err) { sheet = null; }
      if (sheet) {
        var rows = rowsForSheet(sheet);
        if (rows.length) {
          var translated = rows.filter(function (row) { return String(row.target || '').trim(); }).length;
          var stats = window.TideTM.analyze(rows.map(function (r) { return r.source; }), tmIndex());
          var issues = 0;
          try { issues = window.TideQA.runQA(rows.slice(0, 2000), state.matchTerms, {}).length; } catch (err) { issues = 0; }
          results.push({
            name: sheet.Name, total: rows.length, translated: translated,
            empty: rows.length - translated, leverage: stats.leveragePct, issues: issues
          });
        }
      }
      index += 1;
      $('projectSummary').textContent = '分析中 ' + index + '/' + count + ' …';
      window.setTimeout(step, 0);
    }
    step();
  }

  function renderProject() {
    var table = $('projectTable');
    var body = table.tBodies[0];
    body.textContent = '';
    var rows = state.projectScan || [];
    table.hidden = rows.length === 0;
    var totals = rows.reduce(function (acc, row) {
      acc.total += row.total; acc.translated += row.translated; acc.empty += row.empty; acc.issues += row.issues;
      return acc;
    }, { total: 0, translated: 0, empty: 0, issues: 0 });
    $('projectCount').textContent = rows.length
      ? rows.length + ' 张表 · ' + totals.total + ' 条 · 已译 ' + totals.translated + ' · 问题 ' + totals.issues
      : '';
    rows.forEach(function (row) {
      var tr = document.createElement('tr');
      tr.appendChild(textNode('td', row.name, 'hint'));
      tr.appendChild(textNode('td', String(row.total), 'num'));
      tr.appendChild(textNode('td', row.translated + ' / ' + row.total, 'num'));
      tr.appendChild(textNode('td', row.leverage + '%', 'num'));
      tr.appendChild(textNode('td', String(row.issues), 'num'));
      body.appendChild(tr);
    });
    if (rows.length) {
      $('projectSummary').textContent = '合计 ' + totals.total + ' 条，已译 ' + totals.translated +
        '，未译 ' + totals.empty + '，QA 问题 ' + totals.issues + '。';
    }
  }

  function renderPosition() {
    var sheetLabel = state.sheetName ? ' · ' + state.sheetName : '';
    var base = state.entries.length
      ? (state.cursor + 1) + ' / ' + state.entries.length + sheetLabel
      : '0 / 0';
    if (state.sheetMismatch) base += ' · 表已切换，点「刷新条目」';
    $('entryPos').textContent = base;
  }

  function renderEntry() {
    var entry = currentEntry();
    renderPosition();
    if (!entry) {
      $('entrySource').value = '';
      $('entryDraft').value = '';
      $('entryCheck').textContent = '';
      renderMemory(null);
      return;
    }
    $('entrySource').value = entry.source;
    $('entrySource').classList.remove('dirty');
    state.sourceDirty = false;
    $('entryDraft').value = entry.target || '';
    $('entryCheck').textContent = describeEntry(entry, $('entryDraft').value);
    renderMemory(entry, entry.source);
  }

  // WPS's CEF host swallows the usual shortcuts, so copying goes through an
  // explicit path: async clipboard first, hidden textarea + execCommand second.
  function copyText(text) {
    if (!text) return Promise.resolve(false);
    if (navigator.clipboard && navigator.clipboard.writeText) {
      return navigator.clipboard.writeText(text).then(function () { return true; }).catch(function () {
        return legacyCopy(text);
      });
    }
    return Promise.resolve(legacyCopy(text));
  }

  function legacyCopy(text) {
    var area = document.createElement('textarea');
    area.className = 'clip';
    area.value = text;
    document.body.appendChild(area);
    area.focus();
    area.select();
    var ok = false;
    try { ok = document.execCommand('copy'); } catch (err) { ok = false; }
    document.body.removeChild(area);
    return ok;
  }

  // Last resort that never depends on clipboard permissions or key forwarding:
  // the text is put in a plain textarea the user can select / paste into.
  function showManualClip(text, mode) {
    var block = $('manualClip');
    var area = $('manualClipText');
    block.hidden = false;
    if (mode === 'copy') {
      area.readOnly = true;
      area.value = text;
      $('manualClipLabel').textContent = '手动复制：已选中，按 ⌘C / Ctrl+C';
      area.focus();
      area.select();
      $('btnManualApply').hidden = true;
    } else {
      area.readOnly = false;
      area.value = '';
      $('manualClipLabel').textContent = '窗格收不到 ⌘V：请用「从表格选区读取」，或把文字拖进下面的框';
      $('btnManualApply').hidden = false;
      area.focus();
    }
  }

  // WPS swallows ⌘V before the webview sees it, so the reliable paste path is
  // through a cell: the user pastes there (WPS handles it) and we read the text.
  function selectionText() {
    var app = sheetApp();
    if (!app || !app.Selection) return '';
    try {
      var value = app.Selection.Value2;
      if (Array.isArray(value)) {
        var first = Array.isArray(value[1]) ? value[1][1] : value[1];
        return first == null ? '' : String(first);
      }
      return value == null ? '' : String(value);
    } catch (err) {
      return '';
    }
  }

  // target: 'draft' (translation) or 'source' (原文)
  function readSelection(target) {
    var text = selectionText();
    if (!text.trim()) { $('entryCheck').textContent = '选区是空的，先在表格里选中含文字的单元格。'; return; }
    if (target === 'source') {
      $('entrySource').value = text;
      markSourceDirty();
      $('entryCheck').textContent = '已从选区读取 ' + text.length + ' 个字符到原文。';
      return;
    }
    $('entryDraft').value = text;
    $('entryDraft').dispatchEvent(new Event('input'));
    $('entryCheck').textContent = '已从选区读取 ' + text.length + ' 个字符到译文。';
  }

  // Drag-and-drop survives where key events do not.
  function wireDrop(element, target) {
    if (!element) return;
    element.addEventListener('dragover', function (event) { event.preventDefault(); });
    element.addEventListener('drop', function (event) {
      event.preventDefault();
      var text = '';
      try { text = event.dataTransfer.getData('text/plain') || ''; } catch (err) { text = ''; }
      if (!text) return;
      element.value = text;
      element.dispatchEvent(new Event('input'));
      if (target === 'source') markSourceDirty();
      $('entryCheck').textContent = '已拖入 ' + text.length + ' 个字符。';
    });
  }

  function pasteIntoDraft() {
    if (!navigator.clipboard || !navigator.clipboard.readText) {
      $('entryCheck').textContent = '此处不能直接读剪贴板，请用下方手动粘贴。';
      showManualClip('', 'paste');
      return;
    }
    navigator.clipboard.readText().then(function (text) {
      $('entryDraft').value = text;
      $('entryDraft').dispatchEvent(new Event('input'));
      $('entryCheck').textContent = '已粘贴 ' + text.length + ' 个字符。';
    }).catch(function () {
      $('entryCheck').textContent = '读取剪贴板被拒绝，请用下方手动粘贴。';
      showManualClip('', 'paste');
    });
  }

  function describeEntry(entry, draft) {
    var head = '第 ' + entry.row + ' 行 · ' + entryAddress(entry.row);
    if (!String(draft || '').trim()) return head + ' · 尚未翻译';
    var check = window.TideEngine.checkTokens(entry.source, draft);
    if (check.ok) return head + ' · 占位符一致';
    return head + ' · 占位符不一致：缺 ' + (check.missing.join(' ') || '无') +
      ' / 多 ' + (check.extra.join(' ') || '无');
  }

  function currentSourceText() {
    return $('entrySource').value;
  }

  function markSourceDirty() {
    state.sourceDirty = true;
    $('entrySource').classList.add('dirty');
    var entry = currentEntry();
    if (entry) renderMemory(entry, currentSourceText());
    $('entryCheck').textContent = describeEntry(entry, $('entryDraft').value) + ' · 原文已改（未写回）';
  }

  function resetSource() {
    var entry = currentEntry();
    if (!entry) return;
    $('entrySource').value = entry.source;
    $('entrySource').classList.remove('dirty');
    state.sourceDirty = false;
    renderMemory(entry, entry.source);
    $('entryCheck').textContent = describeEntry(entry, $('entryDraft').value);
  }

  function saveSource() {
    var entry = currentEntry();
    var sheet = currentSheet();
    if (!entry || !sheet) return;
    var text = currentSourceText();
    if (text === entry.source) {
      $('entryCheck').textContent = '原文没有变化，无需写回。';
      return;
    }
    if (!window.confirm('写回源文单元格 ' + state.settings.sourceCol + entry.row + '？这会修改游戏源文。')) {
      $('entryCheck').textContent = '已取消写回。';
      return;
    }
    var cell = sheet.Range(state.settings.sourceCol + entry.row);
    var before = cell.Value2;
    cell.Value2 = text;
    if (state.unlocked) {
      state.archive.history.push({
        ts: new Date().toISOString(),
        sheet: currentSheet() ? currentSheet().Name : '',
        cell: state.settings.sourceCol + entry.row,
        before: before == null ? '' : String(before),
        after: text
      });
      saveArchive();
    }
    entry.source = text;
    state.sourceDirty = false;
    $('entrySource').classList.remove('dirty');
    refreshMemory();
    $('entryCheck').textContent = '已写回原文 ' + state.settings.sourceCol + entry.row;
  }

  // Navigation moves the grid too, so the sheet and the pane never disagree
  // about which row is being worked on.
  function move(delta) {
    if (!state.entries.length) return;
    var next = Math.min(Math.max(state.cursor + delta, 0), state.entries.length - 1);
    if (next === state.cursor && delta !== 0) {
      $('entryCheck').textContent = delta < 0 ? '已经是第一条' : '已经是最后一条';
      return;
    }
    state.cursor = next;
    renderEntry();
    var entry = currentEntry();
    if (entry) revealCell(entrySheet(entry), state.settings.sourceCol + entry.row);
  }

  function writeAndNext() {
    if (syncActiveSheet()) {
      $('entryCheck').textContent = '工作表已切换，条目已更新；请确认后再写入。';
      return;
    }
    var entry = currentEntry();
    if (!entry) return;
    var sheet = currentSheet();
    if (!sheet) { $('entryCheck').textContent = '未连接 WPS，无法写入。'; return; }
    var value = $('entryDraft').value;
    if (!value.trim()) { $('entryCheck').textContent = '草稿为空，未写入。'; return; }
    var before = entry.target;
    archivePrevious(entrySheet(entry), entry.row, before);
    if (!writeCell(entry, value)) return;
    var address = entryAddress(entry.row);
    entry.target = value;
    refreshMemory();
    recordWrite(entry, before, value);
    recordRevision(entry.sheet, entry.row, entry.source, before, value);
    move(1);
    $('entryCheck').textContent = '已写入 ' + address + ' · ' + $('entryCheck').textContent;
  }

  function undoSelf() {
    if (!state.lastWrite) { $('entryCheck').textContent = '没有可撤销的写入。'; return; }
    var sheet = sheetByName(state.lastWrite.sheet);
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

  // Look for the text in one sheet: source column first, then target column.
  function findInSheet(sheet, text, prefer) {
    if (!sheet || !text) return null;
    var used;
    try { used = sheet.UsedRange; } catch (err) { return null; }
    if (!used) return null;
    var grid;
    try { grid = toGrid(used.Value2); } catch (err) { return null; }
    if (!grid.length) return null;
    var startRow = Number(used.Row) || 1;
    var startCol = Number(used.Column) || 1;
    var srcIdx = colToNumber(state.settings.sourceCol) - startCol;
    var tgtIdx = colToNumber(state.settings.targetCol) - startCol;
    var wanted = String(text).trim();
    var order = prefer === 'target'
      ? [['target', tgtIdx], ['source', srcIdx]]
      : [['source', srcIdx], ['target', tgtIdx]];
    for (var c = 0; c < order.length; c += 1) {
      var kind = order[c][0];
      var col = order[c][1];
      if (col < 0) continue;
      for (var i = 0; i < grid.length; i += 1) {
        if (cellText(grid, i, col).trim() === wanted) {
          return { row: startRow + i, foundIn: kind };
        }
      }
    }
    return null;
  }

  // Search the visible sheet, then every other worksheet in the workbook.
  function findAcrossWorkbook(text) {
    var current = currentSheet();
    var local = findInSheet(current, text, prefer);
    if (local) return { sheet: current, row: local.row, foundIn: local.foundIn };
    var app = sheetApp();
    var book = app && app.ActiveWorkbook;
    if (!book || !book.Worksheets) return null;
    var sheets = book.Worksheets;
    var count = 0;
    try { count = sheets.Count; } catch (err) { return null; }
    for (var i = 1; i <= count; i += 1) {
      var sheet;
      try { sheet = sheets.Item(i); } catch (err) { continue; }
      if (!sheet || (current && sheet.Name === current.Name)) continue;
      var hit = findInSheet(sheet, text);
      if (hit) return { sheet: sheet, row: hit.row, foundIn: hit.foundIn };
    }
    return null;
  }

  // Non-blocking workbook scan: one sheet per tick, with progress feedback.
  // A synchronous sweep of 93 sheets freezes the pane and reads as "no response".
  function searchWorkbookAsync(text, prefer, onHit, onDone) {
    var app = sheetApp();
    var book = app && app.ActiveWorkbook;
    if (!book || !book.Worksheets) { onDone(); return; }
    var current = currentSheet();
    var local = findInSheet(current, text, prefer);
    if (local) { onHit({ sheet: current, row: local.row, foundIn: local.foundIn }); return; }
    var sheets = book.Worksheets;
    var count = 0;
    try { count = sheets.Count; } catch (err) { onDone(); return; }
    var index = 1;
    function step() {
      if (index > count) { onDone(); return; }
      var sheet = null;
      try { sheet = sheets.Item(index); } catch (err) { sheet = null; }
      if (sheet && current && sheet.Name === current.Name) { index += 1; step(); return; }
      index += 1;
      if (sheet) {
        var hit = findInSheet(sheet, text, prefer);
        if (hit) { onHit({ sheet: sheet, row: hit.row, foundIn: hit.foundIn }); return; }
      }
      if (index % 5 === 0) $('entryCheck').textContent = '搜索中 ' + index + '/' + count;
      window.setTimeout(step, 0);
    }
    step();
  }

  // Find the worksheet row whose source cell matches the text exactly.
  function findRowBySource(sheet, text) {
    if (!sheet || !text) return 0;
    var used = sheet.UsedRange;
    if (!used) return 0;
    var grid = toGrid(used.Value2);
    var startRow = Number(used.Row) || 1;
    var startCol = Number(used.Column) || 1;
    var srcIdx = colToNumber(state.settings.sourceCol) - startCol;
    var wanted = String(text).trim();
    for (var i = 0; i < grid.length; i += 1) {
      if (cellText(grid, i, srcIdx).trim() === wanted) return startRow + i;
    }
    return 0;
  }

  // The row to operate on: the entry's row when it still holds this text, the
  // searched row when the text moved, 0 when it is nowhere in the sheet.
  function resolveRow(entry, text) {
    var sheet = entrySheet(entry);
    var row = findRowBySource(sheet, text);
    if (!row) return 0;
    if (entry && row !== entry.row) {
      entry.row = row;
      state.cursor = state.entries.indexOf(entry);
    }
    return row;
  }

  function locateSource() {
    var entry = currentEntry();
    var text = currentSourceText().trim() || (entry ? entry.source : '');
    if (!text) { $('entryCheck').textContent = '先输入要定位的原文。'; return; }
    $('entryCheck').textContent = '搜索中：' + text.slice(0, 20);
    searchWorkbookAsync(text, 'source', function (hit) {
      var column = hit.foundIn === 'target' ? state.settings.targetCol : state.settings.sourceCol;
      var address = column + hit.row;
      var ok = revealCell(hit.sheet, address);
      if (entry && entry.sheet === hit.sheet.Name) entry.row = hit.row;
      var where = (hit.foundIn === 'target' ? '译文列' : '原文列');
      $('entryCheck').textContent = ok
        ? '已定位 ' + address + '（' + hit.sheet.Name + ' · ' + where + '）'
        : '定位失败：请在表格中手动跳到 ' + address;
    }, function () {
      $('entryCheck').textContent = '整本工作簿都没有这段文字：' + text.slice(0, 20);
    });
  }

  function locateTarget() {
    var entry = currentEntry();
    // Search what the translation box actually holds, not the source.
    var text = ($('entryDraft').value || '').trim();
    if (!text) { $('entryCheck').textContent = '译文草稿为空，没有可定位的内容。'; return; }
    $('entryCheck').textContent = '搜索译文：' + text.slice(0, 20);
    searchWorkbookAsync(text, 'target', function (hit) {
      var column = hit.foundIn === 'target' ? state.settings.targetCol : state.settings.sourceCol;
      var address = column + hit.row;
      var ok = revealCell(hit.sheet, address);
      if (entry && entry.sheet === hit.sheet.Name) entry.row = hit.row;
      $('entryCheck').textContent = ok
        ? '已定位译文 ' + address + '（' + hit.sheet.Name + '）'
        : '定位失败：请在表格中手动跳到 ' + address;
    }, function () {
      $('entryCheck').textContent = '整本工作簿都没有这段文字：' + text.slice(0, 20);
    });
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

  // ---- workbook import (SheetJS) ----

  function mapSheetColumns(header) {
    var map = {};
    (header || []).forEach(function (cell, index) {
      var text = String(cell == null ? '' : cell).trim();
      if (!text) return;
      if (map.zh === undefined && /中文|源文|原文|source|^zh$/i.test(text)) map.zh = index;
      else if (map.approved === undefined && /最终英文|确认英文|approved|final/i.test(text)) map.approved = index;
      else if (map.working === undefined && /现有英文|英文|current|target|^en$/i.test(text)) map.working = index;
      else if (map.note === undefined && /备注|注释|说明|note|comment/i.test(text)) map.note = index;
    });
    if (map.zh === undefined) map.zh = 0;
    if (map.approved === undefined && map.working === undefined) map.working = 1;
    return map;
  }

  function parseWorkbook(buffer) {
    if (typeof XLSX === 'undefined') throw new Error('xlsx 解析库未加载');
    var workbook = XLSX.read(new Uint8Array(buffer), { type: 'array' });
    var terms = [];
    var seen = {};
    workbook.SheetNames.forEach(function (name) {
      var rows = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, blankrows: false, defval: '' });
      if (!rows.length) return;
      var map = mapSheetColumns(rows[0]);
      for (var i = 1; i < rows.length; i += 1) {
        var row = rows[i] || [];
        var zh = String(row[map.zh] == null ? '' : row[map.zh]).trim();
        var en = '';
        var status = 'working';
        if (map.approved !== undefined && String(row[map.approved] == null ? '' : row[map.approved]).trim()) {
          en = String(row[map.approved]).trim();
          status = 'approved';
        } else if (map.working !== undefined) {
          en = String(row[map.working] == null ? '' : row[map.working]).trim();
        }
        if (!zh || !en) continue;
        var key = zh + '\u0000' + en.toLowerCase();
        if (seen[key]) continue;
        seen[key] = 1;
        var looksLikeNote = zh.indexOf('\n') > -1 || zh.length > 30 || zh.indexOf('改用') > -1 || zh.indexOf('建议') > -1;
        if (/[\u4e00-\u9fff]/.test(en) || en.length > 60 || looksLikeNote) status = 'needs_review';
        terms.push({ zh: zh, en: en, status: status, sheet: name, note: '' });
      }
    });
    if (!terms.length) throw new Error('没有读到中英对照行');
    return window.TideEngine.parseTermbase(terms);
  }

  // ---------- free input (no sheet writes) ----------

  var freeTimer = null;

  function adoptIntoFree(text) {
    $('freeDraft').value = text;
    $('freeDraft').focus();
  }

  function renderFree() {
    var text = $('freeSource').value;
    var termBody = $('freeTermTable').tBodies[0];
    var memoryBody = $('freeMemoryTable').tBodies[0];
    termBody.textContent = '';
    memoryBody.textContent = '';
    if (!text.trim()) {
      $('freeTermTable').hidden = true;
      $('freeMemoryTable').hidden = true;
      $('freeTermEmpty').hidden = false;
      $('freeMemoryEmpty').hidden = false;
      $('freeTermCount').textContent = '';
      $('freeMemoryCount').textContent = '';
      $('freeStatus').textContent = '';
      return;
    }

    var hits = window.TideEngine.findMatches(text, state.matchTerms);
    var seen = {};
    var unique = [];
    hits.forEach(function (hit) {
      var key = hit.term.zh + '\u0000' + hit.term.en;
      if (seen[key]) return;
      seen[key] = 1;
      unique.push(hit.term);
    });
    $('freeTermCount').textContent = unique.length ? unique.length + ' 条' : '';
    $('freeTermTable').hidden = unique.length === 0;
    $('freeTermEmpty').hidden = unique.length > 0;
    unique.forEach(function (term) {
      var tr = document.createElement('tr');
      tr.appendChild(textNode('td', term.zh));
      tr.appendChild(textNode('td', term.en || '—'));
      var td = document.createElement('td');
      var use = document.createElement('button');
      use.className = 'act ghost';
      use.textContent = '采用';
      use.addEventListener('click', function () { adoptIntoFree(term.en || ''); });
      td.appendChild(use);
      tr.appendChild(td);
      termBody.appendChild(tr);
    });

    var memory = window.TideEngine.findMemory(text, state.memory, 0.5, 5);
    var rows = memory.exact.map(function (item) { return { item: item, exact: true, score: 1 }; })
      .concat(memory.fuzzy.map(function (item) { return { item: item, exact: false, score: item.score }; }));
    $('freeMemoryCount').textContent = rows.length ? rows.length + ' 条 / ' + state.memory.length + ' 条记忆' : '';
    $('freeMemoryTable').hidden = rows.length === 0;
    $('freeMemoryEmpty').hidden = rows.length > 0;
    rows.forEach(function (row) {
      var tr = document.createElement('tr');
      tr.appendChild(textNode('td', row.exact ? '精确' : Math.round(row.score * 100) + '%', 'num'));
      tr.appendChild(textNode('td', row.item.source.slice(0, 30), 'hint'));
      tr.appendChild(textNode('td', row.item.target.slice(0, 30)));
      var td = document.createElement('td');
      var use = document.createElement('button');
      use.className = 'act ghost';
      use.textContent = '采用';
      use.addEventListener('click', function () { adoptIntoFree(row.item.target); });
      td.appendChild(use);
      tr.appendChild(td);
      memoryBody.appendChild(tr);
    });

    $('freeStatus').textContent = '字符 ' + text.length + ' · 术语 ' + unique.length + ' · 记忆 ' + rows.length;
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
    ['matches', 'translate', 'free', 'find', 'review', 'tm', 'grid', 'project', 'terms', 'history', 'sync', 'settings'].forEach(function (tab) {
      $('panel-' + tab).hidden = tab !== name;
    });
    state.lastTab = name;
    writeJSON(K.lastTab, name);
    if (name === 'translate') { loadEntries(); renderEntry(); }
    if (name === 'free') { if (!state.memory.length) loadEntries(); renderFree(); }
    if (name === 'find') renderFind();
    if (name === 'history') { renderLock(); renderHistory(); }
    if (name === 'terms') { syncTermColumns(); renderTerms(); }
    if (name === 'review') renderReview();
    if (name === 'tm') renderTM();
    if (name === 'grid') renderGrid();
    if (name === 'project') renderProject();
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
        uidCol: $('colUid').value.toUpperCase() || 'A',
        oldCol: ($('colOld') && $('colOld').value.toUpperCase()) || 'E'
      };
      writeJSON(K.settings, state.settings);
      writeJSON(K.repo, $('syncRepo').value);
      writeJSON(K.branch, $('syncBranch').value);
      $('settingsStatus').textContent = '已保存到本机。';
    });
    $('termFile').addEventListener('change', function (event) {
      var file = event.target.files && event.target.files[0];
      if (!file) return;
      var isSheet = /\.(xlsx|xls)$/i.test(file.name);
      var reader = new FileReader();
      reader.onload = function () {
        try {
          var terms = isSheet
            ? parseWorkbook(reader.result)
            : window.TideEngine.parseTermbase(JSON.parse(String(reader.result)));
          var meta = { source: file.name, count: terms.length, importedAt: new Date().toISOString() };
          setTerms(terms, meta);
          persistTerms(meta);
          $('settingsStatus').textContent = '已导入并保存到本机：' + terms.length + ' 条（' + file.name + '）。';
          renderTermSource();
          renderTerms();
        } catch (err) {
          $('settingsStatus').textContent = '导入失败：' + err.message;
        }
      };
      if (isSheet) reader.readAsArrayBuffer(file);
      else reader.readAsText(file, 'utf-8');
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
      setTerms(state.bundledTerms.slice(), null);
      try {
        window.localStorage.removeItem(K.termStore);
        window.localStorage.removeItem(K.terms);
      } catch (err) { /* ignore */ }
      $('settingsStatus').textContent = '已清除本机保存的术语表，恢复内置：' + state.terms.length + ' 条。';
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
    $('btnCopySource').addEventListener('click', function () {
      var entry = currentEntry();
      if (!entry) return;
      copyText(entry.source).then(function (ok) {
        if (ok) { $('entryCheck').textContent = '已复制原文（' + entry.row + ' 行）'; }
        else { $('entryCheck').textContent = '复制被拒绝，请用下方手动复制。'; showManualClip(entry.source, 'copy'); }
      });
    });
    $('btnCopyDraft').addEventListener('click', function () {
      copyText($('entryDraft').value).then(function (ok) {
        if (ok) { $('entryCheck').textContent = '已复制草稿'; }
        else { $('entryCheck').textContent = '复制被拒绝，请用下方手动复制。'; showManualClip($('entryDraft').value, 'copy'); }
      });
    });
    $('btnPasteDraft').addEventListener('click', pasteIntoDraft);
    $('btnReadCell').addEventListener('click', function () { readSelection('draft'); });
    $('btnReadCellSrc').addEventListener('click', function () { readSelection('source'); });
    $('btnManualReadCell').addEventListener('click', function () {
      var text = selectionText();
      if (!text.trim()) { $('entryCheck').textContent = '选区是空的。'; return; }
      $('manualClipText').value = text;
      $('manualClipText').readOnly = false;
    });
    wireDrop($('entryDraft'), 'draft');
    wireDrop($('entrySource'), 'source');
    // If the host ever does forward the key, handle it without relying on focus.
    document.addEventListener('keydown', function (event) {
      if ((event.metaKey || event.ctrlKey) && String(event.key).toLowerCase() === 'v') {
        var tag = (event.target && event.target.tagName) || '';
        if (tag === 'TEXTAREA' || tag === 'INPUT') return;
        $('clipHint').textContent = '⌘V 被 WPS 截获时的替代路径：表格里粘贴 → 选中单元格 → 点「读取选区」。';
      }
    });
    $('btnManualApply').addEventListener('click', function () {
      var text = $('manualClipText').value;
      if (!text) return;
      $('entryDraft').value = text;
      $('entryDraft').dispatchEvent(new Event('input'));
      $('manualClip').hidden = true;
      $('entryCheck').textContent = '已从手动粘贴区应用 ' + text.length + ' 个字符。';
    });
    $('btnManualClose').addEventListener('click', function () { $('manualClip').hidden = true; });
    $('freeSource').addEventListener('input', function () {
      window.clearTimeout(freeTimer);
      freeTimer = window.setTimeout(renderFree, 250);
    });
    $('btnFreeCopy').addEventListener('click', function () {
      copyText($('freeDraft').value).then(function (ok) {
        $('freeStatus').textContent = ok ? '已复制译文' : '复制失败，请手动选择。';
      });
    });
    $('btnFreeClear').addEventListener('click', function () {
      $('freeSource').value = '';
      $('freeDraft').value = '';
      renderFree();
    });
    var colButtons = document.querySelectorAll('button[data-col]');
    for (var c = 0; c < colButtons.length; c += 1) {
      colButtons[c].addEventListener('click', function (event) {
        var target = $(event.currentTarget.getAttribute('data-col'));
        var letter = selectedColumnLetter();
        if (!letter) {
          $('settingsStatus').textContent = '请先在表格里选中一个单元格。';
          return;
        }
        target.value = letter;
        $('settingsStatus').textContent = '已填入第 ' + letter + ' 列，记得点「保存设置」。';
      });
    }
    $('btnFindAll').addEventListener('click', runFind);
    var refreshTerms = $('btnRefreshTerms');
    if (refreshTerms) refreshTerms.addEventListener('click', refreshTermsFromSheet);
    $('btnProjectScan').addEventListener('click', analyzeWorkbook);
    $('btnGridRefresh').addEventListener('click', renderGrid);
    $('gridFilter').addEventListener('change', renderGrid);
    $('btnRunQA').addEventListener('click', runQA);
    $('btnExportQA').addEventListener('click', exportQA);
    $('btnTmImport').addEventListener('click', function () { $('tmFile').click(); });
    $('tmFile').addEventListener('change', function (event) {
      var file = event.target.files && event.target.files[0];
      if (file) tmImport(file);
    });
    $('btnTmExport').addEventListener('click', tmExport);
    $('btnTmCollect').addEventListener('click', tmCollectFromSheet);
    $('btnTmAnalyze').addEventListener('click', tmAnalyze);
    $('tmQuery').addEventListener('input', function () {
      window.clearTimeout(state.tmTimer);
      state.tmTimer = window.setTimeout(renderTM, 200);
    });
    $('btnReviewRefresh').addEventListener('click', function () { loadProject().then(renderReview); });
    $('btnAddComment').addEventListener('click', addComment);
    var termSearch = $('termSearch');
    if (termSearch) {
      var termTimer = null;
      termSearch.addEventListener('input', function () {
        window.clearTimeout(termTimer);
        termTimer = window.setTimeout(renderTerms, 150);
      });
    }
    $('btnRepPreview').addEventListener('click', previewBatch);
    $('btnRepApply').addEventListener('click', applyBatch);
    $('btnRepUndo').addEventListener('click', undoBatch);
    $('findInput').addEventListener('keydown', function (event) {
      if (event.key === 'Enter') runFind();
    });
    $('btnRefreshEntries').addEventListener('click', function () {
      state.sheetMismatch = '';
      loadEntries();
      renderEntry();
      $('entryCheck').textContent = '已刷新到 ' + activeSheetName();
    });
    $('btnPrevEntry').addEventListener('click', function () { move(-1); });
    $('btnNextEntry').addEventListener('click', function () { move(1); });
    $('btnWriteNext').addEventListener('click', writeAndNext);
    $('entrySource').addEventListener('input', markSourceDirty);
    $('btnResetSource').addEventListener('click', resetSource);
    $('btnSaveSource').addEventListener('click', saveSource);
    $('btnLocateSource').addEventListener('click', locateSource);
    $('btnLocate').addEventListener('click', locateTarget);
    $('btnUndoWrite').addEventListener('click', undoSelf);
    $('btnRollback').addEventListener('click', rollbackEntry);
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
    if (!loadSavedTerms()) setTerms([]);
    $('colSource').value = state.settings.sourceCol;
    $('colTarget').value = state.settings.targetCol;
    $('colUid').value = state.settings.uidCol;
    if ($('colOld')) $('colOld').value = state.settings.oldCol || 'E';
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
    loadTM();
    loadProject().then(function () {
      if (state.lastTab === 'review') renderReview();
    });
    watchRibbon();
    if (sheetApp()) window.setInterval(syncActiveSheet, 1200);
    checkForUpdate();
    window.setInterval(checkForUpdate, 120000);
  }

  // Read-only inspection hook for the desktop smoke test.
  Object.defineProperty(window, '__tide', {
    value: {
      get project() { return state.project; },
      get selected() { return state.selectedSegment; },
      get lastTab() { return state.lastTab; }
    },
    configurable: true
  });

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
})();
