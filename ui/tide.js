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
    profiles: 'tide.profiles',
    repo: 'tide.syncRepo',
    branch: 'tide.syncBranch',
    nas: 'tide.nas',
    nasAuth: 'tide.nas.auth'
  };
  var HISTORY_LIMIT = 3000;
  var HELPER = 'http://127.0.0.1:8765';

  var state = {
    terms: [],
    settings: { sourceCol: 'C', targetCol: 'F', oldCol: 'E', uidCol: 'A', targetLang: 'en' },
    perSheetCols: {},
    scanned: [],
    archive: { history: [], snapshot: {} },
    passphrase: null,
    unlocked: false,
    lastTab: 'matches',
    prevTab: '',
    helper: { ok: false, token: false },
    bundledTerms: [],
    matchTerms: [],
    bundledMeta: null,
    termsBlob: null,
    termsMeta: null,
    entries: [],
    profiles: [],
    activeProfile: 'default',
    findResults: [],
    tm: [],
    qaFindings: [],
    qaMode: '全部检查',
    nas: { base: 'https://192.168.23.232:8443', token: '', user: '', projects: [], pid: null,
      langs: [], lang: '', items: [], total: 0, offset: 0, q: '', loaded: false,
      termsCache: {}, jumpStarted: false, jumpBusy: false },
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

  // Click-to-locate for every list that names a row: reveal the cell that the
  // item came from, so the pane always points at something real in the sheet.
  function jumpToRow(row, sheetName, messageEl) {
    var sheet = sheetName ? sheetByName(sheetName) : currentSheet();
    var address = state.settings.targetCol + row;
    var ok = revealCell(sheet, address);
    if (messageEl) {
      messageEl.textContent = ok
        ? '已定位 ' + address + (sheetName ? '（' + sheetName + '）' : '')
        : '定位失败：' + address;
    }
    return ok;
  }

  function jumpToText(text, prefer, messageEl) {
    var sheet = currentSheet();
    var found = findInSheet(sheet, text, prefer);
    if (!found) {
      if (messageEl) messageEl.textContent = '当前表里找不到「' + String(text).slice(0, 20) + '」';
      return false;
    }
    var column = found.foundIn === 'target' ? state.settings.targetCol : state.settings.sourceCol;
    var ok = revealCell(sheet, column + found.row);
    if (messageEl) messageEl.textContent = ok ? '已定位第 ' + found.row + ' 行' : '定位失败：第 ' + found.row + ' 行';
    return ok;
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

  // Real export sheets often carry 2-3 header rows ("导出标记", "标记字段",
  // "字段解释") before data. Treat every leading header-ish row as header.
  var HEADER_RX = /^(source(\s*text)?(\s*\(old\))?|title|desc|weight|en(\s*\(old\))?|字段解释|标记字段|导出标记|中文|原文|备注|uid|ui)$/i;

  function detectHeaderEnd(grid) {
    var last = -1;
    var limit = Math.min(grid.length, 12);
    for (var r = 0; r < limit; r += 1) {
      var line = grid[r] || [];
      var blank = true;
      var headerish = false;
      for (var c = 0; c < line.length; c += 1) {
        var text = String(line[c] == null ? '' : line[c]).trim();
        if (!text) continue;
        blank = false;
        if (HEADER_RX.test(text)) headerish = true;
      }
      if (blank) { last = r; continue; }
      if (headerish) { last = r; continue; }
      break;
    }
    return last;
  }

  // Read the header row and report where "Source Text" and "EN" actually are,
  // so a mismatched column mapping can be corrected instead of silently
  // producing nonsense.
  function detectColumns(grid, startCol) {
    var limit = Math.min(grid.length, 12);
    for (var r = 0; r < limit; r += 1) {
      var line = grid[r] || [];
      var srcIdx = -1;
      var tgtIdx = -1;
      var oldIdx = -1;
      for (var c = 0; c < line.length; c += 1) {
        var text = String(line[c] == null ? '' : line[c]).trim();
        if (srcIdx < 0 && /^source(\s*text)?(\s*\(old\))?$/i.test(text)) srcIdx = c;
        if (tgtIdx < 0 && /^en$/i.test(text)) tgtIdx = c;
        if (oldIdx < 0 && /^en\s*\(old\)$/i.test(text)) oldIdx = c;
      }
      if (srcIdx > -1 && tgtIdx > -1 && tgtIdx > srcIdx) {
        return {
          sourceCol: numberToCol(startCol + srcIdx),
          targetCol: numberToCol(startCol + tgtIdx),
          oldCol: oldIdx > -1 ? numberToCol(startCol + oldIdx) : numberToCol(startCol + tgtIdx - 1)
        };
      }
    }
    return null;
  }

  // Repeated "Source Text / EN" blocks in one sheet are separate segment
  // groups: B–F, G–K, L–P … Each becomes its own set of segments.
  function detectColumnGroups(grid, startCol) {
    var limit = Math.min(grid.length, 12);
    for (var r = 0; r < limit; r += 1) {
      var line = grid[r] || [];
      var groups = [];
      var current = null;
      for (var c = 0; c < line.length; c += 1) {
        var text = String(line[c] == null ? '' : line[c]).trim();
        if (!text) continue;
        if (/^source(\s*text)?(\s*\(old\))?$/i.test(text)) {
          current = { sourceIdx: c, targetIdx: -1, oldIdx: -1 };
          groups.push(current);
        } else if (current && /^en\s*\(old\)$/i.test(text) && current.oldIdx < 0) {
          current.oldIdx = c;
        } else if (current && /^en$/i.test(text) && current.targetIdx < 0) {
          current.targetIdx = c;
        }
      }
      var valid = groups.filter(function (group) {
        return group.sourceIdx > -1 && (group.targetIdx > group.sourceIdx || group.oldIdx > group.sourceIdx);
      });
      if (valid.length) {
        return valid.map(function (group) {
          // A block with EN (old) but no EN still exists: the target is blank,
          // and writing creates the EN column where it belongs (right of old).
          var hasTarget = group.targetIdx > group.sourceIdx;
          var targetIdx = hasTarget ? group.targetIdx : (group.oldIdx + 1);
          // If the slot right of EN (old) belongs to the next block's header,
          // the target is blank but must never be written.
          var occupied = !hasTarget && String(line[targetIdx] == null ? '' : line[targetIdx]).trim() !== '';
          return {
            sourceCol: numberToCol(startCol + group.sourceIdx),
            targetCol: numberToCol(startCol + targetIdx),
            oldCol: group.oldIdx > -1 ? numberToCol(startCol + group.oldIdx) : numberToCol(startCol + targetIdx - 1),
            missingTarget: !hasTarget,
            blocked: occupied
          };
        });
      }
    }
    return null;
  }

  function isTargetHeader(text) {
    var code = state.settings.targetLang || 'en';
    var names = window.TideLangs ? window.TideLangs.headersFor(code) : ['EN'];
    var value = String(text || '').trim();
    for (var i = 0; i < names.length; i += 1) {
      if (value.toLowerCase() === names[i].toLowerCase()) return true;
    }
    return false;
  }

  function buildRows(sheetOverride, silent) {
    var sheet = sheetOverride || currentSheet();
    if (!sheet) return { error: 'not-in-wps' };
    var used = sheet.UsedRange;
    if (!used) return { rows: [], sheetName: sheet.Name };
    var grid = toGrid(used.Value2);
    var startRow = Number(used.Row) || 1;
    var startCol = Number(used.Column) || 1;
    var srcIdx = colToNumber(state.settings.sourceCol) - startCol;
    var tgtIdx = colToNumber(state.settings.targetCol) - startCol;
    var uidIdx = colToNumber(state.settings.uidCol) - startCol;
    var headerAt = detectHeaderEnd(grid);

    // The header is the source of truth: every Source Text / EN pair becomes a
    // group. Sheets without those labels keep the manual mapping.
    var groups = detectColumnGroups(grid, startCol);
    if (groups && groups.length) {
      if (!silent) {
        state.settings.sourceCol = groups[0].sourceCol;
        state.settings.targetCol = groups[0].targetCol;
        state.settings.oldCol = groups[0].oldCol;
        state.perSheetCols[sheet.Name] = groups[0];
        if ($('colSource')) $('colSource').value = groups[0].sourceCol;
        if ($('colTarget')) $('colTarget').value = groups[0].targetCol;
        if ($('colOld')) $('colOld').value = groups[0].oldCol;
        writeJSON(K.settings, state.settings);
      }
    } else {
      groups = [{ sourceCol: state.settings.sourceCol, targetCol: state.settings.targetCol, oldCol: state.settings.oldCol }];
    }

    var rows = [];
    groups.forEach(function (group, groupIndex) {
      var groupSrc = colToNumber(group.sourceCol) - startCol;
      var groupTgt = colToNumber(group.targetCol) - startCol;
      for (var r = headerAt + 1; r < grid.length; r += 1) {
        var source = cellText(grid, r, groupSrc);
        if (!source.trim()) continue;
        rows.push({
          row: startRow + r,
          sheet: sheet.Name,
          group: groupIndex,
          sourceCol: group.sourceCol,
          targetCol: group.targetCol,
          oldCol: group.oldCol,
          missingTarget: !!group.missingTarget,
          blocked: !!group.blocked,
          uid: uidIdx >= 0 ? cellText(grid, r, uidIdx) : '',
          source: source,
          // A blocked group has no EN column; never mirror the neighbour's text.
          target: group.blocked ? '' : (groupTgt >= 0 ? cellText(grid, r, groupTgt) : '')
        });
      }
    });
    rows.sort(function (a, b) { return (a.row - b.row) || (a.group - b.group); });
    return {
      rows: rows,
      sheetName: sheet.Name,
      headerEnd: headerAt,
      groups: groups
    };
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
      tr.appendChild(textNode('td', (row.targetCol || state.settings.targetCol) + row.row, 'num'));
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
      tr.classList.add('rowlink');
      tr.addEventListener('click', function (event) {
        if (event.target && event.target.tagName === 'BUTTON') return;
        revealCell(currentSheet(), state.settings.sourceCol + row.row);
        $('matchCount').textContent = '已定位第 ' + row.row + ' 行';
      });
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
    var list = currentTerms();
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
    var total = currentTerms().length;
    var langName = window.TideLangs ? window.TideLangs.get(state.settings.targetLang || 'en').name : '';
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
    $('termCount').textContent = (query
      ? rows.length + ' / ' + total + ' 条' + (result.fuzzy ? ' · 近似匹配' : '')
      : total + ' 条术语') + (langName ? ' · ' + langName : '');
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
      tr.classList.add('rowlink');
      tr.addEventListener('click', function (event) {
        if (event.target && event.target.tagName === 'BUTTON') return;
        jumpToText(term.zh, 'source', $('termEmpty'));
      });
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
  function currentTerms() {
    var lang = state.settings.targetLang || 'en';
    return (state.terms || []).filter(function (term) {
      return (term.lang || 'en') === lang;
    });
  }

  function setTerms(list, meta) {
    state.terms = list || [];
    state.termsMeta = meta || null;
    state.matchTerms = currentTerms().filter(function (term) {
      return term.status !== 'needs_review';
    });
  }

  function refreshTermLanguage() {
    setTerms(state.terms, state.termsMeta);
    renderTerms();
    loadEntries();
    renderEntry();
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
    ensureWorkbookMemory();
    return state.entries;
  }

  function currentEntry() {
    return state.entries[state.cursor] || null;
  }

  function entryAddress(row) {
    // Callers pass either a row object or a bare row number.
    var column = (row && row.targetCol) ? row.targetCol : state.settings.targetCol;
    var number = (row && row.row) ? row.row : row;
    return column + number;
  }

  // EN (old) holds exactly the previous translation; the current EN is the
  // only official text. Keeping the archive in the sheet lets anyone roll back.
  function archivePrevious(sheet, row, before, oldColumn) {
    var column = oldColumn || state.settings.oldCol;
    if (!column || !sheet) return;
    var previous = String(before == null ? '' : before);
    if (!previous.trim()) return;
    try { sheet.Range(column + row).Value2 = previous; } catch (err) { /* read-only cell */ }
  }

  function rollbackEntry() {
    var entry = currentEntry();
    if (!entry) return;
    var sheet = entrySheet(entry);
    var oldAddress = (entry.oldCol || state.settings.oldCol) + entry.row;
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

  // The batch view is per sheet, but a translator works across the workbook:
  // index every worksheet once, then keep it fresh on writes.
  function refreshMemory() {
    state.memory = window.TideEngine.buildMemory(state.entries);
  }

  function workbookKey() {
    var book = sheetApp() && sheetApp().ActiveWorkbook;
    return book && book.Name ? String(book.Name) : 'default';
  }

  function ensureWorkbookMemory(done) {
    var app = sheetApp();
    if (!app || !app.ActiveWorkbook) { if (done) done(); return; }
    var key = workbookKey();
    if (state.workbookKey === key && state.workbookMemory) {
      state.memory = state.workbookMemory;
      if (done) done();
      return;
    }
    // A build is already running for this workbook; do not start a second one.
    if (state.workbookKey === key && state.workbookBuilding) return;
    var sheets = app.ActiveWorkbook.Worksheets;
    var count = 0;
    try { count = sheets.Count; } catch (err) { count = 0; }
    if (!count) { if (done) done(); return; }

    var memory = [];
    var index = 1;
    state.workbookKey = key;
    state.workbookBuilding = true;
    $('memoryCount').textContent = '正在索引整本工作簿…';

    function step() {
      if (index > count) {
        state.workbookMemory = window.TideEngine.buildMemory(memory);
        state.workbookBuilding = false;
        state.memory = state.workbookMemory;
        renderMemory(currentEntry(), currentSourceText());
        if (done) done();
        return;
      }
      try {
        var sheet = sheets.Item(index);
        var result = buildRows(sheet, true);
        (result.rows || []).forEach(function (row) {
          if (String(row.source || '').trim() && String(row.target || '').trim()) {
            memory.push({
              source: row.source, target: row.target, row: row.row,
              sheet: result.sheetName, uid: row.uid || ''
            });
          }
        });
      } catch (err) { /* a single unreadable sheet must not stop the index */ }
      $('memoryCount').textContent = '正在索引整本工作簿 ' + index + ' / ' + count + ' …';
      index += 1;
      window.setTimeout(step, 0);
    }
    step();
  }

  function rememberPair(sheetName, row, source, target) {
    if (!state.workbookMemory) return;
    var text = String(source || '').trim();
    if (!text) return;
    var found = null;
    for (var i = 0; i < state.workbookMemory.length; i += 1) {
      var item = state.workbookMemory[i];
      if (item.sheet === sheetName && item.row === row && item.source === text) { found = item; break; }
    }
    if (found) found.target = target;
    else state.workbookMemory.push({ source: text, target: target, row: row, sheet: sheetName, uid: '' });
    state.memory = state.workbookMemory;
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
    count.textContent = rows.length ? rows.length + ' 条匹配 / ' + state.memory.length + ' 条整簿记忆' : '';
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

  // Replacement only ever touches the translation column: the source is the
  // reference, and the preview must show target state only.
  function batchColumns() {
    return [{ kind: 'target', letter: state.settings.targetCol }];
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
          writable: cols[c].kind === 'target',
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
        ? '预览完成：译文列命中 ' + hits.length + ' 处，确认后执行。'
        : '译文列没有匹配的单元格。';
    });
  }

  function applyBatch() {
    // Source text is read-only in this tool: matches there are shown in the
    // preview but never written back.
    var hits = state.batchPreview || [];
    if (!hits.length) { $('repStatus').textContent = '没有可替换的译文命中。'; return; }
    if (!window.confirm('将修改 ' + hits.length + ' 个译本单元格，确定执行？')) return;
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

  function segKey(sheet, row, column) {
    return String(sheet) + '!' + (column || state.settings.targetCol) + row;
  }

  function ensureProject() {
    if (!state.project) state.project = { version: 1, segments: {} };
    if (!state.project.segments) state.project.segments = {};
    return state.project;
  }

  function segmentRecord(sheet, row, sourceText, targetText, column) {
    var project = ensureProject();
    var key = segKey(sheet, row, column);
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

  function recordRevision(sheet, row, sourceText, before, after, column) {
    if (String(before == null ? '' : before) === String(after == null ? '' : after)) return;
    var record = segmentRecord(sheet, row, sourceText, after, column);
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
      var match = /^(.*)!([A-Z]+)(\d+)$/.exec(key);
      if (!match) return;
      var sheet = sheetByName(match[1]);
      var address = match[2] + match[3];
      var current = '';
      try { current = String(sheet.Range(address).Value2 == null ? '' : sheet.Range(address).Value2); }
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
      var address = key.split('!')[1];
      tr.appendChild(textNode('td', address, 'num'));
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
      var match = /^(.*)!([A-Z]+)(\d+)$/.exec(key);
      if (match) {
        var sheet = sheetByName(match[1]);
        try { sheet.Range(match[2] + match[3]).Value2 = rev.before; } catch (err) { /* cell gone */ }
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
      tr.classList.add('rowlink');
      tr.addEventListener('click', function (event) {
        if (event.target && event.target.tagName === 'BUTTON') return;
        if (!jumpToText(hit.entry.source, 'source', $('tmStats'))) {
          jumpToText(hit.entry.target, 'target', $('tmStats'));
        }
      });
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
    var locale = window.TideLangs ? window.TideLangs.get(state.settings.targetLang || 'en') : { tmx: 'en-US' };
    var xml = window.TideTM.buildTmx(state.tm, 'zh-CN', locale.tmx);
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

  function runQA(mode) {
    var result = buildRows();
    if (result.error) { $('qaSummary').textContent = '未连接表格。'; return; }
    var lang = state.settings.targetLang || 'en';
    var options = { lang: lang };
    var label = '全部检查';
    if (mode === 'punct') {
      options.only = ['punctuation', 'punct-format'];
      label = '标点符号';
    } else if (mode === 'placeholder') {
      options.checkPlaceholders = true;
      options.only = ['placeholder'];
      label = '占位符';
    }
    state.qaMode = label;
    var findings = window.TideQA.runQA(result.rows, state.matchTerms, options);
    state.qaFindings = findings;
    var summary = window.TideQA.summarize(findings);
    var parts = Object.keys(summary.byCategory).sort().map(function (key) {
      return key + ' ' + summary.byCategory[key];
    });
    $('qaCount').textContent = findings.length ? findings.length + ' 个问题' : '无问题';
    $('qaSummary').textContent = (findings.length
      ? '「' + label + '」「' + result.sheetName + '」源文 ' + state.settings.sourceCol + ' / 译文 ' + state.settings.targetCol +
        '，表头至第 ' + (result.headerEnd + 1) + ' 行，' + result.rows.length + ' 条：高 ' + summary.bySeverity.high +
        ' · 中 ' + summary.bySeverity.medium + ' · 低 ' + summary.bySeverity.low +
        '（' + parts.join('，') + '）'
      : '「' + label + '」「' + result.sheetName + '」' + result.rows.length + ' 条全部通过。');
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
      tr.appendChild(renderQaMessage(finding));
      tr.classList.add('rowlink');
      tr.addEventListener('click', function () {
        var sheetName = finding.sheet || state.sheetName;
        jumpToRow(finding.row, sheetName, $('qaSummary'));
        $('qaSummary').textContent = '第 ' + finding.row + ' 行 · ' + $('qaSummary').textContent +
          ' · ' + finding.category + '：' + finding.message.slice(0, 40);
      });
      body.appendChild(tr);
    });
  }

  function escapeRegExp(text) {
    return String(text).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  function markTokens(text, tokens) {
    var node = document.createElement('span');
    var list = (tokens || []).filter(Boolean);
    if (!list.length) { node.textContent = text; return node; }
    var re = new RegExp('(' + list.map(escapeRegExp).join('|') + ')');
    String(text).split(re).forEach(function (part) {
      if (!part) return;
      if (list.indexOf(part) > -1) {
        var mark = document.createElement('mark');
        mark.className = 'qa-mark';
        mark.textContent = part;
        node.appendChild(mark);
      } else {
        node.appendChild(document.createTextNode(part));
      }
    });
    return node;
  }

  function renderQaMessage(finding) {
    var cell = document.createElement('td');
    var text = document.createElement('div');
    text.appendChild(markTokens(finding.message, finding.highlight));
    cell.appendChild(text);
    if (finding.target && finding.highlight && finding.highlight.length) {
      var snippet = document.createElement('div');
      snippet.className = 'qa-snippet';
      snippet.appendChild(markTokens(String(finding.target).slice(0, 60), finding.highlight));
      cell.appendChild(snippet);
    }
    return cell;
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
    var record = project.segments[segKey(row.sheet, row.row, row.targetCol)];
    return record ? pendingCount(record) : 0;
  }

  function writeGridCell(row, value) {
    var sheet = sheetByName(row.sheet);
    if (!sheet) { $('gridHint').textContent = '未连接表格。'; return false; }
    if (row.blocked) { $('gridHint').textContent = '该组没有 EN 列，右侧被下一组占用，不能写入。'; return false; }
    var before = row.target;
    archivePrevious(sheet, row.row, before, row.oldCol);
    try {
      sheet.Range((row.targetCol || state.settings.targetCol) + row.row).Value2 = value;
    } catch (err) {
      $('gridHint').textContent = '写入失败：' + err.message;
      return false;
    }
    row.target = value;
    rememberPair(row.sheet, row.row, row.source, value);
    recordRevision(row.sheet, row.row, row.source, before, value, row.targetCol);
    recordWrite({ sheet: row.sheet, row: row.row, source: row.source }, before, value);
    return true;
  }

  function locateGridRow(row) {
    var sheet = sheetByName(row.sheet);
    if (!sheet) { $('gridHint').textContent = '未连接表格。'; return; }
    var address = (row.targetCol || state.settings.targetCol) + row.row;
    var ok = revealCell(sheet, address);
    selectGridRow(row.row);
    $('gridHint').textContent = ok
      ? ('已定位 ' + row.sheet + '!' + address + '（译文单元格）')
      : ('定位失败：' + row.sheet + '!' + address);
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
      tr.appendChild(textNode('td', (row.targetCol || state.settings.targetCol) + row.row, 'num'));
      var tdSource = document.createElement('td');
      tdSource.appendChild(textNode('div', row.source.slice(0, 60), 'zh'));
      tr.appendChild(tdSource);
      var tdTarget = document.createElement('td');
      var input = document.createElement('input');
      input.className = 'grid-cell';
      input.value = row.target || '';
      if (row.missingTarget) input.placeholder = 'BLANK';
      if (row.blocked) {
        input.disabled = true;
        input.title = '该组没有 EN 列，右侧是下一组的表头，不能写入';
      }
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
      tr.title = '点击定位到表格中的译文单元格';
      tr.addEventListener('mousedown', function () { selectGridRow(row.row); });
      tr.addEventListener('click', function (event) {
        if (event.target && event.target.tagName === 'INPUT') return;
        locateGridRow(row);
      });
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

  // ---------- project profiles ----------
  // Reserved entry: the selector exists and switching applies a table rule,
  // but profile management (create/edit) is intentionally not wired yet.

  function loadProfiles() {
    var saved = readJSON(K.profiles, null);
    if (saved && saved.length) state.profiles = saved;
    else state.profiles = [{ id: 'default', name: '当前项目', note: '', columns: JSON.parse(JSON.stringify(state.settings)) }];
    state.activeProfile = window.localStorage.getItem('tide.activeProfile') || state.profiles[0].id;
    if (!state.profiles.some(function (p) { return p.id === state.activeProfile; })) {
      state.activeProfile = state.profiles[0].id;
    }
  }

  function persistProfiles() {
    window.localStorage.setItem(K.profiles, JSON.stringify(state.profiles));
    window.localStorage.setItem('tide.activeProfile', state.activeProfile);
  }

  function renderProjects() {
    var select = $('projectSelect');
    if (!select) return;
    select.textContent = '';
    state.profiles.forEach(function (profile) {
      var option = document.createElement('option');
      option.value = profile.id;
      option.textContent = profile.name;
      option.selected = profile.id === state.activeProfile;
      select.appendChild(option);
    });
  }

  function applyProfile(id) {
    var profile = state.profiles.find(function (p) { return p.id === id; });
    if (!profile) return;
    state.activeProfile = id;
    state.settings = Object.assign({}, state.settings, profile.columns || {});
    $('colSource').value = state.settings.sourceCol;
    $('colTarget').value = state.settings.targetCol;
    $('colUid').value = state.settings.uidCol;
    if ($('colOld')) $('colOld').value = state.settings.oldCol || 'E';
    if ($('targetLang') && window.TideLangs) {
      var select = $('targetLang');
      select.textContent = '';
      window.TideLangs.list.forEach(function (lang) {
        var option = document.createElement('option');
        option.value = lang.code;
        option.textContent = lang.name + (lang.native && lang.native !== lang.name ? ' · ' + lang.native : '');
        option.selected = lang.code === (state.settings.targetLang || 'en');
        select.appendChild(option);
      });
    }
    writeJSON(K.settings, state.settings);
    persistProfiles();
    renderProjects();
    loadEntries();
    renderEntry();
  }

  function saveColumnsToProfile() {
    var profile = state.profiles.find(function (p) { return p.id === state.activeProfile; });
    if (!profile) return;
    profile.columns = JSON.parse(JSON.stringify(state.settings));
    persistProfiles();
  }

  // ---------- content scrollbars (vertical + horizontal) ----------

  function updateMainScroll() {
    var el = $('mainScroll');
    if (!el) return;
    updateRail('V', el.scrollHeight, el.clientHeight, el.scrollTop, el);
    updateRail('H', el.scrollWidth, el.clientWidth, el.scrollLeft, el);
  }

  function updateRail(axis, total, view, offset) {
    var rail = $('mainRail' + axis);
    var thumb = $('mainThumb' + axis);
    if (!rail || !thumb) return;
    var overflow = total > view + 1;
    rail.classList.toggle('on', overflow);
    if (!overflow) return;
    var track = axis === 'V' ? rail.clientHeight : rail.clientWidth;
    var size = Math.max(30, Math.round(track * (view / total)));
    var max = track - size;
    var progress = offset / Math.max(1, total - view);
    if (axis === 'V') {
      thumb.style.height = size + 'px';
      thumb.style.top = Math.round(max * progress) + 'px';
    } else {
      thumb.style.width = size + 'px';
      thumb.style.left = Math.round(max * progress) + 'px';
    }
  }

  function wireMainScroll() {
    var el = $('mainScroll');
    if (!el) return;
    el.addEventListener('scroll', updateMainScroll);
    window.addEventListener('resize', updateMainScroll);
    ['V', 'H'].forEach(function (axis) {
      var rail = $('mainRail' + axis);
      var thumb = $('mainThumb' + axis);
      if (!rail || !thumb) return;
      var dragging = false;
      var start = 0;
      var startOffset = 0;
      thumb.addEventListener('mousedown', function (event) {
        dragging = true;
        start = axis === 'V' ? event.clientY : event.clientX;
        startOffset = axis === 'V' ? el.scrollTop : el.scrollLeft;
        event.preventDefault();
      });
      document.addEventListener('mousemove', function (event) {
        if (!dragging) return;
        var track = axis === 'V' ? rail.clientHeight : rail.clientWidth;
        var size = axis === 'V' ? thumb.offsetHeight : thumb.offsetWidth;
        var max = track - size;
        if (max <= 0) return;
        var total = axis === 'V' ? el.scrollHeight - el.clientHeight : el.scrollWidth - el.clientWidth;
        var delta = ((axis === 'V' ? event.clientY : event.clientX) - start) * (total / max);
        if (axis === 'V') el.scrollTop = startOffset + delta;
        else el.scrollLeft = startOffset + delta;
      });
      document.addEventListener('mouseup', function () { dragging = false; });
      rail.addEventListener('mousedown', function (event) {
        if (event.target === thumb) return;
        var rect = rail.getBoundingClientRect();
        var ratio = axis === 'V'
          ? (event.clientY - rect.top) / rect.height
          : (event.clientX - rect.left) / rect.width;
        if (axis === 'V') el.scrollTop = ratio * (el.scrollHeight - el.clientHeight);
        else el.scrollLeft = ratio * (el.scrollWidth - el.clientWidth);
      });
    });
    if (window.MutationObserver) {
      var timer = 0;
      new MutationObserver(function () {
        window.clearTimeout(timer);
        timer = window.setTimeout(updateMainScroll, 60);
      }).observe(el, { childList: true, subtree: true, characterData: true });
    }
    updateMainScroll();
  }

  // ---------- tab strip scrollbar ----------

  function updateTabScroll() {
    var nav = $('tabs');
    var bar = $('tabScroll');
    var thumb = $('tabThumb');
    if (!nav || !bar || !thumb) return;
    var overflow = nav.scrollWidth > nav.clientWidth + 1;
    bar.classList.toggle('on', overflow);
    if (!overflow) return;
    var trackWidth = bar.clientWidth;
    var width = Math.max(32, Math.round(trackWidth * (nav.clientWidth / nav.scrollWidth)));
    var maxLeft = trackWidth - width;
    var progress = nav.scrollLeft / Math.max(1, nav.scrollWidth - nav.clientWidth);
    thumb.style.width = width + 'px';
    thumb.style.left = Math.round(maxLeft * progress) + 'px';
  }

  function wireTabScroll() {
    var nav = $('tabs');
    var thumb = $('tabThumb');
    var bar = $('tabScroll');
    if (!nav || !thumb || !bar) return;
    nav.addEventListener('scroll', updateTabScroll);
    window.addEventListener('resize', updateTabScroll);
    var dragging = false;
    var startX = 0;
    var startLeft = 0;
    thumb.addEventListener('mousedown', function (event) {
      dragging = true;
      startX = event.clientX;
      startLeft = nav.scrollLeft;
      event.preventDefault();
    });
    document.addEventListener('mousemove', function (event) {
      if (!dragging) return;
      var maxLeft = bar.clientWidth - thumb.offsetWidth;
      if (maxLeft <= 0) return;
      var delta = (event.clientX - startX) * ((nav.scrollWidth - nav.clientWidth) / maxLeft);
      nav.scrollLeft = startLeft + delta;
    });
    document.addEventListener('mouseup', function () { dragging = false; });
    bar.addEventListener('mousedown', function (event) {
      if (event.target === thumb) return;
      var rect = bar.getBoundingClientRect();
      var ratio = (event.clientX - rect.left) / rect.width;
      nav.scrollLeft = ratio * (nav.scrollWidth - nav.clientWidth);
    });
    updateTabScroll();
    updateMainScroll();
  }

  // ---------- key lookup ----------

  // Keys live in their own columns ("完整Key" / "完整Key（复制用）"), which are
  // unrelated to the translation columns, so they are detected from the header.
  function keyColumns(grid) {
    var limit = Math.min(grid.length, 12);
    for (var r = 0; r < limit; r += 1) {
      var line = grid[r] || [];
      var keyIdx = -1;
      var copyIdx = -1;
      for (var c = 0; c < line.length; c += 1) {
        var text = String(line[c] == null ? '' : line[c]).trim();
        if (!text) continue;
        if (!/完整\s*key/i.test(text)) continue;
        if (/复制/.test(text)) { if (copyIdx < 0) copyIdx = c; }
        else if (keyIdx < 0) keyIdx = c;
      }
      if (keyIdx > -1 || copyIdx > -1) return { keyIdx: keyIdx, copyIdx: copyIdx, headerRow: r };
    }
    return null;
  }

  function collectKeyRows(sheet) {
    var used;
    try { used = sheet.UsedRange; } catch (err) { return null; }
    if (!used) return null;
    var grid = toGrid(used.Value2);
    if (!grid.length) return null;
    var startRow = Number(used.Row) || 1;
    var startCol = Number(used.Column) || 1;
    var cols = keyColumns(grid);
    if (!cols) return null;
    var headerAt = Math.max(detectHeaderEnd(grid), cols.headerRow);
    var srcIdx = colToNumber(state.settings.sourceCol) - startCol;
    var tgtIdx = colToNumber(state.settings.targetCol) - startCol;
    var rows = [];
    for (var r = headerAt + 1; r < grid.length; r += 1) {
      var key = cols.keyIdx > -1 ? cellText(grid, r, cols.keyIdx).trim() : '';
      var copyKey = cols.copyIdx > -1 ? cellText(grid, r, cols.copyIdx).trim() : '';
      if (!key && !copyKey) continue;
      rows.push({
        row: startRow + r,
        sheet: sheet.Name,
        key: key,
        copyKey: copyKey,
        source: cols.keyIdx > -1 ? (srcIdx >= 0 ? cellText(grid, r, srcIdx) : '') : '',
        target: tgtIdx >= 0 ? cellText(grid, r, tgtIdx) : ''
      });
    }
    return rows;
  }

  function keyScore(item, query) {
    var q = query.toLowerCase();
    var key = (item.key || '').toLowerCase();
    var copy = (item.copyKey || '').toLowerCase();
    if (key === q || copy === q) return 1;
    if (key.indexOf(q) > -1 || copy.indexOf(q) > -1) return 0.9;
    var best = Math.max(window.TideEngine.similarity(item.key || '', query),
                        window.TideEngine.similarity(item.copyKey || '', query));
    return best;
  }

  function runKeySearch() {
    var query = String($('keyQuery').value || '').trim();
    var table = $('keyTable');
    var body = table.tBodies[0];
    body.textContent = '';
    if (!query) {
      table.hidden = true;
      $('keyHint').textContent = '输入 Key 或片段后点「搜索」。';
      $('keyCount').textContent = '';
      return;
    }
    var app = sheetApp();
    if (!app || !app.ActiveWorkbook) { $('keyHint').textContent = '未连接工作簿。'; return; }

    var sheets = app.ActiveWorkbook.Worksheets;
    var whole = $('keyWholeBook').checked;
    var count = 0;
    try { count = sheets.Count; } catch (err) { count = 0; }
    var targets = [];
    if (whole) {
      for (var i = 1; i <= count; i += 1) targets.push(sheets.Item(i));
    } else {
      targets.push(currentSheet());
    }

    var found = [];
    var missing = [];
    var exactOnly = $('keyExactOnly').checked;
    targets.forEach(function (sheet) {
      if (!sheet) return;
      var rows = collectKeyRows(sheet);
      if (!rows) { missing.push(sheet.Name); return; }
      rows.forEach(function (item) {
        var score = keyScore(item, query);
        if (score >= 0.5 && (!exactOnly || score === 1)) {
          found.push({ item: item, score: score, exact: score === 1 });
        }
      });
    });

    found.sort(function (a, b) { return (b.exact - a.exact) || (b.score - a.score); });
    var shown = found.slice(0, 200);
    var exactCount = found.filter(function (hit) { return hit.exact; }).length;
    table.hidden = shown.length === 0;
    $('keyCount').textContent = found.length
      ? '精确 ' + exactCount + ' · 近似 ' + (found.length - exactCount)
      : '无匹配';
    $('keyHint').textContent = missing.length
      ? '这些表没有「完整Key」列，已跳过：' + missing.join('、')
      : '点任意结果跳到对应词条。';

    shown.forEach(function (hit) {
      var item = hit.item;
      var tr = document.createElement('tr');
      var posCell = textNode('td', item.sheet + '!' + item.row, 'num');
      var badge = document.createElement('span');
      badge.className = 'match-badge' + (hit.exact ? ' exact' : '');
      badge.textContent = hit.exact ? '精确' : Math.round(hit.score * 100) + '%';
      posCell.appendChild(badge);
      tr.appendChild(posCell);
      var keyCell = document.createElement('td');
      var keyLine = document.createElement('div');
      keyLine.className = 'keycell';
      keyLine.textContent = item.key || '—';
      keyCell.appendChild(keyLine);
      if (item.copyKey && item.copyKey !== item.key) {
        var copyLine = document.createElement('div');
        copyLine.className = 'copycell';
        copyLine.textContent = '复制用：' + item.copyKey;
        keyCell.appendChild(copyLine);
      }
      tr.appendChild(keyCell);
      var textCell = document.createElement('td');
      var src = document.createElement('div');
      src.textContent = (item.source || '').slice(0, 40) || '—';
      textCell.appendChild(src);
      var tgt = document.createElement('div');
      tgt.className = 'hint';
      tgt.textContent = (item.target || '').slice(0, 40) || '（未译）';
      textCell.appendChild(tgt);
      tr.appendChild(textCell);
      tr.addEventListener('click', function () {
        var sheet = sheetByName(item.sheet);
        var address = (item.source && state.settings.sourceCol ? state.settings.sourceCol : state.settings.sourceCol) + item.row;
        var ok = revealCell(sheet, address);
        $('keyHint').textContent = ok
          ? '已定位 ' + item.sheet + '!' + address
          : '定位失败：' + item.sheet + '!' + address;
      });
      body.appendChild(tr);
    });
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
  // WPS blocks clipboard reads in the pane, but a drag from another app still
  // delivers the text through dataTransfer.
  function wireDrop(el, kind) {
    if (!el) return;
    el.addEventListener('dragover', function (event) {
      event.preventDefault();
      el.classList.add('dropping');
    });
    el.addEventListener('dragleave', function () { el.classList.remove('dropping'); });
    el.addEventListener('drop', function (event) {
      event.preventDefault();
      el.classList.remove('dropping');
      var text = event.dataTransfer ? event.dataTransfer.getData('text') : '';
      if (!text) return;
      el.value = text;
      el.dispatchEvent(new Event('input', { bubbles: true }));
      $('entryCheck').textContent = (kind === 'source' ? '已拖入原文 ' : '已拖入译文 ') + text.length + ' 个字符';
    });
  }

  // Diagnostics go to the local channel server so a real WPS session can be
  // inspected without a devtools window.
  // Every free-text field gets the same pointer affordances as the draft:
  // the host swallows the clipboard shortcuts, so paste must be clickable.
  function decorateFields(ids) {
    ids.forEach(function (id) {
      var input = $(id);
      if (!input || input.dataset.tideActions) return;
      input.dataset.tideActions = '1';
      if (input.style && input.style.width) input.style.width = 'auto';

      var row = document.createElement('div');
      row.className = 'field-row' + (input.tagName === 'TEXTAREA' ? ' tall' : '');
      input.parentNode.insertBefore(row, input);
      row.appendChild(input);

      var actions = document.createElement('div');
      actions.className = 'field-buttons';
      actions.innerHTML =
        '<button type="button" class="act mini" data-paste-for="' + id + '" title="粘贴剪贴板内容">粘贴</button>' +
        '<button type="button" class="act ghost mini" data-clear-for="' + id + '" title="清空">清空</button>';
      row.appendChild(actions);
    });
  }

  function showToast(text) {
    var toast = $('tideToast');
    if (!toast) {
      toast = document.createElement('div');
      toast.id = 'tideToast';
      toast.className = 'tide-toast';
      document.body.appendChild(toast);
    }
    toast.textContent = text;
    toast.classList.add('on');
    window.clearTimeout(showToast.timer);
    showToast.timer = window.setTimeout(function () { toast.classList.remove('on'); }, 1800);
  }

  function paneLog(event, data) {
    try {
      if (!window.fetch) return;
      fetch('/__log', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ ts: new Date().toISOString(), event: event, data: data || {} })
      }).catch(function () { /* ignore */ });
    } catch (err) { /* ignore */ }
  }

  // execCommand('paste') is the last bridge that can work when the async
  // clipboard API is blocked: CEF may still route it through the host.
  function execPaste() {
    var holder = document.createElement('div');
    holder.contentEditable = 'true';
    holder.style.cssText = 'position:fixed;left:-9999px;top:0;width:1px;height:1px;opacity:0';
    document.body.appendChild(holder);
    holder.focus();
    var ok = false;
    try { ok = document.execCommand('paste'); } catch (err) { ok = false; }
    var text = holder.textContent || '';
    document.body.removeChild(holder);
    return ok && text ? text : '';
  }

  // One click, three strategies, and a message that names the failing layer so
  // the next report says exactly which one WPS blocks.
  function pasteWithDiagnostics(target) {
    var el = target || $('entryDraft');
    // Strategy 0: same-origin clipboard bridge (works inside CEF, where the
    // clipboard API is denied). Falls through to the browser APIs elsewhere.
    var bridge = function () {
      return fetch('/__clipboard', { cache: 'no-store' }).then(function (res) {
        if (!res.ok) throw new Error('bridge ' + res.status);
        return res.json();
      }).then(function (data) {
        if (!data || !data.text) return '';
        // base64 of UTF-8 bytes: decoding here avoids CEF's locale-based guess
        var raw = atob(data.text);
        var bytes = new Uint8Array(raw.length);
        for (var i = 0; i < raw.length; i += 1) bytes[i] = raw.charCodeAt(i);
        return new TextDecoder('utf-8').decode(bytes);
      });
    };
    paneLog('paste:start', {
      target: el ? el.id : null,
      hasClipboard: !!(navigator.clipboard),
      hasReadText: !!(navigator.clipboard && navigator.clipboard.readText),
      hasFocus: document.hasFocus(),
      active: document.activeElement ? document.activeElement.id || document.activeElement.tagName : null
    });
    if (!el) return;
    var insert = function (text) {
      el.value = text;
      el.dispatchEvent(new Event('input', { bubbles: true }));
      paneLog('paste:inserted', { length: text.length, via: window.__tidePasteVia || 'clipboard', target: el.id || el.tagName });
      var note = '已粘贴 ' + text.length + ' 个字符（来源：' + (window.__tidePasteVia || 'clipboard') + '）。';
      if (el.id === 'entryDraft') $('entryCheck').textContent = note;
      else showToast(note);
    };
    if (location.protocol === 'http:' || location.protocol === 'https:') {
      bridge().then(function (text) {
        paneLog('paste:bridge', { length: text ? text.length : 0 });
        if (text) { window.__tidePasteVia = '本机通道'; insert(text); return; }
        throw new Error('empty');
      }).catch(function (err) {
        paneLog('paste:bridge-failed', { message: String(err && err.message).slice(0, 80) });
        pasteViaBrowser(el, insert);
      });
      return;
    }
    pasteViaBrowser(el, insert);
  }

  function pasteViaBrowser(el, insert) {
    if (navigator.clipboard && navigator.clipboard.readText) {
      navigator.clipboard.readText().then(function (text) {
        paneLog('paste:readText', { ok: true, length: text ? text.length : 0 });
        if (text) { window.__tidePasteVia = 'clipboard API'; insert(text); return; }
        var fallback = execPaste();
        if (fallback) { window.__tidePasteVia = 'execCommand'; insert(fallback); return; }
        $('entryCheck').textContent = '剪贴板为空，或 WPS 未授权读取。请用右键→粘贴。';
      }).catch(function (err) {
        paneLog('paste:readText', { ok: false, name: err && err.name, message: String(err && err.message).slice(0, 120) });
        var fallback = execPaste();
        paneLog('paste:execCommand', { length: fallback ? fallback.length : 0 });
        if (fallback) { window.__tidePasteVia = 'execCommand'; insert(fallback); return; }
        $('entryCheck').textContent = '剪贴板读取被 WPS 拒绝（clipboard API 与 execCommand 均失败）。请用右键→粘贴，或「读取选区→译文」。';
      });
      return;
    }
    var text = execPaste();
    if (text) { window.__tidePasteVia = 'execCommand'; insert(text); return; }
    $('entryCheck').textContent = '此环境没有可用的剪贴板接口。请用右键→粘贴。';
  }

  function pasteIntoActive() {
    var el = document.activeElement;
    var target = (el && (el.tagName === 'TEXTAREA' || el.tagName === 'INPUT')) ? el : $('entryDraft');
    if (!target) return;
    if (!navigator.clipboard || !navigator.clipboard.readText) {
      $('entryCheck').textContent = '读不到剪贴板：请用右键→粘贴，或「读取选区→译文」。';
      return;
    }
    navigator.clipboard.readText().then(function (text) {
      if (!text) return;
      target.value = text;
      target.dispatchEvent(new Event('input', { bubbles: true }));
      $('entryCheck').textContent = '已粘贴 ' + text.length + ' 个字符。';
    }).catch(function () {
      $('entryCheck').textContent = '剪贴板读取被拒绝：请用右键→粘贴，或「读取选区→译文」。';
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
    archivePrevious(entrySheet(entry), entry.row, before, entry.oldCol);
    if (!writeCell(entry, value)) return;
    var address = entryAddress(entry.row);
    entry.target = value;
    refreshMemory();
    recordWrite(entry, before, value);
    rememberPair(entry.sheet, entry.row, entry.source, value);
    recordRevision(entry.sheet, entry.row, entry.source, before, value, entry.targetCol);
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
    if ($('docName')) $('docName').textContent = result.sheetName || '当前工作簿';
  }

  function nasHeaders() {
    return { 'Authorization': 'Bearer ' + state.nas.token };
  }

  function nasURL(path) {
    return state.nas.base.replace(/\/+$/, '') + path;
  }

  function readNasAuth() {
    try {
      var raw = localStorage.getItem(K.nasAuth);
      if (!raw) return null;
      var auth = JSON.parse(raw);
      return auth && auth.token ? auth : null;
    } catch (err) { return null; }
  }

  function saveNasAuth() {
    try {
      localStorage.setItem(K.nasAuth, JSON.stringify({ token: state.nas.token, username: state.nas.user }));
    } catch (err) { /* storage may be unavailable */ }
  }

  function clearNasAuth() {
    state.nas.token = '';
    state.nas.user = '';
    state.nas.loaded = false;
    state.nas.projects = [];
    state.nas.items = [];
    if (state.nas.jumpAbort) {
      try { state.nas.jumpAbort.abort(); } catch (err) { /* ignore */ }
      state.nas.jumpAbort = null;
    }
    try { localStorage.removeItem(K.nasAuth); } catch (err) { /* ignore */ }
    renderNasAuthUI();
  }

  function renderNasAuthUI() {
    var logged = !!state.nas.token;
    if ($('nasLoginBox')) $('nasLoginBox').hidden = logged;
    if ($('nasSession')) $('nasSession').hidden = !logged;
    if ($('nasWho')) $('nasWho').textContent = logged ? (state.nas.user || '') : '';
    if ($('nasTermsCard')) $('nasTermsCard').hidden = !logged;
  }

  async function nasFetch(path) {
    var response = await fetch(nasURL(path), { headers: nasHeaders() });
    var data = await response.json().catch(function () { return {}; });
    if (response.status === 401) {
      clearNasAuth();
      throw new Error('登录已失效，请重新登录');
    }
    if (!response.ok) throw new Error(data.error || ('HTTP ' + response.status));
    return data;
  }

  function nasFillSelect(select, values, current, labelFn) {
    select.textContent = '';
    values.forEach(function (value) {
      var option = document.createElement('option');
      option.value = value;
      option.textContent = labelFn ? labelFn(value) : value;
      option.selected = value === current;
      select.appendChild(option);
    });
  }

  async function loadNasTerms(pid, source) {
    var card = $('nasTermsCard'), list = $('nasTermList'), count = $('nasTermCount');
    if (!card || !list) return;
    if (!state.nas.token || !pid) { card.hidden = true; return; }
    var terms = state.nas.termsCache[pid];
    if (!terms) {
      try {
        var data = await nasFetch('/api/terms?project_id=' + encodeURIComponent(pid));
        terms = data.terms || [];
        state.nas.termsCache[pid] = terms;
      } catch (err) { card.hidden = true; return; }
    }
    card.hidden = false;
    var hay = String(source || '').toLowerCase();
    var hits = hay ? terms.filter(function (t) {
      return t.source && hay.indexOf(String(t.source).toLowerCase()) >= 0;
    }) : terms.slice(0, 8);
    list.textContent = '';
    if (count) count.textContent = hay ? (hits.length + ' / ' + terms.length) : (terms.length + ' 条');
    if (!terms.length) {
      var empty = document.createElement('div'); empty.className = 'nas-term-empty';
      empty.textContent = '这个项目还没有术语。'; list.appendChild(empty); return;
    }
    if (!hits.length) {
      var miss = document.createElement('div'); miss.className = 'nas-term-empty';
      miss.textContent = '当前片段未命中术语表；项目共 ' + terms.length + ' 条。'; list.appendChild(miss); return;
    }
    hits.slice(0, 12).forEach(function (t) {
      var row = document.createElement('div'); row.className = 'nas-term-row';
      var src = document.createElement('span'); src.className = 'nas-term-src'; src.textContent = t.source || '';
      var arrow = document.createElement('span'); arrow.className = 'nas-term-arrow'; arrow.textContent = '→';
      var tgt = document.createElement('span'); tgt.className = 'nas-term-tgt'; tgt.textContent = t.target || '';
      var badge = document.createElement('span'); badge.className = 'nas-term-badge' + (t.reviewed ? ' ok' : '');
      badge.textContent = t.reviewed ? '已审' : '待审';
      row.append(src, arrow, tgt, badge);
      row.title = t.source + ' → ' + t.target + (t.note ? '\n' + t.note : '');
      row.addEventListener('click', function () { nasLocate(t.source, t.target); });
      list.appendChild(row);
    });
  }

  function nasConsoleURL() {
    try {
      return new URL(state.nas.base).origin + "/";
    } catch (err) {
      return "https://192.168.23.232:8443/";
    }
  }

  function openExternalURL(url, hint, label) {
    if (hint) hint.textContent = label || ('正在打开：' + url);
    try {
      if (typeof wps !== 'undefined' && wps && wps.OAAssist && typeof wps.OAAssist.ShellExecute === 'function') {
        wps.OAAssist.ShellExecute(url);
        return;
      }
    } catch (err) { /* fall through */ }
    try {
      var win = window.open(url, '_blank');
      if (!win && hint) hint.textContent = '浏览器被拦截，请手动打开：' + url;
    } catch (err) {
      if (hint) hint.textContent = '请手动打开：' + url;
    }
  }

  async function openNasConsole() {
    var hint = $('nasConsoleHint');
    if (state.nas.token) {
      if (hint) hint.textContent = '正在验证登录…';
      try {
        var response = await fetch(nasURL('/api/wps/sso'), { method: 'POST', headers: nasHeaders() });
        var data = await response.json().catch(function () { return {}; });
        if (response.ok && data.path) {
          openExternalURL(new URL(data.path, state.nas.base).href, hint, '已登录，正在打开 NAS 操作台…');
          return;
        }
      } catch (err) { /* fall back to normal console */ }
    }
    openExternalURL(nasConsoleURL(), hint);
  }

  function renderNasServer() {
    var el = $('nasServerLabel');
    if (!el) return;
    try { el.textContent = new URL(state.nas.base).host; }
    catch (err) { el.textContent = state.nas.base; }
  }

  async function nasLogin() {
    var user = ($('nasUser').value || '').trim();
    var pass = $('nasPass').value || '';
    if (!user || !pass) { $('nasStatus').textContent = '请输入账号和密码。'; return; }
    $('nasStatus').textContent = '登录中…';
    try {
      var response = await fetch(state.nas.base + '/api/wps/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username: user, password: pass })
      });
      var data = await response.json().catch(function () { return {}; });
      if (!response.ok) throw new Error(data.error || ('HTTP ' + response.status));
      state.nas.token = data.token || '';
      state.nas.user = data.username || user;
      saveNasAuth();
      renderNasAuthUI();
      $('nasPass').value = '';
      $('nasStatus').textContent = '已登录 ' + state.nas.user;
      await nasConnect();
    } catch (err) {
      clearNasAuth();
      $('nasStatus').textContent = '登录失败：' + err.message;
    }
  }

  function nasLogout() {
    clearNasAuth();
    state.nas.pid = null;
    if ($('nasUser')) $('nasUser').value = '';
    if ($('nasPass')) $('nasPass').value = '';
    if ($('nasProject')) $('nasProject').textContent = '';
    if ($('nasLang')) $('nasLang').textContent = '';
    if ($('nasTable')) $('nasTable').hidden = true;
    if ($('nasPage')) $('nasPage').textContent = '';
    $('nasStatus').textContent = '已退出登录。';
  }

  async function nasConnect() {
    if (!state.nas.token) { $('nasStatus').textContent = '请先登录。'; return; }
    $('nasStatus').textContent = '连接中…';
    try {
      var data = await nasFetch('/api/projects');
      state.nas.projects = data.projects || [];
      nasFillSelect($('nasProject'), state.nas.projects.map(function (p) { return String(p.id); }),
        state.nas.pid ? String(state.nas.pid) : null,
        function (id) { var p = state.nas.projects.find(function (x) { return String(x.id) === id; }); return p ? p.name : id; });
      if (!state.nas.projects.length) {
        $('nasStatus').textContent = '已连接，但没有项目。';
        return;
      }
      state.nas.loaded = true;
      $('nasStatus').textContent = '已连接 · ' + state.nas.projects.length + ' 个项目';
      var first = state.nas.projects[0];
      $('nasProject').value = String(first.id);
      await nasLoadProject(first.id);
    } catch (err) {
      state.nas.loaded = false;
      $('nasStatus').textContent = '连接失败：' + err.message;
    }
  }

  async function nasLoadProject(pid, lang) {
    state.nas.pid = Number(pid);
    state.nas.offset = 0;
    state.nas.q = $('nasQuery') ? $('nasQuery').value.trim() : '';
    var query = '?id=' + encodeURIComponent(pid) + '&limit=50&offset=0';
    if (lang) query += '&lang=' + encodeURIComponent(lang);
    try {
      var data = await nasFetch('/api/projects/detail' + query);
      state.nas.langs = data.langs || [];
      state.nas.lang = data.lang;
      state.nas.items = data.segments || [];
      state.nas.total = data.filtered_total || 0;
      nasFillSelect($('nasLang'), state.nas.langs, state.nas.lang);
      renderNasRows();
      loadNasTerms(state.nas.pid, '');
    } catch (err) {
      $('nasStatus').textContent = '读取失败：' + err.message;
    }
  }

  async function nasLoadPage(offset) {
    if (!state.nas.pid) return;
    var query = '?id=' + encodeURIComponent(state.nas.pid) + '&limit=50&offset=' + offset;
    if (state.nas.lang) query += '&lang=' + encodeURIComponent(state.nas.lang);
    if (state.nas.q) query += '&q=' + encodeURIComponent(state.nas.q);
    try {
      var data = await nasFetch('/api/projects/detail' + query);
      state.nas.items = data.segments || [];
      state.nas.total = data.filtered_total || 0;
      state.nas.offset = offset;
      renderNasRows();
    } catch (err) {
      $('nasStatus').textContent = '读取失败：' + err.message;
    }
  }

  function renderNasRows() {
    var table = $('nasTable');
    var body = table.tBodies[0];
    body.textContent = '';
    var items = state.nas.items || [];
    table.hidden = items.length === 0;
    items.forEach(function (item) {
      var tr = document.createElement('tr');
      tr.title = '点击定位到 WPS 表格';
      var c0 = document.createElement('td'); c0.textContent = item.source; tr.appendChild(c0);
      var c1 = document.createElement('td'); c1.textContent = item.target || ''; tr.appendChild(c1);
      var c2 = document.createElement('td'); c2.textContent = item.status; tr.appendChild(c2);
      tr.addEventListener('click', function () {
        nasLocate(item.source, item.target);
        loadNasTerms(state.nas.pid, item.source);
      });
      body.appendChild(tr);
    });
    var start = state.nas.offset + 1;
    var end = state.nas.offset + items.length;
    $('nasPage').textContent = items.length ? (start + '–' + end + ' / ' + state.nas.total) : ('共 ' + state.nas.total + ' 条');
    $('btnNasPrev').disabled = state.nas.offset <= 0;
    $('btnNasNext').disabled = state.nas.offset + 50 >= state.nas.total;
  }

  function nasLocate(text, target) {
    var app = sheetApp();
    if (!app) { $('nasHint').textContent = '未连接 WPS 工作簿。'; return; }
    var book = app.ActiveWorkbook;
    var sheets = book && book.Worksheets;
    if (!sheets) { $('nasHint').textContent = '没有打开的工作簿。'; return; }
    var count = 0;
    try { count = sheets.Count; } catch (err) { count = 0; }
    var needle = String(text || '').trim();
    var wanted = String(target || '').trim();
    var fallback = null;
    for (var i = 1; i <= count; i += 1) {
      var sheet = null;
      try { sheet = sheets.Item(i); } catch (err) { continue; }
      var used = null;
      try { used = sheet.UsedRange; } catch (err2) { continue; }
      if (!used) continue;
      var grid = toGrid(used.Value2);
      var startRow = Number(used.Row) || 1;
      var startCol = Number(used.Column) || 1;
      for (var r = 0; r < grid.length; r += 1) {
        for (var c = 0; c < grid[r].length; c += 1) {
          if (String(cellText(grid, r, c)).trim() !== needle) continue;
          var address = sheet.Cells(startRow + r, startCol + c).Address;
          var targetHit = false;
          if (wanted) {
            for (var tc = 0; tc < grid[r].length; tc += 1) {
              if (String(cellText(grid, r, tc)).trim() === wanted) { targetHit = true; break; }
            }
          }
          if (!wanted || targetHit) {
            var ok = revealCell(sheet, address);
            $('nasHint').textContent = ok
              ? ('已精确定位 ' + sheet.Name + '!' + address)
              : ('定位失败：' + sheet.Name + '!' + address);
            return;
          }
          if (!fallback) fallback = { sheet: sheet, address: address };
        }
      }
    }
    if (fallback) {
      var fallbackOk = revealCell(fallback.sheet, fallback.address);
      $('nasHint').textContent = fallbackOk
        ? ('已按原文定位 ' + fallback.sheet.Name + '!' + fallback.address + '（译文未匹配）')
        : ('定位失败：' + fallback.sheet.Name + '!' + fallback.address);
      return;
    }
    $('nasHint').textContent = '当前工作簿未找到该原文。';
  }

  function startJumpWatch() {
    if (state.nas.jumpStarted) return;
    state.nas.jumpStarted = true;
    jumpStreamLoop();
  }

  async function jumpStreamLoop() {
    var failures = 0;
    while (true) {
      if (!state.nas.token) {
        await new Promise(function (resolve) { window.setTimeout(resolve, 1000); });
        continue;
      }
      try {
        var controller = new AbortController();
        state.nas.jumpAbort = controller;
        var response = await fetch(nasURL('/api/jump/stream'), {
          headers: nasHeaders(), cache: 'no-store', signal: controller.signal
        });
        if (!response.ok || !response.body || !response.body.getReader) throw new Error('stream unavailable');
        failures = 0;
        var reader = response.body.getReader();
        var decoder = new TextDecoder();
        var buffer = '';
        while (true) {
          var chunk = await reader.read();
          if (chunk.done) break;
          buffer += decoder.decode(chunk.value, { stream: true });
          var blocks = buffer.split('\n\n');
          buffer = blocks.pop() || '';
          for (var i = 0; i < blocks.length; i += 1) {
            var line = blocks[i].split('\n').find(function (x) { return x.indexOf('data:') === 0; });
            if (!line) continue;
            try {
              var data = JSON.parse(line.slice(5).trim());
              var jumps = data.jumps || [];
              if (jumps.length && !state.nas.jumpBusy) {
                state.nas.jumpBusy = true;
                await applyJump(jumps[jumps.length - 1]);
                state.nas.jumpBusy = false;
              }
            } catch (err) { /* malformed event */ }
          }
        }
      } catch (err) {
        failures += 1;
        state.nas.jumpAbort = null;
        await new Promise(function (resolve) { window.setTimeout(resolve, failures > 2 ? 4000 : 1200); });
        if (failures > 2) await pollJump();
      }
    }
  }

  async function pollJump() {
    if (!state.nas.token || state.nas.jumpBusy) return;
    try {
      var data = await nasFetch('/api/jump/pending');
      var jumps = data.jumps || [];
      if (jumps.length) {
        state.nas.jumpBusy = true;
        await applyJump(jumps[jumps.length - 1]);
      }
    } catch (err) { /* offline or logged out */ }
    finally { state.nas.jumpBusy = false; }
  }

  async function applyJump(j) {
    showTab('nas');
    state.nas.pid = j.project_id;
    state.nas.lang = j.lang || '';
    state.nas.q = j.source || '';
    if (!state.nas.loaded) await nasConnect();
    if ($('nasProject')) $('nasProject').value = String(j.project_id);
    if ($('nasLang')) $('nasLang').value = j.lang || '';
    await nasLoadPage(0);
    showTab('nas');
    nasLocate(j.source);
    loadNasTerms(j.project_id, j.source);
    $('nasStatus').textContent = '来自 OA：' + j.project_name + ' · ' + (j.job_name || '任务');
  }

  function renderNas() {
    renderNasServer();
    renderNasAuthUI();
    if (!state.nas.loaded && !state.nas.token) {
      $('nasStatus').textContent = '使用 NAS 账号登录后加载项目。';
      return;
    }
    if (!state.nas.loaded && state.nas.token) { nasConnect(); return; }
    if (state.nas.projects.length) {
      nasFillSelect($('nasProject'), state.nas.projects.map(function (p) { return String(p.id); }),
        state.nas.pid ? String(state.nas.pid) : null,
        function (id) { var p = state.nas.projects.find(function (x) { return String(x.id) === id; }); return p ? p.name : id; });
      nasFillSelect($('nasLang'), state.nas.langs, state.nas.lang);
    }
    renderNasRows();
  }

  function showTab(name) {
    var buttons = $('tabs').querySelectorAll('button');
    for (var i = 0; i < buttons.length; i += 1) buttons[i].classList.toggle('on', buttons[i].dataset.tab === name);
    ['matches', 'translate', 'keys', 'free', 'find', 'review', 'tm', 'grid', 'project', 'nas', 'terms', 'history', 'sync', 'settings'].forEach(function (tab) {
      $('panel-' + tab).hidden = tab !== name;
    });
    if (state.lastTab && state.lastTab !== name) state.prevTab = state.lastTab;
    state.lastTab = name;
    writeJSON(K.lastTab, name);
    updateTabScroll();
    updateMainScroll();
    var activeTab = $('tabs').querySelector('button.on');
    if (activeTab && activeTab.scrollIntoView) {
      try { activeTab.scrollIntoView({ inline: 'nearest', block: 'nearest' }); } catch (err) { /* older engine */ }
    }
    if (name === 'translate') { loadEntries(); renderEntry(); }
    if (name === 'free') { if (!state.memory.length) loadEntries(); renderFree(); }
    if (name === 'find') renderFind();
    if (name === 'history') { renderLock(); renderHistory(); }
    if (name === 'terms') { syncTermColumns(); renderTerms(); }
    if (name === 'review') renderReview();
    if (name === 'tm') renderTM();
    if (name === 'grid') renderGrid();
    if (name === 'project') renderProject();
    if (name === 'keys') { $('keyQuery').focus(); }
    if (name === 'sync') refreshHelper();
    if (name === 'nas') renderNas();
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
      var tab = event.target && event.target.closest ? event.target.closest('button[data-tab]') : null;
      if (tab) showTab(tab.dataset.tab);
    });
    $('btnBack').addEventListener('click', function () {
      var target = state.prevTab;
      if (target && target !== state.lastTab) showTab(target);
      else showTab('matches');
    });
    $('btnScan').addEventListener('click', scan);
    $('btnSave').addEventListener('click', function () {
      state.settings = {
        sourceCol: $('colSource').value.toUpperCase() || 'C',
        targetCol: $('colTarget').value.toUpperCase() || 'F',
        uidCol: $('colUid').value.toUpperCase() || 'A',
        oldCol: ($('colOld') && $('colOld').value.toUpperCase()) || 'E',
        targetLang: ($('targetLang') && $('targetLang').value) || state.settings.targetLang || 'en'
      };
      writeJSON(K.settings, state.settings);
      saveColumnsToProfile();
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
          var parsed = isSheet
            ? parseWorkbook(reader.result)
            : window.TideEngine.parseTermbase(JSON.parse(String(reader.result)));
          var langCode = state.settings.targetLang || 'en';
          var terms = parsed.map(function (term) {
            term.lang = term.lang || langCode;
            return term;
          });
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
        else { $('entryCheck').textContent = '复制被拒绝：请在表格里选中原文后按 ⌘C。'; }
      });
    });
    $('btnCopyDraft').addEventListener('click', function () {
      copyText($('entryDraft').value).then(function (ok) {
        if (ok) { $('entryCheck').textContent = '已复制草稿'; }
        else { $('entryCheck').textContent = '复制被拒绝：请手动选中文本后按 ⌘C。'; }
      });
    });
    $('btnReadCell').addEventListener('click', function () { readSelection('draft'); });
    $('btnReadCellSrc').addEventListener('click', function () { readSelection('source'); });
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
    $('btnKeySearch').addEventListener('click', runKeySearch);
    $('keyQuery').addEventListener('keydown', function (event) {
      if (event.key === 'Enter') runKeySearch();
    });
    $('btnGridRefresh').addEventListener('click', renderGrid);
    $('gridFilter').addEventListener('change', renderGrid);
    $('btnRunQA').addEventListener('click', function () { runQA('all'); });
    $('btnPunctQA').addEventListener('click', function () { runQA('punct'); });
    $('btnPlaceholderQA').addEventListener('click', function () { runQA('placeholder'); });
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
    wireTabScroll();
    // Cmd+V never reaches the webview (WPS owns it), but Ctrl+V is not a WPS
    // shortcut on macOS, so the page should receive it.
    window.addEventListener('keydown', function (event) {
      var el = document.activeElement;
      var editable = el && (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA');
      var combo = (event.metaKey ? '⌘' : '') + (event.ctrlKey ? 'Ctrl+' : '') + (event.altKey ? '⌥' : '') + 'V';
      if ((event.metaKey || event.ctrlKey || event.altKey) && (event.key === 'v' || event.key === 'V')) {
        // Reaching this line proves WPS forwarded the key. Focus can be any
        // search field, so paste into whatever is focused (draft by default).
        event.preventDefault();
        event.stopPropagation();
        if (editable) pasteWithDiagnostics(el);
        else pasteWithDiagnostics($('entryDraft'));
        return;
      }
      if (event.key === 'Escape' && editable && el.value) {
        el.value = '';
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
        showToast('已清空');
        event.preventDefault();
      }
    }, true);
    wireMainScroll();
    $('tabs').addEventListener('wheel', function (event) {
      if (event.deltaY && !event.deltaX) {
        this.scrollLeft += event.deltaY;
        event.preventDefault();
      }
    }, { passive: false });
    $('btnReviewRefresh').addEventListener('click', function () { loadProject().then(renderReview); });
    $('btnAddComment').addEventListener('click', addComment);
    decorateFields(['freeSource', 'freeDraft', 'findInput', 'repFind', 'repWith', 'commentInput',
      'tmQuery', 'keyQuery', 'termSearch']);
    document.addEventListener('click', function (event) {
      var pasteBtn = event.target.closest ? event.target.closest('[data-paste-for]') : null;
      if (pasteBtn) {
        pasteWithDiagnostics($(pasteBtn.dataset.pasteFor));
        return;
      }
      var clearBtn = event.target.closest ? event.target.closest('[data-clear-for]') : null;
      if (clearBtn) {
        var field = $(clearBtn.dataset.clearFor);
        if (!field || !field.value) return;
        field.value = '';
        field.dispatchEvent(new Event('input', { bubbles: true }));
        field.dispatchEvent(new Event('change', { bubbles: true }));
        showToast('已清空');
        field.focus();
      }
    });

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
      state.workbookMemory = null;
      state.workbookKey = '';
      state.workbookBuilding = false;
      loadEntries();
      renderEntry();
      $('entryCheck').textContent = '已刷新到 ' + activeSheetName() + '，正在重建整簿记忆…';
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
    $('btnDraftPaste').addEventListener('click', function () { pasteWithDiagnostics($('entryDraft')); });
    $('btnDraftClear').addEventListener('click', function () {
      var draft = $('entryDraft');
      if (!draft.value) return;
      draft.value = '';
      draft.dispatchEvent(new Event('input', { bubbles: true }));
      draft.focus();
      $('entryCheck').textContent = '已清空译文草稿。';
    });
    if ($('targetLang')) {
      $('targetLang').addEventListener('change', function () {
        state.settings.targetLang = $('targetLang').value;
        writeJSON(K.settings, state.settings);
        saveColumnsToProfile();
        refreshTermLanguage();
        $('entryCheck').textContent = '目标语言已切换为 ' + window.TideLangs.get(state.settings.targetLang).name;
      });
    }
    if ($('projectSelect')) $('projectSelect').addEventListener('change', function (event) { applyProfile(event.target.value); });
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
    $('btnNasLogin').addEventListener('click', nasLogin);
    $('btnNasConsole').addEventListener('click', openNasConsole);
    $('btnNasLogout').addEventListener('click', nasLogout);
    $('nasPass').addEventListener('keydown', function (event) { if (event.key === 'Enter') nasLogin(); });
    $('nasProject').addEventListener('change', function () { nasLoadProject($('nasProject').value); });
    $('nasLang').addEventListener('change', function () { nasLoadProject(state.nas.pid, $('nasLang').value); });
    $('nasQuery').addEventListener('input', function () {
      state.nas.q = $('nasQuery').value.trim();
      clearTimeout(state.nas.timer);
      state.nas.timer = setTimeout(function () { nasLoadPage(0); }, 300);
    });
    $('btnNasPrev').addEventListener('click', function () { if (state.nas.offset > 0) nasLoadPage(Math.max(0, state.nas.offset - 50)); });
    $('btnNasNext').addEventListener('click', function () { nasLoadPage(state.nas.offset + 50); });
    $('btnPush').addEventListener('click', pushTerms);
    $('btnPull').addEventListener('click', refreshHelper);
  }

  function loadSettings() {
    var saved = readJSON(K.settings, null);
    var origin = (location.protocol === 'http:' || location.protocol === 'https:') ? location.origin : '';
    // NAS-hosted pane lives under /wps/; GitHub-hosted panes must keep the LAN address.
    if (origin && location.pathname.indexOf('/wps/') === 0) state.nas.base = origin;
    state.nas.token = '';
    if (window.NAS_CONFIG && window.NAS_CONFIG.base) state.nas.base = window.NAS_CONFIG.base;
    renderNasServer();
    if ($('nasPass')) $('nasPass').value = '';
    var savedAuth = readNasAuth();
    if (savedAuth) {
      state.nas.token = savedAuth.token;
      state.nas.user = savedAuth.username || '';
    }
    renderNasAuthUI();
    if (saved) state.settings = saved;
    if (!loadSavedTerms()) setTerms([]);
    // The saved termbase may have been parsed before settings were restored, so
    // re-filter it against the project's language now that the settings exist.
    setTerms(state.terms, state.termsMeta);
    $('colSource').value = state.settings.sourceCol;
    $('colTarget').value = state.settings.targetCol;
    $('colUid').value = state.settings.uidCol;
    if ($('colOld')) $('colOld').value = state.settings.oldCol || 'E';
    if ($('targetLang') && window.TideLangs) {
      var select = $('targetLang');
      select.textContent = '';
      window.TideLangs.list.forEach(function (lang) {
        var option = document.createElement('option');
        option.value = lang.code;
        option.textContent = lang.name + (lang.native && lang.native !== lang.name ? ' · ' + lang.native : '');
        option.selected = lang.code === (state.settings.targetLang || 'en');
        select.appendChild(option);
      });
    }
    loadProfiles();
    var profile = state.profiles.find(function (p) { return p.id === state.activeProfile; });
    if (profile && profile.columns) {
      state.settings = Object.assign({}, state.settings, profile.columns);
      $('colSource').value = state.settings.sourceCol;
      $('colTarget').value = state.settings.targetCol;
      $('colUid').value = state.settings.uidCol;
      if ($('colOld')) $('colOld').value = state.settings.oldCol || 'E';
    if ($('targetLang') && window.TideLangs) {
      var select = $('targetLang');
      select.textContent = '';
      window.TideLangs.list.forEach(function (lang) {
        var option = document.createElement('option');
        option.value = lang.code;
        option.textContent = lang.name + (lang.native && lang.native !== lang.name ? ' · ' + lang.native : '');
        option.selected = lang.code === (state.settings.targetLang || 'en');
        select.appendChild(option);
      });
    }
    }
    renderProjects();
    $('syncRepo').value = readJSON(K.repo, '') || '';
    $('syncBranch').value = readJSON(K.branch, 'main') || 'main';
    $('envHint').textContent = sheetApp() ? '已连接 WPS 表格。' : '当前不在 WPS 中，仅能查看设置与记录。';
    if ($('docName')) $('docName').textContent = sheetApp() && currentSheet() ? currentSheet().Name : '未连接 WPS';
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
      // Follow the visible worksheet: without this the batch and its memory
      // stay on the previous sheet until the user notices and refreshes.
      if (typeof syncActiveSheet === 'function') syncActiveSheet();
    }, 1500);
  }

  function bootUiLanguage() {
    if (!window.TideI18n) return;
    window.TideI18n.boot();

    var select = $('uiLang');
    if (select) {
      select.textContent = '';
      window.TideI18n.langs.forEach(function (lang) {
        var option = document.createElement('option');
        option.value = lang.code;
        option.textContent = lang.label;
        option.selected = lang.code === window.TideI18n.lang;
        select.appendChild(option);
      });
      select.addEventListener('change', function () {
        window.TideI18n.setLang(select.value);
        renderLangMenu();
        $('settingsStatus').textContent = '界面语言：' + select.options[select.selectedIndex].textContent;
      });
    }

    // Header globe: the same setting, available without opening 设置.
    var button = $('btnLang');
    var menu = $('langMenu');
    if (!button || !menu) return;

    function renderLangMenu() {
      menu.textContent = '';
      window.TideI18n.langs.forEach(function (lang) {
        var item = document.createElement('button');
        item.type = 'button';
        item.textContent = lang.label;
        if (lang.code === window.TideI18n.lang) {
          item.className = 'on';
          var tick = document.createElement('span');
          tick.className = 'tick';
          tick.textContent = '✓';
          item.appendChild(tick);
        }
        item.addEventListener('click', function () {
          window.TideI18n.setLang(lang.code);
          if ($('uiLang')) $('uiLang').value = lang.code;
          menu.hidden = true;
          renderLangMenu();
        });
        menu.appendChild(item);
      });
    }

    renderLangMenu();
    button.addEventListener('click', function (event) {
      event.stopPropagation();
      menu.hidden = !menu.hidden;
    });
    document.addEventListener('click', function (event) {
      if (!menu.hidden && !menu.contains(event.target) && event.target !== button) menu.hidden = true;
    });
    document.addEventListener('keydown', function (event) {
      if (event.key === 'Escape') menu.hidden = true;
    });
  }

  function boot() {
    bootUiLanguage();
    bind();
    loadSettings();
    startJumpWatch();
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
    window.setInterval(checkForUpdate, 30000);
    // A pane that was hidden while the app was in the background should check
    // as soon as it becomes visible again, not up to two minutes later.
    document.addEventListener('visibilitychange', function () {
      if (!document.hidden) checkForUpdate();
    });
    window.addEventListener('focus', checkForUpdate);
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
