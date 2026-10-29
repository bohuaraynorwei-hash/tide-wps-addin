/**
 * Tide term-matching engine.
 *
 * Pure functions only: no DOM, no WPS API. Loaded by the add-in and by the
 * Node test runner, so the matching rules are verified outside WPS.
 *
 * Rules:
 *  - Chinese terms match as exact substrings.
 *  - Latin terms match case-insensitively, but only on word boundaries.
 *  - At one position the longest term wins; overlapping hits are dropped.
 *  - Duplicate termbase entries are collapsed by (zh, en).
 */
(function (root) {
  'use strict';

  var CJK = /[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]/;

  function hasCJK(text) {
    return CJK.test(text);
  }

  function escapeRegExp(text) {
    return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  function normaliseTerm(entry) {
    if (!entry || typeof entry !== 'object') return null;
    var zh = entry.zh || entry.source || entry.cn || entry['中文'] || '';
    var en = entry.en || entry.target || entry['英文'] || '';
    if (!zh) return null;
    return {
      uid: entry.uid || entry.id || entry.key || '',
      zh: String(zh),
      en: String(en),
      status: entry.status || entry.state || '',
      variants: Array.isArray(entry.variants) ? entry.variants.slice() : [],
    };
  }

  /** Accept an array, {terms:[...]}, or {entries:[...]} and return clean terms. */
  function parseTermbase(input) {
    var list = [];
    if (Array.isArray(input)) list = input;
    else if (input && Array.isArray(input.terms)) list = input.terms;
    else if (input && Array.isArray(input.entries)) list = input.entries;
    else if (input && Array.isArray(input.glossary)) list = input.glossary;

    var seen = {};
    var out = [];
    for (var i = 0; i < list.length; i += 1) {
      var term = normaliseTerm(list[i]);
      if (!term) continue;
      var key = term.zh + '\u0000' + term.en;
      if (seen[key]) continue;
      seen[key] = true;
      out.push(term);
    }
    out.sort(function (a, b) {
      return b.zh.length - a.zh.length || a.zh.localeCompare(b.zh);
    });
    return out;
  }

  function latinBoundary(text, start, end) {
    var before = start === 0 ? '' : text.charAt(start - 1);
    var after = end >= text.length ? '' : text.charAt(end);
    var word = /[A-Za-z0-9_]/;
    return !word.test(before) && !word.test(after);
  }

  /** Return every non-overlapping occurrence of the terms inside text. */
  function findMatches(text, terms) {
    if (!text || !terms || !terms.length) return [];
    var lower = String(text).toLowerCase();
    var hits = [];

    for (var i = 0; i < terms.length; i += 1) {
      var term = terms[i];
      var forms = [term.zh].concat(term.variants || []);
      for (var f = 0; f < forms.length; f += 1) {
        var needle = String(forms[f] || '');
        if (!needle) continue;
        var haystack = hasCJK(needle) ? text : lower;
        var probe = hasCJK(needle) ? needle : needle.toLowerCase();
        var from = 0;
        var at = haystack.indexOf(probe, from);
        while (at !== -1) {
          var end = at + probe.length;
          if (hasCJK(needle) || latinBoundary(text, at, end)) {
            hits.push({ term: term, start: at, end: end, matched: text.slice(at, end) });
          }
          from = at + 1;
          at = haystack.indexOf(probe, from);
        }
      }
    }

    hits.sort(function (a, b) {
      return a.start - b.start || b.end - a.end;
    });

    var result = [];
    var cursor = -1;
    for (var h = 0; h < hits.length; h += 1) {
      if (hits[h].start >= cursor) {
        result.push(hits[h]);
        cursor = hits[h].end;
      }
    }
    return result;
  }

  /**
   * Scan rows for termbase hits.
   * rows: [{row, uid, source, target}]
   * Returns only rows that contain at least one term.
   */
  function scanRows(rows, terms) {
    var out = [];
    for (var i = 0; i < (rows || []).length; i += 1) {
      var row = rows[i];
      var matches = findMatches(row.source, terms);
      if (!matches.length) continue;
      out.push({
        row: row.row,
        uid: row.uid || '',
        source: row.source,
        target: row.target || '',
        matches: matches.map(function (hit) {
          return {
            uid: hit.term.uid,
            zh: hit.term.zh,
            en: hit.term.en,
            status: hit.term.status,
            matched: hit.matched,
          };
        }),
      });
    }
    return out;
  }

  /** Group scan results per term for the Terms tab. */
  function groupByTerm(scanned) {
    var index = {};
    var order = [];
    for (var i = 0; i < scanned.length; i += 1) {
      var row = scanned[i];
      for (var m = 0; m < row.matches.length; m += 1) {
        var key = row.matches[m].zh + '\u0000' + row.matches[m].en;
        if (!index[key]) {
          index[key] = {
            zh: row.matches[m].zh,
            en: row.matches[m].en,
            uid: row.matches[m].uid,
            status: row.matches[m].status,
            rows: [],
          };
          order.push(key);
        }
        index[key].rows.push({ row: row.row, uid: row.uid, target: row.target });
      }
    }
    return order.map(function (key) {
      return index[key];
    });
  }

  var TOKEN_PATTERNS = [
    /\{\{[^{}]*\}\}/g,
    /\{[A-Za-z0-9_.]+\}/g,
    /%\d*\$?[-+#0 ]*(?:\d+|\*)?(?:\.\d+)?[sdif]/g,
    /<\/?[A-Za-z][^<>]{0,40}>/g
  ];

  /** Runtime tokens that must survive translation unchanged. */
  function tokens(text) {
    if (!text) return [];
    var found = [];
    TOKEN_PATTERNS.forEach(function (pattern) {
      var match;
      pattern.lastIndex = 0;
      while ((match = pattern.exec(text)) !== null) {
        found.push(match[0]);
      }
    });
    return found.sort();
  }

  /** Compare token multisets; returns the differing tokens on each side. */
  function checkTokens(source, target) {
    var a = tokens(source);
    var b = tokens(target);
    var missing = [];
    var extra = [];
    var pool = b.slice();
    a.forEach(function (token) {
      var at = pool.indexOf(token);
      if (at === -1) missing.push(token);
      else pool.splice(at, 1);
    });
    pool.forEach(function (token) { extra.push(token); });
    return { ok: missing.length === 0 && extra.length === 0, missing: missing, extra: extra };
  }

  /** Levenshtein distance over Unicode code points. */
  function levenshtein(a, b) {
    var left = Array.from(String(a || ''));
    var right = Array.from(String(b || ''));
    if (!left.length) return right.length;
    if (!right.length) return left.length;
    var previous = [];
    for (var j = 0; j <= right.length; j += 1) previous[j] = j;
    for (var i = 1; i <= left.length; i += 1) {
      var current = [i];
      for (var k = 1; k <= right.length; k += 1) {
        var cost = left[i - 1] === right[k - 1] ? 0 : 1;
        current[k] = Math.min(current[k - 1] + 1, previous[k] + 1, previous[k - 1] + cost);
      }
      previous = current;
    }
    return previous[right.length];
  }

  function similarity(a, b) {
    var left = String(a || '');
    var right = String(b || '');
    if (!left && !right) return 1;
    var longest = Math.max(Array.from(left).length, Array.from(right).length);
    if (!longest) return 1;
    return 1 - levenshtein(left, right) / longest;
  }

  /** Build a source→target memory from reviewed rows. */
  function buildMemory(rows) {
    var seen = {};
    var memory = [];
    (rows || []).forEach(function (row) {
      var source = String(row.source || '').trim();
      var target = String(row.target || '').trim();
      if (!source || !target) return;
      var key = source + '\u0000' + target;
      if (seen[key]) return;
      seen[key] = 1;
      memory.push({ source: source, target: target, row: row.row, sheet: row.sheet || '', uid: row.uid || '' });
    });
    return memory;
  }

  /**
   * Exact matches first, then fuzzy ones above the threshold.
   * Returns { exact: [...], fuzzy: [{...entry, score}] }.
   */
  function findMemory(source, memory, minScore, limit) {
    var text = String(source || '').trim();
    var floor = typeof minScore === 'number' ? minScore : 0.5;
    var cap = typeof limit === 'number' ? limit : 5;
    var exact = [];
    var fuzzy = [];
    if (!text) return { exact: exact, fuzzy: fuzzy };
    (memory || []).forEach(function (entry) {
      if (entry.source === text) {
        exact.push(entry);
        return;
      }
      // Length alone bounds similarity, so skip entries that cannot reach the floor.
      var longest = Math.max(Array.from(text).length, Array.from(entry.source).length);
      if (longest && 1 - Math.abs(Array.from(text).length - Array.from(entry.source).length) / longest < floor) return;
      var score = similarity(text, entry.source);
      if (score >= floor) fuzzy.push(Object.assign({}, entry, { score: score }));
    });
    fuzzy.sort(function (a, b) { return b.score - a.score; });
    return { exact: exact, fuzzy: fuzzy.slice(0, cap) };
  }

  var api = {
    levenshtein: levenshtein,
    similarity: similarity,
    buildMemory: buildMemory,
    findMemory: findMemory,
    tokens: tokens,
    checkTokens: checkTokens,
    parseTermbase: parseTermbase,
    findMatches: findMatches,
    scanRows: scanRows,
    groupByTerm: groupByTerm,
    hasCJK: hasCJK,
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  if (root) root.TideEngine = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
