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

  var api = {
    parseTermbase: parseTermbase,
    findMatches: findMatches,
    scanRows: scanRows,
    groupByTerm: groupByTerm,
    hasCJK: hasCJK,
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  if (root) root.TideEngine = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
