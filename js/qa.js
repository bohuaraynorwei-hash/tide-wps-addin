/**
 * Quality checks for a localization batch.
 *
 * Every check is deterministic and explainable: no model, no heuristics that
 * cannot be traced to a rule. Each finding names the rule that fired, because a
 * reviewer has to be able to disagree with a specific claim.
 */
(function (root) {
  'use strict';

  var SEVERITY_ORDER = { high: 0, medium: 1, low: 2 };

  function cjk(text) {
    return /[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]/.test(String(text || ''));
  }

  function numbers(text) {
    var found = String(text || '').match(/\d+(?:[.,]\d+)?%?/g) || [];
    return found.map(function (n) { return n.replace(',', '.'); }).sort();
  }

  function sameMultiset(a, b) {
    if (a.length !== b.length) return false;
    for (var i = 0; i < a.length; i += 1) if (a[i] !== b[i]) return false;
    return true;
  }

  // Terms are checked in the direction that matters: a source term present in
  // the source must have its approved target in the translation.
  function terminologyMisses(source, target, terms) {
    var misses = [];
    (terms || []).forEach(function (term) {
      var zh = String(term.zh || '');
      var en = String(term.en || '');
      if (!zh || !en || term.status === 'needs_review') return;
      if (source.indexOf(zh) === -1) return;
      // Accept inflected forms: "Hook" is satisfied by "hooked" / "hooks".
      var present = root.TideMorph
        ? root.TideMorph.containsPhrase(target, en)
        : target.toLowerCase().indexOf(en.toLowerCase()) > -1;
      if (!present) {
        misses.push({ zh: zh, en: en });
      }
    });
    return misses;
  }

  function withFindings(findings, row, category, severity, message) {
    findings.push({
      row: row.row,
      sheet: row.sheet || '',
      category: category,
      severity: severity,
      message: message,
      source: row.source,
      target: row.target
    });
  }

  function checkRow(row, terms, options) {
    var opts = options || {};
    var findings = [];
    var source = String(row.source || '');
    var target = String(row.target || '');

    if (!target.trim()) {
      withFindings(findings, row, 'empty', 'high', '空译文');
      return findings;
    }

    if (target === source && cjk(target)) {
      withFindings(findings, row, 'untranslated', 'high', '译文与原文相同');
    }

    if (root.TideEngine) {
      var tokens = root.TideEngine.checkTokens(source, target);
      if (!tokens.ok) {
        var detail = [];
        if (tokens.missing.length) detail.push('缺 ' + tokens.missing.join(' '));
        if (tokens.extra.length) detail.push('多 ' + tokens.extra.join(' '));
        withFindings(findings, row, 'placeholder', 'high', '占位符不一致：' + detail.join('，'));
      }
    }

    var sourceNumbers = numbers(source);
    var targetNumbers = numbers(target);
    if (sourceNumbers.length || targetNumbers.length) {
      if (!sameMultiset(sourceNumbers, targetNumbers)) {
        withFindings(findings, row, 'number', 'medium',
          '数字不一致：原文 [' + sourceNumbers.join(' ') + ']，译文 [' + targetNumbers.join(' ') + ']');
      }
    }

    var misses = terminologyMisses(source, target, terms);
    if (misses.length) {
      withFindings(findings, row, 'terminology', 'medium',
        '术语未使用：' + misses.map(function (m) { return m.zh + ' → ' + m.en; }).join('；'));
    }

    if (cjk(source) && !cjk(target)) {
      var ratio = target.length / Math.max(source.length, 1);
      if (ratio > 2.4 || ratio < 0.2) {
        withFindings(findings, row, 'length', 'low',
          '长度比例异常：' + ratio.toFixed(2) + '（原文 ' + source.length + ' 字符，译文 ' + target.length + '）');
      }
    }

    if (/[。！？]$/.test(source.trim()) && !/[.!?]$/.test(target.trim())) {
      withFindings(findings, row, 'punctuation', 'low', '原文是陈述/感叹/疑问句，译文缺少句末标点');
    }
    if (/[。，、；：？！]/.test(target)) {
      withFindings(findings, row, 'punctuation', 'medium', '译文里出现中文标点');
    }

    if (/^\s|\s$/.test(target) || /  +/.test(target)) {
      withFindings(findings, row, 'whitespace', 'low', '译文首尾或中间有多余空格');
    }

    if (opts.maxLengthByFunction && row.function && opts.maxLengthByFunction[row.function]) {
      var limit = opts.maxLengthByFunction[row.function];
      if (target.length > limit) {
        withFindings(findings, row, 'length', 'medium',
          '超出该功能长度上限：' + target.length + ' > ' + limit);
      }
    }

    return findings;
  }

  function runQA(rows, terms, options) {
    var findings = [];
    (rows || []).forEach(function (row) {
      findings = findings.concat(checkRow(row, terms, options));
    });
    findings.sort(function (a, b) {
      return (SEVERITY_ORDER[a.severity] - SEVERITY_ORDER[b.severity]) || (a.row - b.row);
    });
    return findings;
  }

  function summarize(findings) {
    var byCategory = {};
    var bySeverity = { high: 0, medium: 0, low: 0 };
    (findings || []).forEach(function (finding) {
      byCategory[finding.category] = (byCategory[finding.category] || 0) + 1;
      bySeverity[finding.severity] = (bySeverity[finding.severity] || 0) + 1;
    });
    return { total: (findings || []).length, byCategory: byCategory, bySeverity: bySeverity };
  }

  var api = { runQA: runQA, checkRow: checkRow, summarize: summarize, terminologyMisses: terminologyMisses };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  if (root) root.TideQA = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
