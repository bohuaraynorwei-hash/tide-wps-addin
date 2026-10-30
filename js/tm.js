/**
 * Translation memory: index, concordance, leverage analysis, TMX 1.4 interop.
 *
 * Pure functions, no DOM — so the same code runs in the renderer and in Node
 * tests. TMX is parsed with a small tolerant scanner rather than DOMParser,
 * because the desktop shell and the test runner need identical behaviour.
 */
(function (root) {
  'use strict';

  function decodeEntities(text) {
    return String(text)
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&#(\d+);/g, function (_m, code) { return String.fromCharCode(Number(code)); })
      .replace(/&amp;/g, '&');
  }

  function encodeEntities(text) {
    return String(text)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function attr(block, name) {
    var match = new RegExp(name + '\\s*=\\s*"([^"]*)"', 'i').exec(block);
    return match ? decodeEntities(match[1]) : '';
  }

  // TMX 1.4 is regular enough for a scanner: <tu> blocks containing <tuv> per
  // language, each with one <seg>.
  function parseTmx(xml) {
    var text = String(xml || '');
    var units = [];
    var tuRe = /<tu\b[^>]*>([\s\S]*?)<\/tu>/gi;
    var match;
    while ((match = tuRe.exec(text)) !== null) {
      var body = match[1];
      var tuvRe = /<tuv\b([^>]*)>([\s\S]*?)<\/tuv>/gi;
      var langs = {};
      var tuv;
      while ((tuv = tuvRe.exec(body)) !== null) {
        var lang = attr(tuv[1], 'xml:lang') || attr(tuv[1], 'lang');
        var segMatch = /<seg\b[^>]*>([\s\S]*?)<\/seg>/i.exec(tuv[2]);
        if (!lang || !segMatch) continue;
        langs[lang.toLowerCase()] = decodeEntities(segMatch[1].replace(/<[^>]+>/g, '').trim());
      }
      var codes = Object.keys(langs);
      if (codes.length < 2) continue;
      units.push({
        source: langs[codes[0]],
        target: langs[codes[1]],
        sourceLang: codes[0],
        targetLang: codes[1]
      });
    }
    return units;
  }

  function buildTmx(units, sourceLang, targetLang) {
    var src = sourceLang || 'zh-CN';
    var tgt = targetLang || 'en-US';
    var now = new Date().toISOString();
    var body = (units || []).map(function (unit) {
      return '  <tu>\n' +
        '    <tuv xml:lang="' + encodeEntities(src) + '"><seg>' + encodeEntities(unit.source) + '</seg></tuv>\n' +
        '    <tuv xml:lang="' + encodeEntities(tgt) + '"><seg>' + encodeEntities(unit.target) + '</seg></tuv>\n' +
        '  </tu>';
    }).join('\n');
    return '<?xml version="1.0" encoding="UTF-8"?>\n' +
      '<tmx version="1.4">\n' +
      '  <header creationtool="Tide" creationtoolversion="0.4" segtype="sentence" ' +
      'o-tmf="Tide" adminlang="en-US" srclang="' + encodeEntities(src) + '" ' +
      'datatype="plaintext" creationdate="' + now + '"/>\n' +
      '  <body>\n' + body + '\n  </body>\n</tmx>\n';
  }

  // Index by source for exact hits; a plain array drives fuzzy scans.
  function buildIndex(units) {
    var bySource = {};
    var entries = [];
    (units || []).forEach(function (unit) {
      var source = String(unit.source || '').trim();
      var target = String(unit.target || '').trim();
      if (!source || !target) return;
      var key = source;
      if (bySource[key]) return;
      var entry = { source: source, target: target };
      bySource[key] = entry;
      entries.push(entry);
    });
    return { entries: entries, bySource: bySource };
  }

  function concordance(index, query, options) {
    var opts = options || {};
    var limit = opts.limit || 20;
    var threshold = typeof opts.threshold === 'number' ? opts.threshold : 0.4;
    var q = String(query || '').trim();
    if (!q || !index || !index.entries) return [];
    var lower = q.toLowerCase();
    var results = [];
    index.entries.forEach(function (entry) {
      var source = entry.source;
      var target = entry.target;
      var score = 0;
      var where = 'fuzzy';
      if (source === q || target === q) {
        score = 1;
        where = source === q ? 'source-exact' : 'target-exact';
      } else if (source.toLowerCase().indexOf(lower) > -1) {
        score = 0.95;
        where = 'source-contains';
      } else if (target.toLowerCase().indexOf(lower) > -1) {
        score = 0.9;
        where = 'target-contains';
      } else if (root.TideEngine && root.TideEngine.similarity) {
        score = Math.max(
          root.TideEngine.similarity(q, source),
          root.TideEngine.similarity(q, target)
        );
      }
      if (score >= threshold) results.push({ entry: entry, score: score, where: where });
    });
    results.sort(function (a, b) { return b.score - a.score; });
    return results.slice(0, limit);
  }

  /**
   * Leverage: how much of a source corpus is already covered.
   * Bands: exact 100%, high 75–99%, low 50–74%, new below 50%.
   */
  function analyze(sources, index) {
    var stats = { total: 0, exact: 0, high: 0, low: 0, new: 0, chars: 0, fuzzyChars: 0 };
    (sources || []).forEach(function (source) {
      var text = String(source || '').trim();
      if (!text) return;
      stats.total += 1;
      stats.chars += text.length;
      if (index && index.bySource[text]) {
        stats.exact += 1;
        return;
      }
      var best = 0;
      if (index && index.entries && root.TideEngine && root.TideEngine.similarity) {
        for (var i = 0; i < index.entries.length; i += 1) {
          var score = root.TideEngine.similarity(text, index.entries[i].source);
          if (score > best) best = score;
          if (best >= 0.99) break;
        }
      }
      if (best >= 0.75) { stats.high += 1; stats.fuzzyChars += text.length; }
      else if (best >= 0.5) { stats.low += 1; stats.fuzzyChars += text.length; }
      else stats.new += 1;
    });
    var total = stats.total || 1;
    stats.exactPct = Math.round((stats.exact / total) * 1000) / 10;
    stats.highPct = Math.round((stats.high / total) * 1000) / 10;
    stats.lowPct = Math.round((stats.low / total) * 1000) / 10;
    stats.newPct = Math.round((stats.new / total) * 1000) / 10;
    stats.leveragePct = Math.round(((stats.exact + stats.high + stats.low) / total) * 1000) / 10;
    return stats;
  }

  var api = { parseTmx: parseTmx, buildTmx: buildTmx, buildIndex: buildIndex, concordance: concordance, analyze: analyze };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  if (root) root.TideTM = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
