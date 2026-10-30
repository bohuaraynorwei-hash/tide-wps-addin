/**
 * Lightweight English morphology for term recognition.
 *
 * Not a stemmer for search engines — a phrase matcher that accepts inflected
 * forms of an approved term: "Hook" also matches "hooked", "fishing rod" also
 * matches "fishing rods". Conservative by design: unknown words are compared
 * literally, because a false terminology pass is worse than a false alarm.
 */
(function (root) {
  'use strict';

  var IRREGULAR = {
    be: ['am', 'is', 'are', 'was', 'were', 'been', 'being'],
    catch: ['caught', 'catches', 'catching'],
    run: ['ran', 'runs', 'running'],
    swim: ['swam', 'swims', 'swimming'],
    bite: ['bit', 'bitten', 'bites', 'biting'],
    hold: ['held', 'holds', 'holding'],
    cast: ['casts', 'casting'],
    feed: ['fed', 'feeds', 'feeding'],
    rise: ['rose', 'risen', 'rises', 'rising'],
    sink: ['sank', 'sunk', 'sinks', 'sinking'],
    throw: ['threw', 'thrown', 'throws', 'throwing'],
    wind: ['wound', 'winds', 'winding'],
    break: ['broke', 'broken', 'breaks', 'breaking'],
    child: ['children'],
    fish: ['fishes', 'fish'],
    man: ['men'],
    foot: ['feet'],
    tooth: ['teeth'],
  };

  var IRREGULAR_BACK = {};
  Object.keys(IRREGULAR).forEach(function (base) {
    IRREGULAR[base].forEach(function (form) { IRREGULAR_BACK[form] = base; });
  });

  function isVowel(ch) { return 'aeiou'.indexOf(ch) > -1; }

  function lemma(word) {
    var w = String(word || '').toLowerCase().replace(/[^a-z'-]/g, '');
    if (!w) return '';
    if (IRREGULAR_BACK[w]) return IRREGULAR_BACK[w];
    if (IRREGULAR[w]) return w;
    if (w.length <= 3) return w;

    if (/ies$/.test(w)) return w.slice(0, -3) + 'y';
    if (/(ches|shes|sses|xes|zes)$/.test(w)) return w.slice(0, -2);
    if (/s$/.test(w) && !/ss$/.test(w)) return w.slice(0, -1);
    if (/ing$/.test(w) && w.length > 5) {
      var base = w.slice(0, -3);
      if (base.length > 2 && base[base.length - 1] === base[base.length - 2]) base = base.slice(0, -1);
      if (base.length > 2 && !isVowel(base[base.length - 1])) return base;   // fishing -> fish
      return base + 'e';                                                     // making -> make
    }
    if (/ed$/.test(w) && w.length > 4) {
      var stem = w.slice(0, -2);
      if (stem[stem.length - 1] === stem[stem.length - 2]) stem = stem.slice(0, -1);
      if (stem.length > 2 && !isVowel(stem[stem.length - 1])) return stem;
      return stem.replace(/i$/, 'y') + (/e$/.test(stem) ? '' : 'e');
    }
    if (/er$/.test(w) && w.length > 5) return w.slice(0, -2);
    if (/est$/.test(w) && w.length > 6) return w.slice(0, -3);
    return w;
  }

  function tokens(text) {
    return String(text || '')
      .toLowerCase()
      .split(/[^a-z0-9']+/)
      .filter(Boolean);
  }

  /** Does `text` contain `phrase`, allowing inflected forms of its words? */
  function containsPhrase(text, phrase) {
    var haystack = tokens(text);
    var needle = tokens(phrase);
    if (!needle.length || haystack.length < needle.length) return false;
    if (String(text).toLowerCase().indexOf(String(phrase).toLowerCase()) > -1) return true;
    for (var i = 0; i + needle.length <= haystack.length; i += 1) {
      var ok = true;
      for (var j = 0; j < needle.length; j += 1) {
        if (lemma(haystack[i + j]) !== lemma(needle[j])) { ok = false; break; }
      }
      if (ok) return true;
    }
    return false;
  }

  var api = { lemma: lemma, tokens: tokens, containsPhrase: containsPhrase, irregular: IRREGULAR };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  if (root) root.TideMorph = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
