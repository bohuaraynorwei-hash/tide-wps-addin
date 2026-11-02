/**
 * Target-language profiles.
 *
 * The source side is always Chinese. Everything language-specific — the column
 * headers an export uses, the punctuation a translation is allowed to contain,
 * and the TMX locale code — is described here so the rest of the app does not
 * grow per-language conditions.
 */
(function (root) {
  'use strict';

  var LANGS = [
    {
      code: 'en', name: '英语', native: 'English', tmx: 'en-US', spaceAfter: true, spaceBefore: false, asciiInCJK: false, cjk: false,
      headers: ['EN', 'English', '英文', '英语'],
      final: /[.!?]$/,
      wrong: /[。，、；：？！]/
    },
    {
      code: 'zh', name: '中文', native: '中文', tmx: 'zh-CN', spaceAfter: false, spaceBefore: false, asciiInCJK: true, cjk: true,
      headers: ['ZH', '中文', 'Chinese'],
      final: /[。！？!?]$/,
      wrong: null
    },
    {
      code: 'ja', name: '日语', native: '日本語', tmx: 'ja-JP', spaceAfter: false, spaceBefore: false, asciiInCJK: true, cjk: true,
      headers: ['JA', 'JP', '日本語', '日语', 'Japanese'],
      final: /[。！？!?]$/,
      wrong: null
    },
    {
      code: 'ko', name: '韩语', native: '한국어', tmx: 'ko-KR', spaceAfter: true, spaceBefore: false, asciiInCJK: false, cjk: true,
      headers: ['KO', 'KR', '한국어', '韩语', 'Korean'],
      final: /[.!?。？！]$/,
      wrong: null
    },
    {
      code: 'ru', name: '俄语', native: 'Русский', tmx: 'ru-RU', spaceAfter: true, spaceBefore: false, asciiInCJK: false, cjk: false,
      headers: ['RU', 'Русский', '俄语', 'Russian'],
      final: /[.!?]$/,
      wrong: /[。，、；：？！]/
    },
    {
      code: 'vi', name: '越南语', native: 'Tiếng Việt', tmx: 'vi-VN', spaceAfter: true, spaceBefore: false, asciiInCJK: false, cjk: false,
      headers: ['VI', 'Tiếng Việt', '越南语', 'Vietnamese'],
      final: /[.!?]$/,
      wrong: /[。，、；：？！]/
    },
    {
      code: 'pt-BR', name: '巴西葡语', native: 'Português (Brasil)', tmx: 'pt-BR', spaceAfter: true, spaceBefore: false, asciiInCJK: false, cjk: false,
      headers: ['PT', 'PT-BR', 'Português', '巴西葡语', '葡萄牙语'],
      final: /[.!?]$/,
      wrong: /[。，、；：？！]/
    },
    {
      code: 'fr', name: '法语', native: 'Français', tmx: 'fr-FR', spaceAfter: true, spaceBefore: true, asciiInCJK: false, cjk: false,
      headers: ['FR', 'Français', '法语', 'French'],
      final: /[.!?]$/,
      wrong: /[。，、；：？！]/
    },
    {
      code: 'es', name: '西班牙语', native: 'Español', tmx: 'es-ES', spaceAfter: true, spaceBefore: false, asciiInCJK: false, cjk: false,
      headers: ['ES', 'Español', '西班牙语', 'Spanish'],
      final: /[.!?]$/,
      wrong: /[。，、；：？！]/
    },
    {
      code: 'th', name: '泰语', native: 'ไทย', tmx: 'th-TH', spaceAfter: false, spaceBefore: false, asciiInCJK: false, cjk: false,
      headers: ['TH', 'ไทย', '泰语', 'Thai'],
      final: /[.!?]$/,
      wrong: /[。，、；：？！]/
    }
  ];

  function get(code) {
    for (var i = 0; i < LANGS.length; i += 1) if (LANGS[i].code === code) return LANGS[i];
    return LANGS[0];
  }

  function headersFor(code) {
    return get(code).headers.slice();
  }

  var api = { list: LANGS, get: get, headersFor: headersFor };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  if (root) root.TideLangs = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
