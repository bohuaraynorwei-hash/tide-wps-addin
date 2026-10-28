/**
 * Archive encryption for the Tide add-in.
 *
 * AES-GCM 256 with a PBKDF2-SHA256 derived key. Runs on WebCrypto, so the same
 * code works in the WPS pane and in Node's test runner. Nothing here touches the
 * network or localStorage; callers decide where the blob goes.
 */
(function (root) {
  'use strict';

  var ITERATIONS = 200000;
  var VERSION = 1;

  function toBase64(bytes) {
    var binary = '';
    var view = new Uint8Array(bytes);
    for (var i = 0; i < view.length; i += 1) binary += String.fromCharCode(view[i]);
    if (typeof btoa === 'function') return btoa(binary);
    return Buffer.from(binary, 'binary').toString('base64');
  }

  function fromBase64(text) {
    if (typeof atob === 'function') {
      var binary = atob(text);
      var out = new Uint8Array(binary.length);
      for (var i = 0; i < binary.length; i += 1) out[i] = binary.charCodeAt(i);
      return out;
    }
    return new Uint8Array(Buffer.from(text, 'base64'));
  }

  function subtle() {
    var api = (typeof crypto !== 'undefined' && crypto.subtle) ? crypto : (root && root.crypto);
    if (!api || !api.subtle) throw new Error('WebCrypto unavailable in this context');
    return api.subtle;
  }

  async function deriveKey(passphrase, salt, iterations) {
    var api = subtle();
    var base = await api.importKey('raw', new TextEncoder().encode(passphrase), 'PBKDF2', false, ['deriveKey']);
    return api.deriveKey(
      { name: 'PBKDF2', salt: salt, iterations: iterations, hash: 'SHA-256' },
      base,
      { name: 'AES-GCM', length: 256 },
      false,
      ['encrypt', 'decrypt']
    );
  }

  function randomBytes(length) {
    var bytes = new Uint8Array(length);
    var api = (typeof crypto !== 'undefined' && crypto.getRandomValues) ? crypto : root.crypto;
    api.getRandomValues(bytes);
    return bytes;
  }

  async function encrypt(passphrase, plaintext) {
    if (!passphrase) throw new Error('passphrase required');
    var salt = randomBytes(16);
    var iv = randomBytes(12);
    var key = await deriveKey(passphrase, salt, ITERATIONS);
    var cipher = await subtle().encrypt({ name: 'AES-GCM', iv: iv }, key, new TextEncoder().encode(plaintext));
    return {
      v: VERSION,
      kdf: 'PBKDF2-SHA256',
      iter: ITERATIONS,
      salt: toBase64(salt),
      iv: toBase64(iv),
      ct: toBase64(cipher)
    };
  }

  async function decrypt(passphrase, blob) {
    if (!blob || blob.v !== VERSION) throw new Error('unsupported archive format');
    var salt = fromBase64(blob.salt);
    var iv = fromBase64(blob.iv);
    var key = await deriveKey(passphrase, salt, blob.iter || ITERATIONS);
    var plain = await subtle().decrypt({ name: 'AES-GCM', iv: iv }, key, fromBase64(blob.ct));
    return new TextDecoder().decode(plain);
  }

  var api = { encrypt: encrypt, decrypt: decrypt, VERSION: VERSION, ITERATIONS: ITERATIONS };

  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  if (root) root.TideCrypto = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
