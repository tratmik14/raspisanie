// Демо-режим: настоящий apps-script/Code.gs, запущенный прямо в браузере.
// Листы, кэш и блокировки Google заменены простыми объектами в localStorage, поэтому
// в демо работают ровно те же правила, что и на сервере, и видны они только в этом браузере.
(function () {
  'use strict';

  var BOOK_KEY = 'raspisanie.demo.book.v2';
  var CACHE_KEY = 'raspisanie.demo.cache.v2';
  var DEMO_ADMIN = { login: 'demo', name: 'Демо-администратор', password: 'demo' };

  function readJSON(key, fallback) {
    try { var v = JSON.parse(localStorage.getItem(key) || 'null'); return v == null ? fallback : v; } catch (e) { return fallback; }
  }
  function writeJSON(key, v) {
    try { localStorage.setItem(key, JSON.stringify(v)); } catch (e) { /* storage full or blocked: the demo keeps working until reload */ }
  }

  // ---------- SHA-256, synchronous like Utilities.computeDigest ----------
  var K = [0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5, 0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3,
    0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174, 0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
    0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967, 0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13,
    0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85, 0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
    0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3, 0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208,
    0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2];
  function sha256(bytes) {
    var h = [0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19];
    var l = bytes.length, withPad = ((l + 9 + 63) >> 6) << 6, m = new Uint8Array(withPad);
    m.set(bytes); m[l] = 0x80;
    var bits = l * 8;
    m[withPad - 4] = (bits >>> 24) & 255; m[withPad - 3] = (bits >>> 16) & 255; m[withPad - 2] = (bits >>> 8) & 255; m[withPad - 1] = bits & 255;
    var w = new Array(64);
    for (var off = 0; off < withPad; off += 64) {
      for (var i = 0; i < 16; i++) w[i] = (m[off + i * 4] << 24) | (m[off + i * 4 + 1] << 16) | (m[off + i * 4 + 2] << 8) | m[off + i * 4 + 3];
      for (i = 16; i < 64; i++) {
        var a = w[i - 15], b = w[i - 2];
        var s0 = ((a >>> 7) | (a << 25)) ^ ((a >>> 18) | (a << 14)) ^ (a >>> 3);
        var s1 = ((b >>> 17) | (b << 15)) ^ ((b >>> 19) | (b << 13)) ^ (b >>> 10);
        w[i] = (w[i - 16] + s0 + w[i - 7] + s1) | 0;
      }
      var A = h[0], B = h[1], C = h[2], D = h[3], E = h[4], F = h[5], G = h[6], H = h[7];
      for (i = 0; i < 64; i++) {
        var S1 = ((E >>> 6) | (E << 26)) ^ ((E >>> 11) | (E << 21)) ^ ((E >>> 25) | (E << 7));
        var t1 = (H + S1 + ((E & F) ^ (~E & G)) + K[i] + w[i]) | 0;
        var S0 = ((A >>> 2) | (A << 30)) ^ ((A >>> 13) | (A << 19)) ^ ((A >>> 22) | (A << 10));
        var t2 = (S0 + ((A & B) ^ (A & C) ^ (B & C))) | 0;
        H = G; G = F; F = E; E = (D + t1) | 0; D = C; C = B; B = A; A = (t1 + t2) | 0;
      }
      h[0] = (h[0] + A) | 0; h[1] = (h[1] + B) | 0; h[2] = (h[2] + C) | 0; h[3] = (h[3] + D) | 0;
      h[4] = (h[4] + E) | 0; h[5] = (h[5] + F) | 0; h[6] = (h[6] + G) | 0; h[7] = (h[7] + H) | 0;
    }
    var out = [];
    h.forEach(function (x) { out.push((x >>> 24) & 255, (x >>> 16) & 255, (x >>> 8) & 255, x & 255); });
    return out.map(function (b) { return b > 127 ? b - 256 : b; }); // signed bytes, as Apps Script returns them
  }

  // ---------- the stand-ins for Google services ----------
  var book = readJSON(BOOK_KEY, { sheets: {} });
  function sheetObj(name) {
    var s = book.sheets[name];
    return {
      getLastRow: function () { var n = s.data.length; while (n > 0 && s.data[n - 1].every(function (v) { return v === '' || v == null; })) n--; return n; },
      getMaxRows: function () { return s.max; },
      insertRowsAfter: function (after, n) { s.max += n; return this; },
      setFrozenRows: function () { return this; },
      getRange: function (r, c, nr, nc) {
        nr = nr || 1; nc = nc || 1;
        if (r + nr - 1 > s.max) throw new Error('Those rows are out of bounds.');
        var self = {
          setValues: function (vals) {
            for (var i = 0; i < nr; i++) {
              while (s.data.length < r + i) s.data.push([]);
              for (var j = 0; j < nc; j++) s.data[r + i - 1][c + j - 1] = vals[i][j];
            }
            return self;
          },
          getValues: function () {
            var out = [];
            for (var i = 0; i < nr; i++) {
              var row = s.data[r + i - 1] || [], o = [];
              for (var j = 0; j < nc; j++) o.push(row[c + j - 1] == null ? '' : row[c + j - 1]);
              out.push(o);
            }
            return out;
          },
          clearContent: function () {
            for (var i = 0; i < nr; i++) { var row = s.data[r + i - 1]; if (row) for (var j = 0; j < nc; j++) row[c + j - 1] = ''; }
            return self;
          },
          setNumberFormat: function () { return self; },
          setFontWeight: function () { return self; }
        };
        return self;
      }
    };
  }
  var SpreadsheetApp = {
    getActiveSpreadsheet: function () {
      return {
        getSheetByName: function (n) { return book.sheets[n] ? sheetObj(n) : null; },
        insertSheet: function (n) { book.sheets[n] = { data: [], max: 1000 }; return sheetObj(n); }
      };
    },
    flush: function () { writeJSON(BOOK_KEY, book); },
    getUi: function () { throw new Error('В демо-режиме меню таблицы нет'); }
  };
  var CacheService = {
    getScriptCache: function () {
      return {
        get: function (k) { var c = readJSON(CACHE_KEY, {}); var v = c[k]; return v && v.until > Date.now() ? v.value : null; },
        put: function (k, value, ttl) { var c = readJSON(CACHE_KEY, {}); c[k] = { value: value, until: Date.now() + (ttl || 600) * 1000 }; writeJSON(CACHE_KEY, c); },
        remove: function (k) { var c = readJSON(CACHE_KEY, {}); delete c[k]; writeJSON(CACHE_KEY, c); }
      };
    }
  };
  var LockService = { getScriptLock: function () { return { waitLock: function () {}, releaseLock: function () {} }; } };
  function uuid() {
    if (window.crypto && crypto.randomUUID) return crypto.randomUUID();
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (ch) { var r = Math.random() * 16 | 0; return (ch === 'x' ? r : (r & 3) | 8).toString(16); });
  }
  var Utilities = {
    getUuid: uuid,
    DigestAlgorithm: { SHA_256: 'SHA_256' },
    Charset: { UTF_8: 'UTF_8' },
    computeDigest: function (alg, value) { return sha256(new TextEncoder().encode(String(value))); },
    formatDate: function (d) { return d.toISOString().slice(0, 10); }
  };
  var Session = { getScriptTimeZone: function () { return 'Europe/Moscow'; } };
  var ContentService = { MimeType: { JSON: 'json' }, createTextOutput: function (s) { return { setMimeType: function () { return { content: s }; } }; } };
  var UrlFetchApp = {
    fetch: function () { throw new Error('в демо-режиме Google-документы по ссылке не открываются, скачайте документ как .docx и загрузите файлом'); }
  };

  var server = null;
  function load() {
    if (server) return server;
    server = fetch('apps-script/Code.gs?v=3.3', { cache: 'no-cache' }).then(function (r) {
      if (!r.ok) throw new Error('Не удалось загрузить apps-script/Code.gs для демо-режима');
      return r.text();
    }).then(function (code) {
      var make = new Function('SpreadsheetApp', 'CacheService', 'LockService', 'Utilities', 'Session', 'ContentService', 'UrlFetchApp',
        code + '\n;return { doGet: doGet, doPost: doPost, addAdmin: addAdmin, readAdmins: readAdmins };');
      var api = make(SpreadsheetApp, CacheService, LockService, Utilities, Session, ContentService, UrlFetchApp);
      if (!api.readAdmins().length) { api.addAdmin(DEMO_ADMIN.login, DEMO_ADMIN.name, DEMO_ADMIN.password); writeJSON(BOOK_KEY, book); }
      return api;
    });
    server.catch(function () { server = null; });
    return server;
  }

  window.RaspSim = {
    call: function (body) {
      return load().then(function (api) {
        var out = JSON.parse(api.doPost({ postData: { contents: JSON.stringify(body) } }).content);
        writeJSON(BOOK_KEY, book);
        return out;
      });
    },
    reset: function () {
      try { localStorage.removeItem(BOOK_KEY); localStorage.removeItem(CACHE_KEY); } catch (e) { /* storage unavailable */ }
      book = { sheets: {} };
      server = null;
    }
  };
})();
