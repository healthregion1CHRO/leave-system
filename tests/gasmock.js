// Minimal in-memory Google Apps Script environment for testing Code.gs.
(function (root) {
  // ---- sync SHA-256 ----
  function sha256(bytes) {
    const K = [0x428a2f98,0x71374491,0xb5c0fbcf,0xe9b5dba5,0x3956c25b,0x59f111f1,0x923f82a4,0xab1c5ed5,0xd807aa98,0x12835b01,0x243185be,0x550c7dc3,0x72be5d74,0x80deb1fe,0x9bdc06a7,0xc19bf174,0xe49b69c1,0xefbe4786,0x0fc19dc6,0x240ca1cc,0x2de92c6f,0x4a7484aa,0x5cb0a9dc,0x76f988da,0x983e5152,0xa831c66d,0xb00327c8,0xbf597fc7,0xc6e00bf3,0xd5a79147,0x06ca6351,0x14292967,0x27b70a85,0x2e1b2138,0x4d2c6dfc,0x53380d13,0x650a7354,0x766a0abb,0x81c2c92e,0x92722c85,0xa2bfe8a1,0xa81a664b,0xc24b8b70,0xc76c51a3,0xd192e819,0xd6990624,0xf40e3585,0x106aa070,0x19a4c116,0x1e376c08,0x2748774c,0x34b0bcb5,0x391c0cb3,0x4ed8aa4a,0x5b9cca4f,0x682e6ff3,0x748f82ee,0x78a5636f,0x84c87814,0x8cc70208,0x90befffa,0xa4506ceb,0xbef9a3f7,0xc67178f2];
    let H = [0x6a09e667,0xbb67ae85,0x3c6ef372,0xa54ff53a,0x510e527f,0x9b05688c,0x1f83d9ab,0x5be0cd19];
    const l = bytes.length, bitLen = l * 8;
    const n = (((l + 9) + 63) >> 6) << 6;
    const m = new Uint8Array(n); m.set(bytes); m[l] = 0x80;
    for (let i = 0; i < 8; i++) m[n - 1 - i] = (i < 4) ? (bitLen >>> (8 * i)) & 0xff : 0;
    const w = new Uint32Array(64);
    const r = (x, k) => (x >>> k) | (x << (32 - k));
    for (let o = 0; o < n; o += 64) {
      for (let i = 0; i < 16; i++) w[i] = (m[o+4*i]<<24)|(m[o+4*i+1]<<16)|(m[o+4*i+2]<<8)|m[o+4*i+3];
      for (let i = 16; i < 64; i++) { const s0 = r(w[i-15],7)^r(w[i-15],18)^(w[i-15]>>>3), s1 = r(w[i-2],17)^r(w[i-2],19)^(w[i-2]>>>10); w[i] = (w[i-16]+s0+w[i-7]+s1)>>>0; }
      let [a,b,c,d,e,f,g,h] = H;
      for (let i = 0; i < 64; i++) {
        const t1 = (h + (r(e,6)^r(e,11)^r(e,25)) + ((e&f)^(~e&g)) + K[i] + w[i])>>>0;
        const t2 = ((r(a,2)^r(a,13)^r(a,22)) + ((a&b)^(a&c)^(b&c)))>>>0;
        h=g; g=f; f=e; e=(d+t1)>>>0; d=c; c=b; b=a; a=(t1+t2)>>>0;
      }
      H = H.map((x, i) => (x + [a,b,c,d,e,f,g,h][i]) >>> 0);
    }
    const out = [];
    H.forEach(x => { for (let k = 24; k >= 0; k -= 8) { const v = (x >>> k) & 0xff; out.push(v > 127 ? v - 256 : v); } });
    return out; // signed bytes, like Apps Script
  }

  function makeEnv() {
    const sheets = []; // {name, rows: [[...]], fmt:{}}
    const isEmptyRow = r => r.every(v => v === '' || v == null);
    function Sheet(name) {
      const s = { name, rows: [], notes: {}, maxRows: 1000 };
      const width = () => Math.max(0, ...s.rows.map(r => r.length));
      const lastRow = () => { for (let i = s.rows.length - 1; i >= 0; i--) if (!isEmptyRow(s.rows[i])) return i + 1; return 0; };
      const cell = (r, c) => (s.rows[r - 1] || [])[c - 1];
      const set = (r, c, v) => { while (s.rows.length < r) s.rows.push([]); const row = s.rows[r - 1]; while (row.length < c) row.push(''); row[c - 1] = v; };
      function Range(r, c, nr, nc) {
        return {
          getValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => { const v = cell(r + i, c + j); return v === undefined ? '' : v; })),
          setValues: vals => { if (vals.length !== nr || vals.some(x => x.length !== nc)) throw new Error('setValues dimension mismatch'); vals.forEach((row, i) => row.forEach((v, j) => set(r + i, c + j, v))); return api_r; },
          setValue: v => { set(r, c, v); return api_r; },
          getValue: () => cell(r, c) ?? '',
          setNumberFormat: () => api_r, setFontWeight: () => api_r, setBackground: () => api_r,
          setNote: t => { s.notes[r + ':' + c] = t; return api_r; }, clearNote: () => { delete s.notes[r + ':' + c]; return api_r; },
        };
      }
      let api_r;
      const A1 = a => { const m = a.match(/^([A-Z]+)(\d+)$/); let c = 0; for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64; return [Number(m[2]), c]; };
      const api = {
        getName: () => name,
        getLastRow: lastRow, getMaxRows: () => Math.max(s.maxRows, s.rows.length),
        insertRowAfter: () => { s.maxRows++; },
        getRange: (a, b, c, d) => { if (typeof a === 'string') { const [r, cc] = A1(a); return (api_r = Range(r, cc, 1, 1)); } return (api_r = Range(a, b, c || 1, d || 1)); },
        getDataRange: () => { const lr = lastRow(); return (api_r = Range(1, 1, Math.max(lr, 1), Math.max(width(), 1))); },
        appendRow: row => { const r = lastRow() + 1; row.forEach((v, j) => set(r, j + 1, v)); return api; },
        deleteRow: r => { s.rows.splice(r - 1, 1); },
        clear: () => { s.rows = []; }, setFrozenRows: () => {},
        _s: s,
      };
      return api;
    }
    const ss = {
      getSheetByName: n => sheets.find(x => x.getName() === n) || null,
      insertSheet: n => { const x = Sheet(n); sheets.push(x); return x; },
      getSheets: () => sheets.slice(),
      deleteSheet: x => { sheets.splice(sheets.indexOf(x), 1); },
      getSpreadsheetTimeZone: () => 'Asia/Bangkok',
    };
    const cache = new Map(), props = new Map();
    const prompts = [];
    let lockHeld = false;
    const env = {
      SpreadsheetApp: {
        getActive: () => ss,
        getUi: () => ({ prompt: () => { const v = prompts.shift(); return { getSelectedButton: () => 'OK', getResponseText: () => v }; }, alert: () => {}, Button: { OK: 'OK' }, ButtonSet: { OK_CANCEL: 1 }, createMenu: () => ({ addItem() { return this; }, addSeparator() { return this; }, addToUi() {} }) }),
      },
      Utilities: {
        getUuid: () => 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => { const r = Math.random() * 16 | 0; return (c === 'x' ? r : (r & 3 | 8)).toString(16); }),
        computeDigest: (alg, str) => sha256(new TextEncoder().encode(str)),
        DigestAlgorithm: { SHA_256: 1 }, Charset: { UTF_8: 1 },
        formatDate: (d, tz, f) => { const t = new Date(d.getTime() + 7 * 3600e3); return t.toISOString().slice(0, 10); },
      },
      CacheService: { getScriptCache: () => ({ get: k => { if (k.length > 250) throw new Error('cache key too long'); return cache.has(k) ? cache.get(k) : null; }, put: (k, v) => { if (k.length > 250) throw new Error('cache key too long'); cache.set(k, v); }, remove: k => cache.delete(k) }) },
      LockService: { getScriptLock: () => ({ tryLock: () => { if (lockHeld) throw new Error('NESTED LOCK'); lockHeld = true; return true; }, releaseLock: () => { lockHeld = false; } }) },
      PropertiesService: { getScriptProperties: () => ({ getProperty: k => props.get(k) || null, setProperty: (k, v) => props.set(k, v) }) },
      ScriptApp: { getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/TEST/exec' }) },
      Session: { getScriptTimeZone: () => 'Asia/Bangkok' },
      HtmlService: {},
    };
    return { env, ss, cache, prompts, sheets };
  }

  // load Code.gs source into an isolated scope; returns all top-level functions
  function loadServer(src) {
    const m = makeEnv();
    const names = [...src.matchAll(/^function\s+([A-Za-z0-9_]+)/gm)].map(x => x[1]);
    const f = new Function(...Object.keys(m.env), src + '\nreturn {' + names.join(',') + ', SH: SH};');
    m.api = f(...Object.values(m.env));
    return m;
  }

  // valid Thai national id from 12 digits
  function cid(prefix12) { let s = 0; for (let i = 0; i < 12; i++) s += Number(prefix12[i]) * (13 - i); return prefix12 + ((11 - (s % 11)) % 10); }

  root.GasMock = { loadServer, cid, sha256 };
})(window);
