/*
 * builtins.js – eingebaute VBA-Funktionen und Konstanten.
 *
 * Jede Funktion bekommt die bereits ausgewerteten Argumente (fehlende = undefined)
 * und einen Kontext { line, interp, argNodes }.
 * MsgBox und InputBox liefern ein Promise (Ausführung wartet auf die Eingabe).
 */
(function (root, factory) {
  const isNode = typeof module === 'object' && module.exports;
  const deps = isNode ? Object.assign({}, require('./errors.js'), require('./values.js')) : root.VBA;
  const mod = factory(deps);
  if (isNode) module.exports = mod;
  else { root.VBA = root.VBA || {}; Object.assign(root.VBA, mod); }
})(typeof self !== 'undefined' ? self : this, function (V) {
  'use strict';
  const {
    runtimeError, EMPTY, NULL, MISSING, VbaDate, VbaArray, TYPE_DISPLAY,
    bankersRound, parseNumberString, serialFromParts, partsFromSerial, nowSerial,
    parseDateString, numberToString, typeOfValue, toNumber, toStr, toBool, toDate, coerce, compare
  } = V;

  const WEEKDAYS = ['Sonntag', 'Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag'];
  const MONTHS = ['Januar', 'Februar', 'März', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Dezember'];

  const CONSTANTS = {
    vbcr: '\r', vblf: '\n', vbcrlf: '\r\n', vbnewline: '\r\n', vbtab: '\t', vbnullstring: '', vbnullchar: '\0', vbback: '\b',
    vbokonly: 0, vbokcancel: 1, vbabortretryignore: 2, vbyesnocancel: 3, vbyesno: 4, vbretrycancel: 5,
    vbcritical: 16, vbquestion: 32, vbexclamation: 48, vbinformation: 64,
    vbdefaultbutton1: 0, vbdefaultbutton2: 256, vbdefaultbutton3: 512,
    vbok: 1, vbcancel: 2, vbabort: 3, vbretry: 4, vbignore: 5, vbyes: 6, vbno: 7,
    vbtrue: -1, vbfalse: 0, vbbinarycompare: 0, vbtextcompare: 1, vbdatabasecompare: 2,
    vbusesystemdayofweek: 0, vbsunday: 1, vbmonday: 2, vbtuesday: 3, vbwednesday: 4, vbthursday: 5, vbfriday: 6, vbsaturday: 7,
    vbempty: 0, vbnull: 1, vbinteger: 2, vblong: 3, vbsingle: 4, vbdouble: 5, vbcurrency: 6, vbdate: 7, vbstring: 8,
    vbboolean: 11, vbvariant: 12, vbbyte: 17, vbarray: 8192,
    vbblack: 0, vbred: 255, vbgreen: 65280, vbblue: 16711680, vbwhite: 16777215, vbyellow: 65535
  };
  const SPECIAL_VALUES = { empty: EMPTY, null: NULL };

  // ---------- Hilfen ----------
  const num = (v, ctx) => toNumber(v, ctx.line);
  const str = (v, ctx) => toStr(v, ctx.line);
  const int = (v, ctx) => coerce(v, 'long', ctx.line);
  function argError(ctx, msg) {
    return runtimeError('Ungültiger Prozeduraufruf oder ungültiges Argument' + (msg ? ': ' + msg : '.'), ctx.line, 5);
  }
  function arrArg(v, ctx, fname) {
    if (!(v instanceof VbaArray)) throw runtimeError(`Typen unverträglich: ${fname} erwartet ein Array.`, ctx.line, 13);
    return v;
  }
  const isMissing = (v) => v === undefined || v === MISSING;

  // ---------- Zufallszahlen (wie VBA: lineare Kongruenz, 24 Bit) ----------
  const rng = { seed: 327680, last: 0 };
  function rnd(n) {
    if (n !== undefined && n < 0) {
      const buf = new DataView(new ArrayBuffer(4));
      buf.setFloat32(0, n);
      const bits = buf.getUint32(0);
      rng.seed = ((bits & 0xFFFFFF) + (bits >>> 24)) & 0xFFFFFF;
    } else if (n === 0) {
      return rng.last;
    }
    rng.seed = (rng.seed * 1140671485 + 12820163) % 16777216;
    rng.last = rng.seed / 16777216;
    return rng.last;
  }
  function randomize(seed) {
    const s = seed === undefined ? Date.now() : seed;
    const buf = new DataView(new ArrayBuffer(8));
    buf.setFloat64(0, s);
    const hi = buf.getUint32(0);
    const lo = buf.getUint32(4);
    const mixed = ((hi ^ lo) & 0xFFFF) << 8;
    rng.seed = ((rng.seed & 0xFF) | mixed) & 0xFFFFFF;
  }

  // ---------- Format ----------
  function splitSections(fmt) {
    const parts = [''];
    let inQ = false;
    for (let i = 0; i < fmt.length; i++) {
      const c = fmt[i];
      if (c === '"') inQ = !inQ;
      if (c === '\\' && i + 1 < fmt.length) { parts[parts.length - 1] += c + fmt[++i]; continue; }
      if (c === ';' && !inQ) { parts.push(''); continue; }
      parts[parts.length - 1] += c;
    }
    return parts;
  }

  function groupThousands(intStr) {
    return intStr.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  }

  function roundHalfAway(x, d) {
    const f = Math.pow(10, d);
    return Math.sign(x) * Math.round(+(Math.abs(x) * f).toPrecision(15)) / f;
  }

  function formatNumberPattern(value, fmt) {
    const sections = splitSections(fmt);
    let pattern = sections[0];
    let v = value;
    let forceNoSign = false;
    if (v < 0 && sections.length > 1 && sections[1] !== '') { pattern = sections[1]; v = -v; forceNoSign = true; }
    else if (v === 0 && sections.length > 2 && sections[2] !== '') { pattern = sections[2]; }

    // Muster in Literale und Platzhalter zerlegen
    let prefix = '', suffix = '';
    let intPh = '', decPh = '';
    let seenPh = false, seenDec = false, thousands = false, percent = false;
    const lit = (s) => { if (seenPh) suffix += s; else prefix += s; };
    for (let i = 0; i < pattern.length; i++) {
      const c = pattern[i];
      if (c === '"') {
        const end = pattern.indexOf('"', i + 1);
        lit(pattern.slice(i + 1, end === -1 ? undefined : end));
        i = end === -1 ? pattern.length : end;
      } else if (c === '\\') {
        lit(pattern[++i] || '');
      } else if (c === '0' || c === '#') {
        if (suffix) { /* Platzhalter nach Literal: Literal gehört in die Mitte → ignorieren */ }
        seenPh = true;
        if (seenDec) decPh += c; else intPh += c;
      } else if (c === '.') {
        if (!seenDec) { seenDec = true; seenPh = true; } else lit(',');
      } else if (c === ',') {
        if (seenPh && !seenDec) thousands = true; else lit('.');
      } else if (c === '%') {
        percent = true; lit('%');
      } else {
        lit(c);
      }
    }
    if (percent) v *= 100;
    const maxDec = decPh.length;
    const minDec = (decPh.match(/0+/) || [''])[0].length === decPh.length ? decPh.length : decPh.replace(/#+$/, '').length;
    const minInt = (intPh.match(/0/g) || []).length;
    const r = roundHalfAway(Math.abs(v), maxDec);
    let [ip, dp = ''] = r.toFixed(maxDec).split('.');
    if (minInt === 0 && ip === '0') ip = '';
    ip = ip.padStart(minInt, '0');
    if (thousands) ip = groupThousands(ip);
    dp = dp.replace(/0+$/, '');
    dp = dp.padEnd(minDec, '0');
    let out = ip + (dp.length > 0 ? ',' + dp : (seenDec && maxDec === 0 && pattern.includes('.') && false ? ',' : ''));
    if (!seenPh) out = '';
    const negative = v < 0 && r !== 0 && !forceNoSign;
    return (negative ? '-' : '') + prefix + out + suffix;
  }

  function formatDatePattern(serial, fmt) {
    const p = partsFromSerial(serial);
    const tokens = [];
    const re = /"[^"]*"|\\.|yyyy|yy|y|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|nn|n|ss|s|AM\/PM|am\/pm|ttttt|./gi;
    let mt;
    while ((mt = re.exec(fmt))) tokens.push(mt[0]);
    let out = '';
    for (let i = 0; i < tokens.length; i++) {
      const t = tokens[i];
      const l = t.toLowerCase();
      if (t[0] === '"') { out += t.slice(1, -1); continue; }
      if (t[0] === '\\') { out += t[1]; continue; }
      switch (l) {
        case 'yyyy': out += p.y; break;
        case 'yy': out += String(p.y % 100).padStart(2, '0'); break;
        case 'y': out += Math.floor(serial) - Math.floor(serialFromParts(p.y, 1, 1)) + 1; break;
        case 'mmmm': out += MONTHS[p.m - 1]; break;
        case 'mmm': out += MONTHS[p.m - 1].slice(0, 3); break;
        case 'mm': case 'm': {
          // nach h/hh oder vor s/ss bedeutet m Minuten
          const prev = tokens.slice(0, i).reverse().find((x) => /^[a-z]/i.test(x) && !/^[\\"]/.test(x));
          const nxt = tokens.slice(i + 1).find((x) => /^[a-z]/i.test(x) && !/^[\\"]/.test(x));
          const isMin = (prev && /^h/i.test(prev)) || (nxt && /^s/i.test(nxt));
          const val = isMin ? p.mi : p.m;
          out += l === 'mm' ? String(val).padStart(2, '0') : val;
          break;
        }
        case 'dddd': out += WEEKDAYS[p.weekday - 1]; break;
        case 'ddd': out += WEEKDAYS[p.weekday - 1].slice(0, 2); break;
        case 'dd': out += String(p.d).padStart(2, '0'); break;
        case 'd': out += p.d; break;
        case 'hh': out += String(p.h).padStart(2, '0'); break;
        case 'h': out += p.h; break;
        case 'nn': out += String(p.mi).padStart(2, '0'); break;
        case 'n': out += p.mi; break;
        case 'ss': out += String(p.s).padStart(2, '0'); break;
        case 's': out += p.s; break;
        case 'am/pm': out += p.h < 12 ? 'AM' : 'PM'; break;
        case 'ttttt': out += `${String(p.h).padStart(2, '0')}:${String(p.mi).padStart(2, '0')}:${String(p.s).padStart(2, '0')}`; break;
        case '/': out += '.'; break;
        default: out += t;
      }
    }
    return out;
  }

  /** Format(x, "@@@@"): @ = Zeichen oder Leerzeichen, & = Zeichen oder nichts, ! = von links füllen, < > = klein/groß */
  function formatStringPattern(text, f) {
    const toks = [];
    let ltr = false, cs = null;
    for (let i = 0; i < f.length; i++) {
      const ch = f[i];
      if (ch === '"') { const e = f.indexOf('"', i + 1); const lit = e === -1 ? f.slice(i + 1) : f.slice(i + 1, e); toks.push({ lit }); i = e === -1 ? f.length : e; continue; }
      if (ch === '\\') { toks.push({ lit: f[i + 1] || '' }); i++; continue; }
      if (ch === '@' || ch === '&') { toks.push({ ph: ch }); continue; }
      if (ch === '!') { ltr = true; continue; }
      if (ch === '<' || ch === '>') { cs = ch; continue; }
      toks.push({ lit: ch });
    }
    let s = cs === '>' ? text.toUpperCase() : cs === '<' ? text.toLowerCase() : text;
    const parts = new Array(toks.length);
    if (ltr) {
      let k = 0;
      toks.forEach((t, i) => { parts[i] = t.lit !== undefined ? t.lit : (k < s.length ? s[k++] : (t.ph === '@' ? ' ' : '')); });
      return parts.join('') + s.slice(k);
    }
    let k = s.length;
    for (let i = toks.length - 1; i >= 0; i--) {
      const t = toks[i];
      parts[i] = t.lit !== undefined ? t.lit : (k > 0 ? s[--k] : (t.ph === '@' ? ' ' : ''));
    }
    return s.slice(0, k) + parts.join('');
  }

  function vbaFormat(value, fmt, ctx) {
    if (value === NULL) return '';
    if (fmt === undefined || fmt === '') return value === EMPTY ? '' : toStr(value, ctx.line);
    const f = String(fmt);
    const named = f.toLowerCase();
    const n = () => toNumber(value, ctx.line);
    switch (named) {
      case 'general number': case 'allgemeine zahl': return numberToString(n());
      case 'currency': case 'währung': return formatNumberPattern(n(), '#,##0.00 €;-#,##0.00 €');
      case 'fixed': case 'festkommazahl': return formatNumberPattern(n(), '0.00');
      case 'standard': case 'standardzahl': return formatNumberPattern(n(), '#,##0.00');
      case 'percent': case 'prozentzahl': return formatNumberPattern(n(), '0.00%');
      case 'scientific': return n().toExponential(2).replace('.', ',').replace(/e([+-])(\d)$/, 'E$10$2');
      case 'yes/no': case 'ja/nein': return toBool(value, ctx.line) ? 'Ja' : 'Nein';
      case 'true/false': case 'wahr/falsch': return toBool(value, ctx.line) ? 'Wahr' : 'Falsch';
      case 'on/off': case 'ein/aus': return toBool(value, ctx.line) ? 'Ein' : 'Aus';
      case 'general date': return toStr(toDate(value, ctx.line), ctx.line);
      case 'short date': case 'datum, kurz': return formatDatePattern(toDate(value, ctx.line).serial, 'dd.mm.yyyy');
      case 'medium date': return formatDatePattern(toDate(value, ctx.line).serial, 'dd. mmm yy');
      case 'long date': case 'datum, lang': return formatDatePattern(toDate(value, ctx.line).serial, 'dddd, d. mmmm yyyy');
      case 'short time': case 'zeit, kurz': return formatDatePattern(toDate(value, ctx.line).serial, 'hh:nn');
      case 'medium time': return formatDatePattern(toDate(value, ctx.line).serial, 'hh:nn');
      case 'long time': case 'zeit, lang': return formatDatePattern(toDate(value, ctx.line).serial, 'hh:nn:ss');
      default: break;
    }
    // Datumsformat, wenn der Wert ein Datum ist oder das Muster Datumszeichen enthält
    const stripped = f.replace(/"[^"]*"|\\./g, '');
    const looksDate = /[dmyhns]/i.test(stripped) && !/[0#]/.test(stripped);
    if (value instanceof VbaDate || (looksDate && typeof value === 'string' && !isNaN(parseDateString(value, false)))) {
      return formatDatePattern(toDate(value, ctx.line).serial, f);
    }
    if (looksDate && typeof value === 'number') return formatDatePattern(value, f);
    if (/[0#]/.test(stripped)) {
      if (typeof value === 'string' && isNaN(parseNumberString(value))) return value;
      return formatNumberPattern(n(), f);
    }
    if (/[@&]/.test(stripped) && !/[0#]/.test(stripped)) return formatStringPattern(toStr(value, ctx.line), f);
    if (/^[@&<>!]+$/.test(stripped)) {
      const s = toStr(value, ctx.line);
      if (stripped.includes('>')) return s.toUpperCase();
      if (stripped.includes('<')) return s.toLowerCase();
      return s;
    }
    return toStr(value, ctx.line);
  }

  // ---------- Datumsrechnung ----------
  const INTERVALS = { yyyy: 'y', q: 'q', m: 'm', y: 'd', d: 'd', w: 'd', ww: 'ww', h: 'h', n: 'n', s: 's' };
  function dateAdd(interval, number, date, ctx) {
    const iv = INTERVALS[String(interval).toLowerCase()];
    if (!iv) throw argError(ctx, `Unbekanntes Intervall "${interval}" (erlaubt: yyyy, q, m, d, ww, h, n, s).`);
    const s = toDate(date, ctx.line).serial;
    const k = Math.trunc(number);
    if (iv === 'y' || iv === 'q' || iv === 'm') {
      const p = partsFromSerial(s);
      const months = iv === 'y' ? k * 12 : (iv === 'q' ? k * 3 : k);
      let y = p.y + Math.floor((p.m - 1 + months) / 12);
      let m = ((p.m - 1 + months) % 12 + 12) % 12 + 1;
      const maxD = new Date(Date.UTC(y, m, 0)).getUTCDate();
      const frac = s - Math.floor(s + 1e-9);
      return new VbaDate(serialFromParts(y, m, Math.min(p.d, maxD)) + frac);
    }
    const factor = { d: 1, ww: 7, h: 1 / 24, n: 1 / 1440, s: 1 / 86400 }[iv];
    return new VbaDate(s + number * factor);
  }
  function dateDiff(interval, d1, d2, ctx) {
    const iv = INTERVALS[String(interval).toLowerCase()];
    if (!iv) throw argError(ctx, `Unbekanntes Intervall "${interval}" (erlaubt: yyyy, q, m, d, ww, h, n, s).`);
    const a = toDate(d1, ctx.line).serial;
    const b = toDate(d2, ctx.line).serial;
    const pa = partsFromSerial(a), pb = partsFromSerial(b);
    switch (iv) {
      case 'y': return pb.y - pa.y;
      case 'q': return (pb.y * 4 + Math.floor((pb.m - 1) / 3)) - (pa.y * 4 + Math.floor((pa.m - 1) / 3));
      case 'm': return (pb.y * 12 + pb.m) - (pa.y * 12 + pa.m);
      case 'd': return Math.floor(b + 1e-9) - Math.floor(a + 1e-9);
      case 'ww': return Math.trunc((Math.floor(b) - Math.floor(a)) / 7);
      case 'h': return Math.trunc(Math.round((b - a) * 86400) / 3600);
      case 'n': return Math.trunc(Math.round((b - a) * 86400) / 60);
      default: return Math.round((b - a) * 86400);
    }
  }

  // ---------- Tabelle der Funktionen ----------
  const F = {};
  const def = (names, min, max, fn) => { for (const n of names.split(',')) F[n] = { min, max, fn, name: n }; };

  // Zeichenketten
  def('len', 1, 1, ([s], c) => (s === NULL ? NULL : str(s, c).length));
  def('left', 2, 2, ([s, n], c) => { const k = int(n, c); if (k < 0) throw argError(c, 'Länge darf nicht negativ sein.'); return str(s, c).slice(0, k); });
  def('right', 2, 2, ([s, n], c) => { const k = int(n, c); if (k < 0) throw argError(c, 'Länge darf nicht negativ sein.'); const t = str(s, c); return k === 0 ? '' : t.slice(-k); });
  def('mid', 2, 3, ([s, st, n], c) => {
    const t = str(s, c);
    const start = int(st, c);
    if (start < 1) throw argError(c, 'Die Startposition bei Mid muss mindestens 1 sein.');
    if (isMissing(n)) return t.slice(start - 1);
    const k = int(n, c);
    if (k < 0) throw argError(c, 'Länge darf nicht negativ sein.');
    return t.substr(start - 1, k);
  });
  def('instr', 2, 4, (args, c) => {
    let start = 1, s1, s2, cmp = c.interp.program.compareText ? 1 : 0;
    if (args.length >= 3 && !isMissing(args[2])) {
      start = int(args[0], c); s1 = str(args[1], c); s2 = str(args[2], c);
      if (!isMissing(args[3])) cmp = int(args[3], c);
    } else { s1 = str(args[0], c); s2 = str(args[1], c); }
    if (start < 1) throw argError(c, 'Die Startposition bei InStr muss mindestens 1 sein.');
    if (cmp) { s1 = s1.toLowerCase(); s2 = s2.toLowerCase(); }
    if (start > s1.length) return 0;
    return s1.indexOf(s2, start - 1) + 1;
  });
  def('instrrev', 2, 4, ([s1, s2, st, cmp], c) => {
    let a = str(s1, c), b = str(s2, c);
    if (!isMissing(cmp) && int(cmp, c)) { a = a.toLowerCase(); b = b.toLowerCase(); }
    const start = isMissing(st) || int(st, c) === -1 ? a.length : int(st, c);
    return a.lastIndexOf(b, start - b.length) + 1;
  });
  def('ucase', 1, 1, ([s], c) => str(s, c).toUpperCase());
  def('lcase', 1, 1, ([s], c) => str(s, c).toLowerCase());
  def('trim', 1, 1, ([s], c) => str(s, c).replace(/^ +| +$/g, ''));
  def('ltrim', 1, 1, ([s], c) => str(s, c).replace(/^ +/, ''));
  def('rtrim', 1, 1, ([s], c) => str(s, c).replace(/ +$/, ''));
  def('replace', 3, 6, ([s, find, repl, st, count, cmp], c) => {
    let t = str(s, c);
    const f = str(find, c), r = str(repl, c);
    const start = isMissing(st) ? 1 : int(st, c);
    let cnt = isMissing(count) ? -1 : int(count, c);
    const textCmp = !isMissing(cmp) ? int(cmp, c) === 1 : false;
    t = t.slice(start - 1);
    if (f === '') return t;
    let out = '', pos = 0;
    const hay = textCmp ? t.toLowerCase() : t;
    const needle = textCmp ? f.toLowerCase() : f;
    while (cnt !== 0) {
      const i = hay.indexOf(needle, pos);
      if (i === -1) break;
      out += t.slice(pos, i) + r;
      pos = i + f.length;
      cnt--;
    }
    return out + t.slice(pos);
  });
  def('space', 1, 1, ([n], c) => ' '.repeat(Math.max(0, int(n, c))));
  def('spc', 1, 1, ([n], c) => ' '.repeat(Math.max(0, int(n, c))));
  // Str: Zahl als Text mit Leerzeichen für das Vorzeichen, immer mit Punkt (unabhängig von der Ländereinstellung)
  def('str', 1, 1, ([v], c) => {
    if (v === NULL) return NULL;
    const t = numberToString(num(v, c)).replace(',', '.');
    return t.startsWith('-') ? t : ' ' + t;
  });
  def('rgb', 3, 3, ([r, g, b], c) => {
    const k = (x) => Math.max(0, Math.min(255, int(x, c)));
    return k(r) + k(g) * 256 + k(b) * 65536;
  });
  def('string', 2, 2, ([n, ch], c) => {
    const k = int(n, c);
    if (k < 0) throw argError(c);
    const chr = typeof ch === 'string' ? ch.charAt(0) : String.fromCharCode(int(ch, c) % 256);
    return chr.repeat(k);
  });
  def('strreverse', 1, 1, ([s], c) => str(s, c).split('').reverse().join(''));
  def('strcomp', 2, 3, ([a, b, cmp], c) => {
    let x = str(a, c), y = str(b, c);
    if (!isMissing(cmp) && int(cmp, c) === 1) { x = x.toLowerCase(); y = y.toLowerCase(); }
    return x < y ? -1 : (x > y ? 1 : 0);
  });
  def('chr', 1, 1, ([n], c) => { const k = int(n, c); if (k < 0 || k > 255) throw argError(c, 'Chr erwartet einen Wert von 0 bis 255.'); return String.fromCharCode(k === 128 ? 8364 : k); });
  def('chrw', 1, 1, ([n], c) => String.fromCharCode(int(n, c) & 0xFFFF));
  def('asc', 1, 1, ([s], c) => { const t = str(s, c); if (t === '') throw argError(c, 'Asc erwartet mindestens ein Zeichen.'); const k = t.charCodeAt(0); return k === 8364 ? 128 : k; });
  def('ascw', 1, 1, ([s], c) => { const t = str(s, c); if (t === '') throw argError(c); return t.charCodeAt(0); });

  // Umwandlungen
  def('val', 1, 1, ([s], c) => {
    const t = str(s, c).replace(/[ \t\r\n]/g, '');
    const h = /^&H([0-9A-F]+)/i.exec(t);
    if (h) return parseInt(h[1], 16);
    const m = /^[+-]?(\d+\.?\d*|\.\d+)([eEdD][+-]?\d+)?/.exec(t);
    return m ? parseFloat(m[0].replace(/[dD]/, 'e')) : 0;
  });
  def('cstr', 1, 1, ([v], c) => str(v, c));
  def('cint', 1, 1, ([v], c) => coerce(v, 'integer', c.line));
  def('clng', 1, 1, ([v], c) => coerce(v, 'long', c.line));
  def('cbyte', 1, 1, ([v], c) => coerce(v, 'byte', c.line));
  def('cdbl', 1, 1, ([v], c) => coerce(v, 'double', c.line));
  def('csng', 1, 1, ([v], c) => coerce(v, 'single', c.line));
  def('ccur', 1, 1, ([v], c) => coerce(v, 'currency', c.line));
  def('cdec', 1, 1, ([v], c) => coerce(v, 'double', c.line));
  def('cbool', 1, 1, ([v], c) => coerce(v, 'boolean', c.line));
  def('cdate,datevalue', 1, 1, ([v], c) => coerce(v, 'date', c.line));
  def('cvar', 1, 1, ([v]) => v);
  def('isnumeric', 1, 1, ([v]) => {
    if (typeof v === 'number' || typeof v === 'boolean' || v === EMPTY) return true;
    if (typeof v === 'string') return !isNaN(parseNumberString(v));
    return false;
  });
  def('isempty', 1, 1, ([v]) => v === EMPTY);
  def('isnull', 1, 1, ([v]) => v === NULL);
  def('isarray', 1, 1, ([v]) => v instanceof VbaArray);
  def('ismissing', 1, 1, ([v]) => v === MISSING);
  def('isdate', 1, 1, ([v]) => v instanceof VbaDate || (typeof v === 'string' && !isNaN(parseDateString(v, false))));
  def('nz', 1, 2, ([v, d]) => (v === NULL || v === EMPTY ? (d === undefined ? '' : d) : v));
  def('iif', 3, 3, ([cond, a, b], c) => (toBool(cond, c.line) ? a : b));
  def('hex', 1, 1, ([v], c) => { const k = int(v, c); return (k < 0 ? (k >>> 0) : k).toString(16).toUpperCase(); });
  def('oct', 1, 1, ([v], c) => { const k = int(v, c); return (k < 0 ? (k >>> 0) : k).toString(8); });
  def('typename', 1, 1, ([v], c) => {
    const node = c.argNodes[0];
    if (node && node.type === 'Ident') {
      const cell = c.interp.findCell(node.lower);
      if (cell && cell.type !== 'variant') return TYPE_DISPLAY[cell.type] + (cell.isArray ? '()' : '');
    }
    return typeOfValue(v);
  });
  def('vartype', 1, 1, ([v], c) => {
    const codes = { Empty: 0, Null: 1, Integer: 2, Long: 3, Single: 4, Double: 5, Currency: 6, Date: 7, String: 8, Boolean: 11, Variant: 12, Byte: 17 };
    let name = F.typename.fn([v], c);
    if (name.endsWith('()')) return 8192 + (codes[name.slice(0, -2)] || 12);
    return codes[name] !== undefined ? codes[name] : 12;
  });

  // Mathematik
  def('round', 1, 2, ([v, d], c) => {
    const digits = isMissing(d) ? 0 : int(d, c);
    if (digits < 0) throw argError(c, 'Die Anzahl der Nachkommastellen darf nicht negativ sein.');
    return bankersRound(num(v, c), digits);
  });
  def('int', 1, 1, ([v], c) => Math.floor(num(v, c)));
  def('fix', 1, 1, ([v], c) => { const r = Math.trunc(num(v, c)); return r === 0 ? 0 : r; });
  def('abs', 1, 1, ([v], c) => Math.abs(num(v, c)));
  def('sgn', 1, 1, ([v], c) => Math.sign(num(v, c)));
  def('sqr', 1, 1, ([v], c) => { const n = num(v, c); if (n < 0) throw argError(c, 'Sqr: Aus einer negativen Zahl kann keine Wurzel gezogen werden.'); return Math.sqrt(n); });
  def('exp', 1, 1, ([v], c) => Math.exp(num(v, c)));
  def('log', 1, 1, ([v], c) => { const n = num(v, c); if (n <= 0) throw argError(c, 'Log ist nur für Zahlen größer als 0 definiert.'); return Math.log(n); });
  def('sin', 1, 1, ([v], c) => Math.sin(num(v, c)));
  def('cos', 1, 1, ([v], c) => Math.cos(num(v, c)));
  def('tan', 1, 1, ([v], c) => Math.tan(num(v, c)));
  def('atn', 1, 1, ([v], c) => Math.atan(num(v, c)));
  def('rnd', 0, 1, ([n], c) => +rnd(isMissing(n) ? undefined : num(n, c)).toPrecision(7));
  def('randomize', 0, 1, ([n], c) => { randomize(isMissing(n) ? undefined : num(n, c)); return EMPTY; });

  // Datum und Zeit
  def('now', 0, 0, () => new VbaDate(nowSerial()));
  def('date', 0, 0, () => new VbaDate(Math.floor(nowSerial())));
  def('time', 0, 0, () => { const s = nowSerial(); return new VbaDate(s - Math.floor(s)); });
  def('timer', 0, 0, () => { const d = new Date(); return +((d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) + d.getMilliseconds() / 1000).toFixed(2); });
  def('year', 1, 1, ([d], c) => partsFromSerial(toDate(d, c.line).serial).y);
  def('month', 1, 1, ([d], c) => partsFromSerial(toDate(d, c.line).serial).m);
  def('day', 1, 1, ([d], c) => partsFromSerial(toDate(d, c.line).serial).d);
  def('hour', 1, 1, ([d], c) => partsFromSerial(toDate(d, c.line).serial).h);
  def('minute', 1, 1, ([d], c) => partsFromSerial(toDate(d, c.line).serial).mi);
  def('second', 1, 1, ([d], c) => partsFromSerial(toDate(d, c.line).serial).s);
  def('weekday', 1, 2, ([d, first], c) => {
    const wd = partsFromSerial(toDate(d, c.line).serial).weekday;
    const f = isMissing(first) || int(first, c) === 0 ? 1 : int(first, c);
    return ((wd - f + 7) % 7) + 1;
  });
  def('weekdayname', 1, 3, ([n, abbr, first], c) => {
    const f = isMissing(first) || int(first, c) === 0 ? 2 : int(first, c);
    const k = int(n, c);
    if (k < 1 || k > 7) throw argError(c);
    const name = WEEKDAYS[(k - 1 + f - 1) % 7];
    return !isMissing(abbr) && toBool(abbr, c.line) ? name.slice(0, 2) : name;
  });
  def('monthname', 1, 2, ([n, abbr], c) => {
    const k = int(n, c);
    if (k < 1 || k > 12) throw argError(c);
    return !isMissing(abbr) && toBool(abbr, c.line) ? MONTHS[k - 1].slice(0, 3) : MONTHS[k - 1];
  });
  def('dateserial', 3, 3, ([y, m, d], c) => {
    const yy = int(y, c), mm = int(m, c), dd = int(d, c);
    const dt = new Date(Date.UTC(fixYearSerial(yy), mm - 1, 1));
    return new VbaDate(dt.getTime() / 86400000 + 25569 + dd - 1);
  });
  def('timeserial', 3, 3, ([h, m, s], c) => new VbaDate((int(h, c) * 3600 + int(m, c) * 60 + int(s, c)) / 86400));
  def('dateadd', 3, 3, ([iv, n, d], c) => dateAdd(str(iv, c), num(n, c), d, c));
  def('datediff', 3, 5, ([iv, a, b], c) => dateDiff(str(iv, c), a, b, c));
  function fixYearSerial(y) { return y >= 0 && y < 100 ? (y < 30 ? 2000 + y : 1900 + y) : y; }

  // Formatierung
  def('format', 1, 2, ([v, f], c) => vbaFormat(v, isMissing(f) ? undefined : str(f, c), c));
  def('formatnumber', 1, 2, ([v, d], c) => {
    const k = isMissing(d) ? 2 : int(d, c);
    return formatNumberPattern(num(v, c), '#,##0' + (k > 0 ? '.' + '0'.repeat(k) : ''));
  });
  def('formatcurrency', 1, 2, ([v, d], c) => {
    const k = isMissing(d) ? 2 : int(d, c);
    return formatNumberPattern(num(v, c), '#,##0' + (k > 0 ? '.' + '0'.repeat(k) : '') + ' €');
  });
  def('formatpercent', 1, 2, ([v, d], c) => {
    const k = isMissing(d) ? 2 : int(d, c);
    return formatNumberPattern(num(v, c), '0' + (k > 0 ? '.' + '0'.repeat(k) : '') + '%');
  });

  // Arrays
  def('lbound', 1, 2, ([a, d], c) => {
    const arr = arrArg(a, c, 'LBound');
    const k = isMissing(d) ? 1 : int(d, c);
    if (!arr.dims) throw runtimeError('Index außerhalb des gültigen Bereichs: Das Array hat noch keine Größe (zuerst ReDim verwenden).', c.line, 9);
    if (k < 1 || k > arr.dims.length) throw runtimeError(`Index außerhalb des gültigen Bereichs: Das Array hat keine Dimension ${k}.`, c.line, 9);
    return arr.dims[k - 1].lb;
  });
  def('ubound', 1, 2, ([a, d], c) => {
    const arr = arrArg(a, c, 'UBound');
    const k = isMissing(d) ? 1 : int(d, c);
    if (!arr.dims) throw runtimeError('Index außerhalb des gültigen Bereichs: Das Array hat noch keine Größe (zuerst ReDim verwenden).', c.line, 9);
    if (k < 1 || k > arr.dims.length) throw runtimeError(`Index außerhalb des gültigen Bereichs: Das Array hat keine Dimension ${k}.`, c.line, 9);
    return arr.dims[k - 1].ub;
  });
  def('array', 0, Infinity, (args) => VbaArray.fromList(args.map((v) => (v === undefined ? MISSING : v)), 'variant'));
  def('split', 1, 4, ([s, d, lim], c) => {
    const t = str(s, c);
    const delim = isMissing(d) ? ' ' : str(d, c);
    if (t === '') { const a = new VbaArray([{ lb: 0, ub: -1 }], 'string'); return a; }
    let parts = delim === '' ? [t] : t.split(delim);
    const l = isMissing(lim) ? -1 : int(lim, c);
    if (l > 0 && parts.length > l) parts = parts.slice(0, l - 1).concat([parts.slice(l - 1).join(delim)]);
    return VbaArray.fromList(parts, 'string');
  });
  def('join', 1, 2, ([a, d], c) => {
    const arr = arrArg(a, c, 'Join');
    const delim = isMissing(d) ? ' ' : str(d, c);
    return arr.data.map((v) => str(v, c)).join(delim);
  });

  // Ein-/Ausgabe (asynchron, über das Host-Objekt)
  def('msgbox', 1, 5, ([p, b, t], c) => {
    const buttons = isMissing(b) ? 0 : int(b, c);
    const title = isMissing(t) ? 'Microsoft Access' : str(t, c);
    return c.interp.hostMsgBox(str(p, c), buttons, title);
  });
  def('inputbox', 1, 7, ([p, t, d], c) => {
    const title = isMissing(t) ? 'Microsoft Access' : str(t, c);
    const deflt = isMissing(d) ? '' : str(d, c);
    return c.interp.hostInputBox(str(p, c), title, deflt);
  });
  def('doevents', 0, 0, (args, c) => c.interp.yieldNow().then(() => 0));
  def('beep', 0, 0, () => EMPTY);

  const ASYNC_BUILTINS = new Set(['msgbox', 'inputbox', 'doevents']);

  function resetRandom() { rng.seed = 327680; rng.last = 0; }

  return { BUILTINS: F, CONSTANTS, SPECIAL_VALUES, ASYNC_BUILTINS, vbaFormat, resetRandom, compareValues: compare };
});
