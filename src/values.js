/*
 * values.js – VBA-Werte, Typumwandlungen und Formatierung.
 *
 * Darstellung der Werte in JavaScript:
 *   Zahlen  → number        Strings → string      Boolean → boolean
 *   Date    → VbaDate (Seriennummer wie in VBA: Tage seit 30.12.1899)
 *   Arrays  → VbaArray      Empty/Null/Nothing → Singleton-Objekte
 *
 * Ausgabe wie in einem deutschen Access: Dezimalkomma, "Wahr"/"Falsch", Datum TT.MM.JJJJ.
 */
(function (root, factory) {
  const isNode = typeof module === 'object' && module.exports;
  const deps = isNode ? require('./errors.js') : root.VBA;
  const mod = factory(deps);
  if (isNode) module.exports = mod;
  else { root.VBA = root.VBA || {}; Object.assign(root.VBA, mod); }
})(typeof self !== 'undefined' ? self : this, function (deps) {
  'use strict';
  const { runtimeError } = deps;

  const EMPTY = Object.freeze({ vba: 'Empty', toString: () => '' });
  const NULL = Object.freeze({ vba: 'Null', toString: () => 'Null' });
  const NOTHING = Object.freeze({ vba: 'Nothing', toString: () => 'Nothing' });
  const MISSING = Object.freeze({ vba: 'Missing', toString: () => '' });

  const TYPE_DISPLAY = {
    integer: 'Integer', long: 'Long', double: 'Double', single: 'Single', currency: 'Currency',
    string: 'String', boolean: 'Boolean', date: 'Date', variant: 'Variant', byte: 'Byte'
  };

  class VbaDate {
    constructor(serial) { this.serial = serial; }
  }

  // ---------- Arrays ----------
  class VbaArray {
    /** @param {Array<{lb:number,ub:number}>|null} dims  null = dynamisches Array ohne Größe */
    constructor(dims, elemType) {
      this.elemType = elemType || 'variant';
      this.dims = dims;
      const size = dims ? dims.reduce((n, d) => n * Math.max(0, d.ub - d.lb + 1), 1) : 0;
      const def = defaultValue(this.elemType);
      this.data = new Array(size).fill(def);
    }
    get allocated() { return this.dims !== null; }
    offset(indices, line) {
      if (!this.dims) throw runtimeError('Index außerhalb des gültigen Bereichs: Das Array hat noch keine Größe (zuerst ReDim verwenden).', line, 9);
      if (indices.length !== this.dims.length) {
        throw runtimeError(`Index außerhalb des gültigen Bereichs: Das Array hat ${this.dims.length} Dimension(en), angegeben wurden ${indices.length} Index/Indizes.`, line, 9);
      }
      let off = 0;
      for (let k = 0; k < indices.length; k++) {
        const d = this.dims[k];
        const i = indices[k];
        if (i < d.lb || i > d.ub) {
          throw runtimeError(`Index außerhalb des gültigen Bereichs: Index ${i} liegt nicht zwischen ${d.lb} und ${d.ub}.`, line, 9);
        }
        off = off * (d.ub - d.lb + 1) + (i - d.lb);
      }
      return off;
    }
    clone() {
      const a = new VbaArray(this.dims ? this.dims.map((d) => ({ lb: d.lb, ub: d.ub })) : null, this.elemType);
      a.data = this.data.map((v) => (v instanceof VbaArray ? v.clone() : v));
      return a;
    }
    static fromList(list, elemType) {
      const a = new VbaArray([{ lb: 0, ub: list.length - 1 }], elemType || 'variant');
      a.data = list.slice();
      return a;
    }
  }

  function defaultValue(type) {
    switch (type) {
      case 'integer': case 'long': case 'double': case 'single': case 'currency': case 'byte': return 0;
      case 'string': return '';
      case 'boolean': return false;
      case 'date': return new VbaDate(0);
      default: return EMPTY;
    }
  }

  // ---------- Hilfsfunktionen ----------
  function isNumberLike(v) {
    return typeof v === 'number' || typeof v === 'boolean' || v instanceof VbaDate || v === EMPTY;
  }

  /** Banker's Rounding (wie VBA bei Round, CInt, CLng). */
  function bankersRound(x, digits) {
    const f = Math.pow(10, digits || 0);
    const n = +(x * f).toPrecision(15);
    const r = Math.round(n);
    const isHalf = Math.abs(n % 1) === 0.5;
    const res = isHalf ? (r % 2 === 0 ? r : r - 1) : r;
    return (res === 0 ? 0 : res) / f;
  }

  /** Zahl aus einem String lesen (Dezimalkomma oder -punkt). NaN bei Fehler. */
  function parseNumberString(s) {
    let t = String(s).trim();
    if (t === '') return NaN;
    const hex = /^&H([0-9A-F]+)$/i.exec(t);
    if (hex) return parseInt(hex[1], 16);
    if (t.includes('.') && t.includes(',')) {
      if (t.lastIndexOf(',') > t.lastIndexOf('.')) t = t.replace(/\./g, '').replace(',', '.');
      else t = t.replace(/,/g, '');
    } else if (t.includes(',')) {
      if ((t.match(/,/g) || []).length > 1) return NaN;
      t = t.replace(',', '.');
    }
    if (!/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(t)) return NaN;
    return parseFloat(t);
  }

  // ---------- Datum ----------
  const MS_PER_DAY = 86400000;
  function serialFromParts(y, m, d, h, mi, s) {
    const days = Date.UTC(y, m - 1, d) / MS_PER_DAY + 25569;
    const time = ((h || 0) * 3600 + (mi || 0) * 60 + (s || 0)) / 86400;
    return days + time;
  }
  function partsFromSerial(serial) {
    const whole = Math.floor(serial + 1e-9);
    let secs = Math.round((serial - whole) * 86400);
    let dayAdj = 0;
    if (secs >= 86400) { secs -= 86400; dayAdj = 1; }
    const dt = new Date((whole + dayAdj - 25569) * MS_PER_DAY);
    return {
      y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate(),
      h: Math.floor(secs / 3600), mi: Math.floor((secs % 3600) / 60), s: secs % 60,
      weekday: dt.getUTCDay() + 1 // 1 = Sonntag
    };
  }
  function nowSerial() {
    const n = new Date();
    return serialFromParts(n.getFullYear(), n.getMonth() + 1, n.getDate(), n.getHours(), n.getMinutes(), n.getSeconds());
  }
  function validDate(y, m, d) {
    return m >= 1 && m <= 12 && d >= 1 && d <= new Date(Date.UTC(y, m, 0)).getUTCDate();
  }
  function fixYear(y) {
    if (y < 100) return y < 30 ? 2000 + y : 1900 + y;
    return y;
  }
  /**
   * Datum aus Text lesen. Bei Literalen (#...#) gilt US-Format M/T/J,
   * bei Strings deutsches Format T.M.J; ISO (J-M-T) geht immer.
   * Liefert Seriennummer oder NaN.
   */
  function parseDateString(str, isLiteral) {
    const s = String(str).trim();
    let datePart = s;
    let timePart = '';
    const tm = /(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?\s*(AM|PM)?$/i.exec(s);
    if (tm) {
      timePart = tm[0];
      datePart = s.slice(0, tm.index).trim();
    }
    let serial = 0;
    if (datePart) {
      let y, m, d, mt;
      if ((mt = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(datePart))) { y = +mt[1]; m = +mt[2]; d = +mt[3]; }
      else if ((mt = /^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$/.exec(datePart))) { d = +mt[1]; m = +mt[2]; y = fixYear(+mt[3]); }
      else if ((mt = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/.exec(datePart))) {
        if (isLiteral) { m = +mt[1]; d = +mt[2]; } else { d = +mt[1]; m = +mt[2]; }
        y = fixYear(+mt[3]);
      } else return NaN;
      if (!validDate(y, m, d)) return NaN;
      serial = serialFromParts(y, m, d, 0, 0, 0);
    }
    if (timePart) {
      let h = +tm[1];
      const mi = +tm[2];
      const sec = tm[3] ? +tm[3] : 0;
      if (tm[4]) {
        const pm = tm[4].toUpperCase() === 'PM';
        if (h === 12) h = pm ? 12 : 0; else if (pm) h += 12;
      }
      if (h > 23 || mi > 59 || sec > 59) return NaN;
      serial += (h * 3600 + mi * 60 + sec) / 86400;
    }
    if (!datePart && !timePart) return NaN;
    return serial;
  }

  const pad2 = (n) => String(n).padStart(2, '0');
  function dateToString(serial) {
    const p = partsFromSerial(serial);
    const hasDate = Math.floor(serial + 1e-9) !== 0;
    const hasTime = Math.abs(serial - Math.floor(serial + 1e-9)) > 1e-9;
    const ds = `${pad2(p.d)}.${pad2(p.m)}.${p.y}`;
    const ts = `${pad2(p.h)}:${pad2(p.mi)}:${pad2(p.s)}`;
    if (hasDate && hasTime) return ds + ' ' + ts;
    if (hasTime) return ts;
    return ds;
  }

  // ---------- Zahlen als Text ----------
  function numberToString(n) {
    if (!isFinite(n)) return n > 0 ? '1,#INF' : (n < 0 ? '-1,#INF' : '-1,#IND');
    if (Object.is(n, -0)) n = 0;
    let s;
    const abs = Math.abs(n);
    if (abs !== 0 && (abs >= 1e15 || abs < 1e-6)) {
      s = n.toExponential(14).replace(/\.?0+e/, 'e');
      s = s.replace(/e([+-])(\d)$/, 'e$10$2').replace('e', 'E');
    } else {
      s = String(+n.toPrecision(15));
      if (/e/.test(s)) { // sehr selten (Rundungsgrenze)
        s = n.toExponential(14).replace(/\.?0+e/, 'e').replace(/e([+-])(\d)$/, 'e$10$2').replace('e', 'E');
      }
    }
    return s.replace('.', ',');
  }

  // ---------- Typumwandlungen ----------
  function typeOfValue(v) {
    if (typeof v === 'string') return 'String';
    if (typeof v === 'boolean') return 'Boolean';
    if (typeof v === 'number') {
      if (Number.isInteger(v) && v >= -32768 && v <= 32767) return 'Integer';
      if (Number.isInteger(v) && v >= -2147483648 && v <= 2147483647) return 'Long';
      return 'Double';
    }
    if (v instanceof VbaDate) return 'Date';
    if (v instanceof VbaArray) return TYPE_DISPLAY[v.elemType] + '()';
    if (v === EMPTY) return 'Empty';
    if (v === NULL) return 'Null';
    if (v === NOTHING) return 'Nothing';
    return 'Variant';
  }

  function describeValue(v) {
    if (typeof v === 'string') return `"${v}"`;
    if (v instanceof VbaArray) return 'ein Array';
    return toStr(v);
  }

  function toNumber(v, line) {
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? -1 : 0;
    if (typeof v === 'string') {
      const n = parseNumberString(v);
      if (isNaN(n)) {
        throw runtimeError(v === ''
          ? 'Typen unverträglich: Ein leerer Text ("") ist keine Zahl.'
          : `Typen unverträglich: "${v}" ist keine Zahl.`, line, 13);
      }
      return n;
    }
    if (v instanceof VbaDate) return v.serial;
    if (v === EMPTY || v === MISSING) return 0;
    if (v === NULL) throw runtimeError('Unzulässige Verwendung von Null.', line, 94);
    if (v instanceof VbaArray) throw runtimeError('Typen unverträglich: Ein ganzes Array kann nicht als Zahl verwendet werden (Index vergessen?).', line, 13);
    throw runtimeError('Typen unverträglich.', line, 13);
  }

  function toStr(v, line) {
    if (typeof v === 'string') return v;
    if (typeof v === 'number') return numberToString(v);
    if (typeof v === 'boolean') return v ? 'Wahr' : 'Falsch';
    if (v instanceof VbaDate) return dateToString(v.serial);
    if (v === EMPTY || v === MISSING) return '';
    if (v === NULL) throw runtimeError('Unzulässige Verwendung von Null.', line, 94);
    if (v instanceof VbaArray) throw runtimeError('Typen unverträglich: Ein ganzes Array kann nicht als Text verwendet werden (Index vergessen?).', line, 13);
    return String(v);
  }

  function toBool(v, line) {
    if (typeof v === 'boolean') return v;
    if (typeof v === 'number') return v !== 0;
    if (typeof v === 'string') {
      const l = v.trim().toLowerCase();
      if (l === 'wahr' || l === 'true') return true;
      if (l === 'falsch' || l === 'false') return false;
      const n = parseNumberString(v);
      if (!isNaN(n)) return n !== 0;
      throw runtimeError(`Typen unverträglich: "${v}" ist kein Wahrheitswert (Wahr/Falsch).`, line, 13);
    }
    if (v instanceof VbaDate) return v.serial !== 0;
    if (v === EMPTY || v === NULL) return false;
    throw runtimeError('Typen unverträglich: Wahrheitswert erwartet.', line, 13);
  }

  function toDate(v, line) {
    if (v instanceof VbaDate) return v;
    if (typeof v === 'number') return new VbaDate(v);
    if (typeof v === 'boolean') return new VbaDate(v ? -1 : 0);
    if (typeof v === 'string') {
      let s = parseDateString(v, false);
      if (isNaN(s)) s = parseNumberString(v);
      if (isNaN(s)) throw runtimeError(`Typen unverträglich: "${v}" ist kein gültiges Datum.`, line, 13);
      return new VbaDate(s);
    }
    if (v === EMPTY) return new VbaDate(0);
    if (v === NULL) throw runtimeError('Unzulässige Verwendung von Null.', line, 94);
    throw runtimeError('Typen unverträglich: Datum erwartet.', line, 13);
  }

  function checkRange(n, min, max, typeName, line) {
    if (n < min || n > max) {
      throw runtimeError(`Überlauf: Der Wert ${numberToString(n)} passt nicht in den Datentyp ${typeName} (${numberToString(min)} bis ${numberToString(max)}).`, line, 6);
    }
    return n;
  }

  /** Wert für eine Variable des Typs `type` umwandeln (Zuweisung, CInt usw.). */
  function coerce(v, type, line) {
    switch (type) {
      case 'variant':
        if (v === MISSING) return v;
        return v instanceof VbaArray ? v.clone() : v;
      case 'string': return toStr(v, line);
      case 'boolean': return toBool(v, line);
      case 'date': return toDate(v, line);
      case 'double': return toNumber(v, line);
      case 'single': {
        const n = toNumber(v, line);
        if (Math.abs(n) > 3.402823e38) checkRange(n, -3.402823e38, 3.402823e38, 'Single', line);
        return n === 0 ? 0 : +Math.fround(n).toPrecision(7);
      }
      case 'currency': {
        const n = bankersRound(toNumber(v, line), 4);
        return checkRange(n, -922337203685477.5808, 922337203685477.5807, 'Currency', line);
      }
      case 'integer': return checkRange(bankersRound(toNumber(v, line)), -32768, 32767, 'Integer', line);
      case 'long': return checkRange(bankersRound(toNumber(v, line)), -2147483648, 2147483647, 'Long', line);
      case 'byte': return checkRange(bankersRound(toNumber(v, line)), 0, 255, 'Byte', line);
      default: return v;
    }
  }

  // ---------- Vergleiche ----------
  /** Vergleicht zwei Werte; Ergebnis -1, 0, 1 oder null (bei Null). */
  function compare(a, b, compareText, line) {
    if (a === NULL || b === NULL) return null;
    if (a instanceof VbaArray || b instanceof VbaArray) {
      throw runtimeError('Typen unverträglich: Arrays können nicht direkt verglichen werden.', line, 13);
    }
    const aStr = typeof a === 'string';
    const bStr = typeof b === 'string';
    if ((aStr && (bStr || b === EMPTY)) || (bStr && a === EMPTY)) {
      let x = aStr ? a : '';
      let y = bStr ? b : '';
      if (compareText) { x = x.toLowerCase(); y = y.toLowerCase(); }
      return x < y ? -1 : (x > y ? 1 : 0);
    }
    // Zahl gegen Text: Text wird als Zahl gelesen, wenn möglich.
    // Sonst gilt wie bei Variant in VBA: jede Zahl ist kleiner als ein Text.
    if (aStr || bStr) {
      const n = parseNumberString(aStr ? a : b);
      if (isNaN(n)) return aStr ? 1 : -1;
      const other = toNumber(aStr ? b : a, line);
      const x = aStr ? n : other, y = aStr ? other : n;
      return x < y ? -1 : (x > y ? 1 : 0);
    }
    const x = toNumber(a, line);
    const y = toNumber(b, line);
    return x < y ? -1 : (x > y ? 1 : 0);
  }

  // ---------- Formatierung für Debug.Print ----------
  function printFormat(v) {
    if (typeof v === 'number') return (v < 0 ? '' : ' ') + numberToString(v) + ' ';
    if (v === NULL) return 'Null';
    return toStr(v);
  }

  return {
    EMPTY, NULL, NOTHING, MISSING, VbaDate, VbaArray, TYPE_DISPLAY,
    defaultValue, isNumberLike, bankersRound, parseNumberString,
    serialFromParts, partsFromSerial, nowSerial, parseDateString, dateToString,
    numberToString, typeOfValue, describeValue,
    toNumber, toStr, toBool, toDate, coerce, compare, printFormat
  };
});
