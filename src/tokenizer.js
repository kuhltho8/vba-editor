/*
 * tokenizer.js – zerlegt VBA-Quelltext in Tokens.
 *
 * Token-Typen:
 *   num   Zahl (value: number)
 *   str   Zeichenkette (value: string)
 *   date  Datumsliteral #...# (value: Rohtext)
 *   ident Bezeichner (value: wie geschrieben, lower: kleingeschrieben)
 *   kw    Schlüsselwort (value: kleingeschrieben)
 *   op    Operator/Satzzeichen
 *   nl    Zeilenende
 *   colon Doppelpunkt als Anweisungstrenner
 *   eof   Ende des Quelltexts
 *
 * Kommentare (' und Rem) werden entfernt, Zeilenfortsetzungen ( _) zusammengefügt.
 * VBA ist nicht case-sensitiv: Schlüsselwörter werden kleingeschrieben abgelegt.
 */
(function (root, factory) {
  const deps = (typeof module === 'object' && module.exports) ? require('./errors.js') : root.VBA;
  const mod = factory(deps);
  if (typeof module === 'object' && module.exports) module.exports = mod;
  else { root.VBA = root.VBA || {}; Object.assign(root.VBA, mod); }
})(typeof self !== 'undefined' ? self : this, function (deps) {
  'use strict';
  const { syntaxError } = deps;

  const KEYWORDS = new Set([
    'and', 'as', 'byref', 'byval', 'call', 'case', 'const', 'dim', 'do', 'each', 'else', 'elseif',
    'end', 'eqv', 'exit', 'false', 'for', 'function', 'goto', 'if', 'imp', 'in', 'is', 'let', 'like',
    'loop', 'mod', 'new', 'next', 'not', 'nothing', 'option', 'optional', 'or', 'paramarray',
    'preserve', 'private', 'public', 'redim', 'rem', 'select', 'set', 'static', 'step', 'sub',
    'then', 'to', 'true', 'until', 'wend', 'while', 'with', 'xor'
  ]);

  const isLetter = (c) => /[A-Za-zÄÖÜäöüß]/.test(c);
  const isIdentChar = (c) => /[A-Za-z0-9_ÄÖÜäöüß]/.test(c);
  const isDigit = (c) => c >= '0' && c <= '9';

  function tokenize(source) {
    const src = source.replace(/\r\n?/g, '\n');
    const tokens = [];
    let i = 0;
    let line = 1;
    let lineStart = 0;

    const push = (type, value, extra) => {
      const t = { type, value, line, col: i - lineStart + 1 };
      if (extra) Object.assign(t, extra);
      tokens.push(t);
      return t;
    };
    const atStatementStart = () => {
      const last = tokens[tokens.length - 1];
      return !last || last.type === 'nl' || last.type === 'colon';
    };
    const skipToLineEnd = () => { while (i < src.length && src[i] !== '\n') i++; };

    while (i < src.length) {
      const c = src[i];

      // Leerraum
      if (c === ' ' || c === '\t') { i++; continue; }

      // Zeilenende
      if (c === '\n') {
        push('nl', '\n');
        i++; line++; lineStart = i;
        continue;
      }

      // Zeilenfortsetzung: " _" am Zeilenende
      if (c === '_' && (i === 0 || src[i - 1] === ' ' || src[i - 1] === '\t')) {
        let j = i + 1;
        while (j < src.length && (src[j] === ' ' || src[j] === '\t')) j++;
        if (j >= src.length || src[j] === '\n') {
          i = j + 1; line++; lineStart = i;
          continue;
        }
        throw syntaxError("Unerwartetes Zeichen '_'. Eine Zeilenfortsetzung ' _' muss am Zeilenende stehen.", line);
      }

      // Kommentar
      if (c === "'") { skipToLineEnd(); continue; }

      // Zeichenkette
      if (c === '"') {
        const startLine = line;
        let s = '';
        i++;
        for (;;) {
          if (i >= src.length || src[i] === '\n') {
            throw syntaxError('Zeichenkette ist nicht abgeschlossen – das schließende " fehlt.', startLine);
          }
          if (src[i] === '"') {
            if (src[i + 1] === '"') { s += '"'; i += 2; continue; }
            i++;
            break;
          }
          s += src[i++];
        }
        tokens.push({ type: 'str', value: s, line: startLine, col: 0 });
        continue;
      }

      // Datumsliteral #...#
      if (c === '#') {
        const end = src.indexOf('#', i + 1);
        const nlPos = src.indexOf('\n', i + 1);
        if (end === -1 || (nlPos !== -1 && nlPos < end)) {
          throw syntaxError("Datumsliteral ist nicht abgeschlossen – das schließende '#' fehlt.", line);
        }
        push('date', src.slice(i + 1, end).trim());
        i = end + 1;
        continue;
      }

      // Hex-/Oktalzahlen &H1F, &O17
      if (c === '&' && /[hHoO]/.test(src[i + 1] || '')) {
        const isHex = /[hH]/.test(src[i + 1]);
        let j = i + 2;
        const re = isHex ? /[0-9A-Fa-f]/ : /[0-7]/;
        while (j < src.length && re.test(src[j])) j++;
        if (j > i + 2) {
          const txt = src.slice(i + 2, j);
          push('num', parseInt(txt, isHex ? 16 : 8));
          i = j;
          if (src[i] === '&') i++; // Typzeichen &H..& ignorieren
          continue;
        }
      }

      // Zahlen
      if (isDigit(c) || (c === '.' && isDigit(src[i + 1] || '') && !prevIsValue(tokens))) {
        let j = i;
        while (j < src.length && isDigit(src[j])) j++;
        if (src[j] === '.' && isDigit(src[j + 1] || '')) {
          j++;
          while (j < src.length && isDigit(src[j])) j++;
        } else if (src[j] === '.' && !isLetter(src[j + 1] || '')) {
          j++;
        }
        if (/[eEdD]/.test(src[j] || '') && /[0-9+-]/.test(src[j + 1] || '')) {
          let k = j + 1;
          if (src[k] === '+' || src[k] === '-') k++;
          if (isDigit(src[k] || '')) {
            while (k < src.length && isDigit(src[k])) k++;
            j = k;
          }
        }
        const txt = src.slice(i, j).replace(/[dD]/, 'e');
        push('num', parseFloat(txt));
        i = j;
        // Typkennzeichen wie 5# oder 3! ignorieren
        if (/[#!@%]/.test(src[i] || '') && !isIdentChar(src[i + 1] || '')) i++;
        continue;
      }

      // Bezeichner und Schlüsselwörter
      if (isLetter(c)) {
        let j = i;
        while (j < src.length && isIdentChar(src[j])) j++;
        let word = src.slice(i, j);
        const lower = word.toLowerCase();
        // Typkennzeichen am Namen (Left$, zahl%, ...) – nur $ wird häufig benutzt
        if (src[j] === '$') j++;
        if (KEYWORDS.has(lower)) {
          if (lower === 'rem' && atStatementStart()) {
            i = j;
            skipToLineEnd();
            continue;
          }
          push('kw', lower, { raw: word });
        } else {
          push('ident', word, { lower });
        }
        i = j;
        continue;
      }

      // Operatoren
      const two = src.substr(i, 2);
      if (two === '<=' || two === '>=' || two === '<>' || two === ':=') {
        push('op', two); i += 2; continue;
      }
      if (c === ':') { push('colon', ':'); i++; continue; }
      if ('+-*/\\^&=<>(),;.!'.includes(c)) { push('op', c); i++; continue; }

      throw syntaxError(`Unerwartetes Zeichen '${c}'.`, line);
    }
    push('nl', '\n');
    push('eof', null);
    return tokens;
  }

  function prevIsValue(tokens) {
    const t = tokens[tokens.length - 1];
    if (!t) return false;
    return t.type === 'ident' || t.type === 'num' || t.type === 'str' || (t.type === 'op' && t.value === ')');
  }

  return { tokenize, KEYWORDS };
});
