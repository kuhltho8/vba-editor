/*
 * highlight.js – Syntaxfarben wie im VBA-Editor:
 * Schlüsselwörter blau, Kommentare grün, alles andere normal.
 * Liefert HTML für eine <pre>-Ebene, die hinter dem (transparenten) Textfeld liegt.
 */
(function (root) {
  'use strict';
  const KEYWORDS = new Set((
    'and as boolean byref byte byval call case const currency date debug ' +
    'dim do double each else elseif empty end eqv error exit false for function get global goto if imp in ' +
    'integer is let like long loop me mod new next not nothing null object on option optional or paramarray ' +
    'preserve print private property public redim rem resume select set single static step stop string sub ' +
    'then to true type until variant wend while with xor'
  ).split(' '));
  // "Debug" ist im VBA-Editor nicht blau, "Print" schon
  KEYWORDS.delete('debug');

  const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

  const OPTION_WORDS = new Set(['explicit', 'compare', 'database', 'binary', 'text', 'base', 'module']);

  function highlightLine(line) {
    const isOption = /^\s*option\s/i.test(line);
    let out = '';
    let i = 0;
    let atStmtStart = true;
    while (i < line.length) {
      const c = line[i];
      if (c === "'") { out += '<span class="hl-comment">' + esc(line.slice(i)) + '</span>'; break; }
      if (c === '"') {
        let j = i + 1;
        while (j < line.length) {
          if (line[j] === '"') { if (line[j + 1] === '"') { j += 2; continue; } j++; break; }
          j++;
        }
        out += esc(line.slice(i, j));
        i = j;
        atStmtStart = false;
        continue;
      }
      if (/[A-Za-zÄÖÜäöüß]/.test(c)) {
        let j = i;
        while (j < line.length && /[A-Za-z0-9_ÄÖÜäöüß]/.test(line[j])) j++;
        const word = line.slice(i, j);
        const lower = word.toLowerCase();
        const prev = line[i - 1];
        if (lower === 'rem' && atStmtStart) { out += '<span class="hl-comment">' + esc(line.slice(i)) + '</span>'; break; }
        // nach einem Punkt ist es eine Eigenschaft (z. B. E_Zahl.Value) – außer Debug.Print
        const afterDot = prev === '.';
        if ((KEYWORDS.has(lower) || (isOption && OPTION_WORDS.has(lower))) && (!afterDot || lower === 'print')) out += '<span class="hl-kw">' + esc(word) + '</span>';
        else out += esc(word);
        i = j;
        atStmtStart = false;
        continue;
      }
      if (c === ':') atStmtStart = true;
      else if (c !== ' ' && c !== '\t') atStmtStart = false;
      out += esc(c);
      i++;
    }
    return out;
  }

  /** @param {string} src Quelltext  @param {number|null} errorLine 1-basiert, wird gelb hinterlegt */
  function highlight(src, errorLine) {
    const lines = src.split('\n');
    let continued = false; // Kommentar mit " _" am Ende setzt sich fort
    return lines.map((l, idx) => {
      let html;
      if (continued) html = '<span class="hl-comment">' + esc(l) + '</span>';
      else html = highlightLine(l);
      continued = /<span class="hl-comment">/.test(html) && / _\s*$/.test(l);
      if (errorLine && idx + 1 === errorLine) html = '<span class="hl-errline">' + (html || ' ') + '</span>';
      return html;
    }).join('\n') + '\n ';
  }

  const api = { highlight, highlightLine, HL_KEYWORDS: KEYWORDS };
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.VBAHighlight = api;
})(typeof self !== 'undefined' ? self : this);
