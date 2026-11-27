/*
 * intellisense.js – Live-Hilfe beim Tippen wie im VBA-Editor:
 *  - gelbe QuickInfo mit der Syntax (MsgBox(Prompt, [Buttons As VbMsgBoxStyle = vbOKOnly], …)),
 *    der aktuelle Parameter ist fett
 *  - Vorschlagsliste (vb-Konstanten, Funktionen, Schlüsselwörter, Steuerelemente, Variablen,
 *    nach "lstAusgabe." die Eigenschaften und Methoden)
 * Auswahl: Doppelklick, Tab oder Enter, Pfeiltasten; auf Touch-Geräten Antippen. Strg+Leertaste öffnet die Liste.
 *
 * Die Analyse (analyze) arbeitet ohne DOM und wird in tests/run-tests.js mitgeprüft.
 */
(function (root) {
  'use strict';

  // ---------- Daten ----------
  const KEYWORDS = ('And As Boolean ByRef Byte ByVal Call Case Const Currency Date Debug Dim Do Double Each Else ElseIf ' +
    'Empty End Exit Explicit False For Function If In Integer Is Long Loop Me Mod Next Not Nothing Null Option Or ' +
    'Preserve Private Public ReDim Select Single Static Step String Sub Then To True Until Variant Wend While Xor ' +
    'GoTo Resume Stop Err DoCmd').split(' ');

  const TYPES = ['Boolean', 'Byte', 'Currency', 'Date', 'Double', 'Integer', 'Long', 'Single', 'String', 'Variant'];

  const CONSTANT_NAMES = ('vbCr vbLf vbCrLf vbNewLine vbTab vbNullString vbNullChar vbBack ' +
    'vbOKOnly vbOKCancel vbAbortRetryIgnore vbYesNoCancel vbYesNo vbRetryCancel vbCritical vbQuestion vbExclamation vbInformation ' +
    'vbDefaultButton1 vbDefaultButton2 vbDefaultButton3 vbOK vbCancel vbAbort vbRetry vbIgnore vbYes vbNo ' +
    'vbTrue vbFalse vbBinaryCompare vbTextCompare vbDatabaseCompare ' +
    'vbUseSystemDayOfWeek vbSunday vbMonday vbTuesday vbWednesday vbThursday vbFriday vbSaturday ' +
    'vbEmpty vbNull vbInteger vbLong vbSingle vbDouble vbCurrency vbDate vbString vbBoolean vbVariant vbByte vbArray ' +
    'vbBlack vbRed vbGreen vbBlue vbWhite vbYellow').split(' ');

  const ENUMS = {
    VbMsgBoxStyle: 'vbOKOnly vbOKCancel vbAbortRetryIgnore vbYesNoCancel vbYesNo vbRetryCancel vbCritical vbQuestion vbExclamation vbInformation vbDefaultButton1 vbDefaultButton2 vbDefaultButton3'.split(' '),
    VbMsgBoxResult: 'vbOK vbCancel vbAbort vbRetry vbIgnore vbYes vbNo'.split(' '),
    VbCompareMethod: 'vbBinaryCompare vbTextCompare vbDatabaseCompare'.split(' '),
    VbDayOfWeek: 'vbUseSystemDayOfWeek vbSunday vbMonday vbTuesday vbWednesday vbThursday vbFriday vbSaturday'.split(' ')
  };

  // Syntax wie in der QuickInfo des VBA-Editors
  const SIGS = [
    'MsgBox(Prompt, [Buttons As VbMsgBoxStyle = vbOKOnly], [Title], [HelpFile], [Context]) As VbMsgBoxResult',
    'InputBox(Prompt, [Title], [Default], [XPos], [YPos], [HelpFile], [Context]) As String',
    'Len(Expression)', 'Left(String, Length As Long)', 'Right(String, Length As Long)', 'Mid(String, Start As Long, [Length])',
    'InStr([Start], [String1], [String2], [Compare As VbCompareMethod = vbBinaryCompare])',
    'InStrRev(StringCheck As String, StringMatch As String, [Start As Long = -1], [Compare As VbCompareMethod = vbBinaryCompare]) As Long',
    'UCase(String)', 'LCase(String)', 'Trim(String)', 'LTrim(String)', 'RTrim(String)',
    'Replace(Expression As String, Find As String, Replace As String, [Start As Long = 1], [Count As Long = -1], [Compare As VbCompareMethod = vbBinaryCompare]) As String',
    'Space(Number As Long)', 'Spc(Count)', 'String(Number As Long, Character)', 'StrReverse(Expression As String) As String',
    'StrComp(String1, String2, [Compare As VbCompareMethod = vbBinaryCompare])',
    'Chr(CharCode As Long)', 'ChrW(CharCode As Long)', 'Asc(String As String) As Integer', 'AscW(String As String) As Integer',
    'Val(String As String) As Double',
    'CStr(Expression)', 'CInt(Expression)', 'CLng(Expression)', 'CByte(Expression)', 'CDbl(Expression)', 'CSng(Expression)',
    'CCur(Expression)', 'CDec(Expression)', 'CBool(Expression)', 'CDate(Expression)', 'CVar(Expression)', 'DateValue(Date As String)',
    'IsNumeric(Expression) As Boolean', 'IsEmpty(Expression) As Boolean', 'IsNull(Expression) As Boolean', 'IsArray(VarName) As Boolean',
    'IsMissing(ArgName) As Boolean', 'IsDate(Expression) As Boolean', 'Nz(Value, [ValueIfNull])', 'IIf(Expression, TruePart, FalsePart)',
    'Hex(Number)', 'Oct(Number)', 'TypeName(VarName) As String', 'VarType(VarName) As VbVarType',
    'Round(Number, [NumDigitsAfterDecimal As Long])', 'Int(Number)', 'Fix(Number)', 'Abs(Number)', 'Sgn(Number)',
    'Sqr(Number As Double) As Double', 'Exp(Number As Double) As Double', 'Log(Number As Double) As Double',
    'Sin(Number As Double) As Double', 'Cos(Number As Double) As Double', 'Tan(Number As Double) As Double', 'Atn(Number As Double) As Double',
    'Rnd([Number]) As Single', 'Randomize([Number])', 'Now()', 'Date()', 'Time()', 'Timer() As Single',
    'Year(Date)', 'Month(Date)', 'Day(Date)', 'Hour(Time)', 'Minute(Time)', 'Second(Time)',
    'Weekday(Date, [FirstDayOfWeek As VbDayOfWeek = vbSunday])',
    'WeekdayName(Weekday As Long, [Abbreviate As Boolean = False], [FirstDayOfWeek As VbDayOfWeek = vbUseSystemDayOfWeek]) As String',
    'MonthName(Month As Long, [Abbreviate As Boolean = False]) As String',
    'DateSerial(Year As Integer, Month As Integer, Day As Integer)', 'TimeSerial(Hour As Integer, Minute As Integer, Second As Integer)',
    'DateAdd(Interval As String, Number As Double, Date)',
    'DateDiff(Interval As String, Date1, Date2, [FirstDayOfWeek As VbDayOfWeek = vbSunday], [FirstWeekOfYear As VbFirstWeekOfYear = vbFirstJan1])',
    'Format(Expression, [Format], [FirstDayOfWeek As VbDayOfWeek = vbSunday], [FirstWeekOfYear As VbFirstWeekOfYear = vbFirstJan1])',
    'FormatNumber(Expression, [NumDigitsAfterDecimal As Long = -1], [IncludeLeadingDigit], [UseParensForNegativeNumbers], [GroupDigits]) As String',
    'FormatCurrency(Expression, [NumDigitsAfterDecimal As Long = -1], [IncludeLeadingDigit], [UseParensForNegativeNumbers], [GroupDigits]) As String',
    'FormatPercent(Expression, [NumDigitsAfterDecimal As Long = -1], [IncludeLeadingDigit], [UseParensForNegativeNumbers], [GroupDigits]) As String',
    'LBound(Arrayname, [Dimension As Long = 1]) As Long', 'UBound(Arrayname, [Dimension As Long = 1]) As Long',
    'Array(ParamArray ArgList() As Variant)',
    'Split(Expression As String, [Delimiter], [Limit As Long = -1], [Compare As VbCompareMethod = vbBinaryCompare])',
    'Join(SourceArray, [Delimiter]) As String', 'DoEvents() As Integer', 'Beep()',
    'Str(Number) As String', 'RGB(Red As Integer, Green As Integer, Blue As Integer) As Long'
  ];

  // Eigenschaften und Methoden der Steuerelemente (Schreibweise wie in Access)
  const MEMBER_NAMES = ['Value', 'Text', 'Caption', 'Visible', 'Enabled', 'Locked', 'ForeColor', 'BackColor', 'Name', 'SetFocus',
    'ListIndex', 'ListCount', 'RowSource', 'RowSourceType', 'ItemData', 'Column', 'AddItem', 'RemoveItem', 'Requery'];
  const MEMBER_SIGS = ['AddItem(Item As String, [Index])', 'RemoveItem(Index)', 'ItemData(Index As Long)', 'Column(Index As Long, [Row])',
    'Raise(Number As Long, [Source], [Description])', 'Close([ObjectType], [ObjectName], [Save])'];
  const METHODS = new Set(['setfocus', 'additem', 'removeitem', 'requery']);

  /** "Name(a, [b As T = x]) As R" → { name, label, params:[{text, enum}], ret } */
  function parseSig(text) {
    const m = /^(\w+)\((.*)\)(.*)$/.exec(text);
    const params = [];
    let depth = 0, cur = '';
    for (const ch of m[2]) {
      if (ch === '[' || ch === '(') depth++;
      if (ch === ']' || ch === ')') depth--;
      if (ch === ',' && depth === 0) { params.push(cur.trim()); cur = ''; continue; }
      cur += ch;
    }
    if (cur.trim()) params.push(cur.trim());
    return {
      name: m[1],
      params: params.map((p) => {
        const e = /As (Vb\w+)/.exec(p);
        return { text: p, enum: e && ENUMS[e[1]] ? e[1] : null, paramArray: /^ParamArray/.test(p) };
      }),
      ret: m[3].trim()
    };
  }
  const BUILTIN_SIGS = new Map(SIGS.map((s) => { const p = parseSig(s); return [p.name.toLowerCase(), p]; }));
  const MEMBER_SIG_MAP = new Map(MEMBER_SIGS.map((s) => { const p = parseSig(s); return [p.name.toLowerCase(), p]; }));
  const BUILTIN_NAMES = SIGS.map((s) => s.slice(0, s.indexOf('(')));

  const KIND_ITEM = { keyword: 'Schlüsselwort', constant: 'Konstante', builtin: 'Funktion', control: 'Steuerelement', proc: 'Prozedur', variable: 'Variable', member: '' };

  // ---------- Analyse des Codes ----------
  /** Eigene Subs/Functions und Variablen aus dem Code */
  function scanCode(code) {
    const procs = new Map();
    const vars = new Set();
    const re = /^[ \t]*(?:(?:Public|Private|Friend)[ \t]+)?(?:Static[ \t]+)?(Sub|Function)[ \t]+(\w+)[ \t]*\(([^)\n]*)\)([ \t]+As[ \t]+\w+)?/gim;
    let m;
    while ((m = re.exec(code))) {
      const params = m[3].split(',').map((t) => t.trim()).filter(Boolean);
      procs.set(m[2].toLowerCase(), {
        name: m[2], kind: m[1], ret: m[4] ? m[4].trim() : '',
        params: params.map((t) => ({ text: /^Optional\b/i.test(t) ? '[' + t.replace(/^Optional\s+/i, '') + ']' : t, enum: null, paramArray: /^ParamArray/i.test(t) }))
      });
      for (const t of params) {
        const v = /^(?:Optional\s+)?(?:ByVal\s+|ByRef\s+)?(?:ParamArray\s+)?(\w+)/i.exec(t);
        if (v) vars.add(v[1]);
      }
    }
    const decl = /^[ \t]*(?:Dim|Private|Public|Static|Const|ReDim(?:[ \t]+Preserve)?|Global)[ \t]+(.+)$/gim;
    while ((m = decl.exec(code))) {
      if (/^\s*(Sub|Function)\b/i.test(m[1])) continue;
      for (const part of m[1].split(',')) {
        const v = /^\s*(?:WithEvents\s+)?([A-Za-zÄÖÜäöüß_]\w*)/.exec(part);
        if (v && !/^(Sub|Function|Const)$/i.test(v[1])) vars.add(v[1]);
      }
    }
    return { procs, vars };
  }

  /** Zeile bis zur Schreibmarke: Strings unkenntlich machen, Kommentar erkennen. */
  function sanitize(line) {
    let out = '', inStr = false;
    for (let i = 0; i < line.length; i++) {
      const c = line[i];
      if (inStr) { out += c === '"' ? '"' : '_'; if (c === '"') inStr = false; continue; }
      if (c === '"') { inStr = true; out += '"'; continue; }
      if (c === "'") return { text: out, comment: true, inStr: false };
      out += c;
    }
    if (/^\s*rem(\s|$)/i.test(out)) return { text: out, comment: true, inStr: false };
    return { text: out, comment: false, inStr };
  }

  function findSig(name, ctx) {
    const l = name.toLowerCase();
    if (l.includes('.')) {
      const [obj, mem] = l.split('.');
      if (obj === 'err' && mem === 'raise') return MEMBER_SIG_MAP.get('raise');
      if (obj === 'docmd' && mem === 'close') return MEMBER_SIG_MAP.get('close');
      return ctx.controls.has(obj) ? MEMBER_SIG_MAP.get(mem) || null : null;
    }
    return ctx.procs.get(l) || BUILTIN_SIGS.get(l) || null;
  }

  /** Index des aktuellen Arguments (Kommas auf Klammerebene 0) */
  function argIndex(text) {
    let depth = 0, n = 0, last = 0;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (c === '(') depth++;
      else if (c === ')') depth--;
      else if (c === ',' && depth === 0) { n++; last = i + 1; }
    }
    return { index: n, current: text.slice(last) };
  }

  /** Welche Prozedur wird gerade aufgerufen? → { sig, index, current } */
  function callContext(t, ctx) {
    // 1) innerste offene Klammer: x = MsgBox("Hallo", |
    let depth = 0;
    for (let i = t.length - 1; i >= 0; i--) {
      const c = t[i];
      if (c === ')') depth++;
      else if (c === '(') {
        if (depth > 0) { depth--; continue; }
        const m = /([A-Za-z_][\w]*(?:\.[A-Za-z_]\w*)?)\s*$/.exec(t.slice(0, i));
        if (m) {
          const sig = findSig(m[1], ctx);
          if (sig) return Object.assign({ sig }, argIndex(t.slice(i + 1)));
        }
        // Klammer ohne bekannte Funktion (z. B. Rechnung) – weiter außen suchen
      }
    }
    // 2) Aufruf ohne Klammern am Anweisungsanfang: MsgBox "Hallo", |
    const stmt = t.split(/:(?!=)/).pop();
    const m = /^\s*(?:Call\s+)?([A-Za-z_][\w]*(?:\.[A-Za-z_]\w*)?)\s+(.*)$/i.exec(stmt);
    if (m && !/^(dim|if|elseif|for|do|while|select|case|const|private|public|sub|function|end|exit|next|loop|redim|static|option)$/i.test(m[1])) {
      const sig = findSig(m[1], ctx);
      if (sig && !/^\s*=/.test(m[2])) return Object.assign({ sig }, argIndex(m[2]));
    }
    return null;
  }

  /**
   * Analysiert den Text bis zur Schreibmarke.
   * @returns {{ tip: {sig, index}|null, list: {items:[{text,kind}], start, prefix}|null }}
   */
  function analyze(code, caret, controls, opts) {
    const o = opts || {};
    const lineStart = code.lastIndexOf('\n', caret - 1) + 1;
    const raw = code.slice(lineStart, caret);
    const s = sanitize(raw);
    if (s.comment) return { tip: null, list: null };
    const scan = scanCode(code);
    const ctrl = new Map((controls || []).map((c) => [c.name.toLowerCase(), c]));
    const ctx = { procs: scan.procs, controls: ctrl };
    const call = callContext(s.text, ctx);
    const tip = call ? { sig: call.sig, index: call.index } : null;
    if (s.inStr) return { tip, list: null };

    const wm = /([A-Za-zÄÖÜäöüß_][\wÄÖÜäöüß]*)?$/.exec(s.text);
    const prefix = wm[1] || '';
    const start = caret - prefix.length;
    const before = s.text.slice(0, s.text.length - prefix.length);
    if (/[\d]$/.test(before) && prefix) return { tip, list: null }; // 12e…
    let items = null;
    let auto = false;

    const dot = /([A-Za-z_][\w]*)\.$/.exec(before);
    if (/\bAs\s+$/i.test(before) && !call) {
      // nach "As": Datentypen
      items = TYPES.map((n) => ({ text: n, kind: 'keyword' }));
      auto = true;
    } else if (dot) {
      // nach "lstAusgabe." oder "Me." oder "Debug."
      const obj = dot[1].toLowerCase();
      if (obj === 'me') items = [...ctrl.values()].map((c) => ({ text: c.name, kind: 'control' }))
        .concat(['Caption', 'Refresh', 'Requery', 'Recalc', 'Repaint'].map((n) => ({ text: n, kind: n === 'Caption' ? 'property' : 'method' })));
      else if (obj === 'err' && !ctrl.has('err')) items = ['Number', 'Description', 'Source', 'Clear', 'Raise'].map((n) => ({ text: n, kind: /^(Clear|Raise)$/.test(n) ? 'method' : 'property' }));
      else if (obj === 'docmd') items = [{ text: 'Close', kind: 'method' }];
      else if (obj === 'debug') items = [{ text: 'Print', kind: 'member' }, { text: 'Assert', kind: 'member' }];
      else if (ctrl.has(obj)) {
        const def = (root.VBA && root.VBA.CONTROL_PROPS) ? root.VBA.CONTROL_PROPS[ctrl.get(obj).kind] : null;
        const allowed = def ? new Set(def.p.concat(def.m)) : null;
        items = MEMBER_NAMES.filter((n) => !allowed || allowed.has(n.toLowerCase()))
          .map((n) => ({ text: n, kind: METHODS.has(n.toLowerCase()) ? 'method' : 'property' }));
      } else return { tip, list: null };
      auto = true;
    } else if (call && call.sig.params.length) {
      // Konstantenliste für Parameter wie Buttons As VbMsgBoxStyle
      const p = call.sig.params[Math.min(call.index, call.sig.params.length - 1)];
      const cur = call.current.slice(0, call.current.length - prefix.length);
      if (p && p.enum && call.index < call.sig.params.length && (/^\s*$/.test(cur) || /(\+|\bOr)\s*$/i.test(cur))) {
        const enumItems = ENUMS[p.enum].map((n) => ({ text: n, kind: 'constant' }));
        if (!prefix || enumItems.some((it) => it.text.toLowerCase().startsWith(prefix.toLowerCase()))) {
          items = enumItems;
          auto = !prefix ? (o.typed === ',' || o.typed === ' ' || o.typed === '+' || o.manual) : true;
        }
      }
    }
    if (!items) {
      if (!o.manual && prefix.length < 2) return { tip, list: null };
      items = [];
      const seen = new Set();
      const add = (text, kind) => { const l = text.toLowerCase(); if (!seen.has(l)) { seen.add(l); items.push({ text, kind }); } };
      for (const c of ctrl.values()) add(c.name, 'control');
      for (const v of scan.vars) add(v, 'variable');
      for (const p of scan.procs.values()) if (!/_(click|afterupdate|change|dblclick|load)$/i.test(p.name)) add(p.name, 'proc');
      for (const n of BUILTIN_NAMES) add(n, 'builtin');
      for (const n of CONSTANT_NAMES) add(n, 'constant');
      for (const n of KEYWORDS) add(n, 'keyword');
      auto = true;
    }
    const lp = prefix.toLowerCase();
    let filtered = items.filter((it) => it.text.toLowerCase().startsWith(lp));
    if (!o.manual && filtered.length === 1 && filtered[0].text.toLowerCase() === lp) filtered = [];
    if (!filtered.length || !auto) return { tip, list: null };
    filtered.sort((a, b) => a.text.localeCompare(b.text, 'de', { sensitivity: 'base' }));
    const exact = filtered.findIndex((it) => it.text.toLowerCase() === lp);
    return { tip, list: { items: filtered, start, prefix, selected: exact >= 0 ? exact : 0 } };
  }

  /** QuickInfo als HTML, aktueller Parameter fett */
  function tipHtml(tip) {
    const esc = (t) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const { sig, index } = tip;
    let active = index;
    if (active >= sig.params.length && sig.params.length && sig.params[sig.params.length - 1].paramArray) active = sig.params.length - 1;
    const ps = sig.params.map((p, i) => (i === active ? '<b>' + esc(p.text) + '</b>' : esc(p.text)));
    return esc(sig.name) + '(' + ps.join(', ') + ')' + (sig.ret ? ' ' + esc(sig.ret) : '');
  }

  // ---------- Oberfläche ----------
  class IntelliSense {
    /** @param {HTMLTextAreaElement} ta  @param {() => object[]} getControls */
    constructor(ta, getControls) {
      this.ta = ta;
      this.getControls = getControls || (() => []);
      this.list = null;
      this.tipEl = document.createElement('div');
      this.tipEl.className = 'is-tip hidden';
      this.tipEl.setAttribute('role', 'tooltip');
      this.listEl = document.createElement('ul');
      this.listEl.className = 'is-list hidden';
      this.listEl.setAttribute('role', 'listbox');
      this.listEl.setAttribute('aria-label', 'Vorschläge');
      document.body.appendChild(this.tipEl);
      document.body.appendChild(this.listEl);
      this.measure = document.createElement('span');
      this.measure.className = 'is-measure';
      this.measure.textContent = 'MMMMMMMMMM';
      document.body.appendChild(this.measure);
      this.lastTyped = '';

      ta.addEventListener('beforeinput', (e) => { this.lastTyped = e.data || (e.inputType === 'insertLineBreak' ? '\n' : ''); });
      ta.addEventListener('input', () => {
        const t = this.lastTyped; this.lastTyped = '';
        if (!this.accepting) this.update({ typed: t, fromInput: true });
      });
      ta.addEventListener('click', () => this.update({}));
      ta.addEventListener('keyup', (e) => { if (!this.list && /^(Arrow\w+|Home|End|PageUp|PageDown)$/.test(e.key)) this.update({}); });
      ta.addEventListener('scroll', () => this.reposition());
      ta.addEventListener('blur', () => setTimeout(() => { if (document.activeElement !== ta) this.hide(); }, 150));
      window.addEventListener('resize', () => this.hide());

      // Liste: Maus – Klick markiert, Doppelklick übernimmt; Touch – Antippen übernimmt
      this.listEl.addEventListener('pointerdown', (e) => {
        e.preventDefault(); // Fokus bleibt im Editor
        const li = e.target.closest('li');
        if (!li) return;
        const i = Number(li.dataset.i);
        if (e.pointerType === 'touch' || e.pointerType === 'pen') { this.accept(i); return; }
        this.select(i);
      });
      this.listEl.addEventListener('dblclick', (e) => {
        const li = e.target.closest('li');
        if (li) this.accept(Number(li.dataset.i));
      });
    }

    /** Für app.js: true, wenn die Taste hier verbraucht wurde */
    handleKey(e) {
      if ((e.ctrlKey || e.metaKey) && (e.key === ' ' || e.code === 'Space')) { e.preventDefault(); this.update({ manual: true }); return true; }
      if (e.key === 'Escape' && (this.list || !this.tipEl.classList.contains('hidden'))) { e.preventDefault(); this.hide(); return true; }
      if (!this.list) return false;
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
        e.preventDefault();
        const n = this.list.items.length;
        this.select((this.list.selected + (e.key === 'ArrowDown' ? 1 : n - 1)) % n);
        return true;
      }
      if (e.key === 'PageDown' || e.key === 'PageUp') {
        e.preventDefault();
        this.select(Math.max(0, Math.min(this.list.items.length - 1, this.list.selected + (e.key === 'PageDown' ? 8 : -8))));
        return true;
      }
      if (e.key === 'Tab') { e.preventDefault(); this.accept(this.list.selected); return true; }
      if (e.key === 'Enter') { this.accept(this.list.selected); return false; } // wie im VBA-Editor: übernehmen und neue Zeile
      return false;
    }

    update(o) {
      const ta = this.ta;
      if (ta.selectionStart !== ta.selectionEnd) { this.hide(); return; }
      const r = analyze(ta.value, ta.selectionStart, this.getControls(), o);
      this.tip = r.tip;
      if (r.tip) {
        this.tipEl.innerHTML = tipHtml(r.tip);
        this.tipEl.classList.remove('hidden');
      } else this.tipEl.classList.add('hidden');
      // Liste nur beim Tippen (oder Strg+Leertaste) öffnen, sonst nur aktualisieren
      if (r.list && (o.fromInput || o.manual || this.list)) {
        this.list = r.list;
        this.renderList();
      } else this.closeList();
      this.reposition();
    }

    renderList() {
      const ul = this.listEl;
      ul.innerHTML = '';
      this.list.items.forEach((it, i) => {
        const li = document.createElement('li');
        li.dataset.i = i;
        li.setAttribute('role', 'option');
        const icon = document.createElement('span');
        icon.className = 'is-icon is-' + it.kind;
        icon.setAttribute('aria-hidden', 'true');
        li.appendChild(icon);
        li.appendChild(document.createTextNode(it.text));
        if (KIND_ITEM[it.kind]) li.title = KIND_ITEM[it.kind];
        ul.appendChild(li);
      });
      ul.classList.remove('hidden');
      this.select(this.list.selected);
    }

    select(i) {
      if (!this.list) return;
      this.list.selected = i;
      [...this.listEl.children].forEach((li, k) => {
        li.classList.toggle('sel', k === i);
        li.setAttribute('aria-selected', String(k === i));
      });
      const li = this.listEl.children[i];
      if (li) {
        const ul = this.listEl;
        if (li.offsetTop < ul.scrollTop) ul.scrollTop = li.offsetTop;
        else if (li.offsetTop + li.offsetHeight > ul.scrollTop + ul.clientHeight) ul.scrollTop = li.offsetTop + li.offsetHeight - ul.clientHeight;
      }
    }

    accept(i) {
      const l = this.list;
      if (!l || !l.items[i]) return;
      const ta = this.ta;
      const text = l.items[i].text;
      const end = ta.selectionStart;
      ta.focus();
      ta.setSelectionRange(l.start, end);
      this.accepting = true;
      let ok = false;
      try { ok = document.execCommand('insertText', false, text); } catch (e) { ok = false; }
      if (!ok) {
        ta.setRangeText(text, l.start, end, 'end');
        ta.dispatchEvent(new Event('input'));
      }
      this.accepting = false;
      this.closeList();
      this.update({});
    }

    closeList() {
      this.list = null;
      this.listEl.classList.add('hidden');
    }

    hide() {
      this.closeList();
      this.tipEl.classList.add('hidden');
    }

    /** Pixelposition der Schreibmarke (Festbreitenschrift, Tabs = 4 Zeichen) */
    caretXY(pos) {
      const ta = this.ta;
      const cs = getComputedStyle(ta);
      this.measure.style.font = cs.font;
      const cw = this.measure.getBoundingClientRect().width / 10 || 8.4;
      const lh = parseFloat(cs.lineHeight) || 21;
      const before = ta.value.slice(0, pos);
      const row = before.split('\n').length - 1;
      const lineText = before.slice(before.lastIndexOf('\n') + 1);
      let col = 0;
      for (const ch of lineText) col = ch === '\t' ? (Math.floor(col / 4) + 1) * 4 : col + 1;
      const rect = ta.getBoundingClientRect();
      const x = rect.left + parseFloat(cs.paddingLeft) + col * cw - ta.scrollLeft;
      const y = rect.top + parseFloat(cs.paddingTop) + row * lh - ta.scrollTop;
      return { x, y, lh, rect };
    }

    reposition() {
      const tipShown = !this.tipEl.classList.contains('hidden');
      if (!tipShown && !this.list) return;
      const ta = this.ta;
      const c = this.caretXY(this.list ? this.list.start : ta.selectionStart);
      if (c.y < c.rect.top - 2 || c.y > c.rect.bottom - c.lh / 2) { this.hide(); return; }
      const vw = document.documentElement.clientWidth;
      const vv = window.visualViewport;
      const vh = vv ? vv.height + vv.offsetTop : document.documentElement.clientHeight;
      let y = c.y + c.lh + 2;
      const place = (el, top) => {
        el.style.left = '0px';
        const w = el.offsetWidth;
        el.style.left = Math.max(4, Math.min(c.x - 4, vw - w - 4)) + 'px';
        el.style.top = top + 'px';
      };
      if (tipShown) { place(this.tipEl, y); y += this.tipEl.offsetHeight + 2; }
      if (this.list) {
        const h = this.listEl.offsetHeight;
        // passt die Liste nicht nach unten, wird sie über der Zeile angezeigt
        if (y + h > vh - 4 && c.y - h - 2 > 0) place(this.listEl, c.y - h - 2);
        else place(this.listEl, y);
      }
    }
  }

  const api = { analyze, tipHtml, scanCode, IntelliSense };
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.VBAIntelliSense = api;
})(typeof self !== 'undefined' ? self : this);
