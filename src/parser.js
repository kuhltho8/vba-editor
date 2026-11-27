/*
 * parser.js – baut aus den Tokens einen abstrakten Syntaxbaum (AST).
 *
 * Jeder Knoten hat eine Eigenschaft `line` (1-basiert), damit Laufzeitfehler
 * die richtige Zeile melden können. Bezeichner werden zusätzlich kleingeschrieben
 * in `lower` abgelegt (VBA ist nicht case-sensitiv).
 */
(function (root, factory) {
  const isNode = typeof module === 'object' && module.exports;
  const deps = isNode ? Object.assign({}, require('./errors.js'), require('./tokenizer.js')) : root.VBA;
  const mod = factory(deps);
  if (isNode) module.exports = mod;
  else { root.VBA = root.VBA || {}; Object.assign(root.VBA, mod); }
})(typeof self !== 'undefined' ? self : this, function (deps) {
  'use strict';
  const { syntaxError, tokenize } = deps;

  // Eigenschaften/Methoden des Formulars selbst (Me.Caption, Me.Refresh …)
  const FORM_MEMBERS = new Set(['caption', 'refresh', 'requery', 'recalc', 'repaint']);
  const TYPE_NAMES = {
    integer: 'integer', long: 'long', double: 'double', single: 'single', currency: 'currency',
    string: 'string', boolean: 'boolean', date: 'date', variant: 'variant', byte: 'byte',
    longlong: 'long', object: 'variant', decimal: 'double'
  };
  const TYPE_DISPLAY = {
    integer: 'Integer', long: 'Long', double: 'Double', single: 'Single', currency: 'Currency',
    string: 'String', boolean: 'Boolean', date: 'Date', variant: 'Variant', byte: 'Byte'
  };

  // Welche Abschluss-Anweisung gehört zu welchem Block?
  const BLOCK_TERMS = {
    for: ['next'], foreach: ['next'], do: ['loop'], while: ['wend'],
    if: ['endif', 'else', 'elseif'], else: ['endif'],
    select: ['case', 'endselect'], sub: ['endsub'], function: ['endfunction']
  };
  const MISSING_MSG = {
    for: "For ohne Next – zu dieser For-Schleife fehlt das 'Next'.",
    foreach: "For Each ohne Next – zu dieser Schleife fehlt das 'Next'.",
    do: "Do ohne Loop – zu dieser Do-Schleife fehlt das 'Loop'.",
    while: "While ohne Wend – zu dieser While-Schleife fehlt das 'Wend'.",
    if: "If-Block ohne End If – zu diesem If fehlt das 'End If'.",
    else: "If-Block ohne End If – zu diesem If fehlt das 'End If'.",
    select: "Select Case ohne End Select – es fehlt das 'End Select'.",
    sub: "Sub ohne End Sub – zu dieser Sub fehlt das 'End Sub'.",
    function: "Function ohne End Function – zu dieser Function fehlt das 'End Function'."
  };
  const STRAY_MSG = {
    next: 'Next ohne For.',
    loop: 'Loop ohne Do.',
    wend: 'Wend ohne While.',
    endif: 'End If ohne If-Block.',
    else: 'Else ohne If.',
    elseif: 'ElseIf ohne If.',
    endselect: 'End Select ohne Select Case.',
    case: 'Case ohne Select Case.',
    endsub: 'End Sub ohne Sub.',
    endfunction: 'End Function ohne Function.',
    procstart: 'Sub oder Function innerhalb einer anderen Prozedur ist nicht erlaubt.'
  };

  class Parser {
    constructor(tokens) {
      this.tokens = tokens;
      this.pos = 0;
      this.blockStack = [];   // offene Blöcke (Art + Zeile)
      this.proc = null;       // aktuell geparste Prozedur
      this.declNames = null;  // Map lower -> Zeile (für doppelte Deklarationen)
      this.inlineIf = 0;      // > 0 innerhalb eines einzeiligen If
    }

    // ---------- Hilfsfunktionen ----------
    peek(o = 0) { return this.tokens[Math.min(this.pos + o, this.tokens.length - 1)]; }
    next() { const t = this.tokens[this.pos]; if (this.pos < this.tokens.length - 1) this.pos++; return t; }
    isKw(k, o = 0) { const t = this.peek(o); return t.type === 'kw' && t.value === k; }
    isOp(v, o = 0) { const t = this.peek(o); return t.type === 'op' && t.value === v; }
    isIdent(lower, o = 0) { const t = this.peek(o); return t.type === 'ident' && t.lower === lower; }
    atEOS(o = 0) {
      const t = this.peek(o);
      if (t.type === 'nl' || t.type === 'colon' || t.type === 'eof') return true;
      // im einzeiligen If beendet auch 'Else' die Anweisung
      return this.inlineIf > 0 && t.type === 'kw' && t.value === 'else';
    }
    acceptKw(k) { if (this.isKw(k)) { return this.next(); } return null; }
    acceptOp(v) { if (this.isOp(v)) { return this.next(); } return null; }
    line() { return this.peek().line; }

    describe(t) {
      if (t.type === 'eof') return 'Ende des Programms';
      if (t.type === 'nl') return 'Zeilenende';
      if (t.type === 'colon') return "':'";
      if (t.type === 'str') return `"${t.value}"`;
      if (t.type === 'kw') return `'${t.raw || t.value}'`;
      return `'${t.value}'`;
    }
    expectKw(k, display) {
      if (this.isKw(k)) return this.next();
      const t = this.peek();
      throw syntaxError(`Erwartet: ${display || k}, gefunden: ${this.describe(t)}.`, t.line);
    }
    expectOp(v) {
      if (this.isOp(v)) return this.next();
      const t = this.peek();
      throw syntaxError(`Erwartet: '${v}', gefunden: ${this.describe(t)}.`, t.line);
    }
    expectIdent(what) {
      const t = this.peek();
      if (t.type === 'ident') { this.next(); return t; }
      if (t.type === 'kw') {
        throw syntaxError(`'${t.raw || t.value}' ist ein reserviertes Schlüsselwort und kann nicht als ${what || 'Name'} verwendet werden.`, t.line);
      }
      throw syntaxError(`Erwartet: ${what || 'Bezeichner'}, gefunden: ${this.describe(t)}.`, t.line);
    }
    expectEOS() {
      if (this.atEOS()) return;
      const t = this.peek();
      throw syntaxError(`Erwartet: Anweisungsende, gefunden: ${this.describe(t)}.`, t.line);
    }
    skipSeparators() {
      while (this.peek().type === 'nl' || this.peek().type === 'colon') this.next();
    }

    parseType() {
      const t = this.peek();
      if (t.type !== 'ident') throw syntaxError(`Erwartet: Datentyp nach 'As', gefunden: ${this.describe(t)}.`, t.line);
      this.next();
      const ty = TYPE_NAMES[t.lower];
      if (!ty) {
        throw syntaxError(`Unbekannter Datentyp '${t.value}'. Erlaubt sind z. B. Integer, Long, Double, String, Boolean, Date, Variant.`, t.line);
      }
      // String * 20 (feste Länge) wird als normaler String behandelt
      if (ty === 'string' && this.isOp('*')) { this.next(); this.next(); }
      return ty;
    }

    declare(name, lower, line) {
      if (!this.declNames) return;
      if (this.declNames.has(lower)) {
        throw syntaxError(`Doppelte Deklaration: '${name}' ist bereits deklariert (Zeile ${this.declNames.get(lower)}).`, line);
      }
      this.declNames.set(lower, line);
    }

    // ---------- Modulebene ----------
    parseProgram() {
      const program = {
        type: 'Program', procs: Object.create(null), procList: [], moduleDecls: [],
        optionExplicit: false, optionBase: 0, compareText: false
      };
      const moduleNames = new Map();
      for (;;) {
        this.skipSeparators();
        const t = this.peek();
        if (t.type === 'eof') break;

        // Exportierte .bas-Dateien beginnen mit "Attribute VB_Name = ..."
        if (t.type === 'ident' && t.lower === 'attribute') {
          while (!this.atEOS()) this.next();
          continue;
        }
        if (this.isKw('option')) { this.parseOption(program); continue; }

        let isPrivate = false;
        let isStatic = false;
        let save = this.pos;
        if (this.isKw('private') || this.isKw('public') || this.isIdent('global')) {
          isPrivate = this.peek().value === 'private';
          this.next();
        }
        if (this.isKw('static')) { isStatic = true; this.next(); }
        if (this.isKw('sub') || this.isKw('function')) {
          const proc = this.parseProc(isPrivate, isStatic);
          if (program.procs[proc.lower] || moduleNames.has(proc.lower)) {
            throw syntaxError(`Mehrdeutiger Name: '${proc.name}' ist bereits vorhanden.`, proc.line);
          }
          program.procs[proc.lower] = proc;
          program.procList.push(proc);
          continue;
        }
        if (isStatic) throw syntaxError("'Static' ist auf Modulebene nicht erlaubt – verwende Dim oder Private.", t.line);
        if (this.isKw('const')) {
          this.next();
          this.declNames = moduleNames;
          program.moduleDecls.push(...this.parseConstList());
          this.declNames = null;
          this.expectEOS();
          continue;
        }
        if (this.isKw('dim') || this.pos > save) {
          if (this.isKw('dim')) this.next();
          this.declNames = moduleNames;
          program.moduleDecls.push(...this.parseVarList(false));
          this.declNames = null;
          this.expectEOS();
          continue;
        }
        if (this.isKw('end') || this.isKw('next') || this.isKw('loop') || this.isKw('wend')) {
          const term = this.getTerminator();
          if (term) throw syntaxError(STRAY_MSG[term], t.line);
        }
        throw syntaxError('Anweisung außerhalb einer Prozedur. Ausführbarer Code muss zwischen Sub ... End Sub (oder Function ... End Function) stehen.', t.line);
      }
      for (const p of program.procList) {
        for (const d of p.decls) {
          // Prüfen, ob lokale Namen mit Prozedurnamen kollidieren, ist in VBA erlaubt (Verdeckung) – nichts zu tun.
          void d;
        }
      }
      return program;
    }

    parseOption(program) {
      const line = this.line();
      this.next();
      const t = this.next();
      const w = (t.lower || t.value || '').toString().toLowerCase();
      if (w === 'explicit') program.optionExplicit = true;
      else if (w === 'base') {
        const n = this.next();
        if (n.type !== 'num' || (n.value !== 0 && n.value !== 1)) throw syntaxError('Option Base erwartet 0 oder 1.', line);
        program.optionBase = n.value;
      } else if (w === 'compare') {
        const m = this.next();
        const mode = (m.lower || '').toLowerCase();
        if (mode === 'text' || mode === 'database') program.compareText = true;
        else if (mode === 'binary') program.compareText = false;
        else throw syntaxError('Option Compare erwartet Binary, Text oder Database.', line);
      } else if (w === 'private') {
        this.next(); // Option Private Module
      } else {
        throw syntaxError(`Unbekannte Option '${t.value}'. Möglich sind Option Explicit, Option Base, Option Compare.`, line);
      }
      this.expectEOS();
    }

    parseProc(isPrivate, isStatic) {
      const kwTok = this.next();
      const kind = kwTok.value; // 'sub' | 'function'
      const nameTok = this.expectIdent(kind === 'sub' ? 'Name der Sub' : 'Name der Function');
      const proc = {
        type: 'Proc', kind, name: nameTok.value, lower: nameTok.lower, params: [],
        returnType: 'variant', body: [], decls: [], line: kwTok.line, isPrivate, isStatic
      };
      this.proc = proc;
      this.declNames = new Map();
      if (this.acceptOp('(')) {
        if (!this.isOp(')')) {
          do { proc.params.push(this.parseParam()); } while (this.acceptOp(','));
        }
        this.expectOp(')');
      }
      if (kind === 'function' && this.acceptKw('as')) proc.returnType = this.parseType();
      else if (kind === 'function' && nameTok.typeChar) proc.returnType = nameTok.typeChar;
      this.expectEOS();
      if (kind === 'function') this.declNames.set(proc.lower, proc.line);

      this.blockStack = [{ kind, line: proc.line }];
      const { body, term, termLine } = this.parseBlock(kind, proc.line);
      void term;
      this.next(); this.next(); // End Sub / End Function
      proc.body = body;
      proc.endLine = termLine;
      this.expectEOS();
      this.blockStack = [];
      this.proc = null;
      this.declNames = null;
      return proc;
    }

    parseParam() {
      const p = { name: '', lower: '', byRef: true, optional: false, defaultExpr: null, type: 'variant', isArray: false, paramArray: false, line: this.line() };
      if (this.acceptKw('optional')) p.optional = true;
      if (this.acceptKw('byval')) p.byRef = false;
      else if (this.acceptKw('byref')) p.byRef = true;
      if (this.acceptKw('paramarray')) p.paramArray = true;
      const nameTok = this.expectIdent('Parametername');
      p.name = nameTok.value; p.lower = nameTok.lower;
      if (this.acceptOp('(')) { this.expectOp(')'); p.isArray = true; }
      if (this.acceptKw('as')) p.type = this.parseType();
      else if (nameTok.typeChar) p.type = nameTok.typeChar;
      if (this.acceptOp('=')) {
        if (!p.optional) throw syntaxError(`Ein Standardwert ist nur bei Optional-Parametern erlaubt ('${p.name}').`, p.line);
        p.defaultExpr = this.parseExpr();
      }
      this.declare(p.name, p.lower, p.line);
      return p;
    }

    // ---------- Blöcke ----------
    /** Liefert den Schlüssel einer Block-Abschlussanweisung am aktuellen Token, sonst null. */
    getTerminator() {
      const t = this.peek();
      if (t.type !== 'kw') return null;
      switch (t.value) {
        case 'next': case 'loop': case 'wend': case 'else': case 'elseif': case 'case':
          return t.value;
        case 'end': {
          const n = this.peek(1);
          if (n.type === 'kw' && ['if', 'select', 'sub', 'function'].includes(n.value)) return 'end' + n.value;
          return null;
        }
        case 'sub': case 'function':
          return 'procstart';
        case 'private': case 'public': {
          const n = this.peek(1);
          if (n.type === 'kw' && (n.value === 'sub' || n.value === 'function' || n.value === 'static')) return 'procstart';
          return null;
        }
        default: return null;
      }
    }

    /**
     * Liest Anweisungen, bis eine zum Block `kind` passende Abschlussanweisung kommt.
     * Die Abschlussanweisung selbst wird NICHT verbraucht.
     */
    parseBlock(kind, startLine) {
      const body = [];
      for (;;) {
        this.skipSeparators();
        const t = this.peek();
        if (t.type === 'eof') throw syntaxError(MISSING_MSG[kind], startLine);
        const term = this.getTerminator();
        if (term) {
          if (BLOCK_TERMS[kind].includes(term)) return { body, term, termLine: t.line };
          // Gehört der Abschluss zu einem äußeren Block? Dann fehlt hier etwas.
          const outer = this.blockStack.slice(0, -1).some((b) => BLOCK_TERMS[b.kind].includes(term));
          if (outer || term === 'procstart') throw syntaxError(MISSING_MSG[kind], startLine);
          if (kind === 'else' && (term === 'else' || term === 'elseif')) {
            throw syntaxError(`${term === 'else' ? 'Else' : 'ElseIf'} nach Else ist nicht erlaubt – Else muss der letzte Zweig sein.`, t.line);
          }
          if ((kind === 'sub' && term === 'endfunction') || (kind === 'function' && term === 'endsub')) {
            throw syntaxError(kind === 'sub' ? "Erwartet: 'End Sub' (die Prozedur beginnt mit Sub)." : "Erwartet: 'End Function' (die Prozedur beginnt mit Function).", t.line);
          }
          throw syntaxError(STRAY_MSG[term], t.line);
        }
        body.push(this.parseStatement());
      }
    }

    withBlock(kind, line, fn) {
      this.blockStack.push({ kind, line });
      try { return fn(); } finally { this.blockStack.pop(); }
    }

    inLoop(kind) {
      return this.blockStack.some((b) => (kind === 'for' ? (b.kind === 'for' || b.kind === 'foreach') : b.kind === kind));
    }

    // ---------- Anweisungen ----------
    parseStatement() {
      const t = this.peek();
      const line = t.line;
      if (t.type === 'kw') {
        switch (t.value) {
          case 'dim': this.next(); return this.parseDimStatement(false, line);
          case 'static': this.next(); return this.parseDimStatement(true, line);
          case 'const': this.next(); return this.parseConstStatement(line);
          case 'redim': this.next(); return this.parseReDim(line);
          case 'if': return this.parseIf();
          case 'select': return this.parseSelect();
          case 'for': return this.parseFor();
          case 'do': return this.parseDo();
          case 'while': return this.parseWhile();
          case 'exit': return this.parseExit();
          case 'call': return this.parseCallStatement();
          case 'let': case 'set': this.next(); return this.parseAssignOrCall(true);
          case 'end': this.next(); this.expectEOS(); return { type: 'End', line };
          case 'option': throw syntaxError('Option-Anweisungen müssen ganz oben im Modul stehen (vor der ersten Sub).', line);
          case 'private': case 'public':
            throw syntaxError(`'${t.raw}' ist innerhalb einer Prozedur nicht erlaubt – verwende Dim.`, line);
          case 'with': throw syntaxError('With-Blöcke werden (noch) nicht unterstützt.', line);
          case 'goto': {
            this.next();
            const lab = this.parseLabelRef('GoTo');
            this.expectEOS();
            return { type: 'GoTo', label: lab.lower, labelName: lab.name, line };
          }
          default:
            throw syntaxError(`Unerwartetes Schlüsselwort '${t.raw || t.value}'.`, line);
        }
      }
      if (t.type === 'ident') {
        // Sprungmarke "Fehler:" am Anfang der Anweisung
        const prev = this.tokens[this.pos - 1];
        if (this.peek(1).type === 'colon' && !t.bracketed && (!prev || prev.type === 'nl') && !this.inlineIf) {
          this.next();
          return { type: 'Label', name: t.value, lower: t.lower, line };
        }
        if (t.lower === 'debug' && this.isOp('.', 1)) return this.parseDebug();
        if (t.lower === 'on' && this.isIdent('error', 1)) return this.parseOnError(line);
        if (t.lower === 'resume' && !this.isOp('=', 1) && !this.isOp('.', 1)) return this.parseResume(line);
        if (t.lower === 'stop' && this.atEOS(1)) { this.next(); return { type: 'Stop', line }; }
        return this.parseAssignOrCall(false);
      }
      throw syntaxError(`Unerwartet: ${this.describe(t)}. Hier wird eine Anweisung erwartet.`, line);
    }

    /** Name einer Sprungmarke nach GoTo/Resume */
    parseLabelRef(what) {
      const t = this.peek();
      if (t.type === 'ident') { this.next(); return { name: t.value, lower: t.lower }; }
      if (t.type === 'num') { this.next(); return { name: String(t.value), lower: String(t.value) }; }
      throw syntaxError(`Erwartet: Name einer Sprungmarke nach '${what}'.`, t.line);
    }

    parseOnError(line) {
      this.next(); this.next(); // On Error
      if (this.acceptKw('goto')) {
        const lab = this.parseLabelRef('On Error GoTo');
        this.expectEOS();
        if (lab.lower === '0') return { type: 'OnError', mode: 'off', line };
        if (lab.lower === '-1') return { type: 'OnError', mode: 'reset', line };
        return { type: 'OnError', mode: 'goto', label: lab.lower, labelName: lab.name, line };
      }
      if (this.isIdent('resume')) {
        this.next();
        if (!this.acceptKw('next')) throw syntaxError("Erwartet: 'Next' nach 'On Error Resume'.", line);
        this.expectEOS();
        return { type: 'OnError', mode: 'next', line };
      }
      throw syntaxError("Erwartet: 'GoTo' oder 'Resume Next' nach 'On Error'.", line);
    }

    parseResume(line) {
      this.next();
      if (this.atEOS()) return { type: 'Resume', mode: 'retry', line };
      if (this.acceptKw('next')) { this.expectEOS(); return { type: 'Resume', mode: 'next', line }; }
      const lab = this.parseLabelRef('Resume');
      this.expectEOS();
      if (lab.lower === '0') return { type: 'Resume', mode: 'retry', line };
      return { type: 'Resume', mode: 'label', label: lab.lower, labelName: lab.name, line };
    }

    /** Me("e_zahl1") → wie Me.e_zahl1 (nur mit festem Text) */
    meIndexName() {
      if (this.isOp('(') && this.peek(1).type === 'str' && this.isOp(')', 2)) {
        this.next();
        const s = this.next();
        this.next();
        return { value: s.value, lower: s.value.toLowerCase(), line: s.line };
      }
      return null;
    }

    /** Me("x") bzw. Me("x").Value */
    parseMeIndexed(meTok, name) {
      let node = { type: 'Ident', name: name.value, lower: name.lower, line: meTok.line, viaMe: true };
      if (this.isOp('.')) {
        this.next();
        const p = this.expectIdent('Eigenschaft');
        node = { type: 'Member', object: name.value, objLower: name.lower, prop: p.value, propLower: p.lower, line: meTok.line };
      }
      return node;
    }

    parseVarList(isStatic) {
      const decls = [];
      do {
        const nameTok = this.expectIdent('Variablenname');
        const d = { type: 'VarDecl', name: nameTok.value, lower: nameTok.lower, varType: 'variant', isArray: false, bounds: null, isStatic, line: nameTok.line };
        if (this.acceptOp('(')) {
          d.isArray = true;
          d.bounds = this.isOp(')') ? [] : this.parseBounds();
          this.expectOp(')');
        }
        if (this.acceptKw('as')) {
          if (this.isKw('new')) throw syntaxError("'As New' (Objekte) wird nicht unterstützt.", this.line());
          d.varType = this.parseType();
        } else if (nameTok.typeChar) d.varType = nameTok.typeChar;
        this.declare(d.name, d.lower, d.line);
        decls.push(d);
      } while (this.acceptOp(','));
      return decls;
    }

    parseBounds() {
      const bounds = [];
      do {
        const first = this.parseExpr();
        if (this.acceptKw('to')) bounds.push({ lower: first, upper: this.parseExpr() });
        else bounds.push({ lower: null, upper: first });
      } while (this.acceptOp(','));
      return bounds;
    }

    parseDimStatement(isStatic, line) {
      const decls = this.parseVarList(isStatic);
      this.expectEOS();
      if (this.proc) this.proc.decls.push(...decls);
      return { type: 'Dim', decls, line };
    }

    parseConstList() {
      const decls = [];
      do {
        const nameTok = this.expectIdent('Konstantenname');
        const d = { type: 'ConstDecl', name: nameTok.value, lower: nameTok.lower, varType: 'variant', expr: null, line: nameTok.line };
        if (this.acceptKw('as')) d.varType = this.parseType();
        else if (nameTok.typeChar) d.varType = nameTok.typeChar;
        this.expectOp('=');
        d.expr = this.parseExpr();
        this.declare(d.name, d.lower, d.line);
        decls.push(d);
      } while (this.acceptOp(','));
      return decls;
    }

    parseConstStatement(line) {
      const decls = this.parseConstList();
      this.expectEOS();
      if (this.proc) this.proc.decls.push(...decls);
      return { type: 'Dim', decls, line };
    }

    parseReDim(line) {
      const preserve = !!this.acceptKw('preserve');
      const items = [];
      do {
        const nameTok = this.expectIdent('Arrayname');
        this.expectOp('(');
        const bounds = this.parseBounds();
        this.expectOp(')');
        let varType = null;
        if (this.acceptKw('as')) varType = this.parseType();
        items.push({ name: nameTok.value, lower: nameTok.lower, bounds, varType, line: nameTok.line });
      } while (this.acceptOp(','));
      this.expectEOS();
      return { type: 'ReDim', preserve, items, line };
    }

    parseIf() {
      const ifTok = this.next();
      const line = ifTok.line;
      const cond = this.parseExpr();
      this.expectKw('then', "'Then'");
      const after = this.peek();
      if (after.type !== 'nl' && after.type !== 'eof') {
        // einzeiliges If
        if (after.type === 'colon') this.next();
        this.inlineIf++;
        let thenBody;
        try { thenBody = this.parseInlineStatements(true); } finally { this.inlineIf--; }
        let elseBody = null;
        if (this.acceptKw('else')) elseBody = this.parseInlineStatements(false);
        if (this.isKw('elseif')) throw syntaxError('ElseIf ist im einzeiligen If nicht erlaubt – verwende die Blockform mit End If.', this.line());
        return { type: 'If', branches: [{ cond, body: thenBody, line }], elseBody, line, singleLine: true };
      }
      const branches = [];
      let elseBody = null;
      let current = { cond, body: null, line };
      return this.withBlock('if', line, () => {
        for (;;) {
          const { body, term } = this.parseBlock('if', line);
          current.body = body;
          branches.push(current);
          if (term === 'endif') { this.next(); this.next(); this.expectEOS(); break; }
          if (term === 'elseif') {
            const l = this.next().line;
            const c = this.parseExpr();
            this.expectKw('then', "'Then'");
            this.expectEOS();
            current = { cond: c, body: null, line: l };
            continue;
          }
          // Else
          this.next();
          if (this.isKw('if') && !this.atEOS(1)) {
            // "Else If" auf derselben Zeile → verschachteltes If im Else-Zweig
          } else {
            this.expectEOS();
          }
          this.blockStack[this.blockStack.length - 1].kind = 'else';
          const res = this.parseBlock('else', line);
          elseBody = res.body;
          this.next(); this.next(); this.expectEOS();
          break;
        }
        return { type: 'If', branches, elseBody, line };
      });
    }

    parseInlineStatements(stopAtElse) {
      const list = [];
      for (;;) {
        const t = this.peek();
        if (t.type === 'nl' || t.type === 'eof') break;
        if (t.type === 'colon') { this.next(); continue; }
        if (stopAtElse && this.isKw('else')) break;
        list.push(this.parseStatement());
      }
      return list;
    }

    parseSelect() {
      const line = this.next().line;
      this.expectKw('case', "'Case' nach 'Select'");
      const expr = this.parseExpr();
      this.expectEOS();
      const cases = [];
      let elseBody = null;
      return this.withBlock('select', line, () => {
        this.skipSeparators();
        let term = this.getTerminator();
        if (term !== 'case' && term !== 'endselect') {
          if (this.peek().type === 'eof') throw syntaxError(MISSING_MSG.select, line);
          throw syntaxError("Nach 'Select Case' muss zuerst ein 'Case' folgen.", this.line());
        }
        for (;;) {
          if (term === 'endselect') { this.next(); this.next(); this.expectEOS(); break; }
          const caseLine = this.next().line; // Case
          if (elseBody) throw syntaxError("Nach 'Case Else' ist kein weiteres 'Case' erlaubt.", caseLine);
          if (this.acceptKw('else')) {
            const res = this.parseBlock('select', line);
            elseBody = res.body;
            term = res.term;
            continue;
          }
          const tests = [];
          do {
            if (this.acceptKw('is')) {
              const opTok = this.next();
              if (opTok.type !== 'op' || !['=', '<>', '<', '>', '<=', '>='].includes(opTok.value)) {
                throw syntaxError("Nach 'Case Is' wird ein Vergleichsoperator erwartet (z. B. Case Is > 10).", caseLine);
              }
              tests.push({ kind: 'is', op: opTok.value, expr: this.parseExpr() });
            } else {
              const a = this.parseExpr();
              if (this.acceptKw('to')) tests.push({ kind: 'range', from: a, to: this.parseExpr() });
              else tests.push({ kind: 'value', expr: a });
            }
          } while (this.acceptOp(','));
          const res = this.parseBlock('select', line);
          cases.push({ tests, body: res.body, line: caseLine });
          term = res.term;
        }
        return { type: 'Select', expr, cases, elseBody, line };
      });
    }

    parseFor() {
      const line = this.next().line;
      if (this.acceptKw('each')) {
        const v = this.expectIdent('Schleifenvariable');
        this.expectKw('in', "'In'");
        const collection = this.parseExpr();
        this.expectEOS();
        return this.withBlock('foreach', line, () => {
          const { body } = this.parseBlock('foreach', line);
          this.parseNext(v, line);
          return { type: 'ForEach', varName: v.value, lower: v.lower, collection, body, line };
        });
      }
      const v = this.expectIdent('Zählvariable');
      this.expectOp('=');
      const start = this.parseExpr();
      this.expectKw('to', "'To'");
      const end = this.parseExpr();
      let step = null;
      if (this.acceptKw('step')) step = this.parseExpr();
      this.expectEOS();
      return this.withBlock('for', line, () => {
        const { body } = this.parseBlock('for', line);
        this.parseNext(v, line);
        return { type: 'For', varName: v.value, lower: v.lower, start, end, step, body, line };
      });
    }

    parseNext(v, forLine) {
      const nextTok = this.next();
      if (this.peek().type === 'ident') {
        const n = this.next();
        if (n.lower !== v.lower) {
          throw syntaxError(`Next ${n.value} passt nicht zur For-Schleife mit der Variablen '${v.value}' (Zeile ${forLine}).`, nextTok.line);
        }
        if (this.isOp(',')) throw syntaxError('Bitte für jede For-Schleife ein eigenes Next schreiben.', nextTok.line);
      }
      this.expectEOS();
    }

    parseDo() {
      const line = this.next().line;
      let preCond = null;
      if (this.isKw('while') || this.isKw('until')) {
        const kind = this.next().value;
        preCond = { kind, expr: this.parseExpr() };
      }
      this.expectEOS();
      return this.withBlock('do', line, () => {
        const { body } = this.parseBlock('do', line);
        const loopLine = this.next().line; // Loop
        let postCond = null;
        if (this.isKw('while') || this.isKw('until')) {
          if (preCond) throw syntaxError("Die Bedingung darf nur bei 'Do' oder bei 'Loop' stehen, nicht bei beiden.", loopLine);
          const kind = this.next().value;
          postCond = { kind, expr: this.parseExpr() };
        }
        this.expectEOS();
        return { type: 'Do', preCond, postCond, body, line };
      });
    }

    parseWhile() {
      const line = this.next().line;
      const cond = this.parseExpr();
      this.expectEOS();
      return this.withBlock('while', line, () => {
        const { body } = this.parseBlock('while', line);
        this.next(); // Wend
        this.expectEOS();
        return { type: 'While', cond, body, line };
      });
    }

    parseExit() {
      const line = this.next().line;
      const t = this.next();
      const what = t.type === 'kw' ? t.value : '';
      switch (what) {
        case 'for':
          if (!this.inLoop('for')) throw syntaxError('Exit For ist nur innerhalb einer For-Schleife erlaubt.', line);
          break;
        case 'do':
          if (!this.inLoop('do')) throw syntaxError('Exit Do ist nur innerhalb einer Do-Schleife erlaubt.', line);
          break;
        case 'sub':
          if (!this.proc || this.proc.kind !== 'sub') throw syntaxError('Exit Sub ist nur in einer Sub erlaubt – in einer Function heißt es Exit Function.', line);
          break;
        case 'function':
          if (!this.proc || this.proc.kind !== 'function') throw syntaxError('Exit Function ist nur in einer Function erlaubt – in einer Sub heißt es Exit Sub.', line);
          break;
        default:
          throw syntaxError("Nach 'Exit' wird For, Do, Sub oder Function erwartet.", line);
      }
      this.expectEOS();
      return { type: 'Exit', what, line };
    }

    parseCallStatement() {
      const line = this.next().line;
      const nameTok = this.expectIdent('Name der Prozedur');
      if (this.isOp('.')) {
        const target = this.parseMemberChain(nameTok);
        this.expectEOS();
        if (target.type === 'Member') { target.args = target.args || []; return { type: 'MemberCall', target, line }; }
        return { type: 'Call', name: target.name, lower: target.lower, args: [], line };
      }
      let args = [];
      if (this.acceptOp('(')) {
        args = this.parseArgList(')');
        this.expectOp(')');
      }
      this.expectEOS();
      return { type: 'Call', name: nameTok.value, lower: nameTok.lower, args, line };
    }

    parseDebug() {
      const line = this.next().line;
      this.next(); // .
      const m = this.expectIdent("'Print'");
      if (m.lower !== 'print') {
        if (m.lower === 'assert') {
          const expr = this.parseExpr();
          this.expectEOS();
          return { type: 'Assert', expr, line };
        }
        throw syntaxError(`Debug.${m.value} wird nicht unterstützt – verwende Debug.Print.`, line);
      }
      const items = [];
      while (!this.atEOS()) {
        if (this.isOp(';') || this.isOp(',')) { items.push({ sep: this.next().value }); continue; }
        items.push({ expr: this.parseExpr() });
      }
      return { type: 'Print', items, line };
    }

    /** Argumentliste; leere Argumente (z. B. MsgBox "x", , "Titel") werden als null abgelegt. */
    parseArgList(closer) {
      const args = [];
      if (closer && this.isOp(closer)) return args;
      if (!closer && this.atEOS()) return args;
      for (;;) {
        if (this.isOp(',')) { args.push(null); this.next(); continue; }
        if ((closer && this.isOp(closer)) || (!closer && this.atEOS())) { args.push(null); break; }
        if (this.peek().type === 'ident' && this.isOp(':=', 1)) {
          throw syntaxError('Benannte Argumente (name:=wert) werden nicht unterstützt. Bitte die Argumente der Reihe nach angeben.', this.line());
        }
        args.push(this.parseExpr());
        if (!this.acceptOp(',')) break;
      }
      return args;
    }

    parseAssignOrCall(isLetSet) {
      const nameTok = this.peek();
      const line = nameTok.line;
      if (nameTok.type !== 'ident') {
        throw syntaxError(`Unerwartet: ${this.describe(nameTok)}. Hier wird eine Anweisung erwartet.`, line);
      }
      this.next();
      if (nameTok.lower === 'me' && this.isOp('(')) {
        const nm = this.meIndexName();
        if (nm) {
          const target = this.parseMeIndexed(nameTok, nm);
          if (this.acceptOp('=')) {
            const value = this.parseExpr();
            this.expectEOS();
            return { type: 'Assign', target, value, line };
          }
          if (target.type === 'Member' && this.atEOS()) return { type: 'MemberCall', target, line };
          throw syntaxError("Erwartet: '='.", line);
        }
      }
      if (this.isOp('.') || (this.isOp('!') && nameTok.lower === 'me')) {
        // Steuerelement-Eigenschaft oder -Methode: E_Zahl.Visible = False, E_Zahl.SetFocus, Me.E_Zahl = 5
        const startPos = this.pos;
        let target = this.parseMemberChain(nameTok);
        if (target.args && !this.atEOS() && !this.isOp('=')) {
          // lstAusgabe.AddItem (i) & " Euro" – die Klammer gehört zum Argument
          this.pos = startPos;
          target = this.parseMemberChain(nameTok, true);
        }
        if (this.acceptOp('=')) {
          const value = this.parseExpr();
          this.expectEOS();
          return { type: 'Assign', target, value, line };
        }
        if (target.type === 'Member' && this.atEOS()) return { type: 'MemberCall', target, line };
        if (target.type === 'Member' && !target.args) {
          // Methode mit Argumenten ohne Klammern: lstAusgabe.AddItem "Text", 0
          target.args = this.parseArgList(null);
          this.expectEOS();
          return { type: 'MemberCall', target, line };
        }
        if (target.type === 'Ident' && this.atEOS()) return { type: 'Call', name: target.name, lower: target.lower, args: [], line, viaMe: target.viaMe };
        throw syntaxError("Erwartet: '='.", line);
      }

      // Zuweisung an einfache Variable
      if (this.acceptOp('=')) {
        const value = this.parseExpr();
        this.expectEOS();
        return { type: 'Assign', target: { type: 'Ident', name: nameTok.value, lower: nameTok.lower, line }, value, line };
      }
      if (isLetSet) throw syntaxError(`Erwartet: '=' nach '${nameTok.value}'.`, line);

      if (this.isOp('(')) {
        const save = this.pos;
        let args = null;
        try {
          this.next();
          args = this.parseArgList(')');
          this.expectOp(')');
        } catch (e) {
          args = null;
        }
        if (args) {
          if (this.isOp('.')) throw syntaxError(`Methode oder Datenobjekt nicht gefunden: '${nameTok.value}(…).' wird nicht unterstützt.`, line);
          if (this.acceptOp('=')) {
            const value = this.parseExpr();
            this.expectEOS();
            return { type: 'Assign', target: { type: 'Index', name: nameTok.value, lower: nameTok.lower, args, line }, value, line };
          }
          if (this.atEOS()) {
            if (args.length > 1) {
              throw syntaxError(`Beim Aufruf ohne 'Call' stehen die Argumente ohne Klammern: ${nameTok.value} a, b – oder mit Call: Call ${nameTok.value}(a, b).`, line);
            }
            // Foo (x) – die Klammer gehört zum Ausdruck (Übergabe als Wert)
            const a = args.length === 1 && args[0] ? [{ type: 'Paren', expr: args[0], line }] : [];
            return { type: 'Call', name: nameTok.value, lower: nameTok.lower, args: a, line };
          }
        }
        this.pos = save;
      }
      if (this.atEOS()) return { type: 'Call', name: nameTok.value, lower: nameTok.lower, args: [], line };
      const args = this.parseArgList(null);
      this.expectEOS();
      return { type: 'Call', name: nameTok.value, lower: nameTok.lower, args, line };
    }

    /**
     * Liest "obj.eigenschaft" (das aktuelle Token ist der Punkt).
     * Me.X wird zu X (mit viaMe), Me.X.Y zu Member(X, Y).
     * Ob es das Objekt/die Eigenschaft gibt, prüft erst der Analyzer.
     */
    parseMemberChain(objTok, noParens) {
      const line = objTok.line;
      const sep = this.next(); // . oder !
      if (sep.value === '!' && objTok.lower !== 'me') throw syntaxError(`Unerwartet: '!' nach '${objTok.value}'. Verwende den Punkt, z. B. ${objTok.value}.Value.`, line);
      const m = this.peek();
      if (m.type !== 'ident' && m.type !== 'kw') throw syntaxError('Erwartet: Bezeichner nach dem Punkt.', line);
      this.next();
      let mName = m.type === 'kw' ? (m.raw || m.value) : m.value;
      let mLower = mName.toLowerCase();
      let node;
      if (objTok.lower === 'me' && sep.value === '.' && mLower === 'controls') {
        // Me.Controls("e_zahl1")
        const nm = this.meIndexName();
        if (!nm) throw syntaxError('Erwartet: Me.Controls("Name") mit dem Namen in Anführungszeichen.', line);
        mName = nm.value; mLower = nm.lower;
      }
      if (objTok.lower === 'me' && sep.value === '.' && FORM_MEMBERS.has(mLower) && !this.isOp('.') && !this.isOp('!')) {
        // Eigenschaften und Methoden des Formulars: Me.Caption, Me.Refresh
        node = { type: 'Member', object: 'Me', objLower: 'me', prop: mName, propLower: mLower, line };
      } else if (objTok.lower === 'me') {
        node = { type: 'Ident', name: mName, lower: mLower, line, viaMe: true };
        if (this.isOp('.')) {
          this.next();
          const p = this.expectIdent('Eigenschaft');
          node = { type: 'Member', object: mName, objLower: mLower, prop: p.value, propLower: p.lower, line };
        }
      } else {
        node = { type: 'Member', object: objTok.value, objLower: objTok.lower, prop: mName, propLower: mLower, line };
      }
      if (this.isOp('.')) throw syntaxError(`Methode oder Datenobjekt nicht gefunden: '${objTok.value}.${mName}.…' wird nicht unterstützt.`, line);
      if (this.isOp('(') && node.type === 'Member' && !noParens) {
        // lstAusgabe.ItemData(0), lstAusgabe.Column(0, 1), lstAusgabe.AddItem ("Text")
        this.next();
        node.args = this.parseArgList(')');
        this.expectOp(')');
        if (this.isOp('.')) throw syntaxError(`Methode oder Datenobjekt nicht gefunden: '${node.object}.${node.prop}(…).…' wird nicht unterstützt.`, line);
      }
      return node;
    }

    // ---------- Ausdrücke (Rangfolge wie in VBA) ----------
    parseExpr() { return this.parseImp(); }

    binLeft(next, ops) {
      let left = next();
      for (;;) {
        const t = this.peek();
        const v = t.type === 'kw' ? t.value : (t.type === 'op' ? t.value : null);
        if (v === null || !ops.includes(v)) return left;
        this.next();
        const right = next();
        left = { type: 'Bin', op: v, left, right, line: t.line };
      }
    }
    parseImp() { return this.binLeft(() => this.parseEqv(), ['imp']); }
    parseEqv() { return this.binLeft(() => this.parseXor(), ['eqv']); }
    parseXor() { return this.binLeft(() => this.parseOr(), ['xor']); }
    parseOr() { return this.binLeft(() => this.parseAnd(), ['or']); }
    parseAnd() { return this.binLeft(() => this.parseNot(), ['and']); }
    parseNot() {
      if (this.isKw('not')) {
        const t = this.next();
        return { type: 'Unary', op: 'not', operand: this.parseNot(), line: t.line };
      }
      return this.parseCompare();
    }
    parseCompare() { return this.binLeft(() => this.parseConcat(), ['=', '<>', '<', '>', '<=', '>=', 'like', 'is']); }
    parseConcat() { return this.binLeft(() => this.parseAdd(), ['&']); }
    parseAdd() { return this.binLeft(() => this.parseMod(), ['+', '-']); }
    parseMod() { return this.binLeft(() => this.parseIntDiv(), ['mod']); }
    parseIntDiv() { return this.binLeft(() => this.parseMul(), ['\\']); }
    parseMul() { return this.binLeft(() => this.parseNeg(), ['*', '/']); }
    parseNeg() {
      if (this.isOp('-') || this.isOp('+')) {
        const t = this.next();
        const operand = this.parseNeg();
        return t.value === '-' ? { type: 'Unary', op: '-', operand, line: t.line } : operand;
      }
      return this.parsePow();
    }
    parsePow() {
      let left = this.parsePostfix();
      while (this.isOp('^')) {
        const t = this.next();
        let right;
        if (this.isOp('-') || this.isOp('+')) {
          const s = this.next();
          const operand = this.parsePostfix();
          right = s.value === '-' ? { type: 'Unary', op: '-', operand, line: s.line } : operand;
        } else right = this.parsePostfix();
        left = { type: 'Bin', op: '^', left, right, line: t.line };
      }
      return left;
    }

    parsePostfix() {
      let expr = this.parsePrimary();
      while (this.isOp('(') && (expr.type === 'CallOrIndex' || expr.type === 'IndexExpr')) {
        const line = this.next().line;
        const args = this.parseArgList(')');
        this.expectOp(')');
        expr = { type: 'IndexExpr', target: expr, args, line };
      }
      return expr;
    }

    parsePrimary() {
      const t = this.peek();
      const line = t.line;
      switch (t.type) {
        case 'num': this.next(); return { type: 'Num', value: t.value, numType: t.numType || null, line };
        case 'str': this.next(); return { type: 'Str', value: t.value, line };
        case 'date': this.next(); return { type: 'DateLit', raw: t.value, line };
        case 'op':
          if (t.value === '(') {
            this.next();
            const e = this.parseExpr();
            this.expectOp(')');
            return { type: 'Paren', expr: e, line };
          }
          break;
        case 'kw':
          if (t.value === 'true' || t.value === 'false') { this.next(); return { type: 'Bool', value: t.value === 'true', line }; }
          if (t.value === 'nothing') { this.next(); return { type: 'Nothing', line }; }
          if (t.value === 'new') throw syntaxError("'New' (Objekte erzeugen) wird nicht unterstützt.", line);
          throw syntaxError(`'${t.raw || t.value}' ist ein Schlüsselwort und kann hier nicht stehen.`, line);
        case 'ident': {
          this.next();
          if (t.lower === 'me' && this.isOp('(')) {
            const nm = this.meIndexName();
            if (nm) return this.parseMeIndexed(t, nm);
          }
          if (this.isOp('.') || (this.isOp('!') && t.lower === 'me')) return this.parseMemberChain(t);
          if (this.isOp('(')) {
            this.next();
            const args = this.parseArgList(')');
            this.expectOp(')');
            return { type: 'CallOrIndex', name: t.value, lower: t.lower, args, line };
          }
          return { type: 'Ident', name: t.value, lower: t.lower, line };
        }
        default: break;
      }
      if (t.type === 'nl' || t.type === 'eof' || t.type === 'colon') {
        throw syntaxError('Ausdruck erwartet – die Anweisung ist unvollständig.', line);
      }
      throw syntaxError(`Ausdruck erwartet, gefunden: ${this.describe(t)}.`, line);
    }
  }

  function parse(source) {
    const tokens = tokenize(source);
    const p = new Parser(tokens);
    return p.parseProgram();
  }

  return { parse, Parser, TYPE_NAMES, TYPE_DISPLAY, FORM_MEMBERS };
});
