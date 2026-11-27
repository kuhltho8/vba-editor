/*
 * interpreter.js – führt den AST asynchron aus.
 *
 * Ein-/Ausgabe läuft über ein Host-Objekt (Browser-UI oder Testumgebung):
 *   host.print(zeile)                         – eine Zeile in die Konsole
 *   host.msgBox(text, buttons, titel)         – Promise<number> (vbOK, vbYes, ...)
 *   host.inputBox(text, titel, vorgabe)       – Promise<string|null> (null = Abbrechen)
 *
 * Schutz vor Endlosschleifen: Stopp über stop(), Schrittlimit (maxSteps).
 * Der Interpreter gibt regelmäßig die Kontrolle an den Browser ab, damit die
 * Oberfläche (z. B. der Stopp-Button) bedienbar bleibt.
 */
(function (root, factory) {
  const isNode = typeof module === 'object' && module.exports;
  const deps = isNode
    ? Object.assign({}, require('./errors.js'), require('./values.js'), require('./builtins.js'))
    : root.VBA;
  const mod = factory(deps);
  if (isNode) module.exports = mod;
  else { root.VBA = root.VBA || {}; Object.assign(root.VBA, mod); }
})(typeof self !== 'undefined' ? self : this, function (V) {
  'use strict';
  const {
    VbaError, runtimeError, EMPTY, NULL, NOTHING, MISSING, VbaDate, VbaArray, TYPE_DISPLAY,
    defaultValue, bankersRound, toNumber, toStr, toBool, coerce, compare, printFormat,
    BUILTINS, CONSTANTS, SPECIAL_VALUES, resetRandom, numberToString, RUNTIME_TEXTS, parseNumberString
  } = V;

  // Steuersignale für Exit-Anweisungen
  const EXIT_FOR = { signal: 'for' };
  const EXIT_DO = { signal: 'do' };
  const EXIT_PROC = { signal: 'proc' };
  class EndSignal { }
  // Sprung zu einer Sprungmarke (GoTo, Resume Marke) und Verlassen einer Prozedur aus einer Fehlerbehandlung
  class GotoSignal { constructor(label, frame) { this.label = label; this.frame = frame; } }
  class ProcExitSignal { constructor(frame) { this.frame = frame; } }
  class ResumeSignal { constructor(mode, label, frame) { this.mode = mode; this.label = label; this.frame = frame; } }

  const OVERFLOW_RANGE = { byte: [0, 255], integer: [-32768, 32767], long: [-2147483648, 2147483647] };
  const TYPE_DE = { byte: 'Byte', integer: 'Integer', long: 'Long' };

  class Cell {
    constructor(type, value) {
      this.type = type;        // Datentyp (bei Arrays: Typ der Elemente)
      this.value = value;
      this.isConst = false;
      this.isArray = false;    // als Array deklariert
      this.fixed = false;      // Array mit fester Größe (Dim a(1 To 10))
    }
  }

  /** Steuerelement als Variable: Lesen/Schreiben geht an das Formular (Value). */
  class ControlCell {
    constructor(adapter) {
      this.adapter = adapter;
      this.type = 'variant';
      this.isConst = false;
      this.isArray = false;
      this.fixed = false;
      this.isControl = true;
    }
    get value() { return this.adapter.getValue(null); }
    set value(v) { this.adapter.setValue(v, null); }
  }

  // Möglichst schnelles "Kontrolle abgeben" je nach Umgebung
  const schedule = (function () {
    if (typeof setImmediate === 'function') return (fn) => setImmediate(fn);
    if (typeof MessageChannel === 'function') {
      const ch = new MessageChannel();
      const queue = [];
      ch.port1.onmessage = () => { const f = queue.shift(); if (f) f(); };
      return (fn) => { queue.push(fn); ch.port2.postMessage(0); };
    }
    return (fn) => setTimeout(fn, 0);
  })();

  function likeToRegex(pattern, ignoreCase) {
    let re = '';
    for (let i = 0; i < pattern.length; i++) {
      const c = pattern[i];
      if (c === '?') re += '.';
      else if (c === '*') re += '[\\s\\S]*';
      else if (c === '#') re += '\\d';
      else if (c === '[') {
        const end = pattern.indexOf(']', i + 1);
        if (end === -1) { re += '\\['; continue; }
        let cls = pattern.slice(i + 1, end);
        if (cls[0] === '!') cls = '^' + cls.slice(1);
        re += '[' + cls.replace(/\\/g, '\\\\') + ']';
        i = end;
      } else re += c.replace(/[.*+?^${}()|[\]\\\/]/g, '\\$&');
    }
    return new RegExp('^' + re + '$', ignoreCase ? 'i' : '');
  }

  class Interpreter {
    constructor(program, host, options) {
      this.program = program;
      this.host = host || {};
      const o = options || {};
      this.maxSteps = o.maxSteps || 10000000;
      this.yieldMs = o.yieldMs || 25;
      this.maxDepth = o.maxDepth || 800;
      this.globals = new Map();
      this.frames = [];
      this.frame = null;
      this.steps = 0;
      this.stopped = false;
      this.running = false;
      this.currentLine = 0;
      this.printBuffer = '';
      this.lastYield = Date.now();
      this.controls = new Map();
      for (const c of o.controls || []) this.controls.set(c.lower, c);
      this.formAdapter = o.form || null; // Me.Caption, Me.Refresh
      this.err = { number: 0, description: '', source: '' };
      this.opened = false;
    }

    /** Liste der startbaren Subs (ohne Pflichtparameter) – z. B. für ein Dropdown. */
    static startableSubs(program) {
      return program.procList.filter((p) => p.kind === 'sub' && p.params.every((x) => x.optional || x.paramArray));
    }

    stop() { this.stopped = true; }

    // ---------- Start ----------
    /** Startet eine Sub (Programm im Konsolenmodus). Modulvariablen werden neu angelegt. */
    async run(entryName) {
      let proc;
      try {
        proc = this.chooseEntry(entryName);
      } catch (e) { throw this.normalizeError(e); }
      this.open();
      return this.invoke(proc.name);
    }

    /**
     * Öffnet eine Sitzung (z. B. Formularansicht): Modulvariablen anlegen.
     * Danach können mit invoke() beliebig viele Prozeduren aufgerufen werden,
     * die Modulvariablen bleiben dazwischen erhalten – wie in Access.
     */
    open() {
      resetRandom();
      this.steps = 0;
      try {
        this.initGlobals();
      } catch (e) { throw this.normalizeError(e); }
      this.opened = true;
    }

    hasProc(name) { return !!this.program.procs[String(name).toLowerCase()]; }

    /** Ruft eine parameterlose Prozedur in der offenen Sitzung auf (z. B. btnOk_Click). */
    async invoke(name) {
      if (!this.opened) this.open();
      const proc = this.program.procs[String(name).toLowerCase()];
      if (!proc) throw new VbaError(`Die Prozedur '${name}' gibt es nicht.`, null, 'runtime');
      this.running = true;
      this.stopped = false;
      this.steps = 0;
      this.lastYield = Date.now();
      try {
        await this.callUserProc(proc, [], proc.line);
      } catch (e) {
        if (e instanceof EndSignal) {
          // End setzt alle Modulvariablen zurück
          try { this.initGlobals(); } catch (e2) { /* ignorieren */ }
        } else throw this.explainControls(this.normalizeError(e));
      } finally {
        this.flushPrint();
        this.running = false;
        this.frames = [];
        this.frame = null;
      }
      return { steps: this.steps };
    }

    chooseEntry(entryName) {
      const p = this.program;
      if (entryName) {
        const proc = p.procs[String(entryName).toLowerCase()];
        if (!proc) throw new VbaError(`Die Prozedur '${entryName}' gibt es nicht.`, null, 'runtime');
        if (proc.kind !== 'sub' || !proc.params.every((x) => x.optional || x.paramArray)) {
          throw new VbaError(`'${proc.name}' kann nicht direkt gestartet werden – nur eine Sub ohne Parameter ist startbar.`, proc.line, 'runtime');
        }
        return proc;
      }
      const subs = Interpreter.startableSubs(p);
      if (subs.length === 0) {
        throw new VbaError('Keine startbare Sub gefunden. Lege eine Sub ohne Parameter an, z. B.:  Sub Main() ... End Sub', null, 'runtime');
      }
      return subs.find((s) => s.lower === 'main') || subs[0];
    }

    /**
     * Fehler 94/13 mit einem Steuerelement in der Fehlerzeile: Hinweis mit dem Feldnamen,
     * z. B. „Das Feld e_zahl1 ist leer (Null).“
     */
    explainControls(e) {
      if (!(e instanceof VbaError) || e.kind !== 'runtime' || (e.number !== 94 && e.number !== 13) || !e.line || !this.controls.size) return e;
      const names = new Set();
      const seen = new Set();
      const walk = (n) => {
        if (!n || typeof n !== 'object' || seen.has(n)) return;
        seen.add(n);
        if (Array.isArray(n)) { n.forEach(walk); return; }
        if (n.type === 'Ident' && n.line === e.line && this.controls.has(n.lower)) names.add(n.lower);
        for (const k of Object.keys(n)) if (k !== 'labels' && n[k] && typeof n[k] === 'object') walk(n[k]);
      };
      for (const p of this.program.procList) for (const st of p.body) walk(st);
      const tips = [];
      for (const l of names) {
        const c = this.controls.get(l);
        let v;
        try { v = c.getValue(null); } catch (x) { continue; }
        if (e.number === 94 && v === NULL) tips.push(`Das Feld ${c.name} ist leer (Null). Prüfe vorher mit IsNull(${c.name}) oder nutze Nz(${c.name}, 0).`);
        if (e.number === 13 && typeof v === 'string' && isNaN(parseNumberString(v))) tips.push(`Das Feld ${c.name} enthält „${v}“ – das ist keine Zahl. Prüfe vorher mit IsNumeric(${c.name}).`);
      }
      if (tips.length) e.hint = tips.join(' ');
      return e;
    }

    normalizeError(e) {
      if (e instanceof VbaError) {
        if (!e.line && this.currentLine) {
          return new VbaError(e.plainMessage, this.currentLine, e.kind, e.number);
        }
        return e;
      }
      if (e instanceof RangeError && /call stack/i.test(e.message)) {
        return new VbaError('Nicht genügend Stapelspeicher – zu viele verschachtelte Aufrufe (Rekursion ohne Abbruchbedingung?).', this.currentLine, 'runtime', 28);
      }
      return new VbaError('Interner Fehler des Interpreters: ' + (e && e.message ? e.message : String(e)), this.currentLine, 'runtime');
    }

    // ---------- Variablen ----------
    initGlobals() {
      this.globals = new Map();
      this.frame = { proc: null, locals: this.globals };
      for (const [lower, c] of this.controls) this.globals.set(lower, new ControlCell(c));
      for (const d of this.program.moduleDecls) this.declare(d, this.globals);
      this.frame = null;
    }

    declare(d, map) {
      this.currentLine = d.line;
      if (d.type === 'ConstDecl') {
        const v = this.evalSyncOrFail(d.expr);
        const cell = new Cell(d.varType, coerce(v, d.varType, d.line));
        cell.isConst = true;
        map.set(d.lower, cell);
        return;
      }
      if (d.isStatic && this.frame && this.frame.proc) {
        const store = this.frame.proc.staticCells || (this.frame.proc.staticCells = new Map());
        if (!store.has(d.lower)) store.set(d.lower, this.makeVarCell(d));
        map.set(d.lower, store.get(d.lower));
        return;
      }
      map.set(d.lower, this.makeVarCell(d));
    }

    makeVarCell(d) {
      if (!d.isArray) return new Cell(d.varType, defaultValue(d.varType));
      const cell = new Cell(d.varType, null);
      cell.isArray = true;
      if (d.bounds.length === 0) {
        cell.value = new VbaArray(null, d.varType);
      } else {
        cell.fixed = true;
        cell.value = new VbaArray(this.evalBounds(d.bounds, d.line), d.varType);
      }
      return cell;
    }

    evalBounds(bounds, line) {
      return bounds.map((b) => {
        const lb = b.lower ? coerce(this.evalSyncOrFail(b.lower), 'long', line) : this.program.optionBase;
        const ub = coerce(this.evalSyncOrFail(b.upper), 'long', line);
        if (ub < lb) throw runtimeError(`Index außerhalb des gültigen Bereichs: Die Obergrenze (${ub}) ist kleiner als die Untergrenze (${lb}).`, line, 9);
        return { lb, ub };
      });
    }

    evalSyncOrFail(node) {
      if (node.async) throw runtimeError('Hier ist nur ein konstanter Ausdruck erlaubt (kein Funktionsaufruf).', node.line);
      return this.evalSync(node);
    }

    findCell(lower) {
      if (this.frame) {
        const c = this.frame.locals.get(lower);
        if (c) return c;
      }
      return this.globals.get(lower) || null;
    }

    implicitCell(lower) {
      const cell = new Cell('variant', EMPTY);
      (this.frame ? this.frame.locals : this.globals).set(lower, cell);
      return cell;
    }

    // ---------- Ablaufsteuerung ----------
    tick() {
      this.steps++;
      if (this.stopped) throw new VbaError('Programm wurde gestoppt.', this.currentLine, 'abort');
      if (this.steps > this.maxSteps) {
        const e = runtimeError(`Das Programm wurde nach ${this.maxSteps.toLocaleString('de-DE')} Schritten abgebrochen – vermutlich eine Endlosschleife. Prüfe die Abbruchbedingung deiner Schleife.`, this.currentLine);
        e.untrappable = true;
        throw e;
      }
      if ((this.steps & 1023) === 0 && Date.now() - this.lastYield > this.yieldMs) return this.yieldNow();
      return null;
    }

    yieldNow() {
      const now = Date.now();
      // etwa alle 100 ms richtig abgeben (Bildaufbau, Klick auf Stoppen), sonst nur kurz
      if (typeof requestAnimationFrame === 'function' && now - (this.lastPaint || 0) > 100) {
        return new Promise((resolve) => {
          let done = false;
          const go = () => { if (done) return; done = true; this.lastPaint = this.lastYield = Date.now(); resolve(); };
          requestAnimationFrame(() => setTimeout(go, 0));
          setTimeout(go, 50); // falls der Tab im Hintergrund ist
        });
      }
      return new Promise((resolve) => schedule(() => { this.lastYield = Date.now(); resolve(); }));
    }

    async execBlock(list) {
      for (let i = 0; i < list.length; i++) {
        const sig = await this.exec(list[i]);
        if (sig) return sig;
      }
      return null;
    }

    /**
     * Führt eine Anweisung aus. Mit On Error … wird ein Laufzeitfehler hier abgefangen:
     * Resume Next macht mit der nächsten Anweisung weiter, On Error GoTo Marke führt die
     * Fehlerbehandlung an Ort und Stelle aus (damit Resume / Resume Next zurückfinden).
     */
    async exec(s) {
      const frame = this.frame;
      if (!frame || !frame.onError || frame.inHandler) return this.execStmt(s);
      for (;;) {
        try {
          return await this.execStmt(s);
        } catch (e) {
          if (!(e instanceof VbaError) || e.kind !== 'runtime' || e.untrappable || !frame.onError || frame.inHandler || this.stopped) throw e;
          this.frame = frame;
          this.setErr(e);
          if (frame.onError.mode === 'next') return null;
          // On Error GoTo Marke
          frame.inHandler = true;
          frame.errorLine = e.line;
          try {
            await this.runFromLabel(frame, frame.onError.label);
            throw new ProcExitSignal(frame); // Ende der Prozedur in der Fehlerbehandlung erreicht
          } catch (r) {
            if (!(r instanceof ResumeSignal) || r.frame !== frame) throw r;
            frame.inHandler = false;
            this.clearErr();
            if (r.mode === 'next') return null;
            if (r.mode === 'label') throw new GotoSignal(r.label, frame);
            // Resume: Anweisung noch einmal ausführen
          }
        }
      }
    }

    /** Anweisungen der Prozedur ab einer Sprungmarke ausführen (für die Fehlerbehandlung) */
    async runFromLabel(frame, label) {
      const body = frame.proc.body;
      for (let i = frame.proc.labels.get(label); i < body.length; i++) {
        const sig = await this.exec(body[i]);
        if (sig === EXIT_PROC) throw new ProcExitSignal(frame);
        if (sig) return;
      }
    }

    /** Rumpf einer Prozedur: Sprünge zu Marken auf oberster Ebene (GoTo, Resume Marke) */
    async runProcBody(frame) {
      const body = frame.proc.body;
      let i = 0;
      for (;;) {
        try {
          for (; i < body.length; i++) {
            const sig = await this.exec(body[i]);
            if (sig) return;
          }
          return;
        } catch (e) {
          if (e instanceof ProcExitSignal && e.frame === frame) return;
          if (e instanceof GotoSignal && e.frame === frame) { this.frame = frame; i = frame.proc.labels.get(e.label); continue; }
          throw e;
        }
      }
    }

    async execStmt(s) {
      this.currentLine = s.line;
      const y = this.tick();
      if (y) await y;
      switch (s.type) {
        case 'Assign': return this.execAssign(s);
        case 'Call': await this.execCall(s); return null;
        case 'MemberCall': {
          const args = [];
          for (const a of s.target.args || []) args.push(a ? await this.ev(a) : MISSING);
          this.currentLine = s.line;
          this.control(s.target).call(s.target.propLower, s.line, args);
          return null;
        }
        case 'Print': return this.execPrint(s);
        case 'If': {
          for (const b of s.branches) {
            this.currentLine = b.line;
            const c = await this.ev(b.cond);
            if (toBool(c, b.line)) return this.execBlock(b.body);
          }
          return s.elseBody ? this.execBlock(s.elseBody) : null;
        }
        case 'Select': return this.execSelect(s);
        case 'For': return this.execFor(s);
        case 'ForEach': return this.execForEach(s);
        case 'Do': return this.execDo(s);
        case 'While': {
          for (;;) {
            this.currentLine = s.line;
            if (!toBool(await this.ev(s.cond), s.line)) return null;
            const sig = await this.execBlock(s.body);
            if (sig) return sig;
            const y2 = this.tick(); if (y2) await y2;
          }
        }
        case 'Exit':
          return s.what === 'for' ? EXIT_FOR : (s.what === 'do' ? EXIT_DO : EXIT_PROC);
        case 'Dim': return null; // Deklarationen werden beim Prozedurstart angelegt
        case 'ReDim': this.execReDim(s); return null;
        case 'End': throw new EndSignal();
        case 'Assert':
          if (!toBool(await this.ev(s.expr), s.line)) throw runtimeError('Debug.Assert: Die Bedingung ist nicht erfüllt.', s.line);
          return null;
        case 'Label': return null;
        case 'GoTo': throw new GotoSignal(s.label, this.frame);
        case 'OnError': {
          const f = this.frame;
          this.clearErr();
          if (s.mode === 'off' || s.mode === 'reset') { f.onError = null; if (s.mode === 'reset') f.inHandler = false; }
          else f.onError = { mode: s.mode, label: s.label };
          return null;
        }
        case 'Resume': {
          const f = this.frame;
          if (!f.inHandler) throw runtimeError('Resume ohne Fehler: Resume ist nur in einer Fehlerbehandlung (nach On Error GoTo) erlaubt.', s.line, 20);
          throw new ResumeSignal(s.mode, s.label, f);
        }
        case 'Stop': {
          // wie ein Haltepunkt: anhalten, Zeile markieren, auf Fortsetzen/Beenden warten
          this.flushPrint();
          if (typeof this.host.pause === 'function') {
            const r = await this.host.pause(s.line);
            this.lastYield = Date.now();
            if (r === 'end') throw new VbaError('Programm wurde beendet.', s.line, 'abort');
          }
          return null;
        }
        default:
          throw runtimeError(`Unbekannte Anweisung (${s.type}).`, s.line);
      }
    }

    async execAssign(s) {
      const value = s.valueAsync ? await this.ev(s.value) : this.evalSync(s.value);
      this.currentLine = s.line;
      const t = s.target;
      if (t.type === 'Member') {
        this.control(t).set(t.propLower, value, s.line);
        return null;
      }
      if (t.type === 'Ident') {
        let cell = this.findCell(t.lower);
        if (!cell) cell = this.implicitCell(t.lower);
        this.assignCell(cell, value, t.name, s.line);
        return null;
      }
      // Array-Element
      const cell = this.findCell(t.lower);
      if (!cell || !(cell.value instanceof VbaArray)) {
        throw runtimeError(`'${t.name}' ist kein Array.`, s.line, 13);
      }
      const arr = cell.value;
      const idx = await this.evalIndices(t.args, s.line);
      arr.data[arr.offset(idx, s.line)] = coerce(value, arr.elemType, s.line);
      return null;
    }

    assignCell(cell, value, name, line) {
      if (cell.isConst) throw runtimeError(`Einer Konstanten ('${name}') kann kein Wert zugewiesen werden.`, line);
      if (cell.isArray) {
        if (!(value instanceof VbaArray)) {
          throw runtimeError(`Typen unverträglich: '${name}' ist ein Array – zugewiesen werden kann nur ein einzelnes Element, z. B. ${name}(1) = ...`, line, 13);
        }
        if (cell.fixed) throw runtimeError(`Zuweisung an das Array '${name}' nicht möglich, weil es eine feste Größe hat.`, line);
        const copy = value.clone();
        if (cell.type !== 'variant') { copy.data = copy.data.map((v) => coerce(v, cell.type, line)); copy.elemType = cell.type; }
        cell.value = copy;
        return;
      }
      cell.value = coerce(value, cell.type, line);
    }

    async evalIndices(args, line) {
      const out = [];
      for (const a of args) {
        if (!a) throw runtimeError('Index fehlt.', line, 9);
        out.push(coerce(await this.ev(a), 'long', line));
      }
      return out;
    }

    async execSelect(s) {
      const v = await this.ev(s.expr);
      const ct = this.program.compareText;
      for (const c of s.cases) {
        this.currentLine = c.line;
        for (const t of c.tests) {
          let hit = false;
          if (t.kind === 'value') hit = compare(v, await this.ev(t.expr), ct, c.line) === 0;
          else if (t.kind === 'range') {
            const from = await this.ev(t.from);
            const to = await this.ev(t.to);
            hit = compare(v, from, ct, c.line) >= 0 && compare(v, to, ct, c.line) <= 0;
          } else {
            const r = compare(v, await this.ev(t.expr), ct, c.line);
            hit = r !== null && { '=': r === 0, '<>': r !== 0, '<': r < 0, '>': r > 0, '<=': r <= 0, '>=': r >= 0 }[t.op];
          }
          if (hit) return this.execBlock(c.body);
        }
      }
      return s.elseBody ? this.execBlock(s.elseBody) : null;
    }

    loopCell(s) {
      let cell = this.findCell(s.lower);
      if (!cell) cell = this.implicitCell(s.lower);
      if (cell.isConst) throw runtimeError(`Die Konstante '${s.varName}' kann nicht als Zählvariable verwendet werden.`, s.line);
      if (cell.isArray) throw runtimeError(`Das Array '${s.varName}' kann nicht als Zählvariable verwendet werden.`, s.line);
      return cell;
    }

    async execFor(s) {
      const cell = this.loopCell(s);
      const start = toNumber(await this.ev(s.start), s.line);
      const end = toNumber(await this.ev(s.end), s.line);
      const step = s.step ? toNumber(await this.ev(s.step), s.line) : 1;
      cell.value = coerce(start, cell.type === 'variant' ? 'double' : cell.type, s.line);
      for (;;) {
        const v = toNumber(cell.value, s.line);
        if (step >= 0 ? v > end : v < end) return null;
        const sig = await this.execBlock(s.body);
        if (sig === EXIT_FOR) return null;
        if (sig) return sig;
        this.currentLine = s.line;
        cell.value = coerce(toNumber(cell.value, s.line) + step, cell.type === 'variant' ? 'double' : cell.type, s.line);
        const y = this.tick(); if (y) await y;
      }
    }

    async execForEach(s) {
      const cell = this.loopCell(s);
      const coll = await this.ev(s.collection);
      if (!(coll instanceof VbaArray)) throw runtimeError('For Each funktioniert nur mit Arrays.', s.line, 13);
      if (!coll.dims) throw runtimeError('Index außerhalb des gültigen Bereichs: Das Array hat noch keine Größe.', s.line, 9);
      const items = coll.data.slice();
      for (const item of items) {
        cell.value = coerce(item, cell.type, s.line);
        const sig = await this.execBlock(s.body);
        if (sig === EXIT_FOR) return null;
        if (sig) return sig;
        this.currentLine = s.line;
        const y = this.tick(); if (y) await y;
      }
      return null;
    }

    async checkCond(cond, line) {
      this.currentLine = line;
      const v = toBool(await this.ev(cond.expr), line);
      return cond.kind === 'while' ? v : !v;
    }

    async execDo(s) {
      for (;;) {
        if (s.preCond && !(await this.checkCond(s.preCond, s.line))) return null;
        const sig = await this.execBlock(s.body);
        if (sig === EXIT_DO) return null;
        if (sig) return sig;
        if (s.postCond && !(await this.checkCond(s.postCond, s.line))) return null;
        const y = this.tick(); if (y) await y;
      }
    }

    execReDim(s) {
      for (const it of s.items) {
        this.currentLine = it.line;
        let cell = this.findCell(it.lower);
        if (!cell) {
          cell = new Cell(it.varType || 'variant', null);
          cell.isArray = true;
          cell.value = new VbaArray(null, cell.type);
          this.frame.locals.set(it.lower, cell);
        }
        if (cell.isConst) throw runtimeError(`Einer Konstanten ('${it.name}') kann kein Wert zugewiesen werden.`, it.line);
        if (cell.fixed) {
          throw runtimeError(`Das Array '${it.name}' hat eine feste Größe und kann nicht mit ReDim geändert werden. Deklariere es ohne Größe: Dim ${it.name}() As ...`, it.line, 10);
        }
        if (!cell.isArray && cell.type !== 'variant') {
          throw runtimeError(`'${it.name}' ist kein Array. Deklariere es z. B. mit Dim ${it.name}() As ${TYPE_DISPLAY[cell.type]}.`, it.line, 13);
        }
        const elemType = cell.isArray ? cell.type : (it.varType || 'variant');
        const dims = this.evalBounds(it.bounds, it.line);
        const fresh = new VbaArray(dims, elemType);
        const old = cell.value instanceof VbaArray ? cell.value : null;
        if (s.preserve && old && old.dims) {
          if (old.dims.length !== dims.length || old.dims.some((d, k) => k < dims.length - 1 && (d.lb !== dims[k].lb || d.ub !== dims[k].ub))) {
            throw runtimeError('Index außerhalb des gültigen Bereichs: Bei ReDim Preserve darf nur die letzte Dimension geändert werden.', it.line, 9);
          }
          // Elemente übernehmen, die in beiden Arrays existieren
          const idx = dims.map((d) => d.lb);
          for (let off = 0; off < fresh.data.length; off++) {
            if (idx.every((v, k) => v >= old.dims[k].lb && v <= old.dims[k].ub)) {
              fresh.data[off] = old.data[old.offset(idx, it.line)];
            }
            for (let k = idx.length - 1; k >= 0; k--) {
              if (++idx[k] <= dims[k].ub) break;
              idx[k] = dims[k].lb;
            }
          }
        }
        cell.value = fresh;
      }
    }

    // ---------- Debug.Print ----------
    execPrint(s) {
      return (async () => {
        let buf = this.printBuffer;
        let endsWithSep = false;
        for (const it of s.items) {
          if (it.sep) {
            endsWithSep = true;
            if (it.sep === ',') {
              const col = buf.length - Math.max(buf.lastIndexOf('\n') + 1, 0);
              buf += ' '.repeat(14 - (col % 14));
            }
            continue;
          }
          endsWithSep = false;
          const e = it.expr;
          if (e.inPrint) {
            // Tab(n): zur Spalte n springen, Tab ohne Zahl: nächste Druckzone; Spc(n): n Leerzeichen
            const col = buf.length - (buf.lastIndexOf('\n') + 1);
            const arg = e.args && e.args[0] ? coerce(await this.ev(e.args[0]), 'long', e.line) : null;
            if (e.lower === 'spc') buf += ' '.repeat(Math.max(0, arg));
            else if (arg === null) buf += ' '.repeat(14 - (col % 14));
            else {
              const target = Math.max(1, arg) - 1;
              if (col > target) buf += '\n' + ' '.repeat(target);
              else buf += ' '.repeat(target - col);
            }
            endsWithSep = true; // nach Tab/Spc geht es in derselben Zeile weiter
            continue;
          }
          const v = e.async ? await this.ev(e) : this.evalSync(e);
          buf += printFormat(v);
        }
        if (endsWithSep) {
          this.printBuffer = buf;
        } else {
          this.printBuffer = '';
          this.emit(buf);
        }
        return null;
      })();
    }

    emit(text) {
      const lines = String(text).split(/\r\n|\r|\n/);
      for (const l of lines) {
        if (typeof this.host.print === 'function') this.host.print(l);
      }
    }

    flushPrint() {
      if (this.printBuffer !== '') {
        const b = this.printBuffer;
        this.printBuffer = '';
        this.emit(b);
      }
    }

    // ---------- MsgBox / InputBox ----------
    async hostMsgBox(prompt, buttons, title) {
      this.flushPrint();
      let r = 1;
      if (typeof this.host.msgBox === 'function') r = await this.host.msgBox(prompt, buttons, title);
      else this.emit(`[MsgBox] ${prompt}`);
      if (this.stopped) throw new VbaError('Programm wurde gestoppt.', this.currentLine, 'abort');
      this.lastYield = Date.now();
      return typeof r === 'number' ? r : 1;
    }

    async hostInputBox(prompt, title, deflt) {
      this.flushPrint();
      let r = deflt;
      if (typeof this.host.inputBox === 'function') r = await this.host.inputBox(prompt, title, deflt);
      if (this.stopped) throw new VbaError('Programm wurde gestoppt.', this.currentLine, 'abort');
      this.lastYield = Date.now();
      return r === null || r === undefined ? '' : String(r);
    }

    // ---------- Aufrufe ----------
    async execCall(s) {
      const proc = this.program.procs[s.lower];
      if (proc) { await this.callUserProc(proc, s.args, s.line); return; }
      const b = BUILTINS[s.lower];
      if (b) { await this.callBuiltin(b, s.args, s.line); return; }
      throw runtimeError(`Sub oder Function '${s.name}' ist nicht definiert.`, s.line, 35);
    }

    async callBuiltin(b, argNodes, line) {
      const args = [];
      for (const a of argNodes) args.push(a ? await this.ev(a) : undefined);
      this.currentLine = line;
      return b.fn(args, { line, interp: this, argNodes });
    }

    callBuiltinSync(b, argNodes, line) {
      const args = argNodes.map((a) => (a ? this.evalSync(a) : undefined));
      return b.fn(args, { line, interp: this, argNodes });
    }

    /** Liefert eine Referenz (Zelle) für ByRef-Übergabe oder null, wenn nur ein Wert übergeben werden kann. */
    async refFor(node) {
      if (node.type === 'Ident') {
        const cell = this.findCell(node.lower);
        if (cell) return cell.isConst ? null : cell;
        if (node.lower in CONSTANTS || node.lower in SPECIAL_VALUES || this.program.procs[node.lower] || BUILTINS[node.lower]) return null;
        return this.implicitCell(node.lower);
      }
      if (node.type === 'CallOrIndex') {
        const cell = this.findCell(node.lower);
        if (cell && cell.value instanceof VbaArray) {
          const arr = cell.value;
          const off = arr.offset(await this.evalIndices(node.args, node.line), node.line);
          return {
            type: arr.elemType, isArray: false, isConst: false, element: true,
            get value() { return arr.data[off]; },
            set value(v) { arr.data[off] = v; }
          };
        }
      }
      return null;
    }

    async callUserProc(proc, argNodes, line) {
      if (this.frames.length >= this.maxDepth) {
        throw runtimeError('Nicht genügend Stapelspeicher – zu viele verschachtelte Aufrufe (Rekursion ohne Abbruchbedingung?).', line, 28);
      }
      const locals = new Map();
      const params = proc.params;
      for (let i = 0; i < params.length; i++) {
        const p = params[i];
        if (p.paramArray) {
          const rest = [];
          for (let k = i; k < argNodes.length; k++) rest.push(argNodes[k] ? await this.ev(argNodes[k]) : MISSING);
          const c = new Cell('variant', VbaArray.fromList(rest, 'variant'));
          c.isArray = true;
          locals.set(p.lower, c);
          break;
        }
        const node = i < argNodes.length ? argNodes[i] : null;
        if (!node) {
          if (!p.optional) throw runtimeError(`Argument nicht optional: Beim Aufruf von '${proc.name}' fehlt der Parameter '${p.name}'.`, line, 449);
          let v;
          if (p.defaultExpr) v = this.evalSyncOrFail(p.defaultExpr);
          else v = p.type === 'variant' ? MISSING : defaultValue(p.type);
          const c = new Cell(p.type, p.type === 'variant' ? v : coerce(v, p.type, line));
          locals.set(p.lower, c);
          continue;
        }
        if (p.byRef) {
          let ref = await this.refFor(node);
          // Ein Steuerelement ist keine Variable: an einen Parameter mit Datentyp wird sein Wert übergeben (wie in Access)
          if (ref instanceof ControlCell && p.type !== 'variant') ref = null;
          if (ref) {
            const refIsArray = ref.isArray || ref.value instanceof VbaArray;
            if (p.isArray && !refIsArray) {
              throw runtimeError(`Typen unverträglich: Für den Parameter '${p.name}' von '${proc.name}' wird ein Array erwartet.`, line, 13);
            }
            if (p.type !== 'variant' && ref.type !== p.type) {
              throw runtimeError(`ByRef-Argumenttyp unverträglich – das Argument für '${p.name}' ist vom Typ ${TYPE_DISPLAY[ref.type]}, der Parameter erwartet ${TYPE_DISPLAY[p.type]}. Tipp: gleichen Datentyp deklarieren oder den Parameter mit ByVal übergeben.`, line, 13);
            }
            locals.set(p.lower, ref);
            continue;
          }
        }
        const v = await this.ev(node);
        let c;
        if (p.isArray || v instanceof VbaArray) {
          if (p.isArray && !(v instanceof VbaArray)) throw runtimeError(`Typen unverträglich: Für den Parameter '${p.name}' von '${proc.name}' wird ein Array erwartet.`, line, 13);
          c = new Cell(p.type, v.clone());
          c.isArray = p.isArray;
        } else {
          c = new Cell(p.type, coerce(v, p.type, line));
        }
        locals.set(p.lower, c);
      }

      const savedFrame = this.frame;
      const savedLine = this.currentLine;
      const frame = { proc, locals, onError: null, inHandler: false };
      this.frames.push(frame);
      this.frame = frame;
      try {
        if (proc.kind === 'function') locals.set(proc.lower, new Cell(proc.returnType, defaultValue(proc.returnType)));
        for (const d of proc.decls) this.declare(d, locals);
        if (proc.labels && proc.labels.size) await this.runProcBody(frame);
        else await this.execBlock(proc.body);
        if (frame.onError) this.clearErr(); // Err wird beim Verlassen einer Prozedur mit Fehlerbehandlung zurückgesetzt
        this.currentLine = savedLine; // nur bei Erfolg: bei Fehlern bleibt die Fehlerzeile stehen
        return proc.kind === 'function' ? locals.get(proc.lower).value : EMPTY;
      } finally {
        this.frames.pop();
        this.frame = savedFrame;
      }
    }

    // ---------- Ausdrücke ----------
    ev(node) {
      return node.async ? this.evalAsync(node) : this.evalSync(node);
    }

    evalSync(node) {
      switch (node.type) {
        case 'Num': case 'Str': case 'Bool': return node.value;
        case 'DateLit': return node.value;
        case 'Nothing': return NOTHING;
        case 'Paren': return this.evalSync(node.expr);
        case 'Unary': {
          const r = this.applyUnary(node.op, this.evalSync(node.operand), node.line);
          return node.ovf ? this.checkOverflow(r, node) : r;
        }
        case 'Bin': {
          const r = this.applyBinary(node.op, this.evalSync(node.left), this.evalSync(node.right), node.line);
          return node.ovf ? this.checkOverflow(r, node) : r;
        }
        case 'Ident': return node.errNumber ? this.err.number : this.readIdentSync(node);
        case 'Member': return this.control(node).get(node.propLower, node.line, (node.args || []).map((a) => (a ? this.evalSync(a) : MISSING)));
        case 'CallOrIndex': {
          if (node.lenBytes) return node.lenBytes;
          const cell = this.findCell(node.lower);
          if (cell && cell.value instanceof VbaArray) {
            const idx = node.args.map((a) => {
              if (!a) throw runtimeError('Index fehlt.', node.line, 9);
              return coerce(this.evalSync(a), 'long', node.line);
            });
            return cell.value.data[cell.value.offset(idx, node.line)];
          }
          if (cell) throw runtimeError(`'${node.name}' ist kein Array – runde Klammern gibt es nur bei Arrays und Funktionsaufrufen.`, node.line, 13);
          const b = BUILTINS[node.lower];
          if (b) return this.callBuiltinSync(b, node.args, node.line);
          throw runtimeError(`Sub oder Function '${node.name}' ist nicht definiert.`, node.line, 35);
        }
        case 'IndexExpr': {
          const target = this.evalSync(node.target);
          const idx = node.args.map((a) => coerce(this.evalSync(a), 'long', node.line));
          return this.indexValue(target, idx, node.line);
        }
        default: throw runtimeError(`Unbekannter Ausdruck (${node.type}).`, node.line);
      }
    }

    async evalAsync(node) {
      switch (node.type) {
        case 'Paren': return this.ev(node.expr);
        case 'Unary': {
          const r = this.applyUnary(node.op, await this.ev(node.operand), node.line);
          return node.ovf ? this.checkOverflow(r, node) : r;
        }
        case 'Bin': {
          const l = await this.ev(node.left);
          const r = await this.ev(node.right);
          const v = this.applyBinary(node.op, l, r, node.line);
          return node.ovf ? this.checkOverflow(v, node) : v;
        }
        case 'Ident': {
          const cell = this.findCell(node.lower);
          if (cell) return cell.value;
          const proc = this.program.procs[node.lower];
          if (proc) return this.callUserProc(proc, [], node.line);
          const b = BUILTINS[node.lower];
          if (b) return this.callBuiltin(b, [], node.line);
          return this.readIdentSync(node);
        }
        case 'CallOrIndex': {
          const cell = this.findCell(node.lower);
          const proc = this.program.procs[node.lower];
          if (cell && cell.value instanceof VbaArray) {
            const idx = await this.evalIndices(node.args, node.line);
            return cell.value.data[cell.value.offset(idx, node.line)];
          }
          if (proc) return this.callUserProc(proc, node.args, node.line);
          if (cell) throw runtimeError(`'${node.name}' ist kein Array – runde Klammern gibt es nur bei Arrays und Funktionsaufrufen.`, node.line, 13);
          const b = BUILTINS[node.lower];
          if (b) return this.callBuiltin(b, node.args, node.line);
          throw runtimeError(`Sub oder Function '${node.name}' ist nicht definiert.`, node.line, 35);
        }
        case 'IndexExpr': {
          const target = await this.ev(node.target);
          const idx = await this.evalIndices(node.args, node.line);
          return this.indexValue(target, idx, node.line);
        }
        case 'Member': {
          const args = [];
          for (const a of node.args || []) args.push(a ? await this.ev(a) : MISSING);
          return this.control(node).get(node.propLower, node.line, args);
        }
        default: return this.evalSync(node);
      }
    }

    /** Überlauf wie in Access: Integer * Integer bleibt Integer (200 * 200 → Laufzeitfehler 6) */
    checkOverflow(v, node) {
      if (typeof v !== 'number') return v;
      const [lo, hi] = OVERFLOW_RANGE[node.ovf];
      if (v < lo || v > hi) {
        const t = TYPE_DE[node.ovf];
        throw runtimeError(`Überlauf: Beide Werte sind vom Typ ${t}, daher rechnet VBA auch das Ergebnis (${numberToString(v)}) als ${t} – und das passt nicht hinein (${t} reicht von ${numberToString(lo)} bis ${numberToString(hi)}). ` +
          (node.ovf === 'long' ? 'Rechne mit Double, z. B. CDbl(a) * b.' : 'Rechne mit Long, z. B. CLng(a) * b oder 200& * 200.'), node.line, 6);
      }
      return v;
    }

    control(node) {
      if (node.special) return this.specialObject(node);
      const c = this.controls.get(node.objLower);
      if (!c) throw runtimeError(`Objekt erforderlich: Es gibt kein Steuerelement '${node.object}'.`, node.line, 424);
      return c;
    }

    /** Me (Formular), Err und DoCmd mit derselben Schnittstelle wie Steuerelemente */
    specialObject(node) {
      const self = this;
      const host = this.host;
      if (node.special === 'err') {
        return {
          get(p) { return p === 'number' ? self.err.number : p === 'description' ? self.err.description : self.err.source; },
          set(p, v, line) {
            if (p === 'number') { self.err.number = coerce(v, 'long', line); if (!self.err.description) self.err.description = RUNTIME_TEXTS[self.err.number] || ''; }
            else if (p === 'description') self.err.description = toStr(v, line);
            else self.err.source = toStr(v, line);
          },
          call(p, line, args) {
            if (p === 'clear') { self.clearErr(); return; }
            // Err.Raise Nummer, [Quelle], [Beschreibung]
            const num = coerce(args[0], 'long', line);
            const desc = args.length > 2 && args[2] !== MISSING ? toStr(args[2], line) : (RUNTIME_TEXTS[num] || 'Anwendungs- oder objektdefinierter Fehler');
            const e = runtimeError(desc, line, num);
            e.accessMessage = desc;
            e.raisedSource = args.length > 1 && args[1] !== MISSING ? toStr(args[1], line) : '';
            throw e;
          }
        };
      }
      if (node.special === 'docmd') {
        return {
          get() { return EMPTY; }, set() {},
          call(p) {
            if (p === 'beep') return;
            // DoCmd.Close: Formular schließen (die Oberfläche wechselt nach dem Ereignis in den Entwurf)
            if (typeof host.closeForm === 'function') host.closeForm();
          }
        };
      }
      // Me: das Formular selbst
      const form = this.formAdapter;
      return {
        get(p) { return p === 'caption' ? (form ? form.getCaption() : '') : EMPTY; },
        set(p, v, line) { if (p === 'caption' && form) form.setCaption(toStr(v, line)); },
        call(p) { if (form && form.refresh) form.refresh(); }
      };
    }

    setErr(e) {
      this.err.number = e.number || 5;
      this.err.description = e.accessMessage || e.plainMessage || '';
      this.err.source = e.raisedSource || '';
    }
    clearErr() { this.err.number = 0; this.err.description = ''; this.err.source = ''; }

    indexValue(target, idx, line) {
      if (!(target instanceof VbaArray)) throw runtimeError('Typen unverträglich: Hier wird ein Array erwartet.', line, 13);
      return target.data[target.offset(idx, line)];
    }

    readIdentSync(node) {
      const cell = this.findCell(node.lower);
      if (cell) return cell.value;
      if (node.lower in CONSTANTS) return CONSTANTS[node.lower];
      if (node.lower in SPECIAL_VALUES) return SPECIAL_VALUES[node.lower];
      const b = BUILTINS[node.lower];
      if (b) return this.callBuiltinSync(b, [], node.line);
      if (this.program.optionExplicit) {
        throw runtimeError(`Variable '${node.name}' wurde nicht deklariert.`, node.line);
      }
      return this.implicitCell(node.lower).value;
    }

    // ---------- Operatoren ----------
    applyUnary(op, v, line) {
      if (op === '-') {
        if (v === NULL) return NULL;
        if (v instanceof VbaDate) return new VbaDate(-v.serial);
        return -toNumber(v, line) || 0;
      }
      // Not
      if (v === NULL) return NULL;
      if (typeof v === 'boolean') return !v;
      return ~coerce(v, 'long', line);
    }

    applyBinary(op, a, b, line) {
      switch (op) {
        case '+': {
          if (typeof a === 'string' && typeof b === 'string') return a + b;
          if (a === NULL || b === NULL) return NULL;
          if (typeof a === 'string' && b === EMPTY) return a;
          if (typeof b === 'string' && a === EMPTY) return b;
          if (a instanceof VbaDate || b instanceof VbaDate) return new VbaDate(toNumber(a, line) + toNumber(b, line));
          return toNumber(a, line) + toNumber(b, line);
        }
        case '-': {
          if (a === NULL || b === NULL) return NULL;
          if (a instanceof VbaDate && !(b instanceof VbaDate)) return new VbaDate(a.serial - toNumber(b, line));
          return toNumber(a, line) - toNumber(b, line);
        }
        case '*':
          if (a === NULL || b === NULL) return NULL;
          return toNumber(a, line) * toNumber(b, line);
        case '/': {
          if (a === NULL || b === NULL) return NULL;
          const x = toNumber(a, line), y = toNumber(b, line);
          if (y === 0) {
            if (x === 0) throw runtimeError('Überlauf: 0 / 0 ist nicht definiert.', line, 6);
            throw runtimeError('Division durch Null.', line, 11);
          }
          return x / y;
        }
        case '\\': case 'mod': {
          if (a === NULL || b === NULL) return NULL;
          const x = bankersRound(toNumber(a, line)), y = bankersRound(toNumber(b, line));
          if (y === 0) throw runtimeError('Division durch Null.', line, 11);
          const r = op === '\\' ? Math.trunc(x / y) : x % y;
          return r === 0 ? 0 : r;
        }
        case '^': {
          if (a === NULL || b === NULL) return NULL;
          const r = Math.pow(toNumber(a, line), toNumber(b, line));
          if (isNaN(r)) throw runtimeError('Ungültiger Prozeduraufruf oder ungültiges Argument: Potenz einer negativen Zahl mit gebrochenem Exponenten.', line, 5);
          return r;
        }
        case '&':
          return (a === NULL ? '' : toStr(a, line)) + (b === NULL ? '' : toStr(b, line));
        case '=': case '<>': case '<': case '>': case '<=': case '>=': {
          const r = compare(a, b, this.program.compareText, line);
          if (r === null) return NULL;
          switch (op) {
            case '=': return r === 0;
            case '<>': return r !== 0;
            case '<': return r < 0;
            case '>': return r > 0;
            case '<=': return r <= 0;
            default: return r >= 0;
          }
        }
        case 'like': {
          if (a === NULL || b === NULL) return NULL;
          return likeToRegex(toStr(b, line), this.program.compareText).test(toStr(a, line));
        }
        case 'is':
          return a === b;
        case 'and': case 'or': case 'xor': case 'eqv': case 'imp': {
          if (typeof a === 'boolean' && typeof b === 'boolean') {
            switch (op) {
              case 'and': return a && b;
              case 'or': return a || b;
              case 'xor': return a !== b;
              case 'eqv': return a === b;
              default: return !a || b;
            }
          }
          if (a === NULL || b === NULL) {
            // vereinfachte Null-Logik
            if (op === 'and' && (a === false || b === false)) return false;
            if (op === 'or' && (a === true || b === true)) return true;
            return NULL;
          }
          const x = typeof a === 'boolean' ? (a ? -1 : 0) : coerce(a, 'long', line);
          const y = typeof b === 'boolean' ? (b ? -1 : 0) : coerce(b, 'long', line);
          switch (op) {
            case 'and': return x & y;
            case 'or': return x | y;
            case 'xor': return x ^ y;
            case 'eqv': return ~(x ^ y);
            default: return (~x) | y;
          }
        }
        default:
          throw runtimeError(`Unbekannter Operator '${op}'.`, line);
      }
    }
  }

  return { Interpreter, Cell, ControlCell };
});
