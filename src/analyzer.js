/*
 * analyzer.js – Prüfung nach dem Parsen ("Kompilieren"), bevor das Programm läuft.
 *
 * - Option Explicit: nicht deklarierte Variablen werden schon vor dem Start gemeldet.
 * - Unbekannte Subs/Functions, falsche Argumentanzahl, Zuweisung an Konstanten.
 * - Markiert Ausdrucksknoten mit `async = true`, wenn darin eine eigene Prozedur
 *   oder MsgBox/InputBox aufgerufen wird (nur diese brauchen den langsamen async-Pfad).
 * - Wandelt Datumsliterale in Werte um.
 */
(function (root, factory) {
  const isNode = typeof module === 'object' && module.exports;
  const deps = isNode ? Object.assign({}, require('./errors.js'), require('./values.js'), require('./builtins.js'), require('./forms.js')) : root.VBA;
  const mod = factory(deps);
  if (isNode) module.exports = mod;
  else { root.VBA = root.VBA || {}; Object.assign(root.VBA, mod); }
})(typeof self !== 'undefined' ? self : this, function (V) {
  'use strict';
  const { syntaxError, BUILTINS, CONSTANTS, SPECIAL_VALUES, ASYNC_BUILTINS, VbaDate, parseDateString, CONTROL_PROPS, READ_ONLY, MEMBER_ARGS, KIND_NAMES } = V;

  // Objekte, die es ohne Formular-Steuerelement gibt
  const SPECIAL_OBJECTS = {
    me: {
      name: 'Me', hint: 'Steuerelemente sprichst du direkt mit ihrem Namen an.',
      members: {
        caption: { kind: 'prop' }, refresh: { kind: 'method' }, requery: { kind: 'method' },
        recalc: { kind: 'method' }, repaint: { kind: 'method' }
      }
    },
    err: {
      name: 'Err',
      members: {
        number: { kind: 'prop' }, description: { kind: 'prop' }, source: { kind: 'prop' },
        clear: { kind: 'method' }, raise: { kind: 'method', args: [1, 3] }
      }
    },
    docmd: {
      name: 'DoCmd', hint: 'Im Editor gibt es nur DoCmd.Close (Formular schließen).',
      members: { close: { kind: 'method', args: [0, 3] }, beep: { kind: 'method' } },
      unsupported: {
        openform: 'DoCmd.OpenForm gibt es hier nicht – der Editor hat nur ein Formular.',
        openreport: 'DoCmd.OpenReport gibt es hier nicht – Berichte werden nicht unterstützt.',
        runsql: 'DoCmd.RunSQL gibt es hier nicht – der Editor hat keine Datenbank.',
        openquery: 'DoCmd.OpenQuery gibt es hier nicht – der Editor hat keine Datenbank.',
        opentable: 'DoCmd.OpenTable gibt es hier nicht – der Editor hat keine Datenbank.',
        quit: 'DoCmd.Quit gibt es hier nicht. Zum Schließen des Formulars: DoCmd.Close.'
      }
    }
  };

  function analyze(program) {
    const procs = program.procs;
    const moduleNames = new Set();
    const moduleConsts = new Set();
    for (const d of program.moduleDecls) {
      moduleNames.add(d.lower);
      if (d.type === 'ConstDecl') moduleConsts.add(d.lower);
    }

    // Steuerelemente des Formulars verhalten sich wie modulweite Variablen
    const controls = new Map();
    for (const c of program.controls || []) controls.set(c.lower, c);
    for (const d of program.moduleDecls) {
      if (controls.has(d.lower)) throw syntaxError(`Mehrdeutiger Name: '${d.name}' ist bereits der Name eines Steuerelements im Formular.`, d.line);
    }
    for (const p of program.procList) {
      if (controls.has(p.lower)) throw syntaxError(`Mehrdeutiger Name: '${p.name}' ist bereits der Name eines Steuerelements im Formular.`, p.line);
    }
    for (const l of controls.keys()) moduleNames.add(l);

    function meCheck(node) {
      if (!node.viaMe || controls.has(node.lower)) return;
      if (controls.size === 0) throw syntaxError("'Me.' wird hier nicht benötigt – es gibt noch kein Formular. Variablen werden direkt über ihren Namen angesprochen.", node.line);
      throw syntaxError(`Methode oder Datenobjekt nicht gefunden: Im Formular gibt es kein Steuerelement '${node.name}'.`, node.line);
    }

    function member(node, mode) {
      const c = controls.get(node.objLower);
      if (!c && SPECIAL_OBJECTS[node.objLower]) return specialMember(node, mode);
      if (!c) {
        if (node.objLower === 'debug') throw syntaxError(`Debug.${node.prop} wird nicht unterstützt – verwende Debug.Print als eigene Anweisung.`, node.line);
        throw syntaxError(controls.size === 0
          ? `Methode oder Datenobjekt nicht gefunden: '${node.object}.${node.prop}' – Objekte gibt es nur als Steuerelemente im Formular.`
          : `Methode oder Datenobjekt nicht gefunden: Im Formular gibt es kein Steuerelement '${node.object}'.`, node.line);
      }
      const def = CONTROL_PROPS[c.kind];
      const isProp = def.p.includes(node.propLower);
      const isMethod = def.m.includes(node.propLower);
      if (mode === 'call' ? !isMethod : !isProp) {
        throw syntaxError(`Methode oder Datenobjekt nicht gefunden: ${KIND_NAMES[c.kind]} '${c.name}' hat ${mode === 'call' ? 'keine Methode' : 'keine Eigenschaft'} '${node.prop}'.`, node.line);
      }
      if (mode === 'write' && READ_ONLY.has(node.propLower)) {
        throw syntaxError(`Die Eigenschaft '${node.prop}' ist schreibgeschützt.`, node.line);
      }
      const range = MEMBER_ARGS[node.propLower] || [0, 0];
      const given = node.args ? node.args.length : 0;
      if (given < range[0]) {
        throw syntaxError(`Argument nicht optional: '${node.prop}' braucht mindestens ${range[0]} Argument(e)` +
          (node.propLower === 'additem' ? `, z. B. ${c.name}.AddItem "Text".` : '.'), node.line);
      }
      if (given > range[1]) throw syntaxError(`Falsche Anzahl an Argumenten für '${node.prop}': höchstens ${range[1]} erlaubt, ${given} angegeben.`, node.line);
      let async = false;
      (node.args || []).forEach((a, i) => {
        if (!a) { if (i < range[0]) throw syntaxError(`Argument nicht optional: '${node.prop}'.`, node.line); return; }
        async = expr(a) || async;
      });
      node.async = async;
      return async;
    }

    // Me (Formular), Err und DoCmd
    function specialMember(node, mode) {
      const obj = SPECIAL_OBJECTS[node.objLower];
      const p = node.propLower;
      if (obj.unsupported && obj.unsupported[p]) throw syntaxError(obj.unsupported[p], node.line);
      const def = obj.members[p];
      if (!def) {
        throw syntaxError(`Methode oder Datenobjekt nicht gefunden: '${obj.name}.${node.prop}' gibt es hier nicht.` +
          (obj.hint ? ' ' + obj.hint : ''), node.line);
      }
      const isMethod = def.kind === 'method';
      if (mode === 'call' && !isMethod) throw syntaxError(`Ungültige Verwendung einer Eigenschaft: '${obj.name}.${node.prop}' ist keine Methode.`, node.line);
      if (mode !== 'call' && isMethod && !(mode === 'read' && def.returns)) throw syntaxError(`'${obj.name}.${node.prop}' ist eine Methode und liefert keinen Wert.`, node.line);
      if (mode === 'write' && def.readOnly) throw syntaxError(`Die Eigenschaft '${obj.name}.${node.prop}' ist schreibgeschützt.`, node.line);
      const range = def.args || [0, 0];
      const given = node.args ? node.args.length : 0;
      if (given < range[0]) throw syntaxError(`Argument nicht optional: '${obj.name}.${node.prop}' braucht mindestens ${range[0]} Argument(e).`, node.line);
      if (given > range[1]) throw syntaxError(`Falsche Anzahl an Argumenten für '${obj.name}.${node.prop}': höchstens ${range[1]} erlaubt.`, node.line);
      let async = false;
      for (const a of node.args || []) if (a) async = expr(a) || async;
      node.special = node.objLower;
      node.async = async;
      return async;
    }

    let scope = null; // { proc, declared:Set, consts:Set, types:Map }

    // ---------- statische Datentypen (für Überlauf wie in Access: 200 * 200) ----------
    const moduleTypes = new Map();
    for (const d of program.moduleDecls) moduleTypes.set(d.lower, d.isArray ? 'array' : (d.varType || 'variant'));
    const INT_TYPES = { byte: 1, boolean: 2, integer: 2, long: 3 };
    const BUILTIN_TYPES = {
      cint: 'integer', clng: 'long', cbyte: 'byte', cbool: 'boolean', cdbl: 'double', csng: 'single', ccur: 'currency',
      cstr: 'string', len: 'long', asc: 'integer', ascw: 'integer', ubound: 'long', lbound: 'long', sgn: 'integer',
      year: 'integer', month: 'integer', day: 'integer', hour: 'integer', minute: 'integer', second: 'integer', weekday: 'integer'
    };
    function varType(lower) {
      if (controls.has(lower)) return 'variant';
      if (scope.types && scope.types.has(lower)) return scope.types.get(lower);
      if (scope.declared.has(lower)) return 'variant';
      return moduleTypes.get(lower) || 'variant';
    }
    function staticType(node) {
      if (!node) return 'variant';
      switch (node.type) {
        case 'Num':
          if (node.numType) return node.numType;
          if (Number.isInteger(node.value)) return Math.abs(node.value) <= 32767 ? 'integer' : (Math.abs(node.value) <= 2147483647 ? 'long' : 'double');
          return 'double';
        case 'Str': return 'string';
        case 'Bool': return 'boolean';
        case 'Paren': return staticType(node.expr);
        case 'Ident': {
          const t = varType(node.lower);
          if (t === 'variant' && procs[node.lower] && procs[node.lower].kind === 'function' && !isVar(node.lower)) return procs[node.lower].returnType;
          return t;
        }
        case 'CallOrIndex': {
          if (isVar(node.lower) && !(scope.proc && scope.proc.lower === node.lower)) {
            const t = varType(node.lower);
            return t === 'array' ? 'variant' : t; // Element eines Arrays
          }
          if (procs[node.lower]) return procs[node.lower].kind === 'function' ? procs[node.lower].returnType : 'variant';
          if (BUILTIN_TYPES[node.lower]) return BUILTIN_TYPES[node.lower];
          if ((node.lower === 'int' || node.lower === 'fix' || node.lower === 'abs') && node.args.length === 1) return staticType(node.args[0]);
          return 'variant';
        }
        case 'Unary': return node.op === '-' ? staticType(node.operand) : (node.ovf || 'variant');
        case 'Bin': return node.stype || 'variant';
        default: return 'variant';
      }
    }
    // Ergebnistyp ganzzahliger Rechnungen: Byte < Integer < Long (Boolean zählt als Integer)
    function intResult(a, b) {
      if (!(a in INT_TYPES) || !(b in INT_TYPES)) return null;
      if (a === 'byte' && b === 'byte') return 'byte';
      return Math.max(INT_TYPES[a], INT_TYPES[b], 2) === 3 ? 'long' : 'integer';
    }
    function markOverflow(node) {
      if (node.type === 'Bin') {
        const a = staticType(node.left), b = staticType(node.right);
        if (['+', '-', '*', '\\', 'mod'].includes(node.op)) {
          const r = intResult(a, b);
          if (r) { node.ovf = r; node.stype = r; return; }
          if ((node.op === '\\' || node.op === 'mod')) node.stype = (a === 'variant' || b === 'variant') ? 'variant' : 'long';
          else if (a === 'double' || b === 'double' || a === 'single' || b === 'single') node.stype = (a === 'string' || b === 'string' || a === 'variant' || b === 'variant') ? 'variant' : 'double';
        } else if (node.op === '/' || node.op === '^') {
          node.stype = 'double';
        } else if (node.op === '&') node.stype = 'string';
      } else if (node.type === 'Unary' && node.op === '-') {
        const t = staticType(node.operand);
        if (t in INT_TYPES && t !== 'boolean') node.ovf = t === 'byte' ? 'integer' : t;
      }
    }
    // Len(i) bei Integer & Co. liefert wie in Access die Größe in Bytes
    const BYTE_SIZES = { byte: 1, boolean: 2, integer: 2, long: 4, single: 4, double: 8, currency: 8, date: 8 };

    const isVar = (lower) => scope.declared.has(lower) || moduleNames.has(lower);
    const isConst = (lower) => scope.consts.has(lower) || (!scope.declared.has(lower) && moduleConsts.has(lower));

    function checkArgCount(name, min, max, given, line) {
      if (given > max) {
        throw syntaxError(`Falsche Anzahl an Argumenten für '${name}': höchstens ${max} erlaubt, ${given} angegeben.`, line);
      }
      if (given < min) {
        throw syntaxError(`Argument nicht optional: '${name}' braucht mindestens ${min} Argument(e), angegeben ${given}.`, line);
      }
    }
    function checkProcArgs(proc, args, line) {
      const hasParamArray = proc.params.some((p) => p.paramArray);
      const max = hasParamArray ? Infinity : proc.params.length;
      if (args.length > max) {
        throw syntaxError(`Falsche Anzahl an Argumenten für '${proc.name}': erwartet ${proc.params.length}, angegeben ${args.length}.`, line);
      }
      proc.params.forEach((p, i) => {
        if (!p.optional && !p.paramArray && (i >= args.length || args[i] === null)) {
          throw syntaxError(`Argument nicht optional: Beim Aufruf von '${proc.name}' fehlt der Parameter '${p.name}'.`, line);
        }
      });
    }
    function builtinArgs(lower, name, args, line) {
      const b = BUILTINS[lower];
      // Leere Argumente am Ende zählen nicht
      let n = args.length;
      while (n > 0 && args[n - 1] === null) n--;
      checkArgCount(name, b.min, b.max, n, line);
    }

    function undeclared(name, line) {
      return syntaxError(`Variable '${name}' wurde nicht deklariert. (Option Explicit ist aktiv – bitte mit Dim deklarieren.)`, line);
    }

    // ---------- Ausdrücke ----------
    function expr(node) {
      if (!node) return false;
      let async = false;
      switch (node.type) {
        case 'Num': case 'Str': case 'Bool': case 'Nothing': break;
        case 'DateLit': {
          const s = parseDateString(node.raw, true);
          if (isNaN(s)) throw syntaxError(`Ungültiges Datumsliteral #${node.raw}#. Format: #Monat/Tag/Jahr#, z. B. #12/24/2026#.`, node.line);
          node.value = new VbaDate(s);
          break;
        }
        case 'Paren': async = expr(node.expr); break;
        case 'Unary': async = expr(node.operand); markOverflow(node); break;
        case 'Bin': async = expr(node.left); async = expr(node.right) || async; markOverflow(node); break;
        case 'Ident':
          if (node.lower === 'err' && !isVar(node.lower) && !controls.has('err')) { node.errNumber = true; break; }
          meCheck(node); async = ident(node); break;
        case 'Member': async = member(node, 'read'); break;
        case 'CallOrIndex': async = callOrIndex(node); break;
        case 'IndexExpr':
          async = expr(node.target);
          for (const a of node.args) async = expr(a) || async;
          break;
        default: break;
      }
      node.async = async;
      return async;
    }

    function ident(node) {
      const l = node.lower;
      if (isVar(l)) return !!procs[l];
      if (l in CONSTANTS || l in SPECIAL_VALUES) return false;
      if (procs[l]) {
        const p = procs[l];
        if (p.kind === 'sub') throw syntaxError(`'${p.name}' ist eine Sub und liefert keinen Wert. Nur eine Function kann in einem Ausdruck verwendet werden.`, node.line);
        checkProcArgs(p, [], node.line);
        return true;
      }
      if (BUILTINS[l]) {
        builtinArgs(l, node.name, [], node.line);
        return ASYNC_BUILTINS.has(l);
      }
      if (l === 'tab' && node.inPrint) return false;
      if (program.optionExplicit) throw undeclared(node.name, node.line);
      scope.declared.add(l); // implizite Variant-Variable
      return false;
    }

    function callOrIndex(node) {
      const l = node.lower;
      let async = false;
      for (const a of node.args) async = expr(a) || async;
      const ownFunction = scope.proc && scope.proc.lower === l && scope.proc.kind === 'function';
      if (isVar(l) && !ownFunction) return async || !!procs[l];
      if (procs[l]) {
        const p = procs[l];
        if (p.kind === 'sub') throw syntaxError(`'${p.name}' ist eine Sub und liefert keinen Wert. Nur eine Function kann in einem Ausdruck verwendet werden.`, node.line);
        checkProcArgs(p, node.args, node.line);
        return true;
      }
      if (BUILTINS[l]) {
        builtinArgs(l, node.name, node.args, node.line);
        if (l === 'len' && node.args[0] && node.args[0].type === 'Ident' && BYTE_SIZES[varType(node.args[0].lower)]) {
          node.lenBytes = BYTE_SIZES[varType(node.args[0].lower)];
        }
        return async || ASYNC_BUILTINS.has(l);
      }
      if ((l === 'tab' || l === 'spc') && node.inPrint) return async;
      if (l in CONSTANTS) throw syntaxError(`'${node.name}' ist eine Konstante und kein Array.`, node.line);
      throw syntaxError(`Sub oder Function '${node.name}' ist nicht definiert.`, node.line);
    }

    // ---------- Anweisungen ----------
    function writeTarget(name, lower, line) {
      if (isConst(lower)) throw syntaxError(`Einer Konstanten ('${name}') kann kein Wert zugewiesen werden.`, line);
      if (isVar(lower)) return;
      if (lower in CONSTANTS) throw syntaxError(`'${name}' ist eine eingebaute Konstante – ihr kann kein Wert zugewiesen werden.`, line);
      if (procs[lower]) {
        throw syntaxError(`'${name}' ist der Name einer Prozedur. Den Rückgabewert einer Function setzt man nur innerhalb dieser Function.`, line);
      }
      if (BUILTINS[lower]) {
        throw syntaxError(`'${name}' ist eine eingebaute Funktion und kann nicht als Variable verwendet werden. Bitte einen anderen Namen wählen oder die Variable mit Dim deklarieren.`, line);
      }
      if (program.optionExplicit) throw undeclared(name, line);
      scope.declared.add(lower);
    }

    function block(list) { for (const s of list) stmt(s); }

    function stmt(s) {
      switch (s.type) {
        case 'Dim':
          for (const d of s.decls) {
            if (d.type === 'ConstDecl') expr(d.expr);
            else if (d.bounds) for (const b of d.bounds) { expr(b.lower); expr(b.upper); }
          }
          break;
        case 'ReDim':
          for (const it of s.items) {
            if (isConst(it.lower)) throw syntaxError(`Einer Konstanten ('${it.name}') kann kein Wert zugewiesen werden.`, it.line);
            for (const b of it.bounds) { expr(b.lower); expr(b.upper); }
          }
          break;
        case 'MemberCall': member(s.target, 'call'); break;
        case 'Assign':
          if (s.target.type === 'Ident') { meCheck(s.target); writeTarget(s.target.name, s.target.lower, s.line); }
          else if (s.target.type === 'Member') member(s.target, 'write');
          else {
            const l = s.target.lower;
            if (!isVar(l)) {
              if (procs[l] || BUILTINS[l]) throw syntaxError(`'${s.target.name}' ist eine Prozedur bzw. Funktion – einem Aufruf kann kein Wert zugewiesen werden.`, s.line);
              throw syntaxError(program.optionExplicit
                ? `Variable '${s.target.name}' wurde nicht deklariert. Arrays bitte mit Dim deklarieren, z. B. Dim ${s.target.name}(1 To 10).`
                : `'${s.target.name}' ist kein Array. Arrays bitte mit Dim deklarieren, z. B. Dim ${s.target.name}(1 To 10).`, s.line);
            }
            for (const a of s.target.args) { if (!a) throw syntaxError('Index fehlt.', s.line); expr(a); }
          }
          s.valueAsync = expr(s.value);
          break;
        case 'Call': {
          const l = s.lower;
          for (const a of s.args) expr(a);
          if (procs[l]) { checkProcArgs(procs[l], s.args, s.line); break; }
          if (BUILTINS[l]) { builtinArgs(l, s.name, s.args, s.line); break; }
          if (isVar(l)) {
            throw syntaxError(`'${s.name}' ist eine Variable. Für eine Zuweisung fehlt das '=' (z. B. ${s.name} = 5).`, s.line);
          }
          throw syntaxError(`Sub oder Function '${s.name}' ist nicht definiert.`, s.line);
        }
        case 'Print':
          for (const it of s.items) {
            if (!it.expr) continue;
            const e = it.expr;
            // Tab(n) und Spc(n) gibt es nur in Debug.Print
            if ((e.type === 'CallOrIndex' && (e.lower === 'tab' || e.lower === 'spc') && !isVar(e.lower) && !procs[e.lower]) ||
                (e.type === 'Ident' && e.lower === 'tab' && !isVar(e.lower) && !procs[e.lower])) {
              e.inPrint = true;
              if (e.type === 'CallOrIndex' && e.args.length > 1) throw syntaxError(`Falsche Anzahl an Argumenten für '${e.name}'.`, e.line);
              if (e.type === 'CallOrIndex' && e.lower === 'spc' && e.args.length !== 1) throw syntaxError("Argument nicht optional: 'Spc' braucht die Anzahl der Leerzeichen.", e.line);
            }
            expr(e);
          }
          break;
        case 'Label': case 'Stop': break;
        case 'GoTo': case 'OnError': case 'Resume':
          if (s.label !== undefined && s.label !== null && !scope.labels.has(s.label)) {
            throw syntaxError(`Sprungmarke nicht definiert: '${s.labelName}'. Lege sie in dieser Prozedur an, z. B. ${s.labelName}: in einer eigenen Zeile.`, s.line);
          }
          break;
        case 'Assert': expr(s.expr); break;
        case 'If':
          for (const b of s.branches) { expr(b.cond); block(b.body); }
          if (s.elseBody) block(s.elseBody);
          break;
        case 'Select':
          expr(s.expr);
          for (const c of s.cases) {
            for (const t of c.tests) { expr(t.expr); expr(t.from); expr(t.to); }
            block(c.body);
          }
          if (s.elseBody) block(s.elseBody);
          break;
        case 'For':
          writeTarget(s.varName, s.lower, s.line);
          expr(s.start); expr(s.end); expr(s.step);
          block(s.body);
          break;
        case 'ForEach':
          writeTarget(s.varName, s.lower, s.line);
          expr(s.collection);
          block(s.body);
          break;
        case 'Do':
          if (s.preCond) expr(s.preCond.expr);
          if (s.postCond) expr(s.postCond.expr);
          block(s.body);
          break;
        case 'While': expr(s.cond); block(s.body); break;
        default: break;
      }
    }

    function checkNestedLabels(list, top) {
      for (const st of list) {
        if (st.type === 'Label' && !top) throw syntaxError(`Die Sprungmarke '${st.name}' steht innerhalb eines Blocks (If, Schleife …). Hier werden Sprungmarken nur direkt in der Prozedur unterstützt.`, st.line);
        for (const key of ['body', 'elseBody']) if (Array.isArray(st[key])) checkNestedLabels(st[key], false);
        if (st.branches) for (const b of st.branches) checkNestedLabels(b.body, false);
        if (st.cases) for (const c of st.cases) checkNestedLabels(c.body, false);
      }
    }

    function collectReDims(list, out) {
      for (const s of list) {
        if (s.type === 'ReDim') for (const it of s.items) out.add(it.lower);
        for (const key of ['body', 'elseBody']) if (Array.isArray(s[key])) collectReDims(s[key], out);
        if (s.branches) for (const b of s.branches) collectReDims(b.body, out);
        if (s.cases) for (const c of s.cases) collectReDims(c.body, out);
      }
    }

    // Modulebene
    scope = { proc: null, declared: new Set(), consts: new Set(), types: new Map(), labels: new Map() };
    for (const d of program.moduleDecls) {
      if (d.type === 'ConstDecl') expr(d.expr);
      else if (d.bounds) for (const b of d.bounds) { expr(b.lower); expr(b.upper); }
    }

    for (const proc of program.procList) {
      scope = { proc, declared: new Set(), consts: new Set() };
      for (const p of proc.params) {
        scope.declared.add(p.lower);
        if (p.defaultExpr) expr(p.defaultExpr);
      }
      if (proc.kind === 'function') scope.declared.add(proc.lower);
      scope.types = new Map();
      for (const p of proc.params) scope.types.set(p.lower, p.isArray || p.paramArray ? 'array' : p.type);
      if (proc.kind === 'function') scope.types.set(proc.lower, proc.returnType);
      for (const d of proc.decls) {
        scope.declared.add(d.lower);
        if (d.type === 'ConstDecl') scope.consts.add(d.lower);
        scope.types.set(d.lower, d.isArray ? 'array' : (d.type === 'ConstDecl' && d.varType === 'variant' ? staticType(d.expr) : d.varType));
      }
      collectReDims(proc.body, scope.declared);
      // Sprungmarken (nur auf oberster Ebene der Prozedur)
      scope.labels = new Map();
      proc.labels = scope.labels;
      proc.body.forEach((st, i) => {
        if (st.type !== 'Label') return;
        if (scope.labels.has(st.lower)) throw syntaxError(`Doppelte Sprungmarke: '${st.name}' gibt es in dieser Prozedur schon.`, st.line);
        scope.labels.set(st.lower, i);
      });
      checkNestedLabels(proc.body, true);
      block(proc.body);
    }
    return program;
  }

  return { analyze, SPECIAL_OBJECTS };
});
