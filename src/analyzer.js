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
  const { syntaxError, BUILTINS, CONSTANTS, SPECIAL_VALUES, ASYNC_BUILTINS, VbaDate, parseDateString, CONTROL_PROPS, READ_ONLY, KIND_NAMES } = V;

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
    }

    let scope = null; // { proc, declared:Set, consts:Set }

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
        case 'Unary': async = expr(node.operand); break;
        case 'Bin': async = expr(node.left); async = expr(node.right) || async; break;
        case 'Ident': meCheck(node); async = ident(node); break;
        case 'Member': member(node, 'read'); break;
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
        return async || ASYNC_BUILTINS.has(l);
      }
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
          for (const it of s.items) if (it.expr) expr(it.expr);
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

    function collectReDims(list, out) {
      for (const s of list) {
        if (s.type === 'ReDim') for (const it of s.items) out.add(it.lower);
        for (const key of ['body', 'elseBody']) if (Array.isArray(s[key])) collectReDims(s[key], out);
        if (s.branches) for (const b of s.branches) collectReDims(b.body, out);
        if (s.cases) for (const c of s.cases) collectReDims(c.body, out);
      }
    }

    // Modulebene
    scope = { proc: null, declared: new Set(), consts: new Set() };
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
      for (const d of proc.decls) {
        scope.declared.add(d.lower);
        if (d.type === 'ConstDecl') scope.consts.add(d.lower);
      }
      collectReDims(proc.body, scope.declared);
      block(proc.body);
    }
    return program;
  }

  return { analyze };
});
