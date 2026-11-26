/*
 * vba.js – öffentliche Schnittstelle.
 *
 * Browser: nach errors.js, tokenizer.js, parser.js, values.js, builtins.js,
 *          analyzer.js und interpreter.js einbinden → window.VBA
 * Node:    const VBA = require('./src/vba.js')
 *
 *   const prog = VBA.compile(quelltext);         // wirft VbaError bei Syntaxfehlern
 *   VBA.startableSubs(prog)                      // [{name, line}, ...] für das Start-Dropdown
 *   const run = VBA.createRun(prog, host, opts); // run.start('Main'), run.stop()
 */
(function (root, factory) {
  const isNode = typeof module === 'object' && module.exports;
  const deps = isNode
    ? Object.assign({},
      require('./errors.js'), require('./tokenizer.js'), require('./parser.js'), require('./values.js'),
      require('./builtins.js'), require('./forms.js'), require('./analyzer.js'), require('./interpreter.js'))
    : root.VBA;
  const mod = factory(deps);
  if (isNode) module.exports = Object.assign({}, deps, mod);
  else { root.VBA = root.VBA || {}; Object.assign(root.VBA, mod); }
})(typeof self !== 'undefined' ? self : this, function (V) {
  'use strict';

  /**
   * @param {string} source Quelltext
   * @param {{controls?: Array<{name:string, kind:string}>}} [opts] Steuerelemente des Formulars
   */
  function compile(source, opts) {
    const program = V.parse(source);
    program.controls = ((opts && opts.controls) || []).map((c) => ({ name: c.name, lower: c.name.toLowerCase(), kind: c.kind }));
    V.analyze(program);
    return program;
  }

  function startableSubs(program) {
    return V.Interpreter.startableSubs(program).map((p) => ({ name: p.name, line: p.line }));
  }

  function createRun(program, host, options) {
    const interp = new V.Interpreter(program, host, options);
    return {
      interpreter: interp,
      start: (entry) => interp.run(entry),
      open: () => interp.open(),
      invoke: (name) => interp.invoke(name),
      has: (name) => interp.hasProc(name),
      get running() { return interp.running; },
      stop: () => interp.stop()
    };
  }

  /** Bequemer Komplettaufruf: übersetzen und ausführen. */
  async function run(source, host, options) {
    const program = compile(source);
    const r = createRun(program, host, options);
    return r.start(options && options.entry);
  }

  return { compile, startableSubs, createRun, run, version: '0.1.0' };
});
