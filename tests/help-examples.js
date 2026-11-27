#!/usr/bin/env node
/* Prüft, dass alle Beispiele aus der Hilfe fehlerfrei übersetzt werden und laufen. */
'use strict';
const path = require('path');
const VBA = require(path.join(__dirname, '..', 'src', 'vba.js'));
const { TOPICS } = require(path.join(__dirname, '..', 'js', 'help.js'));
(async () => {
  let fail = 0;
  for (const t of TOPICS) {
    if (!t.code) continue;
    const out = [];
    try {
      if (t.form) {
        const rt = new VBA.FormRuntime(t.form);
        const prog = VBA.compile(t.code, { controls: t.form.controls });
        const r = VBA.createRun(prog, { print: (l) => out.push(l), msgBox: async () => 1 }, { controls: rt.adapters() });
        r.open();
        if (r.has('Form_Load')) await r.invoke('Form_Load');
        const tc = t.test || {};
        for (const [n, v] of Object.entries(tc.input || {})) rt.userInput(n, v);
        for (const [n, v] of Object.entries(tc.select || {})) rt.userSelect(n, v);
        if (tc.click) await r.invoke(tc.click + '_Click');
        for (const [n, want] of Object.entries(tc.expect || {})) {
          const st = rt.get(n);
          const got = st.items ? st.items.join('|') : rt.displayText(n);
          out.push(n + ' = ' + got);
          if (got !== want) throw new Error(`falsches Ergebnis in ${n}: ${got}`);
        }
        console.log(`✓ ${t.title}\n    ${out.join('\n    ')}`);
        continue;
      }
      await VBA.run(t.code, { print: (l) => out.push(l), msgBox: async () => 6, inputBox: async (p, ti, d) => d || '42' });
      console.log(`✓ ${t.title}\n    ${out.join('\n    ')}`);
    } catch (e) {
      if (t.id === 'errors') { console.log(`✓ ${t.title} (erwarteter Fehler: ${e.message})`); continue; }
      fail++; console.log(`✗ ${t.title}: ${e.message}`);
    }
  }
  process.exit(fail ? 1 : 0);
})();
