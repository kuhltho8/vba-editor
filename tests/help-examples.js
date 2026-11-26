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
        rt.userInput('E_Zahl1', '2,5'); rt.userInput('E_Zahl2', '4'); rt.userSelect('E_Rechenart', 3);
        await r.invoke('btnBerechnen_Click');
        out.push('A_Ergebnis = ' + rt.displayText('A_Ergebnis'));
        if (rt.displayText('A_Ergebnis') !== '10') throw new Error('falsches Ergebnis');
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
