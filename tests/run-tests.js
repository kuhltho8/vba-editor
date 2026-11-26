#!/usr/bin/env node
/* Unit-Tests für den Interpreter: node tests/run-tests.js */
'use strict';
const path = require('path');
const VBA = require(path.join(__dirname, '..', 'src', 'vba.js'));
const { cases, runCase } = require('./cases.js');

(async () => {
  let passed = 0;
  const failed = [];
  for (const c of cases) {
    const r = await runCase(VBA, c);
    if (r.ok) passed++;
    else failed.push({ name: c.name, details: r.details });
  }
  for (const f of failed) console.log(`✗ ${f.name}\n  ${f.details}`);
  console.log(`\n${passed} von ${cases.length} Tests bestanden.`);
  process.exit(failed.length ? 1 : 0);
})();
