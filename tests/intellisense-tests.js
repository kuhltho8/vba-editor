#!/usr/bin/env node
/* Tests für die Live-Hilfe (QuickInfo und Vorschlagsliste): node tests/intellisense-tests.js */
'use strict';
const path = require('path');
const IS = require(path.join(__dirname, '..', 'js', 'intellisense.js'));

const CTL = [{ name: 'lstAusgabe', kind: 'listbox' }, { name: 'e_zahl', kind: 'textbox' }];
// | markiert die Schreibmarke; tip = Text des fetten Parameters, list = erste Einträge
const cases = [
  { code: 'MsgBox "Hallo", |', typed: ',', tip: '[Buttons As VbMsgBoxStyle = vbOKOnly]', list: ['vbAbortRetryIgnore', 'vbCritical'] },
  { code: 'x = MsgBox("Hallo, Welt", vbYes|', tip: '[Buttons As VbMsgBoxStyle = vbOKOnly]', list: ['vbYesNo', 'vbYesNoCancel'] },
  { code: 'MsgBox "a", vbYesNo + |', typed: ' ', list: ['vbAbortRetryIgnore'] },
  { code: 'x = InputBox(|', tip: 'Prompt', list: null },
  { code: 'x = InputBox("Zahl?", |', tip: '[Title]' },
  { code: 'Debug.Print vbInf|', tip: null, list: ['vbInformation'] },
  { code: 'Debug.Print Left(s, |', tip: 'Length As Long' },
  { code: 'Debug.Print Left(Mid(s, 2, |', tip: '[Length]' },
  { code: 'Debug.Print Left(Mid(s, 2), |', tip: 'Length As Long' },
  { code: 'lstAusgabe.|', list: ['AddItem', 'BackColor'] },
  { code: 'lstAusgabe.AddItem i & " x ", |', tip: '[Index]' },
  { code: 'Dim summe As Integer\nsu|', list: ['Sub', 'summe'] },
  { code: 'Dim x As Lo|', list: ['Long'] },
  { code: "' MsgBox |", tip: null, list: null },
  { code: 'Debug.Print "vb|', tip: null, list: null },
  { code: 'Gruss "Anna", |\nSub Gruss(ByVal name As String, Optional alter As Integer)\nEnd Sub', tip: '[alter As Integer]' },
  { code: 'Next|', list: null },
  { code: 'e_|', list: ['e_zahl'] }
];
let fail = 0;
for (const c of cases) {
  const pos = c.code.indexOf('|');
  const src = c.code.replace('|', '');
  const r = IS.analyze(src, pos, CTL, { fromInput: true, typed: c.typed || 'x' });
  const bold = r.tip ? (/<b>(.*?)<\/b>/.exec(IS.tipHtml(r.tip)) || [])[1] || '' : null;
  const problems = [];
  if ('tip' in c && (bold === null ? null : bold.replace(/&lt;/g, '<')) !== c.tip) problems.push(`QuickInfo: ${bold} statt ${c.tip}`);
  if ('list' in c) {
    const got = r.list ? r.list.items.map((x) => x.text) : null;
    if (c.list === null ? got !== null : !got || c.list.some((x, i) => got[i] !== x)) problems.push(`Liste: ${got && got.slice(0, 5)} statt ${c.list}`);
  }
  if (problems.length) { fail++; console.log(`✗ ${c.code.split('\n')[0]}\n  ${problems.join('\n  ')}`); }
}
console.log(`${cases.length - fail} von ${cases.length} Live-Hilfe-Tests bestanden.`);
process.exit(fail ? 1 : 0);
