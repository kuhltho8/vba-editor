/*
 * app.js – Oberfläche: Editor, Konsole, Dialoge, Seitenleiste mit Formular (Entwurf/Formularansicht)
 * und Hilfe, Neu/Rückgängig, Speichern/Öffnen.
 */
(function () {
  'use strict';
  const $ = (id) => document.getElementById(id);
  const code = $('code'), gutter = $('gutter'), consoleEl = $('console'), entry = $('entry');
  const btnRun = $('btnRun'), btnStop = $('btnStop');
  const STORAGE_KEY = 'vba-editor-code';
  const FORM_KEY = 'vba-editor-form';
  const store = {
    get(k) { try { return localStorage.getItem(k); } catch (e) { return null; } },
    set(k, v) { try { localStorage.setItem(k, v); } catch (e) { /* ignorieren */ } }
  };

  // Grundgerüst, das beim ersten Öffnen im Editor steht
  const SKELETON = [
    "Private Sub Main()",
    "    Debug.Print(\"Hello World\")",
    "End Sub",
    ""
  ].join('\n');

  let errorLine = null;
  let current = null;

  // ---------- Editor ----------
  const hl = $('hl');
  function renderGutter() {
    const n = code.value.split('\n').length;
    let html = '';
    for (let i = 1; i <= n; i++) html += (i === errorLine ? `<span class="err">${i}</span>` : i) + '\n';
    gutter.innerHTML = html;
    hl.innerHTML = VBAHighlight.highlight(code.value, errorLine);
    syncScroll();
  }
  function syncScroll() {
    gutter.scrollTop = code.scrollTop;
    hl.scrollTop = code.scrollTop;
    hl.scrollLeft = code.scrollLeft;
  }
  code.addEventListener('scroll', syncScroll);
  code.addEventListener('input', () => {
    errorLine = null;
    renderGutter();
    refreshEntries();
    store.set(STORAGE_KEY, code.value);
    if (formSession) formSession.stale = true;
    if (designer) designer.renderProps();
  });
  // Live-Hilfe (QuickInfo und Vorschlagsliste) wie im VBA-Editor
  const intelli = new VBAIntelliSense.IntelliSense(code, () => formModel.controls);
  code.addEventListener('keydown', (e) => {
    if (intelli.handleKey(e)) return;
    if (e.key === 'Tab') {
      e.preventDefault();
      const s = code.selectionStart, t = code.selectionEnd;
      code.setRangeText('    ', s, t, 'end');
      code.dispatchEvent(new Event('input'));
    } else if (e.key === 'Enter') {
      // Einrückung der aktuellen Zeile übernehmen
      const s = code.selectionStart;
      const lineStart = code.value.lastIndexOf('\n', s - 1) + 1;
      const indent = /^[ \t]*/.exec(code.value.slice(lineStart, s))[0];
      e.preventDefault();
      code.setRangeText('\n' + indent, s, code.selectionEnd, 'end');
      code.dispatchEvent(new Event('input'));
    }
  });
  document.addEventListener('keydown', (e) => {
    // Strg+Pause hält das Programm an wie in Access
    if ((e.key === 'Pause' || e.key === 'Cancel') && current) { e.preventDefault(); current.stop(); if (closeModal) closeModal(null); return; }
    if (e.key === 'F5') { e.preventDefault(); if (!current) { showPane(hasForm() ? 'formPane' : 'consolePane'); run(); } }
  });

  function markStop(line) {
    if (line) { markError(line); showPane('editorPane'); } else markError(null);
  }

  function markError(line) {
    errorLine = line || null;
    renderGutter();
    if (!line) return;
    const lines = code.value.split('\n');
    let pos = 0;
    for (let i = 0; i < line - 1 && i < lines.length; i++) pos += lines[i].length + 1;
    const lh = parseFloat(getComputedStyle(code).lineHeight) || 21;
    code.scrollTop = Math.max(0, (line - 4) * lh);
    try { code.setSelectionRange(pos, pos); } catch (e) { /* ignorieren */ }
    syncScroll();
  }

  function refreshEntries() {
    let subs = [];
    try { subs = VBA.startableSubs(VBA.compile(code.value, { controls: formModel.controls })); } catch (e) {
      // beim Tippen: grob nach "Sub Name()" suchen
      const re = /^\s*(?:(?:Public|Private)\s+)?Sub\s+([A-Za-zÄÖÜäöüß][\wÄÖÜäöüß]*)\s*\(\s*\)/gim;
      let m;
      while ((m = re.exec(code.value))) subs.push({ name: m[1] });
    }
    subs = subs.filter((x) => !isEventProc(x.name));
    const prev = entry.value;
    entry.innerHTML = '';
    if (subs.length === 0) {
      const o = document.createElement('option');
      o.value = ''; o.textContent = hasForm() ? '(nur Formular)' : '(keine Sub)';
      entry.appendChild(o);
    }
    for (const s of subs) {
      const o = document.createElement('option');
      o.value = o.textContent = s.name;
      entry.appendChild(o);
    }
    document.body.classList.toggle('one-entry', entry.options.length <= 1); // Handy: Auswahl nur bei mehreren Subs
    if (subs.some((s) => s.name === prev)) entry.value = prev;
    else { const main = subs.find((s) => s.name.toLowerCase() === 'main'); if (main) entry.value = main.name; }
  }

  // ---------- Konsole ----------
  // Direktfenster: Ausgaben sammeln und höchstens einmal pro Bildaufbau schreiben.
  // Wie in Access bleiben nur die letzten 200 Zeilen stehen.
  const CONSOLE_MAX = 200;
  let pendingOut = [];
  let outScheduled = false;
  function out(text, cls) {
    pendingOut.push({ text, cls });
    if (pendingOut.length > CONSOLE_MAX) pendingOut.splice(0, pendingOut.length - CONSOLE_MAX);
    if (!outScheduled) {
      outScheduled = true;
      const go = () => { if (outScheduled) flushOut(); };
      if (typeof requestAnimationFrame === 'function') requestAnimationFrame(go);
      setTimeout(go, 60); // auch im Hintergrund-Tab
    }
  }
  function flushOut() {
    outScheduled = false;
    if (!pendingOut.length) return;
    const frag = document.createDocumentFragment();
    for (const o of pendingOut) {
      const span = document.createElement('span');
      if (o.cls) span.className = o.cls;
      span.textContent = o.text + '\n';
      frag.appendChild(span);
    }
    pendingOut = [];
    consoleEl.appendChild(frag);
    let extra = consoleEl.childElementCount - CONSOLE_MAX;
    while (extra-- > 0) consoleEl.firstChild.remove();
    consoleEl.scrollTop = consoleEl.scrollHeight;
  }
  function clearConsole() { pendingOut = []; consoleEl.textContent = ''; }

  // ---------- Tabs (Handy) ----------
  function showPane(id) {
    document.querySelectorAll('.tab').forEach((t) => t.classList.toggle('active', t.dataset.pane === id));
    document.querySelectorAll('.pane').forEach((p) => p.classList.toggle('active', p.id === id));
  }
  document.querySelectorAll('.tab').forEach((t) => t.addEventListener('click', () => showPane(t.dataset.pane)));

  // ---------- Trenner zwischen Code und Direktfenster ----------
  const SPLIT_KEY = 'vba-editor-split';
  const splitter = $('splitter');
  function setSplit(ratio, save) {
    const r = Math.max(0.2, Math.min(0.85, ratio));
    $('editorPane').style.flex = `${r} 1 0`;
    $('consolePane').style.flex = `${1 - r} 1 0`;
    if (save) store.set(SPLIT_KEY, String(r));
    return r;
  }
  let splitRatio = parseFloat(store.get(SPLIT_KEY));
  splitRatio = setSplit(isFinite(splitRatio) ? splitRatio : 0.6, false);
  splitter.addEventListener('pointerdown', (e) => {
    e.preventDefault();
    splitter.setPointerCapture(e.pointerId);
    const col = splitter.parentElement.getBoundingClientRect();
    const move = (ev) => { splitRatio = setSplit((ev.clientY - col.top) / col.height, false); };
    const up = () => {
      splitter.removeEventListener('pointermove', move);
      splitter.removeEventListener('pointerup', up);
      splitter.removeEventListener('pointercancel', up);
      setSplit(splitRatio, true);
    };
    splitter.addEventListener('pointermove', move);
    splitter.addEventListener('pointerup', up);
    splitter.addEventListener('pointercancel', up);
  });
  splitter.addEventListener('keydown', (e) => {
    const d = { ArrowUp: -0.05, ArrowDown: 0.05 }[e.key];
    if (d) { e.preventDefault(); splitRatio = setSplit(splitRatio + d, true); }
  });
  splitter.addEventListener('dblclick', () => { splitRatio = setSplit(0.6, true); });

  // ---------- Dialoge (MsgBox / InputBox) ----------
  const modal = $('modal'), mTitle = $('modalTitle'), mText = $('modalText'), mIcon = $('modalIcon');
  const mInput = $('modalInput'), mButtons = $('modalButtons'), mNote = $('modalNote');
  let closeModal = null;

  function openDialog({ title, text, icon, input, buttons, note }) {
    return new Promise((resolve) => {
      mTitle.textContent = title;
      mText.textContent = text;
      mNote.textContent = note || '';
      mNote.classList.toggle('hidden', !note);
      mIcon.textContent = icon || '';
      mInput.classList.toggle('hidden', input === undefined);
      if (input !== undefined) mInput.value = input;
      mButtons.innerHTML = '';
      const finish = (val) => {
        modal.classList.add('hidden');
        document.removeEventListener('keydown', onKey, true);
        closeModal = null;
        resolve(val);
      };
      buttons.forEach((b) => {
        const el = document.createElement('button');
        el.textContent = b.label;
        if (b.primary) el.className = 'primary';
        el.addEventListener('click', () => finish(b.value()));
        mButtons.appendChild(el);
      });
      const onKey = (e) => {
        if (e.key === 'Enter') { e.preventDefault(); finish(buttons.find((b) => b.primary).value()); }
        if (e.key === 'Escape') { e.preventDefault(); finish((buttons.find((b) => b.cancel) || buttons[buttons.length - 1]).value(true)); }
      };
      document.addEventListener('keydown', onKey, true);
      closeModal = finish;
      modal.classList.remove('hidden');
      if (input !== undefined) { mInput.focus(); mInput.select(); } else mButtons.querySelector('button').focus();
    });
  }

  const BUTTON_SETS = {
    0: [['OK', 1]], 1: [['OK', 1], ['Abbrechen', 2]], 2: [['Abbrechen', 3], ['Wiederholen', 4], ['Ignorieren', 5]],
    3: [['Ja', 6], ['Nein', 7], ['Abbrechen', 2]], 4: [['Ja', 6], ['Nein', 7]], 5: [['Wiederholen', 4], ['Abbrechen', 2]]
  };
  const ICONS = { 16: '⛔', 32: '❓', 48: '⚠️', 64: 'ℹ️' };

  const host = {
    print: (line) => out(line),
    msgBox: (prompt, buttons, title) => {
      const set = BUTTON_SETS[buttons & 7] || BUTTON_SETS[0];
      const def = (buttons & 768) >> 8;
      return openDialog({
        title, text: prompt, icon: ICONS[buttons & 112],
        buttons: set.map(([label, value], i) => ({ label, primary: i === Math.min(def, set.length - 1), value: () => value }))
      });
    },
    // Stop-Anweisung: anhalten wie an einem Haltepunkt
    pause: async (line) => {
      flushOut();
      markError(null);
      markStop(line);
      const r = await openDialog({
        title: 'Microsoft Visual Basic', icon: '⏸',
        text: `Die Ausführung wurde durch die Anweisung Stop in Zeile ${line} angehalten.`,
        note: 'In Access wäre der Code jetzt im Haltemodus. Hier kannst du fortsetzen oder beenden.',
        buttons: [
          { label: 'Beenden', cancel: true, value: () => 'end' },
          { label: 'Fortsetzen', primary: true, value: () => 'continue' }
        ]
      });
      markStop(null);
      return r || 'end';
    },
    // DoCmd.Close: Formular nach dem laufenden Ereignis schließen
    closeForm: () => { if (formSession) formSession.closeRequested = true; },
    inputBox: (prompt, title, deflt) => openDialog({
      title, text: prompt, input: deflt,
      buttons: [
        { label: 'OK', primary: true, value: () => mInput.value },
        { label: 'Abbrechen', value: () => null }
      ]
    })
  };

  // ---------- Ausführen ----------
  // Fehler anzeigen wie in Access: eigener Dialog mit dem Originaltext
  async function showError(e) {
    const where = e.line ? `Zeile ${e.line}` : '';
    const note = e.hint ? (where ? `Hinweis (${where}): ` : 'Hinweis: ') + e.hint : (where ? `Fehler in ${where}` : '');
    if (e.kind === 'syntax') {
      markError(e.line);
      await openDialog({
        title: 'Microsoft Visual Basic für Applikationen', icon: '⚠️',
        text: 'Fehler beim Kompilieren:\n\n' + e.accessMessage, note,
        buttons: [{ label: 'OK', primary: true, value: () => 'ok' }]
      });
      showPane('editorPane');
      return;
    }
    const interrupted = e.accessMessage === 'Ausführung des Codes wurde unterbrochen';
    const head = interrupted ? '' : (e.number ? `Laufzeitfehler '${e.number}':\n\n` : 'Laufzeitfehler:\n\n');
    const choice = await openDialog({
      title: 'Microsoft Visual Basic', icon: '',
      text: head + (e.accessMessage || e.message), note,
      buttons: [
        { label: 'Beenden', cancel: true, value: () => 'end' },
        { label: 'Debuggen', primary: true, value: () => 'debug' }
      ]
    });
    if (choice === 'debug') { markError(e.line); showPane('editorPane'); }
  }

  // ---------- Formular ----------
  let formModel = VBA.emptyForm();
  let formSession = null; // { runtime, run, stale }
  let designer = null;
  const hasForm = () => formModel.controls.length > 0;

  function isEventProc(name) {
    const l = name.toLowerCase();
    if (l === 'form_load') return true;
    return formModel.controls.some((c) => {
      const p = c.name.toLowerCase() + '_';
      return l.startsWith(p) && (VBA.CONTROL_EVENTS[c.kind] || []).some(([ev]) => l === p + ev.toLowerCase());
    });
  }

  function procRegex(name) {
    return new RegExp('^[ \\t]*(?:(?:Private|Public)[ \\t]+)?Sub[ \\t]+' + name + '[ \\t]*\\(', 'im');
  }

  /** Ereignisprozedur anlegen (falls nötig) und im Editor anspringen – wie [Ereignisprozedur] in Access. */
  function gotoEventProc(name) {
    let src = code.value;
    let m = procRegex(name).exec(src);
    if (!m) {
      if (src.length && !src.endsWith('\n')) src += '\n';
      src += `\nPrivate Sub ${name}()\n    \nEnd Sub\n`;
      code.value = src;
      code.dispatchEvent(new Event('input'));
      m = procRegex(name).exec(src);
    }
    const lineStart = m.index + m[0].length;
    const nextLine = code.value.indexOf('\n', lineStart) + 1;
    const lineEnd = code.value.indexOf('\n', nextLine);
    showPane('editorPane');
    code.focus();
    code.setSelectionRange(lineEnd, lineEnd);
    const line = code.value.slice(0, nextLine).split('\n').length;
    const lh = parseFloat(getComputedStyle(code).lineHeight) || 21;
    code.scrollTop = Math.max(0, (line - 4) * lh);
  }

  function setFormModel(m) {
    formModel = m && Array.isArray(m.controls) ? m : VBA.emptyForm();
    store.set(FORM_KEY, JSON.stringify(formModel));
    closeFormView();
    designer.setModel(formModel);
    refreshEntries();
  }

  const btnDesign = $('btnDesign'), btnFormView = $('btnFormView');
  function setViewButtons(view) {
    btnDesign.classList.toggle('active', !view);
    btnFormView.classList.toggle('active', view);
    btnDesign.setAttribute('aria-pressed', String(!view));
    btnFormView.setAttribute('aria-pressed', String(view));
    $('toolbox').classList.toggle('hidden', view);
  }

  function newSession(runtime) {
    const program = VBA.compile(code.value, { controls: formModel.controls });
    const r = VBA.createRun(program, host, {
      controls: runtime.adapters(),
      form: {
        getCaption: () => runtime.caption,
        setCaption: (t) => { runtime.caption = t; designer.setTitle(t); },
        refresh: () => {}
      }
    });
    r.open();
    return r;
  }

  // Änderungen an Steuerelementen gesammelt anzeigen (AddItem in langen Schleifen bleibt schnell)
  const pendingUpdates = new Set();
  let updateTimer = null;
  function scheduleUpdate(name) {
    pendingUpdates.add(name);
    if (updateTimer === null) {
      updateTimer = setTimeout(() => {
        updateTimer = null;
        const names = [...pendingUpdates];
        pendingUpdates.clear();
        for (const n of names) designer.updateControl(n);
      }, 0);
    }
  }

  /** Formularansicht öffnen: Code übersetzen, Sitzung anlegen, Form_Load ausführen. */
  async function openFormView() {
    if (formSession) return true;
    markError(null);
    const runtime = new VBA.FormRuntime(formModel, scheduleUpdate);
    runtime.onFocus = (name) => designer.focusControl(name);
    let r;
    try {
      r = newSession(runtime);
    } catch (e) {
      await showError(e);
      return false;
    }
    formSession = { runtime, run: r, stale: false };
    designer.setMode('view', runtime);
    setViewButtons(true);
    if (r.has('Form_Load')) await invoke('Form_Load');
    return true;
  }

  function closeFormView() {
    if (current) current.stop();
    formSession = null;
    if (designer) designer.setMode('design');
    setViewButtons(false);
  }

  /** Prozedur in der Formular-Sitzung ausführen (Ereignis oder Start-Sub). */
  async function invoke(procName) {
    if (!formSession) return;
    if (current) return; // während ein Ereignis läuft, werden weitere ignoriert
    if (formSession.stale) {
      try {
        formSession.run = newSession(formSession.runtime);
        formSession.stale = false;
      } catch (e) { await showError(e); return; }
    }
    const r = formSession.run;
    if (!r.has(procName)) return 'missing';
    current = r;
    btnRun.disabled = true; btnStop.disabled = false;
    try {
      await r.invoke(procName);
    } catch (e) {
      current = null;
      if (e.kind !== 'abort') await showError(e);
    } finally {
      current = null;
      btnRun.disabled = false; btnStop.disabled = true;
      if (formSession && formSession.closeRequested) { closeFormView(); showToast('Das Formular wurde mit DoCmd.Close geschlossen.'); }
    }
  }

  async function fireEvent(controlName, ev) {
    const res = await invoke(controlName + '_' + ev);
    const c = formModel.controls.find((x) => x.name === controlName);
    if (res === 'missing' && ev === 'Click' && c && c.kind === 'button') {
      showToast(`Zu „${controlName}“ gibt es noch keinen Code. Lege im Code „Private Sub ${controlName}_Click()“ an (im Entwurf: Doppelklick auf die Schaltfläche).`);
    }
  }

  // ---------- Seitenleiste: Reiter „Formular“ und „Hilfe“ ----------
  // Am PC liegen Formular und Hilfe rechts in einer gemeinsamen Spalte, damit der Code breit bleibt.
  // Auf dem Handy ist das Formular ein eigener Reiter und die Hilfe ein Vollbild.
  const btnForm = $('btnForm'), btnHelp = $('btnHelp'), help = $('helpPanel');
  const sideTabForm = $('sideTabForm'), sideTabHelp = $('sideTabHelp');
  const FORM_VISIBLE_KEY = 'vba-editor-form-visible';
  let sideTab = null;      // 'form' | 'help' | null (Seitenleiste zu)
  let helpReturn = null;   // 'form', wenn die Hilfe aus dem Formular heraus geöffnet wurde
  let formTouched = false; // Formular-Reiter seit dem Öffnen der Seitenleiste schon gezeigt?

  function showSide(tab) {
    sideTab = tab;
    document.body.classList.toggle('side-closed', !tab);
    document.body.classList.toggle('form-hidden', tab !== 'form');
    help.classList.toggle('hidden', tab !== 'help');
    btnForm.setAttribute('aria-pressed', String(tab === 'form'));
    btnHelp.setAttribute('aria-pressed', String(tab === 'help'));
    sideTabForm.setAttribute('aria-selected', String(tab === 'form'));
    sideTabHelp.setAttribute('aria-selected', String(tab === 'help'));
    store.set(FORM_VISIBLE_KEY, tab === 'form' || (tab === 'help' && helpReturn === 'form') ? '1' : '0');
    if (!tab && document.querySelector('.tab.active[data-pane="formPane"]')) showPane('editorPane');
  }
  function closeSide() {
    helpReturn = null;
    formTouched = false;
    showSide(null);
    closeFormView();
  }
  /** Formular-Reiter zeigen. Beim ersten Zeigen (oder mit forceView) gleich in der Formularansicht,
   *  damit Schaltflächen sofort reagieren; beim Zurückwechseln aus der Hilfe bleibt die Ansicht, wie sie war. */
  async function openFormTab(forceView) {
    const fresh = forceView || !formTouched;
    helpReturn = null;
    formTouched = true;
    showSide('form');
    if (fresh && hasForm()) await openFormView();
  }
  async function setFormVisible(show) {
    if (show) await openFormTab(true); else closeSide();
  }
  btnForm.addEventListener('click', () => {
    if (sideTab === 'form') { closeSide(); return; }
    openFormTab(false);
    showPane('formPane');
  });
  sideTabForm.addEventListener('click', () => { if (sideTab !== 'form') openFormTab(false); });
  sideTabHelp.addEventListener('click', () => toggleHelp(true));
  $('btnSideClose').addEventListener('click', closeSide);

  // ---------- Trenner zwischen Code und Seitenleiste ----------
  const SIDE_W_KEY = 'vba-editor-side-width';
  const sidebar = $('sidebar'), sideSplitter = $('sideSplitter');
  function setSideWidth(px, save) {
    const w = Math.round(Math.max(280, px));
    sidebar.style.flexBasis = w + 'px'; // Grenzen (min/max) setzt das CSS
    if (save) store.set(SIDE_W_KEY, String(w));
  }
  const savedSideW = parseFloat(store.get(SIDE_W_KEY));
  if (isFinite(savedSideW) && savedSideW > 0) setSideWidth(savedSideW, false);
  sideSplitter.addEventListener('pointerdown', (e) => {
    e.preventDefault();
    sideSplitter.setPointerCapture(e.pointerId);
    const right = sidebar.getBoundingClientRect().right;
    const move = (ev) => setSideWidth(right - ev.clientX - sideSplitter.offsetWidth / 2, false);
    const up = () => {
      sideSplitter.removeEventListener('pointermove', move);
      sideSplitter.removeEventListener('pointerup', up);
      sideSplitter.removeEventListener('pointercancel', up);
      setSideWidth(sidebar.getBoundingClientRect().width, true);
    };
    sideSplitter.addEventListener('pointermove', move);
    sideSplitter.addEventListener('pointerup', up);
    sideSplitter.addEventListener('pointercancel', up);
  });
  sideSplitter.addEventListener('keydown', (e) => {
    const d = { ArrowLeft: 40, ArrowRight: -40 }[e.key];
    if (d) { e.preventDefault(); setSideWidth(sidebar.getBoundingClientRect().width + d, true); }
  });
  sideSplitter.addEventListener('dblclick', () => { sidebar.style.flexBasis = ''; store.set(SIDE_W_KEY, ''); });

  // ---------- Ausführen ----------
  async function run() {
    markError(null);
    if (hasForm()) {
      // Mit Formular: Formular einblenden, Formularansicht öffnen und ggf. die Start-Sub darin ausführen
      if (sideTab !== 'form') { helpReturn = null; formTouched = true; showSide('form'); }
      const ok = await openFormView();
      if (!ok) return;
      if (entry.value) await invoke(entry.value);
      return;
    }
    let program;
    try {
      program = VBA.compile(code.value);
    } catch (e) {
      await showError(e);
      return;
    }
    refreshEntries();
    const r = VBA.createRun(program, host, {});
    current = r;
    btnRun.disabled = true; btnStop.disabled = false;
    try {
      await r.start(entry.value || undefined);
    } catch (e) {
      current = null;
      if (e.kind !== 'abort') await showError(e);
    } finally {
      current = null;
      btnRun.disabled = false; btnStop.disabled = true;
    }
  }

  btnRun.addEventListener('click', () => { showPane(hasForm() ? 'formPane' : 'consolePane'); run(); });
  btnDesign.addEventListener('click', () => closeFormView());
  btnFormView.addEventListener('click', () => openFormView());
  btnStop.addEventListener('click', () => {
    if (current) current.stop();
    if (closeModal) closeModal(null);
  });
  $('btnClear').addEventListener('click', clearConsole);
  // Menü „Datei“ (Neu, Öffnen, Speichern)
  const btnMore = $('btnMore');
  function setMore(open) { document.body.classList.toggle('more-open', open); btnMore.setAttribute('aria-expanded', String(open)); }
  btnMore.addEventListener('click', (e) => { e.stopPropagation(); setMore(!document.body.classList.contains('more-open')); });
  document.addEventListener('click', (e) => { if (!e.target.closest('#btnMore')) setMore(false); });
  $('moreMenu').addEventListener('focusout', (e) => { if (!e.relatedTarget || !e.relatedTarget.closest('.file-group')) setMore(false); });
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape') setMore(false); });
  // ---------- Schriftgröße (A− / A+) für Code und Direktfenster, z. B. für den Beamer ----------
  const FONT_KEY = 'vba-editor-font-size';
  const FONT_MIN = 10, FONT_MAX = 32;
  function codeFontSize() { return parseFloat(getComputedStyle(code).fontSize) || 14; }
  function setCodeFont(px, save) {
    const v = Math.max(FONT_MIN, Math.min(FONT_MAX, Math.round(px)));
    document.documentElement.style.setProperty('--code-size', v + 'px');
    if (save) store.set(FONT_KEY, String(v));
    document.querySelectorAll('[data-font]').forEach((b) => {
      b.disabled = b.dataset.font < 0 ? v <= FONT_MIN : v >= FONT_MAX;
    });
    renderGutter();
    return v;
  }
  document.querySelectorAll('[data-font]').forEach((b) => b.addEventListener('click', (e) => {
    e.stopPropagation(); // Datei-Menü bleibt offen, man kann mehrmals tippen
    setCodeFont(codeFontSize() + 2 * Math.sign(Number(b.dataset.font)), true);
  }));
  const savedFont = parseFloat(store.get(FONT_KEY));
  if (isFinite(savedFont) && savedFont > 0) setCodeFont(savedFont, false);

  // ---------- Zeichenleiste über der Bildschirmtastatur (Handy/Tablet) ----------
  // Auf der Handytastatur liegen " ( ) = & _ ' auf der zweiten Ebene. Die Leiste erscheint, solange
  // der Code-Editor auf einem Touch-Gerät den Fokus hat, und schlägt das passende Schlusswort vor.
  const keyBar = $('keyBar');
  const touchMq = window.matchMedia ? window.matchMedia('(pointer: coarse)') : { matches: false };
  const KEYS = [
    { label: '⇥', text: '    ', title: 'Einrücken (Tab)' },
    { label: '"' }, { label: '(' }, { label: ')' }, { label: '=' }, { label: '&' }, { label: '_' }, { label: "'" },
    { label: ',' }, { label: '.' }, { label: '<' }, { label: '>' }, { label: '+' }, { label: '-' }, { label: '*' }, { label: '/' }, { label: ':' },
    { label: 'Debug.Print', text: 'Debug.Print ', word: true }, { label: 'Dim', text: 'Dim ', word: true },
    { label: 'As', text: 'As ', word: true }, { label: 'Then', text: 'Then', word: true }
  ];
  const closeBtn = document.createElement('button');
  closeBtn.type = 'button';
  closeBtn.className = 'kb-word kb-close hidden';
  keyBar.appendChild(closeBtn);
  for (const k of KEYS) {
    const b = document.createElement('button');
    b.type = 'button';
    b.textContent = k.label;
    if (k.word) b.className = 'kb-word';
    b.title = k.title || `„${k.text || k.label}“ einfügen`;
    b.dataset.text = k.text || k.label;
    keyBar.appendChild(b);
  }
  // Fokus im Editor lassen, damit die Tastatur offen bleibt
  keyBar.addEventListener('pointerdown', (e) => { if (e.target.closest('button')) e.preventDefault(); });
  keyBar.addEventListener('mousedown', (e) => e.preventDefault());
  keyBar.addEventListener('click', (e) => {
    const b = e.target.closest('button');
    if (!b) return;
    if (b === closeBtn) insertCloser(); else insertText(b.dataset.text);
    code.focus();
  });
  function insertText(t) {
    code.setRangeText(t, code.selectionStart, code.selectionEnd, 'end');
    code.dispatchEvent(new Event('input'));
  }
  /** Offene Blöcke (If, For, Do, …) bis zum Cursor bestimmen und den innersten liefern. */
  function openBlock() {
    const before = code.value.slice(0, code.selectionStart);
    const lines = before.split('\n'); // inkl. aktueller Zeile bis zum Cursor (gerade eingefügtes End If zählt schon)
    const stack = [];
    const RULES = [
      [/^(?:(?:public|private|static)\s+)*sub\s+\w/, 'End Sub'], [/^(?:(?:public|private|static)\s+)*function\s+\w/, 'End Function'],
      [/^if\b.*\bthen$/, 'End If'], [/^for\b/, 'Next'], [/^do\b/, 'Loop'], [/^while\b/, 'Wend'], [/^select\s+case\b/, 'End Select']
    ];
    const CLOSE = [[/^end\s+sub\b/, 'End Sub'], [/^end\s+function\b/, 'End Function'], [/^end\s+if\b/, 'End If'], [/^next\b/, 'Next'],
      [/^loop\b/, 'Loop'], [/^wend\b/, 'Wend'], [/^end\s+select\b/, 'End Select']];
    for (const raw of lines) {
      const indent = /^[ \t]*/.exec(raw)[0];
      const l = raw.replace(/"[^"]*"/g, '""').replace(/'.*$/, '').trim().toLowerCase();
      if (!l || /^rem\b/.test(l)) continue;
      const c = CLOSE.find(([re]) => re.test(l));
      if (c) {
        const i = stack.map((x) => x.word).lastIndexOf(c[1]);
        if (i >= 0) stack.length = i;
        continue;
      }
      const o = RULES.find(([re]) => re.test(l));
      if (o) stack.push({ word: o[1], indent });
    }
    return stack[stack.length - 1] || null;
  }
  function updateCloser() {
    const b = openBlock();
    closeBtn.classList.toggle('hidden', !b);
    if (b) { closeBtn.textContent = b.word; closeBtn.title = `„${b.word}“ einfügen (schließt den offenen Block)`; }
  }
  /** Schlusswort mit der Einrückung des öffnenden Blocks einfügen, wie im VBA-Editor üblich. */
  function insertCloser() {
    const b = openBlock();
    if (!b) return;
    const s = code.selectionStart;
    const lineStart = code.value.lastIndexOf('\n', s - 1) + 1;
    const head = code.value.slice(lineStart, s);
    if (/^[ \t]*$/.test(head)) code.setRangeText(b.indent + b.word, lineStart, code.selectionEnd, 'end');
    else code.setRangeText('\n' + b.indent + b.word, s, code.selectionEnd, 'end');
    code.dispatchEvent(new Event('input'));
  }
  function setKeyBar(on) {
    keyBar.classList.toggle('hidden', !on);
    document.body.classList.toggle('keybar-on', on);
    if (on) updateCloser();
  }
  code.addEventListener('focus', () => { if (touchMq.matches) setKeyBar(true); });
  code.addEventListener('blur', () => setTimeout(() => { if (document.activeElement !== code) setKeyBar(false); }, 150));
  code.addEventListener('input', () => { if (!keyBar.classList.contains('hidden')) updateCloser(); });
  document.addEventListener('selectionchange', () => { if (document.activeElement === code && !keyBar.classList.contains('hidden')) updateCloser(); });
  // Höhe der Bildschirmtastatur, damit Leiste und ▶-Knopf darüber stehen (iOS verkleinert nur den sichtbaren Bereich)
  if (window.visualViewport) {
    const vv = window.visualViewport;
    const updateKb = () => {
      const kb = Math.max(0, Math.round(window.innerHeight - vv.height - vv.offsetTop));
      document.documentElement.style.setProperty('--kb', kb + 'px');
    };
    vv.addEventListener('resize', updateKb);
    vv.addEventListener('scroll', updateKb);
    updateKb();
  }

  // ---------- Code ersetzen mit Rückgängig ----------
  const toast = $('toast'), toastText = $('toastText');
  let undoState = null, toastTimer = null;
  function snapshot() { return { code: code.value, form: JSON.parse(JSON.stringify(formModel)) }; }
  /** Code (und ggf. Formular) ersetzen; der alte Stand kann per „Rückgängig“ zurückgeholt werden. */
  function replaceProject(newCode, newForm, message) {
    const old = snapshot();
    code.value = newCode;
    if (newForm !== undefined) {
      setFormModel(newForm ? JSON.parse(JSON.stringify(newForm)) : VBA.emptyForm());
      if (newForm && newForm.controls && newForm.controls.length) setFormVisible(true);
    }
    code.dispatchEvent(new Event('input'));
    showPane('editorPane');
    if ((old.code.trim() && old.code !== newCode) || old.form.controls.length) {
      undoState = old;
      $('toastUndo').classList.remove('hidden');
      toastText.textContent = message;
      toast.classList.remove('hidden');
      clearTimeout(toastTimer);
      toastTimer = setTimeout(() => toast.classList.add('hidden'), 10000);
    }
  }
  $('toastUndo').addEventListener('click', () => {
    if (!undoState) return;
    const u = undoState;
    undoState = null;
    code.value = u.code;
    setFormModel(u.form);
    code.dispatchEvent(new Event('input'));
    toast.classList.add('hidden');
    showPane('editorPane');
  });
  $('btnNew').addEventListener('click', () => replaceProject(SKELETON, null, 'Neues Programm (Grundgerüst, leeres Formular).'));

  // ---------- Speichern / Öffnen (eine .bas-Datei mit Code und Formular) ----------
  // In der claude.ai-Vorschau gehen Downloads nur über die Plattform (als .txt); lokal normal als .bas
  let platformDownloads = null;
  if (window.claude && typeof window.claude.use === 'function') {
    window.claude.use('downloads').then((d) => { platformDownloads = d; }).catch(() => {});
  }
  function showToast(message) {
    undoState = null;
    toastText.textContent = message;
    $('toastUndo').classList.add('hidden');
    toast.classList.remove('hidden');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { toast.classList.add('hidden'); $('toastUndo').classList.remove('hidden'); }, 6000);
  }
  $('btnSave').addEventListener('click', () => {
    const text = VBA.serializeProject(code.value, formModel);
    const base = (formModel.caption || 'Programm').replace(/[^\wÄÖÜäöüß-]+/g, '_');
    if (platformDownloads) {
      platformDownloads.save({ filename: base + '.txt', data: text })
        .then(() => showToast(`Gespeichert als ${base}.txt (lässt sich wieder öffnen).`))
        .catch((err) => { if (!err || err.code !== 'declined') showToast('Speichern ist hier nicht möglich.'); });
      return;
    }
    const blob = new Blob([text], { type: 'text/plain;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = base + '.bas';
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  });
  const fileInput = $('fileInput');
  $('btnOpen').addEventListener('click', () => fileInput.click());
  fileInput.addEventListener('change', () => {
    const f = fileInput.files && fileInput.files[0];
    if (!f) return;
    const reader = new FileReader();
    reader.onload = () => {
      const { code: src, form } = VBA.deserializeProject(String(reader.result));
      replaceProject(src, form, `Datei „${f.name}“ geöffnet.`);
      fileInput.value = '';
    };
    reader.readAsText(f, 'utf-8');
  });

  // ---------- Hilfe ----------
  const helpBody = $('helpBody'), helpToc = $('helpToc');
  function buildHelp() {
    for (const t of VBAHelp.TOPICS) {
      const sec = document.createElement('section');
      sec.id = 'hilfe-' + t.id;
      const h = document.createElement('h3');
      h.textContent = t.title;
      sec.appendChild(h);
      const div = document.createElement('div');
      div.innerHTML = t.html;
      sec.appendChild(div);
      if (t.code) {
        const pre = document.createElement('pre');
        pre.textContent = t.code;
        sec.appendChild(pre);
        const btn = document.createElement('button');
        btn.textContent = 'In Editor übernehmen';
        btn.addEventListener('click', () => {
          replaceProject(t.code, t.form || undefined, `Beispiel „${t.title}“ übernommen.`);
          if (window.matchMedia('(max-width: 760px)').matches) toggleHelp(false);
        });
        sec.appendChild(btn);
      }
      helpBody.appendChild(sec);
      const link = document.createElement('button');
      link.textContent = t.title;
      link.addEventListener('click', () => sec.scrollIntoView({ behavior: 'smooth', block: 'start' }));
      helpToc.appendChild(link);
    }
  }
  /** Hilfe-Reiter öffnen oder schließen. Kam man aus dem Formular, geht es beim Schließen dorthin zurück. */
  function toggleHelp(show) {
    const open = show === undefined ? sideTab !== 'help' : show;
    if (open) {
      if (sideTab === 'help') return;
      helpReturn = sideTab === 'form' ? 'form' : null;
      if (!helpReturn) formTouched = false;
      showSide('help');
    } else {
      if (sideTab !== 'help') return;
      if (helpReturn === 'form') { helpReturn = null; showSide('form'); } else closeSide();
    }
  }
  btnHelp.addEventListener('click', () => toggleHelp());
  $('btnHelpClose').addEventListener('click', () => toggleHelp(false));
  document.addEventListener('keydown', (e) => {
    if (e.key === 'F1') { e.preventDefault(); toggleHelp(); }
    if (e.key === 'Escape' && sideTab === 'help' && modal.classList.contains('hidden') && !document.body.classList.contains('more-open')) toggleHelp(false);
  });
  buildHelp();

  // ---------- Start ----------
  designer = new FormDesigner({
    canvas: $('formCanvas'),
    props: $('formProps'),
    title: $('formTitle'), stage: $('formStage'), fit: $('formFit'), scroller: $('formScroll'),
    onChange: (m) => { store.set(FORM_KEY, JSON.stringify(m)); refreshEntries(); },
    onEvent: (name, ev) => fireEvent(name, ev),
    onEventProc: (name) => gotoEventProc(name),
    hasProc: (name) => procRegex(name).test(code.value)
  });
  document.querySelectorAll('#toolbox [data-add]').forEach((b) => {
    b.addEventListener('click', () => { closeFormView(); designer.add(b.dataset.add, b.dataset.variant); });
  });

  const saved = store.get(STORAGE_KEY);
  code.value = saved && saved.trim() ? saved : SKELETON;
  let savedForm = null;
  try { savedForm = JSON.parse(store.get(FORM_KEY) || 'null'); } catch (e) { savedForm = null; }
  formModel = savedForm && Array.isArray(savedForm.controls) ? savedForm : VBA.emptyForm();
  designer.setModel(formModel);
  setViewButtons(false);
  const formWasVisible = store.get(FORM_VISIBLE_KEY) === '1';
  showSide(null);
  renderGutter();
  refreshEntries();
  if (formWasVisible) setFormVisible(true);
})();
