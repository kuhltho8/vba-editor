/*
 * designer.js – Formular-Designer (Entwurfsansicht) und Formularansicht.
 *
 * Entwurf: Steuerelemente anlegen, verschieben (Maus/Touch), Größe ändern, benennen.
 * Formularansicht: echte Eingabefelder/Buttons; Ereignisse gehen an onEvent(name, ereignis).
 */
(function (root) {
  'use strict';
  const NAME_RE = /^[A-Za-zÄÖÜäöüß][A-Za-z0-9_ÄÖÜäöüß]*$/;
  const GRID = 4;
  const snap = (v) => Math.round(v / GRID) * GRID;
  const el = (tag, cls, text) => {
    const e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text !== undefined) e.textContent = text;
    return e;
  };

  const DEFAULTS = {
    textbox: { w: 140, h: 26 },
    label: { w: 160, h: 22 },
    button: { w: 130, h: 32 },
    optiongroup: { w: 170, h: 0 }
  };

  class FormDesigner {
    constructor(o) {
      this.canvas = o.canvas;
      this.props = o.props;
      this.onChange = o.onChange || (() => {});
      this.onEvent = o.onEvent || (() => {});
      this.onEventProc = o.onEventProc || (() => {});
      this.hasProc = o.hasProc || (() => false);
      this.model = VBA.emptyForm();
      this.mode = 'design';
      this.selected = null;
      this.runtime = null;
      this.elements = new Map(); // lower name → DOM-Element
      this.canvas.addEventListener('pointerdown', (e) => {
        if (this.mode === 'design' && e.target === this.canvas) this.select(null);
      });
      this.canvas.tabIndex = 0;
      this.canvas.addEventListener('keydown', (e) => {
        if (this.mode !== 'design' || !this.selected) return;
        if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); this.remove(this.selected); return; }
        const d = { ArrowLeft: [-GRID, 0], ArrowRight: [GRID, 0], ArrowUp: [0, -GRID], ArrowDown: [0, GRID] }[e.key];
        if (d) {
          e.preventDefault();
          const c = this.find(this.selected);
          c.x = Math.max(0, c.x + d[0]); c.y = Math.max(0, c.y + d[1]);
          this.changed(); this.render();
        }
      });
    }

    find(name) { return this.model.controls.find((c) => c.name.toLowerCase() === String(name).toLowerCase()) || null; }

    setModel(m) {
      this.model = m && Array.isArray(m.controls) ? m : VBA.emptyForm();
      this.selected = null;
      this.render();
    }

    changed() { this.onChange(this.model); }

    // ---------- Steuerelemente anlegen ----------
    uniqueName(base) {
      let i = 1;
      while (this.find(base + i)) i++;
      return base + i;
    }

    add(kind, variant) {
      const bases = { textbox: variant === 'output' ? 'A_Ausgabe' : 'E_Eingabe', label: 'lblText', button: 'btnBefehl', optiongroup: 'E_Auswahl' };
      const name = this.uniqueName(bases[kind]);
      const bottom = this.model.controls.reduce((m, c) => Math.max(m, c.y + c.h), 0);
      const c = {
        kind, name,
        caption: kind === 'textbox' ? (variant === 'output' ? 'Ergebnis:' : 'Eingabe:')
          : kind === 'label' ? 'Bezeichnung' : kind === 'button' ? 'Schaltfläche' : 'Auswahl',
        x: kind === 'textbox' ? 104 : 16, y: snap(bottom + 12), w: DEFAULTS[kind].w, h: DEFAULTS[kind].h
      };
      if (kind === 'optiongroup') {
        c.options = [{ caption: 'Option 1', value: 1 }, { caption: 'Option 2', value: 2 }];
        c.h = this.groupHeight(c);
      }
      if (c.y + c.h > this.model.height) this.model.height = snap(c.y + c.h + 16);
      this.model.controls.push(c);
      this.selected = c.name;
      this.changed();
      this.render();
    }

    groupHeight(c) { return 30 + 26 * (c.options || []).length; }

    remove(name) {
      this.model.controls = this.model.controls.filter((c) => c.name.toLowerCase() !== name.toLowerCase());
      this.selected = null;
      this.changed();
      this.render();
    }

    select(name) {
      this.selected = name;
      if (this.mode === 'design') {
        for (const [l, e] of this.elements) e.classList.toggle('selected', !!name && l === name.toLowerCase());
        this.renderProps();
      }
    }

    // ---------- Darstellung ----------
    setMode(mode, runtime) {
      this.mode = mode;
      this.runtime = runtime || null;
      this.render();
    }

    render() {
      const m = this.model;
      this.canvas.innerHTML = '';
      this.elements.clear();
      this.canvas.style.width = m.width + 'px';
      this.canvas.style.height = m.height + 'px';
      this.canvas.classList.toggle('design', this.mode === 'design');
      this.canvas.classList.toggle('view', this.mode === 'view');
      if (m.controls.length === 0 && this.mode === 'design') {
        const hint = el('div', 'form-empty');
        hint.innerHTML = 'Noch leer. Lege oben ein <b>Eingabefeld (E_)</b>, ein <b>Ausgabefeld (A_)</b> oder eine <b>Schaltfläche</b> an.';
        this.canvas.appendChild(hint);
      }
      for (const c of m.controls) {
        const e = this.mode === 'design' ? this.renderDesign(c) : this.renderView(c);
        e.style.left = c.x + 'px';
        e.style.top = c.y + 'px';
        e.style.width = c.w + 'px';
        e.style.height = c.h + 'px';
        this.canvas.appendChild(e);
        this.elements.set(c.name.toLowerCase(), e);
        if (this.mode === 'view') this.updateControl(c.name);
      }
      if (this.mode === 'design') this.select(this.selected && this.find(this.selected) ? this.selected : null);
      else this.props.innerHTML = '';
    }

    inner(c, interactive) {
      switch (c.kind) {
        case 'textbox': {
          const box = el('div', 'ctl-box');
          if (c.caption) box.appendChild(el('span', 'ctl-caption', c.caption));
          const inp = el('input', 'ctl-input');
          inp.type = 'text';
          inp.setAttribute('aria-label', c.caption || c.name);
          if (!interactive) { inp.readOnly = true; inp.tabIndex = -1; inp.value = c.name; }
          box.appendChild(inp);
          return box;
        }
        case 'label': return el('div', 'ctl-label', c.caption);
        case 'button': {
          const b = el('button', 'ctl-button', c.caption);
          b.type = 'button';
          if (!interactive) b.tabIndex = -1;
          return b;
        }
        case 'optiongroup': {
          const fs = el('fieldset', 'ctl-group');
          fs.appendChild(el('legend', '', c.caption));
          (c.options || []).forEach((o) => {
            const lab = el('label', 'ctl-option');
            const r = el('input');
            r.type = 'radio';
            r.name = 'grp_' + c.name;
            r.value = String(o.value);
            if (!interactive) r.tabIndex = -1;
            lab.appendChild(r);
            lab.appendChild(document.createTextNode(' ' + o.caption));
            fs.appendChild(lab);
          });
          return fs;
        }
        default: return el('div');
      }
    }

    renderDesign(c) {
      const wrap = el('div', 'ctl ctl-' + c.kind);
      wrap.appendChild(this.inner(c, false));
      const tag = el('span', 'ctl-name', c.name);
      wrap.appendChild(tag);
      const handle = el('span', 'ctl-resize');
      wrap.appendChild(handle);

      // Doppelklick: zur Ereignisprozedur springen (wie in Access)
      wrap.addEventListener('dblclick', () => {
        const ev = { button: 'Click', label: 'Click', textbox: 'AfterUpdate', optiongroup: 'AfterUpdate' }[c.kind];
        this.onEventProc(c.name + '_' + ev);
      });
      // Verschieben
      wrap.addEventListener('pointerdown', (e) => {
        if (e.target === handle) return;
        e.preventDefault();
        this.select(c.name);
        this.canvas.focus({ preventScroll: true });
        const sx = e.clientX, sy = e.clientY, ox = c.x, oy = c.y;
        let moved = false;
        wrap.setPointerCapture(e.pointerId);
        const move = (ev) => {
          const nx = Math.max(0, Math.min(this.model.width - 20, snap(ox + ev.clientX - sx)));
          const ny = Math.max(0, Math.min(this.model.height - 10, snap(oy + ev.clientY - sy)));
          if (nx !== c.x || ny !== c.y) { moved = true; c.x = nx; c.y = ny; wrap.style.left = nx + 'px'; wrap.style.top = ny + 'px'; }
        };
        const up = () => {
          wrap.removeEventListener('pointermove', move);
          wrap.removeEventListener('pointerup', up);
          wrap.removeEventListener('pointercancel', up);
          if (moved) { this.changed(); this.renderProps(); }
        };
        wrap.addEventListener('pointermove', move);
        wrap.addEventListener('pointerup', up);
        wrap.addEventListener('pointercancel', up);
      });
      // Größe ändern
      handle.addEventListener('pointerdown', (e) => {
        e.preventDefault(); e.stopPropagation();
        this.select(c.name);
        const sx = e.clientX, sy = e.clientY, ow = c.w, oh = c.h;
        handle.setPointerCapture(e.pointerId);
        const move = (ev) => {
          c.w = Math.max(40, snap(ow + ev.clientX - sx));
          c.h = Math.max(c.kind === 'optiongroup' ? this.groupHeight(c) : 18, snap(oh + ev.clientY - sy));
          wrap.style.width = c.w + 'px'; wrap.style.height = c.h + 'px';
        };
        const up = () => {
          handle.removeEventListener('pointermove', move);
          handle.removeEventListener('pointerup', up);
          this.changed(); this.renderProps();
        };
        handle.addEventListener('pointermove', move);
        handle.addEventListener('pointerup', up);
      });
      return wrap;
    }

    renderView(c) {
      const wrap = el('div', 'ctl ctl-' + c.kind);
      const content = this.inner(c, true);
      wrap.appendChild(content);
      const rt = this.runtime;
      if (c.kind === 'textbox') {
        const inp = content.querySelector('input');
        let before = '';
        inp.addEventListener('focus', () => { before = inp.value; });
        inp.addEventListener('input', () => {
          rt.userInput(c.name, inp.value);
          this.onEvent(c.name, 'Change');
        });
        inp.addEventListener('keydown', (e) => {
          if (e.key === 'Enter') { e.preventDefault(); this.focusNext(c.name); }
        });
        inp.addEventListener('blur', () => {
          if (inp.value !== before) { before = inp.value; this.onEvent(c.name, 'AfterUpdate'); }
        });
      } else if (c.kind === 'button') {
        content.addEventListener('click', () => this.onEvent(c.name, 'Click'));
      } else if (c.kind === 'label') {
        content.addEventListener('click', () => this.onEvent(c.name, 'Click'));
      } else if (c.kind === 'optiongroup') {
        content.querySelectorAll('input').forEach((r) => {
          r.addEventListener('change', () => {
            if (!r.checked) return;
            rt.userSelect(c.name, Number(r.value));
            this.onEvent(c.name, 'Click').then(() => this.onEvent(c.name, 'AfterUpdate'));
          });
        });
      }
      return wrap;
    }

    focusNext(name) {
      const order = this.model.controls.filter((c) => c.kind !== 'label').sort((a, b) => a.y - b.y || a.x - b.x);
      const i = order.findIndex((c) => c.name === name);
      const nxt = order[(i + 1) % order.length];
      if (nxt) this.focusControl(nxt.name);
    }

    focusControl(name) {
      const e = this.elements.get(String(name).toLowerCase());
      if (!e) return;
      const f = e.querySelector('input:checked') || e.querySelector('input, button');
      if (f) f.focus();
    }

    /** Formularansicht: Zustand eines Steuerelements aus der Laufzeit übernehmen. */
    updateControl(name) {
      if (this.mode !== 'view' || !this.runtime) return;
      const s = this.runtime.get(name);
      const e = this.elements.get(String(name).toLowerCase());
      if (!s || !e) return;
      e.style.visibility = s.visible ? '' : 'hidden';
      const kind = s.def.kind;
      const css = (v) => (v === null ? '' : VBA.vbaColorToCss(v));
      if (kind === 'textbox') {
        const inp = e.querySelector('input');
        const t = this.runtime.displayText(name);
        if (inp.value !== t) inp.value = t;
        inp.disabled = !s.enabled;
        inp.readOnly = s.locked;
        inp.style.color = css(s.foreColor);
        inp.style.backgroundColor = css(s.backColor);
      } else if (kind === 'label') {
        const d = e.firstChild;
        d.textContent = s.caption;
        d.style.color = css(s.foreColor);
        d.style.backgroundColor = css(s.backColor);
      } else if (kind === 'button') {
        const b = e.firstChild;
        b.textContent = s.caption;
        b.disabled = !s.enabled;
        b.style.color = css(s.foreColor);
        b.style.backgroundColor = css(s.backColor);
      } else if (kind === 'optiongroup') {
        e.querySelectorAll('input').forEach((r) => {
          r.checked = typeof s.value === 'number' && Number(r.value) === s.value;
          r.disabled = !s.enabled;
        });
      }
    }

    // ---------- Eigenschaften ----------
    renderProps() {
      const p = this.props;
      p.innerHTML = '';
      if (this.mode !== 'design') return;
      const c = this.selected ? this.find(this.selected) : null;
      const row = (label, input) => {
        const r = el('label', 'prop-row');
        r.appendChild(el('span', 'prop-label', label));
        r.appendChild(input);
        p.appendChild(r);
        return input;
      };
      const text = (value, onCommit, id) => {
        const i = el('input');
        i.type = 'text'; i.value = value; i.id = id;
        i.addEventListener('change', () => onCommit(i.value, i));
        return i;
      };
      const num = (value, onCommit, id) => {
        const i = el('input');
        i.type = 'number'; i.value = value; i.id = id; i.step = GRID;
        i.addEventListener('change', () => onCommit(Math.max(0, Math.round(Number(i.value) || 0))));
        return i;
      };

      if (!c) {
        p.appendChild(el('div', 'prop-title', 'Formular'));
        row('Breite', num(this.model.width, (v) => { this.model.width = Math.max(200, v); this.changed(); this.render(); }, 'propFormW'));
        row('Höhe', num(this.model.height, (v) => { this.model.height = Math.max(120, v); this.changed(); this.render(); }, 'propFormH'));
        const ev = el('div', 'prop-events');
        ev.appendChild(this.eventButton('Form_Load', 'Beim Laden'));
        p.appendChild(ev);
        p.appendChild(el('div', 'prop-hint', 'Tipp: Ein Steuerelement anklicken, um es zu bearbeiten. Verschieben mit der Maus, dem Finger oder den Pfeiltasten.'));
        return;
      }
      p.appendChild(el('div', 'prop-title', VBA.KIND_NAMES[c.kind]));
      const err = el('div', 'prop-error');
      const nameInput = row('Name', text(c.name, (v, i) => {
        const msg = this.checkName(v, c);
        err.textContent = msg || '';
        if (msg) { i.value = c.name; return; }
        c.name = v; this.selected = v; this.changed(); this.render();
      }, 'propName'));
      void nameInput;
      p.appendChild(err);
      const capLabel = c.kind === 'textbox' ? 'Beschriftung' : c.kind === 'optiongroup' ? 'Überschrift' : 'Beschriftung (Caption)';
      row(capLabel, text(c.caption || '', (v) => { c.caption = v; this.changed(); this.render(); }, 'propCaption'));
      if (c.kind === 'optiongroup') {
        const ta = el('textarea');
        ta.id = 'propOptions';
        ta.rows = Math.max(2, c.options.length);
        ta.value = c.options.map((o) => o.caption).join('\n');
        ta.addEventListener('change', () => {
          const lines = ta.value.split('\n').map((s) => s.trim()).filter(Boolean);
          c.options = lines.map((cap, i) => ({ caption: cap, value: i + 1 }));
          c.h = Math.max(c.h, this.groupHeight(c));
          this.changed(); this.render();
        });
        row('Optionen (eine pro Zeile, Wert 1, 2, 3 …)', ta);
      }
      const pos = el('div', 'prop-grid');
      [['X', 'x'], ['Y', 'y'], ['Breite', 'w'], ['Höhe', 'h']].forEach(([lab, k]) => {
        const r = el('label', 'prop-row small');
        r.appendChild(el('span', 'prop-label', lab));
        r.appendChild(num(c[k], (v) => { c[k] = k === 'w' || k === 'h' ? Math.max(18, v) : v; this.changed(); this.render(); }, 'prop' + k));
        pos.appendChild(r);
      });
      p.appendChild(pos);
      if (c.kind === 'textbox' && !/^[EA]_/i.test(c.name)) {
        p.appendChild(el('div', 'prop-hint', 'Tipp: Eingabefelder beginnen mit E_ (z. B. E_Zahl1), Ausgabefelder mit A_ (z. B. A_Ergebnis).'));
      }
      const events = VBA.CONTROL_EVENTS[c.kind] || [];
      if (events.length) {
        const ev = el('div', 'prop-events');
        for (const [evName, label] of events) ev.appendChild(this.eventButton(c.name + '_' + evName, label));
        p.appendChild(ev);
      }
      const del = el('button', 'prop-delete', 'Löschen');
      del.type = 'button';
      del.addEventListener('click', () => this.remove(c.name));
      p.appendChild(del);
    }

    eventButton(procName, label) {
      const b = el('button', 'prop-event');
      b.type = 'button';
      const exists = this.hasProc(procName);
      b.textContent = `${label}: ${exists ? '[Ereignisprozedur] → Code' : 'Ereignisprozedur anlegen'}`;
      b.title = `Private Sub ${procName}()`;
      b.addEventListener('click', () => this.onEventProc(procName));
      return b;
    }

    checkName(v, c) {
      if (!NAME_RE.test(v)) return 'Ungültiger Name: nur Buchstaben, Ziffern und _, am Anfang ein Buchstabe.';
      if (VBA.KEYWORDS.has(v.toLowerCase())) return `'${v}' ist ein reserviertes Wort.`;
      const other = this.find(v);
      if (other && other !== c) return `Den Namen '${v}' gibt es im Formular schon.`;
      if (v.toLowerCase() === 'form' || v.toLowerCase() === 'me') return `'${v}' ist als Name nicht erlaubt.`;
      return null;
    }
  }

  root.FormDesigner = FormDesigner;
})(typeof self !== 'undefined' ? self : this);
