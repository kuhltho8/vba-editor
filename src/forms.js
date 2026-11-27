/*
 * forms.js – Formularmodell und Laufzeit der Steuerelemente (ohne DOM).
 *
 * Modell (wird mit dem Code gespeichert):
 *   { width, height, caption, controls: [
 *       { kind: 'textbox'|'label'|'button'|'optiongroup', name, caption, x, y, w, h,
 *         options?: [{ caption, value }] } ] }
 *
 * Zugriff aus VBA wie in Access, aber ohne Me.:
 *   zahl1 = E_Eingabe1        (Value eines Textfelds, leeres Feld = Null)
 *   A_Ergebnis = summe        (Wert anzeigen)
 *   lblInfo.Caption = "..."   E_Eingabe1.SetFocus   btnOk.Enabled = False
 *   lstAusgabe.AddItem "Zeile"   lstAusgabe.RemoveItem 0   lstAusgabe.RowSource = ""  (Listenfeld leeren)
 * Ereignisse: btnX_Click, e_x_AfterUpdate, e_x_Change, o_gruppe_AfterUpdate/_Click, lstX_Click/_DblClick, Form_Load
 */
(function (root, factory) {
  const isNode = typeof module === 'object' && module.exports;
  const deps = isNode ? Object.assign({}, require('./errors.js'), require('./values.js')) : root.VBA;
  const mod = factory(deps);
  if (isNode) module.exports = mod;
  else { root.VBA = root.VBA || {}; Object.assign(root.VBA, mod); }
})(typeof self !== 'undefined' ? self : this, function (V) {
  'use strict';
  const { runtimeError, NULL, EMPTY, MISSING, toStr, toBool, toNumber, coerce } = V;
  const MISSING_ARG = MISSING;

  const KIND_NAMES = { textbox: 'Textfeld', label: 'Bezeichnungsfeld', button: 'Befehlsschaltfläche', optiongroup: 'Optionsgruppe', listbox: 'Listenfeld' };

  // Erlaubte Eigenschaften (p) und Methoden (m) je Steuerelementtyp
  const CONTROL_PROPS = {
    textbox: { p: ['value', 'text', 'visible', 'enabled', 'locked', 'forecolor', 'backcolor', 'name'], m: ['setfocus'], hasValue: true },
    label: { p: ['caption', 'visible', 'forecolor', 'backcolor', 'name'], m: [], hasValue: false },
    button: { p: ['caption', 'visible', 'enabled', 'forecolor', 'backcolor', 'name'], m: ['setfocus'], hasValue: false },
    optiongroup: { p: ['value', 'visible', 'enabled', 'name'], m: ['setfocus'], hasValue: true },
    listbox: {
      p: ['value', 'listindex', 'listcount', 'rowsource', 'rowsourcetype', 'itemdata', 'column', 'visible', 'enabled', 'locked', 'forecolor', 'backcolor', 'name'],
      m: ['additem', 'removeitem', 'setfocus', 'requery'], hasValue: true
    }
  };
  const READ_ONLY = new Set(['name', 'listcount', 'itemdata', 'column']);
  // Anzahl Argumente [min, max] für Methoden und Eigenschaften mit Klammern (ItemData(i), Column(0, i))
  const MEMBER_ARGS = { additem: [1, 2], removeitem: [1, 1], itemdata: [1, 1], column: [1, 2] };

  // Access begrenzt die Datensatzherkunft (Wertliste) auf 32.750 Zeichen
  const ROWSOURCE_MAX = 32750;

  /** Wertliste "a;b;'c d'" → ['a', 'b', 'c d'] */
  function splitValueList(text) {
    const out = [];
    let cur = '', q = null, any = false;
    for (const ch of String(text)) {
      if (q) { if (ch === q) q = null; else cur += ch; continue; }
      if (ch === '"' || ch === "'") { q = ch; any = true; continue; }
      if (ch === ';') { out.push(cur.trim()); cur = ''; any = true; continue; }
      cur += ch;
    }
    if (cur.trim() !== '' || any) out.push(cur.trim());
    return out;
  }
  const quoteItem = (t) => (/[;"']/.test(t) || t !== t.trim() ? '"' + t.replace(/"/g, '') + '"' : t);
  const joinValueList = (items) => items.map(quoteItem).join(';');

  const EVENTS = {
    textbox: [['AfterUpdate', 'Nach Aktualisierung'], ['Change', 'Bei Änderung']],
    label: [['Click', 'Beim Klicken']],
    button: [['Click', 'Beim Klicken']],
    optiongroup: [['AfterUpdate', 'Nach Aktualisierung'], ['Click', 'Beim Klicken']],
    listbox: [['Click', 'Beim Klicken'], ['AfterUpdate', 'Nach Aktualisierung'], ['DblClick', 'Beim Doppelklicken']]
  };

  function emptyForm() { return { caption: 'Formular1', width: 480, height: 320, controls: [] }; }

  /** VBA-Farbwert (Long, BGR) → CSS-Farbe. */
  function vbaColorToCss(v) {
    const n = Math.max(0, Math.min(16777215, Math.round(v)));
    return `rgb(${n & 255}, ${(n >> 8) & 255}, ${(n >> 16) & 255})`;
  }

  class FormRuntime {
    /** @param {object} model Formularmodell  @param {function} onChange (controlName) => void */
    constructor(model, onChange) {
      this.model = model;
      this.onChange = onChange || (() => {});
      this.onFocus = () => {};
      this.caption = model.caption || 'Formular1';
      this.state = new Map();
      for (const c of model.controls) {
        this.state.set(c.name.toLowerCase(), {
          def: c, value: NULL, caption: c.caption || '', visible: true, enabled: true, locked: false,
          foreColor: null, backColor: null,
          items: c.kind === 'listbox' ? splitValueList(c.rowSource || '').filter((t, i, a) => a.length > 1 || t !== '') : null,
          listIndex: -1, rowSourceType: 'Wertliste'
        });
      }
    }

    get(name) { return this.state.get(String(name).toLowerCase()) || null; }

    /** Anzeige-Text eines Textfelds */
    displayText(name) {
      const s = this.get(name);
      if (!s) return '';
      if (s.value === NULL || s.value === EMPTY) return '';
      try { return toStr(s.value); } catch (e) { return ''; }
    }

    // ---------- Eingaben aus der Oberfläche ----------
    userInput(name, text) {
      const s = this.get(name);
      if (!s) return;
      s.value = text === '' ? NULL : String(text);
    }
    userSelect(name, value) {
      const s = this.get(name);
      if (!s) return;
      if (s.items) { s.listIndex = value; return; }
      s.value = value;
    }

    /** Wert eines Listenfelds: markierter Eintrag oder Null */
    listValue(s) {
      return s.listIndex >= 0 && s.listIndex < s.items.length ? s.items[s.listIndex] : NULL;
    }

    // ---------- Schnittstelle für den Interpreter ----------
    adapters() {
      const list = [];
      for (const [lower, s] of this.state) {
        const kind = s.def.kind;
        const self = this;
        const name = s.def.name;
        const tooLong = (line) => runtimeError(
          `Die Einstellung für diese Eigenschaft ist zu lang: Ein Listenfeld fasst wie in Access höchstens ${ROWSOURCE_MAX} Zeichen. ` +
          `Leeren mit ${name}.RowSource = ""`, line, 2176);
        const notSupported = (what, line) => runtimeError(
          `Objekt unterstützt diese Eigenschaft oder Methode nicht: ${KIND_NAMES[kind]} '${name}' hat ${what}.` +
          (kind === 'label' ? ` Für Bezeichnungsfelder: ${name}.Caption = ...` : '') +
          (kind === 'button' ? ` Für Schaltflächen: ${name}.Caption = ...` : ''), line, 438);
        list.push({
          name, lower, kind,
          getValue(line) {
            if (!CONTROL_PROPS[kind].hasValue) throw notSupported('keinen Wert (Value)', line);
            if (kind === 'listbox') return self.listValue(s);
            return s.value;
          },
          setValue(v, line) {
            if (!CONTROL_PROPS[kind].hasValue) throw notSupported('keinen Wert (Value)', line);
            if (kind === 'optiongroup') {
              s.value = (v === NULL || v === EMPTY) ? NULL : toNumber(v, line);
            } else if (kind === 'listbox') {
              // wie in Access: Value setzen markiert den passenden Eintrag
              const t = (v === NULL || v === EMPTY) ? null : toStr(v, line);
              s.listIndex = t === null ? -1 : s.items.findIndex((x) => x.toLowerCase() === t.toLowerCase());
            } else {
              s.value = v === EMPTY ? NULL : v;
              if (s.value !== NULL) toStr(s.value, line); // prüft, ob anzeigbar
            }
            self.onChange(name);
          },
          get(prop, line, args) {
            switch (prop) {
              case 'listindex': return s.listIndex;
              case 'listcount': return s.items.length;
              case 'rowsource': return joinValueList(s.items);
              case 'rowsourcetype': return s.rowSourceType;
              case 'itemdata': {
                const i = Math.round(toNumber(args[0], line));
                return i >= 0 && i < s.items.length ? s.items[i] : NULL;
              }
              case 'column': {
                const col = Math.round(toNumber(args[0], line));
                const row = args.length > 1 ? Math.round(toNumber(args[1], line)) : s.listIndex;
                return col === 0 && row >= 0 && row < s.items.length ? s.items[row] : NULL;
              }
              case 'value': return this.getValue(line);
              case 'text': return self.displayText(name);
              case 'caption': return s.caption;
              case 'visible': return s.visible;
              case 'enabled': return s.enabled;
              case 'locked': return s.locked;
              case 'forecolor': return s.foreColor === null ? 0 : s.foreColor;
              case 'backcolor': return s.backColor === null ? 16777215 : s.backColor;
              case 'name': return name;
              default: throw notSupported(`keine Eigenschaft '${prop}'`, line);
            }
          },
          set(prop, v, line) {
            switch (prop) {
              case 'value': this.setValue(v, line); return;
              case 'listindex': {
                const i = Math.round(toNumber(v, line));
                if (i < -1 || i >= s.items.length) {
                  throw runtimeError(`Sie haben die ListIndex-Eigenschaft falsch verwendet: '${name}' hat ${s.items.length} Einträge (Index 0 bis ${s.items.length - 1}).`, line, 7777);
                }
                s.listIndex = i;
                break;
              }
              case 'rowsource': {
                const t = toStr(v, line);
                if (t.length > ROWSOURCE_MAX) throw tooLong(line);
                s.items = t === '' ? [] : splitValueList(t);
                s.listIndex = -1;
                s.rsLen = undefined;
                break;
              }
              case 'rowsourcetype': s.rowSourceType = toStr(v, line); break;
              case 'text': s.value = toStr(v, line) === '' ? NULL : toStr(v, line); break;
              case 'caption': s.caption = toStr(v, line); break;
              case 'visible': s.visible = toBool(v, line); break;
              case 'enabled': s.enabled = toBool(v, line); break;
              case 'locked': s.locked = toBool(v, line); break;
              case 'forecolor': s.foreColor = coerce(v, 'long', line); break;
              case 'backcolor': s.backColor = coerce(v, 'long', line); break;
              default: throw notSupported(`keine Eigenschaft '${prop}'`, line);
            }
            self.onChange(name);
          },
          call(method, line, args) {
            if (method === 'additem') {
              const parts = splitValueList(toStr(args[0], line));
              const hasIndex = args.length > 1 && args[1] !== MISSING_ARG;
              let at = s.items.length;
              if (hasIndex) {
                at = Math.round(toNumber(args[1], line));
                if (at < 0 || at > s.items.length) {
                  throw runtimeError(`Ungültiger Index ${at} für AddItem: '${name}' hat ${s.items.length} Einträge.`, line, 5);
                }
              }
              const before = s.rsLen === undefined ? joinValueList(s.items).length : s.rsLen;
              const len = before + parts.reduce((n, t) => n + quoteItem(t).length + 1, 0) - (s.items.length ? 0 : 1);
              if (len > ROWSOURCE_MAX) throw tooLong(line);
              if (at === s.items.length) for (const t of parts) s.items.push(t);
              else s.items.splice(at, 0, ...parts);
              s.rsLen = len;
              if (s.listIndex >= at && hasIndex) s.listIndex += parts.length;
              self.onChange(name);
              return;
            }
            if (method === 'removeitem') {
              let i;
              if (typeof args[0] === 'string' && !/^\s*-?\d+\s*$/.test(args[0])) {
                i = s.items.findIndex((x) => x.toLowerCase() === args[0].toLowerCase());
              } else i = Math.round(toNumber(args[0], line));
              if (i < 0 || i >= s.items.length) {
                throw runtimeError(`Ungültiger Index für RemoveItem: '${name}' hat ${s.items.length} Einträge (Index 0 bis ${s.items.length - 1}).`, line, 5);
              }
              s.items.splice(i, 1);
              s.rsLen = undefined;
              if (s.listIndex === i) s.listIndex = -1;
              else if (s.listIndex > i) s.listIndex--;
              self.onChange(name);
              return;
            }
            if (method === 'requery') { self.onChange(name); return; }
            if (method === 'setfocus') {
              if (!s.visible || !s.enabled) {
                throw runtimeError(`Der Fokus kann nicht auf '${name}' gesetzt werden, weil es ausgeblendet oder deaktiviert ist.`, line, 2110);
              }
              self.onFocus(name);
              return;
            }
            throw notSupported(`keine Methode '${method}'`, line);
          }
        });
      }
      return list;
    }
  }

  /** Formular als Kommentarzeile vor den Code schreiben (eine Datei für beides). */
  const FORM_MARK = "' @Formular ";
  function serialize(code, form) {
    const body = code.replace(/^' @Formular .*\r?\n?/m, '');
    if (!form || !form.controls || form.controls.length === 0) return body;
    return FORM_MARK + JSON.stringify(form) + '\n' + body;
  }
  function deserialize(text) {
    const src = String(text).replace(/\r\n?/g, '\n');
    const m = /^' @Formular (.*)\n?/m.exec(src);
    if (!m) return { code: src, form: null };
    let form = null;
    try { form = JSON.parse(m[1]); } catch (e) { form = null; }
    return { code: src.replace(m[0], ''), form };
  }

  return { FormRuntime, CONTROL_PROPS, READ_ONLY, MEMBER_ARGS, splitValueList, CONTROL_EVENTS: EVENTS, KIND_NAMES, emptyForm, vbaColorToCss, serializeProject: serialize, deserializeProject: deserialize };
});
