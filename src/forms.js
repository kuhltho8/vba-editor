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
 * Ereignisse: btnX_Click, E_X_AfterUpdate, E_X_Change, E_Gruppe_AfterUpdate/_Click, Form_Load
 */
(function (root, factory) {
  const isNode = typeof module === 'object' && module.exports;
  const deps = isNode ? Object.assign({}, require('./errors.js'), require('./values.js')) : root.VBA;
  const mod = factory(deps);
  if (isNode) module.exports = mod;
  else { root.VBA = root.VBA || {}; Object.assign(root.VBA, mod); }
})(typeof self !== 'undefined' ? self : this, function (V) {
  'use strict';
  const { runtimeError, NULL, EMPTY, toStr, toBool, toNumber, coerce } = V;

  const KIND_NAMES = { textbox: 'Textfeld', label: 'Bezeichnungsfeld', button: 'Befehlsschaltfläche', optiongroup: 'Optionsgruppe' };

  // Erlaubte Eigenschaften (p) und Methoden (m) je Steuerelementtyp
  const CONTROL_PROPS = {
    textbox: { p: ['value', 'text', 'visible', 'enabled', 'locked', 'forecolor', 'backcolor', 'name'], m: ['setfocus'], hasValue: true },
    label: { p: ['caption', 'visible', 'forecolor', 'backcolor', 'name'], m: [], hasValue: false },
    button: { p: ['caption', 'visible', 'enabled', 'forecolor', 'backcolor', 'name'], m: ['setfocus'], hasValue: false },
    optiongroup: { p: ['value', 'visible', 'enabled', 'name'], m: ['setfocus'], hasValue: true }
  };
  const READ_ONLY = new Set(['name']);

  const EVENTS = {
    textbox: [['AfterUpdate', 'Nach Aktualisierung'], ['Change', 'Bei Änderung']],
    label: [['Click', 'Beim Klicken']],
    button: [['Click', 'Beim Klicken']],
    optiongroup: [['AfterUpdate', 'Nach Aktualisierung'], ['Click', 'Beim Klicken']]
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
      this.state = new Map();
      for (const c of model.controls) {
        this.state.set(c.name.toLowerCase(), {
          def: c, value: NULL, caption: c.caption || '', visible: true, enabled: true, locked: false,
          foreColor: null, backColor: null
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
      s.value = value;
    }

    // ---------- Schnittstelle für den Interpreter ----------
    adapters() {
      const list = [];
      for (const [lower, s] of this.state) {
        const kind = s.def.kind;
        const self = this;
        const name = s.def.name;
        const notSupported = (what, line) => runtimeError(
          `Objekt unterstützt diese Eigenschaft oder Methode nicht: ${KIND_NAMES[kind]} '${name}' hat ${what}.` +
          (kind === 'label' ? ` Für Bezeichnungsfelder: ${name}.Caption = ...` : '') +
          (kind === 'button' ? ` Für Schaltflächen: ${name}.Caption = ...` : ''), line, 438);
        list.push({
          name, lower, kind,
          getValue(line) {
            if (!CONTROL_PROPS[kind].hasValue) throw notSupported('keinen Wert (Value)', line);
            return s.value;
          },
          setValue(v, line) {
            if (!CONTROL_PROPS[kind].hasValue) throw notSupported('keinen Wert (Value)', line);
            if (kind === 'optiongroup') {
              s.value = (v === NULL || v === EMPTY) ? NULL : toNumber(v, line);
            } else {
              s.value = v === EMPTY ? NULL : v;
              if (s.value !== NULL) toStr(s.value, line); // prüft, ob anzeigbar
            }
            self.onChange(name);
          },
          get(prop, line) {
            switch (prop) {
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
          call(method, line) {
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

  return { FormRuntime, CONTROL_PROPS, READ_ONLY, CONTROL_EVENTS: EVENTS, KIND_NAMES, emptyForm, vbaColorToCss, serializeProject: serialize, deserializeProject: deserialize };
});
