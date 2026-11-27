/*
 * errors.js – Fehlerklasse für den VBA-Interpreter.
 *
 * Jede Meldung wird auf den Originaltext von Access/VBA abgebildet
 * (z. B. "Next ohne For", Laufzeitfehler 13 "Typen unverträglich").
 * Der ausführlichere Text des Interpreters bleibt als Hinweis (hint) erhalten.
 *
 *   e.kind          'syntax' (Fehler beim Kompilieren) | 'runtime' (Laufzeitfehler) | 'abort'
 *   e.accessMessage Text wie in Access
 *   e.number        Fehlernummer bei Laufzeitfehlern (Err.Number)
 *   e.hint          zusätzliche Erklärung für Schüler (kann leer sein)
 *   e.line          Zeilennummer
 *   e.message       "Zeile 5: Laufzeitfehler '13': Typen unverträglich"
 */
(function (root, factory) {
  const mod = factory();
  if (typeof module === 'object' && module.exports) module.exports = mod;
  else { root.VBA = root.VBA || {}; Object.assign(root.VBA, mod); }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // Laufzeitfehler: Nummer → Originaltext
  const RUNTIME_TEXTS = {
    5: 'Ungültiger Prozeduraufruf oder ungültiges Argument',
    20: 'Resume ohne Fehler',
    6: 'Überlauf',
    9: 'Index außerhalb des gültigen Bereichs',
    10: 'Datenfeld ist unveränderlich oder momentan gesperrt',
    11: 'Division durch Null',
    13: 'Typen unverträglich',
    28: 'Nicht genügend Stapelspeicher',
    35: 'Sub oder Function nicht definiert',
    94: 'Unzulässige Verwendung von Null',
    424: 'Objekt erforderlich',
    438: 'Objekt unterstützt diese Eigenschaft oder Methode nicht',
    2110: 'Microsoft Access kann den Fokus nicht auf dieses Steuerelement verschieben',
    2176: 'Die Einstellung für diese Eigenschaft ist zu lang',
    7777: 'Sie haben die ListIndex-Eigenschaft falsch verwendet',
    449: 'Argument ist nicht optional',
    450: 'Falsche Anzahl an Argumenten oder ungültige Zuweisung zu einer Eigenschaft'
  };

  // Kompilierfehler: Muster der internen Meldung → Originaltext
  const COMPILE_RULES = [
    [/^Das Array '.*' hat eine feste Größe/, 'Datenfeld wurde bereits dimensioniert'],
    [/^ByRef-Argumenttyp unverträglich/, 'ByRef-Argumenttyp unverträglich'],
    [/^Next ohne For/, 'Next ohne For'],
    [/^For( Each)? ohne Next/, 'For ohne Next'],
    [/^Do ohne Loop/, 'Do ohne Loop'],
    [/^Loop ohne Do/, 'Loop ohne Do'],
    [/^While ohne Wend/, 'While ohne Wend'],
    [/^Wend ohne While/, 'Wend ohne While'],
    [/^If-Block ohne End If/, 'Block If ohne End If'],
    [/^End If ohne If/, 'End If ohne Block If'],
    [/^Select Case ohne End Select/, 'Select Case ohne End Select'],
    [/^End Select ohne Select Case/, 'End Select ohne Select Case'],
    [/^Case ohne Select Case/, 'Case ohne Select Case'],
    [/^(Else|ElseIf) ohne If/, 'Else ohne If'],
    [/^Sub ohne End Sub|^Erwartet: 'End Sub'|^Sub oder Function innerhalb/, 'Erwartet: End Sub'],
    [/^Function ohne End Function|^Erwartet: 'End Function'/, 'Erwartet: End Function'],
    [/^End (Sub|Function) ohne/, 'Ungültig außerhalb einer Prozedur'],
    [/^Anweisung außerhalb einer Prozedur|^'Static' ist auf Modulebene/, 'Ungültig außerhalb einer Prozedur'],
    [/^Option-Anweisungen müssen|ist innerhalb einer Prozedur nicht erlaubt/, 'Ungültig innerhalb einer Prozedur'],
    [/^Erwartet: 'Then'/, 'Erwartet: Then oder GoTo'],
    [/^Erwartet: Anweisungsende/, 'Erwartet: Anweisungsende'],
    [/^Ausdruck erwartet|ist ein Schlüsselwort und kann hier nicht stehen/, 'Erwartet: Ausdruck'],
    [/^Erwartet: '='|^Beim Aufruf ohne 'Call'/, 'Erwartet: ='],
    [/^Erwartet: '\)'/, 'Erwartet: Listentrennzeichen oder )'],
    [/^Erwartet: '\('/, 'Erwartet: ('],
    [/^Erwartet: 'To'/, 'Erwartet: To'],
    [/^Erwartet: 'In'/, 'Erwartet: In'],
    [/reserviertes Schlüsselwort|^Erwartet: (Bezeichner|Variablenname|Name|Zählvariable|Schleifenvariable|Parametername|Arrayname|Konstantenname)/, 'Erwartet: Bezeichner'],
    [/^Unerwartetes Zeichen/, 'Ungültiges Zeichen'],
    [/^Unbekannter Datentyp/, 'Benutzerdefinierter Typ nicht definiert'],
    [/^Doppelte Deklaration/, 'Doppelte Deklaration im aktuellen Gültigkeitsbereich'],
    [/^Mehrdeutiger Name: '([^']*)'/, (m) => `Mehrdeutiger Name erkannt: ${m[1]}`],
    [/^Variable '.*' wurde nicht deklariert/, 'Variable nicht definiert'],
    [/^Sub oder Function '.*' ist nicht definiert|ist kein Array\. Arrays bitte mit Dim/, 'Sub oder Function nicht definiert'],
    [/^Falsche Anzahl an Argumenten/, 'Falsche Anzahl an Argumenten oder ungültige Zuweisung zu einer Eigenschaft'],
    [/^Argument nicht optional/, 'Argument ist nicht optional'],
    [/^Einer Konstanten|eingebaute Konstante|kann nicht als Zählvariable/, 'Zuweisung zu einer Konstanten nicht zulässig'],
    [/ist eine Sub und liefert keinen Wert|ist der Name einer Prozedur|ist eine eingebaute Funktion und kann nicht|ist eine Prozedur bzw\. Funktion/, 'Erwartet: Funktion oder Variable'],
    [/ist eine Konstante und kein Array/, 'Erwartet: Datenfeld'],
    [/ist eine Variable\. Für eine Zuweisung/, 'Erwartet: Sub, Function oder Property'],
    [/^'Me\.'/, 'Ungültige Verwendung des Me-Schlüsselwortes'],
    [/^Next .* passt nicht zur For-Schleife/, 'Ungültiger Verweis auf Next-Steuervariable'],
    [/^Nach 'Select Case' muss zuerst/, 'Anweisungen und Bezeichnungen zwischen Select Case und erstem Case ungültig'],
    [/^Exit For ist nur/, 'Exit For nicht innerhalb For...Next'],
    [/^Exit Do ist nur/, 'Exit Do nicht innerhalb Do...Loop'],
    [/^Exit Sub ist nur/, 'Exit Sub nicht zulässig in Function oder Property'],
    [/^Exit Function ist nur/, 'Exit Function nicht zulässig in Sub oder Property'],
    [/^Debug\..* wird nicht unterstützt|^Methode oder Datenobjekt nicht gefunden/, 'Methode oder Datenobjekt nicht gefunden'],
    [/^Erwartet: Bezeichner nach dem Punkt/, 'Erwartet: Bezeichner'],
    [/^Die Eigenschaft '.*' ist schreibgeschützt/, 'Zuweisung zu einer schreibgeschützten Eigenschaft nicht möglich'],
    [/^Hier ist nur ein konstanter Ausdruck/, 'Konstanter Ausdruck erforderlich'],
    [/^Zuweisung an das Array/, 'Zuweisen zu Datenfeld nicht möglich'],
    [/^Zeichenkette ist nicht abgeschlossen|Datumsliteral|^Unerwartet|^Erwartet|^Ein Standardwert|^Die Bedingung darf|ElseIf ist im einzeiligen|nach Else ist nicht erlaubt|^Nach 'Case Else'|^Nach 'Case Is'|^Nach 'Exit'|^Option Base|^Option Compare|^Unbekannte Option|^Index fehlt/, 'Syntaxfehler']
  ];

  // Interne Laufzeitmeldungen, die in Access Kompilierfehler wären
  const COMPILE_AT_RUNTIME = /^(Das Array '.*' hat eine feste Größe|ByRef-Argumenttyp|Einer Konstanten|Variable '.*' wurde nicht deklariert|Hier ist nur ein konstanter Ausdruck|Zuweisung an das Array|Die Konstante '.*' kann nicht als Zählvariable)/;

  function classify(message, kind, number) {
    let k = kind || 'runtime';
    if (k === 'abort') return { kind: k, access: message, hint: '' };
    if (k === 'runtime' && COMPILE_AT_RUNTIME.test(message)) k = 'syntax';
    let access = null;
    if (k === 'runtime') {
      if (/Schritten abgebrochen/.test(message)) access = 'Ausführung des Codes wurde unterbrochen';
      else if (number && RUNTIME_TEXTS[number]) access = RUNTIME_TEXTS[number];
    } else {
      for (const [re, text] of COMPILE_RULES) {
        const m = re.exec(message);
        if (m) { access = typeof text === 'function' ? text(m) : text; break; }
      }
    }
    if (!access) return { kind: k, access: message.replace(/\.$/, ''), hint: '' };
    let hint = message;
    if (hint.toLowerCase().startsWith(access.toLowerCase())) hint = hint.slice(access.length).replace(/^[\s:.,;–-]+/, '');
    if (hint.replace(/\.$/, '') === access) hint = '';
    if (hint) hint = hint.charAt(0).toUpperCase() + hint.slice(1); // Hinweis beginnt groß
    return { kind: k, access, hint };
  }

  class VbaError extends Error {
    /**
     * @param {string} message  interne deutsche Meldung (ausführlich)
     * @param {number|null} line 1-basierte Zeilennummer im Quelltext
     * @param {'syntax'|'runtime'|'abort'} kind Art des Fehlers
     * @param {number} [number] VBA-Fehlernummer (Err.Number), falls vorhanden
     */
    constructor(message, line, kind, number) {
      const c = classify(message, kind, number);
      const prefix = c.kind === 'syntax'
        ? 'Fehler beim Kompilieren: '
        : (c.kind === 'runtime' ? (number ? `Laufzeitfehler '${number}': ` : 'Laufzeitfehler: ') : '');
      super((line ? `Zeile ${line}: ` : '') + prefix + c.access);
      this.name = 'VbaError';
      this.plainMessage = message;
      this.accessMessage = c.access;
      this.hint = c.hint;
      this.line = line || null;
      this.kind = c.kind;
      this.number = c.kind === 'runtime' ? (number || 0) : 0;
    }
  }

  function syntaxError(msg, line) { return new VbaError(msg, line, 'syntax'); }
  function runtimeError(msg, line, number) { return new VbaError(msg, line, 'runtime', number); }

  return { VbaError, syntaxError, runtimeError, RUNTIME_TEXTS };
});
