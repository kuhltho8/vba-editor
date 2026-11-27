# CLAUDE.md – VBA-Web-Editor für den Unterricht

## Projektziel

Ein browserbasierter Mini-VBA-Editor mit Interpreter, mit dem Schülerinnen und Schüler einfache VBA-Programme schreiben und ausführen können, ohne dass Microsoft Access installiert sein muss. Hintergrund: Bei Schulausfällen (z. B. Glätte im Winter) haben viele Schüler zu Hause keinen PC oder kein Access. Ein echtes Online-Tool, das VBA (nicht VB.NET) ausführt, gibt es nicht.

Einsatz: Berufsschule (KBS Nordhorn), u. a. Berufliches Gymnasium und Fachinformatiker, überall dort, wo mit Access/VBA gearbeitet wird.

## Technische Rahmenbedingungen

- Rein statische Web-App: HTML, CSS, JavaScript. Kein Server, keine Datenbank, kein PHP, kein Build-Zwang.
- Läuft offline und im Browser auf Handy, Tablet und PC (responsive, touch-tauglich).
- Veröffentlichung über GitHub Pages; muss ohne Änderungen auf beliebiges Webhosting oder einen Schulserver kopierbar sein.
- Keine externen CDNs zur Laufzeit nötig (alle Abhängigkeiten lokal einbinden), damit es auch offline funktioniert.
- Code sauber getrennt: Tokenizer, Parser (AST), Interpreter, UI.
- Unit-Tests für den Interpreter (z. B. mit einer kleinen Testseite oder Node).

## Phase 1: Editor und Konsole

### Editor

- Code-Editor mit Syntaxhervorhebung für VBA, Zeilennummern, Einrückung.
- Buttons: Ausführen, Stoppen, Konsole leeren, Beispiel laden, Code speichern/laden (localStorage und Datei-Download/-Upload als .bas oder .txt).
- Auf dem Handy gut bedienbar (z. B. Editor und Konsole als Tabs).

### Unterstützter Sprachumfang (VBA, nicht VB.NET)

- Sub und Function, Aufruf mit und ohne Call, Parameter (ByVal/ByRef), Rückgabewerte.
- Dim, Const, Datentypen: Integer, Long, Double, Single, Currency, String, Boolean, Date, Variant. Option Explicit optional.
- Operatoren: `+ - * / \ Mod ^ &` sowie `= <> < > <= >=`, `And Or Not Xor`.
- If ... Then ... ElseIf ... Else ... End If (auch einzeilig), Select Case (inkl. Is, To, Listen).
- For ... To ... Step ... Next, For Each (später), Do While/Until ... Loop, Do ... Loop While/Until, While ... Wend, Exit For/Do/Sub/Function.
- Arrays (Dim a(1 To 10), ReDim, LBound, UBound) – kann Phase 1b sein.
- Wichtige eingebaute Funktionen: Len, Left, Right, Mid, InStr, UCase, LCase, Trim, Replace, Val, CStr, CInt, CLng, CDbl, CBool, IsNumeric, Round, Int, Fix, Abs, Sqr, Rnd, Randomize, Now, Date, Format (Grundfälle), Chr, Asc.
- Kommentare mit `'` und `Rem`, Zeilenfortsetzung mit ` _`.
- VBA ist nicht case-sensitiv: Bezeichner und Schlüsselwörter unabhängig von Groß-/Kleinschreibung behandeln.

### Ein- und Ausgabe

- Debug.Print gibt in die Konsole aus (inkl. `;` und `,` als Trenner wie im Direktfenster).
- MsgBox zeigt einen Dialog (eigenes Modal, kein alert) und schreibt optional zusätzlich in die Konsole.
- InputBox(prompt, title, default) öffnet ein Eingabefeld (eigenes Modal oder Eingabezeile in der Konsole) und liefert einen String zurück.
- Ausführung muss asynchron laufen, damit InputBox/MsgBox auf Eingaben warten können, ohne den Browser zu blockieren.
- Schutz vor Endlosschleifen (Stopp-Button, Schritt-/Zeitlimit).

### Fehlermeldungen

- Verständliche deutsche Fehlermeldungen mit Zeilennummer, z. B. "Zeile 5: Variable 'summe' wurde nicht deklariert" oder "Zeile 8: Next ohne For".
- Fehlerzeile im Editor markieren.

### Hilfefenster

Die Schüler arbeiten im Unterricht selten mit Debug.Print, deshalb ist eine gut sichtbare Hilfe wichtig:

- Kurzreferenz: Ausgabe mit Debug.Print und MsgBox, Eingabe mit InputBox, Variablen mit Dim, Verzweigungen, Schleifen, Subs/Functions.
- Zu jedem Thema ein kurzes Beispiel mit Button "In Editor übernehmen".
- Hinweis, wie man ein Programm startet (welche Sub ausgeführt wird, z. B. Auswahl der Start-Sub per Dropdown).

## Phase 2: Formularfenster

- Einfacher Formular-Designer neben dem Editor: Textfelder, Bezeichnungsfelder und Buttons per Klick anlegen, verschieben, benennen.
- Ereignisprozeduren wie in Access: `Private Sub btnBerechnen_Click() ... End Sub`.
- Namenskonvention der Klasse (unbedingt als Standard vorschlagen und in der Hilfe erklären):
  - Standardnamen kleingeschrieben (Wunsch Thomas, 26.11.2026):
  - Eingabefelder beginnen mit `e_` (z. B. `e_eingabe1`)
  - Ausgabefelder beginnen mit `a_` (z. B. `a_ausgabe1`)
  - Optionsgruppen beginnen mit `o_` (z. B. `o_option1`)
  - Schaltflächen `btnBefehl1`, Bezeichnungsfelder `lblText1`, Listenfelder `lstAusgabe1`
- Zugriff auf Felder ohne `Me.`, direkt über den Namen:
  - Lesen: `zahl1 = e_eingabe1`
  - Schreiben: `a_ausgabe1 = summe`
- Listenfeld (Wertliste wie in Access) für mehrere Ausgaben in Schleifen: `AddItem`, `RemoveItem`, `RowSource = ""` zum Leeren, `ListCount`, `ListIndex`, `ItemData(i)`, `Value`.
- Layout: Code oben, Direktfenster darunter (kleiner, Trenner verschiebbar); Formular per Button „Formular“ rechts über die volle Höhe einblendbar.
- Werte aus Textfeldern kommen als String/Variant; Verhalten möglichst wie in Access (z. B. Val/CDbl nötig bzw. automatische Typumwandlung wie bei Variant).
- Formular und Code gemeinsam speichern und laden.

## Spätere Ideen (optional)

- Mini-Recordset mit Beispieltabellen für einfache Datenzugriffe.
- Aufgabensammlung mit Musterlösungen, ähnlich dem bestehenden SQL-Übungstool.
- Teilen von Code per Link.

## Arbeitsweise

- Schrittweise vorgehen: zuerst Tokenizer und Interpreter mit Konsole, dann Editor-UI, dann Hilfe, dann Formulare.
- Jede Erweiterung mit Testprogrammen absichern.
- Oberfläche und Meldungen auf Deutsch.

## Projektstruktur (Stand Phase 2)

- `src/errors.js` – Fehlerklasse `VbaError` mit Zeilennummer
- `src/tokenizer.js` – zerlegt Quelltext in Tokens
- `src/parser.js` – baut den AST
- `src/interpreter.js` – führt den AST asynchron aus (I/O über Host-Objekt); Sitzungen (open/invoke) für Formulare
- `src/forms.js` – Formularmodell, Laufzeit der Steuerelemente (Value, Caption, Visible …), Speichern als .bas mit Kopfzeile `' @Formular {json}`
- `js/designer.js` – Formular-Designer (Entwurf) und Formularansicht
- `js/highlight.js` – Syntaxfarben (Schlüsselwörter blau, Kommentare grün) als Ebene hinter dem Textfeld
- `js/intellisense.js` – Live-Hilfe beim Tippen (QuickInfo mit Syntax, Vorschlagsliste); Tests: `node tests/intellisense-tests.js`
- `src/vba.js` – bündelt alles für Browser (`window.VBA`) und Node (`require`)
- `tests/run-tests.js` – Unit-Tests: `node tests/run-tests.js`; `tests/help-examples.js` prüft alle Hilfe-Beispiele
- `js/help.js` – Inhalte des Hilfefensters (Themen mit Beispiel)
- `js/app.js` – Oberfläche (Editor, Konsole, Dialoge, Hilfe, Neu/Rückgängig)
- `index.html` – Oberfläche (Editor, Direktfenster, Formularbereich, Hilfe)

## Entscheidungen (Fehlerbehebung nach Volltest, 26.11.2026)

- Verhalten wie in Access: Integer-Überlauf schon bei `200 * 200`, Punkt in Zahlentexten = Tausendertrennzeichen, Direktfenster max. 200 Zeilen, `Stop` hält an (Fortsetzen/Beenden).
- `On Error`/`Resume`/`Err` unterstützt; Sprungmarken nur auf oberster Prozedurebene.
- Debug-Modus (Einzelschritt, Haltepunkte) vorerst nicht.

Die Module funktionieren ohne Build: im Browser als normale `<script>`-Dateien, in Node per `require`.
