# VBA-Editor für den Unterricht

Browserbasierter Mini-VBA-Editor mit Interpreter – läuft komplett offline, ohne Server und ohne Installation.

- **Starten:** `index.html` im Browser öffnen (Doppelklick genügt) oder den Ordner auf GitHub Pages / einen Schulserver kopieren.
- **Tests:** `node tests/run-tests.js` oder `tests/tests.html` im Browser öffnen.
- Spezifikation und Projektstruktur: siehe `CLAUDE.md`.

## Abweichungen von echtem VBA (bewusst)

- Ausgabe wie in einem deutschen Access: Dezimalkomma, `Wahr`/`Falsch`, Datum `TT.MM.JJJJ`.
- `CDbl`, `CInt` usw. akzeptieren Dezimalkomma **und** Dezimalpunkt (`Val` nur den Punkt – wie in VBA).
- Integer-Überlauf wird erst bei der Zuweisung an eine Variable geprüft, nicht schon bei `200 * 200`.
- Formular über den Button „Formular“ ein-/ausblenden. Eingeblendet startet es in der Formularansicht: ein Klick auf eine Schaltfläche führt sofort ihre `…_Click`-Prozedur aus (Doppelklick im Entwurf legt sie an).
- Syntaxfarben wie im VBA-Editor: Schlüsselwörter blau, Kommentare grün, Fehlerzeile gelb.
- Formulare: Textfelder, Bezeichnungsfelder, Schaltflächen, Optionsgruppen; Zugriff ohne `Me.` (`Me.` geht aber auch). Leeres Textfeld = `Null`, Werte sind Text – wie in Access.
- „Speichern“ lädt eine .bas-Datei herunter (Code + Formular in einer Datei). In der claude.ai-Vorschau blockiert der Browser Downloads, lokal funktioniert es.
- Noch nicht unterstützt: `On Error`, `With`, `GoTo`, Objekte/Klassen, benannte Argumente (`name:=wert`).
