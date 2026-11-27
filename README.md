# VBA-Editor für den Unterricht

Browserbasierter Mini-VBA-Editor mit Interpreter – läuft komplett offline, ohne Server und ohne Installation.

- **Starten:** `index.html` im Browser öffnen (Doppelklick genügt) oder den Ordner auf GitHub Pages / einen Schulserver kopieren.
- **Tests:** `node tests/run-tests.js` oder `tests/tests.html` im Browser öffnen.
- Spezifikation und Projektstruktur: siehe `CLAUDE.md`.

## Abweichungen von echtem VBA (bewusst)

- Ausgabe wie in einem deutschen Access: Dezimalkomma, `Wahr`/`Falsch`, Datum `TT.MM.JJJJ`.
- `CDbl`, `CInt` usw. akzeptieren Dezimalkomma **und** Dezimalpunkt (`Val` nur den Punkt – wie in VBA).
- Zahlen wie in Access: `200 * 200` ergibt Überlauf (Integer * Integer), `CDbl("3.5")` ergibt 35 (Punkt = Tausendertrennzeichen).
- Direktfenster zeigt wie in Access nur die letzten 200 Zeilen; `Stop` hält an (Fortsetzen/Beenden).
- Code oben, Direktfenster darunter (Trenner verschiebbar). Formular und Hilfe teilen sich rechts eine Seitenleiste mit Reitern (Breite per Trenner verschiebbar, Doppelklick setzt zurück), damit der Code breit bleibt. Datei-Befehle (Neu, Öffnen, Speichern) im Menü „Datei“, „Leeren“ im Kopf des Direktfensters. Formular über den Button „Formular“ ein-/ausblenden. A−/A+ ändert die Schriftgröße von Code und Direktfenster (wird gespeichert; am Handy im Menü „Datei“).
- Handy/Tablet: eine Kopfzeile, ▶ Ausführen schwebt unten rechts. Beim Tippen im Code erscheint über der Bildschirmtastatur eine Zeichenleiste (`" ( ) = & _ '` usw., Debug.Print, Dim) mit Knopf für das passende Schlusswort (End If, Next, Loop …). Eingeblendet startet es in der Formularansicht: ein Klick auf eine Schaltfläche führt sofort ihre `…_Click`-Prozedur aus (Doppelklick im Entwurf legt sie an).
- Live-Hilfe beim Tippen wie im VBA-Editor: gelbe QuickInfo mit der Syntax (aktueller Parameter fett) und Vorschlagsliste (vb-Konstanten, Funktionen, Steuerelemente, Variablen, nach `ctl.` Eigenschaften/Methoden). Übernehmen mit Tab/Enter/Doppelklick bzw. Antippen, Strg+Leertaste öffnet die Liste. Tests: `node tests/intellisense-tests.js`.
- Syntaxfarben wie im VBA-Editor: Schlüsselwörter blau, Kommentare grün, Fehlerzeile gelb.
- Formulare: Textfelder, Bezeichnungsfelder, Schaltflächen, Optionsgruppen, Listenfelder (`AddItem`, `RowSource = ""` zum Leeren); Standardnamen `e_eingabe1`, `a_ausgabe1`, `o_option1`, `lstAusgabe1`; Zugriff ohne `Me.` (`Me.` geht aber auch). Leeres Textfeld = `Null`, Werte sind Text – wie in Access.
- „Speichern“ lädt eine .bas-Datei herunter (Code + Formular in einer Datei). In der claude.ai-Vorschau blockiert der Browser Downloads, lokal funktioniert es.
- Fehlerbehandlung: `On Error GoTo Marke`, `On Error Resume Next`, `On Error GoTo 0`, `Resume`/`Resume Next`, `Err.Number`/`Err.Description`/`Err.Raise`. Sprungmarken nur auf oberster Ebene einer Prozedur (nicht in If/Schleifen).
- Formular: `Me!feld`, `Me("feld")`, `Me.Caption`, `DoCmd.Close`.
- Noch nicht unterstützt: Debug-Modus (Einzelschritt, Haltepunkte), `With`, Objekte/Klassen, benannte Argumente (`name:=wert`).
