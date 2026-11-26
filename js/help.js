/*
 * help.js – Inhalte des Hilfefensters.
 * Jedes Thema: { id, title, html, code? }. Beispiele werden in tests/help-examples.js automatisch geprüft.
 */
(function (root) {
  'use strict';
  const L = (s) => s.replace(/^\n/, '');

  const TOPICS = [
    {
      id: 'start', title: 'So startest du ein Programm',
      html: `<p>Ein VBA-Programm besteht aus <b>Prozeduren</b>. Gestartet wird immer eine <code>Sub</code> ohne Parameter.</p>
<ol>
<li>Schreibe deinen Code zwischen <code>Sub Main()</code> und <code>End Sub</code>.</li>
<li>Wähle oben bei <b>Start-Sub</b> aus, welche Sub laufen soll. Hast du mehrere Subs, stehen alle in der Liste.</li>
<li>Klicke auf <b>▶ Ausführen</b> (oder drücke <kbd>F5</kbd>).</li>
<li>Hängt das Programm in einer Endlosschleife, beendest du es mit <b>■ Stoppen</b>.</li>
</ol>
<p>Mit <b>Neu</b> bekommst du wieder das leere Grundgerüst.</p>`,
      code: L(`
Private Sub Main()
    Debug.Print("Hello World")
End Sub
`)
    },
    {
      id: 'io', title: 'Ein- und Ausgabe im Überblick',
      html: `<p>Es gibt zwei Wege, mit dem Programm zu „reden“:</p>
<table class="help-table">
<tr><th></th><th>Konsole (Direktfenster)</th><th>Dialogfenster</th></tr>
<tr><th>Ausgabe</th><td><code>Debug.Print</code> schreibt eine Zeile in die <b>Konsole</b>. Gut für viele Werte, z. B. in Schleifen.</td><td><code>MsgBox</code> zeigt eine Meldung, das Programm wartet auf <b>OK</b>.</td></tr>
<tr><th>Eingabe</th><td colspan="2"><code>InputBox</code> öffnet ein Eingabefeld und liefert den eingegebenen Text zurück.</td></tr>
</table>
<p>Die Konsole entspricht dem <b>Direktfenster</b> in Access (<kbd>Strg</kbd>+<kbd>G</kbd>). Auf dem Handy findest du sie im Reiter <b>Konsole</b>.</p>`,
      code: L(`
Sub Main()
    Dim name As String
    name = InputBox("Wie heißt du?")      ' Eingabe
    Debug.Print "Hallo " & name           ' Ausgabe in die Konsole
    MsgBox "Hallo " & name                ' Ausgabe als Dialog
End Sub
`)
    },
    {
      id: 'print', title: 'Ausgabe mit Debug.Print',
      html: `<p><code>Debug.Print</code> schreibt in die Konsole. Mehrere Werte trennst du so:</p>
<ul>
<li><code>;</code> schreibt direkt weiter (Zahlen bekommen ein Leerzeichen davor),</li>
<li><code>,</code> springt zur nächsten Spalte (Tabelle),</li>
<li><code>&amp;</code> verbindet Texte zu einem Text.</li>
</ul>
<p>Steht am Ende ein <code>;</code>, bleibt die nächste Ausgabe in derselben Zeile.</p>`,
      code: L(`
Sub Main()
    Dim i As Integer
    Debug.Print "Zahl", "Quadrat"
    For i = 1 To 5
        Debug.Print i, i * i
    Next i
    Debug.Print "Summe: " & (1 + 2 + 3)
    Debug.Print "a"; "b"; 3
End Sub
`)
    },
    {
      id: 'msgbox', title: 'Ausgabe mit MsgBox',
      html: `<p><code>MsgBox Text, Schaltflächen, Titel</code> zeigt einen Dialog. Nur der Text ist Pflicht.</p>
<p>Mit Schaltflächen wie <code>vbYesNo</code> wird MsgBox zur Frage. Dann setzt du die Argumente in Klammern und prüfst das Ergebnis (<code>vbYes</code>, <code>vbNo</code>).</p>
<p>Symbole: <code>vbInformation</code>, <code>vbExclamation</code>, <code>vbCritical</code>, <code>vbQuestion</code>. Kombinieren mit <code>+</code>.</p>`,
      code: L(`
Sub Main()
    MsgBox "Willkommen!", vbInformation, "Start"
    If MsgBox("Weitermachen?", vbYesNo + vbQuestion) = vbYes Then
        MsgBox "Gut, weiter geht's."
    Else
        MsgBox "Dann eben nicht."
    End If
End Sub
`)
    },
    {
      id: 'inputbox', title: 'Eingabe mit InputBox',
      html: `<p><code>InputBox(Text, Titel, Vorgabe)</code> fragt einen Wert ab. Bei <b>Abbrechen</b> kommt ein leerer Text <code>""</code> zurück.</p>
<p class="help-warn"><b>Achtung:</b> InputBox liefert immer einen <b>Text</b>. <code>"2" + "3"</code> ergibt <code>"23"</code>! Speichere die Eingabe in einer Zahlen-Variablen (<code>As Integer</code>, <code>As Double</code>) oder wandle sie mit <code>CInt</code>, <code>CDbl</code> bzw. <code>Val</code> um.</p>`,
      code: L(`
Sub Main()
    Dim a As Double, b As Double
    a = InputBox("Erste Zahl:", "Addition", "0")
    b = InputBox("Zweite Zahl:", "Addition", "0")
    Debug.Print a; "+"; b; "="; a + b
    MsgBox "Ergebnis: " & (a + b)
End Sub
`)
    },
    {
      id: 'vars', title: 'Variablen, Datentypen, Konstanten',
      html: `<p>Variablen deklarierst du mit <code>Dim name As Datentyp</code>.</p>
<table class="help-table">
<tr><th>Datentyp</th><th>für</th></tr>
<tr><td><code>Integer</code> / <code>Long</code></td><td>ganze Zahlen (bis 32.767 bzw. ca. 2 Mrd.)</td></tr>
<tr><td><code>Double</code> / <code>Single</code></td><td>Kommazahlen</td></tr>
<tr><td><code>Currency</code></td><td>Geldbeträge (4 Nachkommastellen)</td></tr>
<tr><td><code>String</code></td><td>Text</td></tr>
<tr><td><code>Boolean</code></td><td><code>True</code> / <code>False</code></td></tr>
<tr><td><code>Date</code></td><td>Datum und Uhrzeit, z. B. <code>#12/24/2026#</code></td></tr>
<tr><td><code>Variant</code></td><td>alles (Standard, wenn kein Typ angegeben ist)</td></tr>
</table>
<p><code>Option Explicit</code> ganz oben erzwingt, dass jede Variable deklariert wird. Das findet Tippfehler sofort.</p>
<p>Werte, die sich nie ändern: <code>Const MWST As Double = 0.19</code>.</p>`,
      code: L(`
Option Explicit
Const MWST As Double = 0.19

Sub Main()
    Dim netto As Currency
    Dim brutto As Currency
    netto = 100
    brutto = netto * (1 + MWST)
    Debug.Print "Brutto: "; brutto
End Sub
`)
    },
    {
      id: 'if', title: 'Verzweigungen: If und Select Case',
      html: `<p><code>If ... Then ... ElseIf ... Else ... End If</code> prüft Bedingungen nacheinander. Kurze Fälle passen in eine Zeile: <code>If x &gt; 0 Then Debug.Print "positiv"</code>.</p>
<p>Vergleiche: <code>=  &lt;&gt;  &lt;  &gt;  &lt;=  &gt;=</code>, verknüpft mit <code>And</code>, <code>Or</code>, <code>Not</code>.</p>
<p><code>Select Case</code> lohnt sich bei vielen Fällen. Möglich sind Listen (<code>Case 1, 2</code>), Bereiche (<code>Case 3 To 5</code>) und Vergleiche (<code>Case Is &gt;= 90</code>).</p>`,
      code: L(`
Sub Main()
    Dim punkte As Integer
    punkte = InputBox("Punkte (0 bis 100):", "Note", "75")

    If punkte >= 50 Then
        Debug.Print "bestanden"
    Else
        Debug.Print "nicht bestanden"
    End If

    Select Case punkte
        Case Is >= 92: Debug.Print "Note 1"
        Case 81 To 91: Debug.Print "Note 2"
        Case 67 To 80: Debug.Print "Note 3"
        Case 50 To 66: Debug.Print "Note 4"
        Case Else:     Debug.Print "Note 5 oder 6"
    End Select
End Sub
`)
    },
    {
      id: 'loops', title: 'Schleifen',
      html: `<ul>
<li><code>For i = 1 To 10 [Step 2] ... Next i</code>: feste Anzahl Durchläufe.</li>
<li><code>Do While Bedingung ... Loop</code>: solange die Bedingung wahr ist.</li>
<li><code>Do Until Bedingung ... Loop</code>: bis die Bedingung wahr ist.</li>
<li><code>Do ... Loop While/Until</code>: prüft erst am Ende, läuft also mindestens einmal.</li>
<li><code>Exit For</code> bzw. <code>Exit Do</code> verlässt eine Schleife vorzeitig.</li>
</ul>
<p>Hängt dein Programm, ändert sich die Bedingung vermutlich nie. Dann hilft <b>■ Stoppen</b>.</p>`,
      code: L(`
Sub Main()
    Dim i As Integer, summe As Long

    For i = 1 To 10
        summe = summe + i
    Next i
    Debug.Print "Summe 1 bis 10:"; summe

    i = 1
    Do While i < 100
        i = i * 2
    Loop
    Debug.Print "Erste Zweierpotenz ab 100:"; i
End Sub
`)
    },
    {
      id: 'procs', title: 'Subs und Functions',
      html: `<p>Eine <code>Sub</code> erledigt eine Aufgabe. Eine <code>Function</code> liefert zusätzlich einen Wert zurück: Du weist ihn dem Namen der Function zu.</p>
<p>Aufruf einer Sub: <code>Ausgabe "Text", 3</code> (ohne Klammern) oder <code>Call Ausgabe("Text", 3)</code>.<br>
Aufruf einer Function: <code>x = Quadrat(5)</code> (mit Klammern).</p>
<p>Parameter sind standardmäßig <code>ByRef</code>: Änderungen wirken auf die übergebene Variable zurück. Mit <code>ByVal</code> bekommt die Prozedur nur eine Kopie.</p>`,
      code: L(`
Sub Main()
    Gruss "Anna"
    Call Gruss("Ben")
    Debug.Print "5 zum Quadrat ist"; Quadrat(5)
End Sub

Sub Gruss(ByVal name As String)
    Debug.Print "Hallo " & name
End Sub

Function Quadrat(ByVal x As Double) As Double
    Quadrat = x * x
End Function
`)
    },
    {
      id: 'arrays', title: 'Arrays (Datenfelder)',
      html: `<p><code>Dim noten(1 To 5) As Integer</code> legt fünf Plätze an. Zugriff mit <code>noten(3)</code>.</p>
<p><code>LBound</code> und <code>UBound</code> liefern den kleinsten und größten Index. Ein Array ohne Größe (<code>Dim a() As Integer</code>) bekommt seine Größe mit <code>ReDim</code>. <code>ReDim Preserve</code> behält dabei die alten Werte.</p>`,
      code: L(`
Sub Main()
    Dim noten(1 To 5) As Integer
    Dim i As Integer, summe As Integer
    noten(1) = 2: noten(2) = 3: noten(3) = 1: noten(4) = 4: noten(5) = 2
    For i = LBound(noten) To UBound(noten)
        summe = summe + noten(i)
    Next i
    Debug.Print "Durchschnitt:"; summe / 5
End Sub
`)
    },
    {
      id: 'functions', title: 'Wichtige eingebaute Funktionen',
      html: `<table class="help-table">
<tr><th>Text</th><td><code>Len</code>, <code>Left</code>, <code>Right</code>, <code>Mid</code>, <code>InStr</code>, <code>UCase</code>, <code>LCase</code>, <code>Trim</code>, <code>Replace</code>, <code>Chr</code>, <code>Asc</code></td></tr>
<tr><th>Umwandeln</th><td><code>Val</code>, <code>CStr</code>, <code>CInt</code>, <code>CLng</code>, <code>CDbl</code>, <code>CBool</code>, <code>IsNumeric</code></td></tr>
<tr><th>Mathe</th><td><code>Round</code>, <code>Int</code>, <code>Fix</code>, <code>Abs</code>, <code>Sqr</code>, <code>Rnd</code>, <code>Randomize</code>, <code>Mod</code>, <code>\\</code> (ganzzahlig teilen), <code>^</code> (Potenz)</td></tr>
<tr><th>Datum</th><td><code>Now</code>, <code>Date</code>, <code>Year</code>, <code>Month</code>, <code>Day</code>, <code>Weekday</code>, <code>DateAdd</code>, <code>DateDiff</code></td></tr>
<tr><th>Format</th><td><code>Format(1234.5, "#,##0.00")</code> ergibt <code>1.234,50</code></td></tr>
</table>`,
      code: L(`
Sub Main()
    Dim wort As String
    wort = "Informatik"
    Debug.Print Len(wort); Left(wort, 4); " "; UCase(wort)
    Debug.Print Round(3.14159, 2); Int(7.9); 17 Mod 5
    Debug.Print Format(1234.5, "#,##0.00 €")
    Randomize
    Debug.Print "Würfel:"; Int(Rnd * 6) + 1
End Sub
`)
    },
    {
      id: 'forms', title: 'Formulare wie in Access',
      html: `<p>Im Bereich <b>Formular</b> baust du eine Oberfläche wie in Access. In der <b>Entwurfsansicht</b> legst du Steuerelemente an, verschiebst sie und gibst ihnen Namen. In der <b>Formularansicht</b> (oder mit <b>▶ Ausführen</b>) benutzt du das Formular.</p>
<p class="help-warn"><b>Namenskonvention:</b> Eingabefelder beginnen mit <code>E_</code> (z. B. <code>E_Zahl1</code>), Ausgabefelder mit <code>A_</code> (z. B. <code>A_Ergebnis</code>). Schaltflächen beginnen mit <code>btn</code> (z. B. <code>btnBerechnen</code>).</p>
<p><b>Zugriff ohne <code>Me.</code></b>, direkt über den Namen:</p>
<ul>
<li>Lesen: <code>zahl1 = E_Zahl1</code></li>
<li>Schreiben: <code>A_Ergebnis = summe</code></li>
<li>Bezeichnungsfeld: <code>lblInfo.Caption = "Fertig"</code></li>
<li>Weitere Eigenschaften: <code>.Visible</code>, <code>.Enabled</code>, <code>.Locked</code>, <code>.ForeColor</code>, <code>.BackColor</code>, Methode <code>.SetFocus</code></li>
</ul>
<p><b>Ereignisprozeduren</b> wie in Access: Klickt man auf <code>btnBerechnen</code>, läuft <code>Private Sub btnBerechnen_Click()</code>. Weitere Ereignisse: <code>_AfterUpdate</code> und <code>_Change</code> bei Textfeldern, <code>_AfterUpdate</code> bei Optionsgruppen und <code>Form_Load</code> beim Öffnen. Im Entwurf legt der Button <i>Ereignisprozedur anlegen</i> den Rahmen im Code an.</p>
<p class="help-warn"><b>Wie in Access:</b> Werte aus Textfeldern sind <b>Text</b>, ein leeres Feld ist <code>Null</code>. Deshalb ergibt <code>E_Zahl1 + E_Zahl2</code> bei 2 und 3 den Text „23“. Speichere die Werte in <code>Double</code>-Variablen oder wandle sie mit <code>CDbl</code>/<code>Val</code> um. Für leere Felder hilft <code>Nz(E_Zahl1, 0)</code>.</p>
<p>Eine <b>Optionsgruppe</b> liefert die Nummer der gewählten Option (1, 2, 3 …), solange nichts gewählt ist <code>Null</code>.</p>
<p>Das Beispiel übernimmt Code <b>und</b> Formular.</p>`,
      form: {
        caption: 'Rechner', width: 440, height: 250, controls: [
          { kind: 'textbox', name: 'E_Zahl1', caption: 'Zahl 1:', x: 96, y: 20, w: 120, h: 26 },
          { kind: 'textbox', name: 'E_Zahl2', caption: 'Zahl 2:', x: 96, y: 56, w: 120, h: 26 },
          { kind: 'optiongroup', name: 'E_Rechenart', caption: 'Rechenart', x: 244, y: 12, w: 170, h: 108, options: [{ caption: 'Addieren', value: 1 }, { caption: 'Subtrahieren', value: 2 }, { caption: 'Multiplizieren', value: 3 }] },
          { kind: 'button', name: 'btnBerechnen', caption: 'Berechnen', x: 96, y: 100, w: 120, h: 32 },
          { kind: 'textbox', name: 'A_Ergebnis', caption: 'Ergebnis:', x: 96, y: 152, w: 120, h: 26 },
          { kind: 'label', name: 'lblInfo', caption: '', x: 20, y: 196, w: 400, h: 24 }
        ]
      },
      code: L(`
Option Explicit

Private Sub Form_Load()
    E_Rechenart = 1
    lblInfo.Caption = "Zahlen eingeben und auf Berechnen klicken."
End Sub

Private Sub btnBerechnen_Click()
    Dim zahl1 As Double, zahl2 As Double, ergebnis As Double

    If IsNull(E_Zahl1) Or IsNull(E_Zahl2) Then
        MsgBox "Bitte beide Zahlen eingeben.", vbExclamation
        Exit Sub
    End If

    zahl1 = E_Zahl1          ' Text wird in Double umgewandelt
    zahl2 = E_Zahl2

    Select Case E_Rechenart
        Case 1: ergebnis = zahl1 + zahl2
        Case 2: ergebnis = zahl1 - zahl2
        Case 3: ergebnis = zahl1 * zahl2
    End Select

    A_Ergebnis = ergebnis
    lblInfo.Caption = "Berechnet um " & Format(Now, "hh:nn:ss")
End Sub
`)
    },
    {
      id: 'errors', title: 'Fehlermeldungen verstehen',
      html: `<p>Die Meldungen entsprechen denen in Access:</p>
<ul>
<li><b>Fehler beim Kompilieren</b> erscheinen schon vor dem Start, z. B. <i>Next ohne For</i>, <i>Block If ohne End If</i>, <i>Variable nicht definiert</i>. Die Zeile wird markiert.</li>
<li><b>Laufzeitfehler</b> passieren während das Programm läuft, z. B. <i>'13' Typen unverträglich</i> (Text statt Zahl), <i>'6' Überlauf</i> (Zahl zu groß für den Datentyp), <i>'11' Division durch Null</i>, <i>'9' Index außerhalb des gültigen Bereichs</i>. Mit <b>Debuggen</b> springst du zur Zeile.</li>
</ul>
<p>Der kleine Hinweis unter der Meldung erklärt genauer, was los ist.</p>`,
      code: L(`
Sub Main()
    Dim zahl As Integer
    zahl = InputBox("Gib etwas ein, das keine Zahl ist:")
    Debug.Print zahl
End Sub
`)
    }
  ];

  const api = { TOPICS };
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.VBAHelp = api;
})(typeof self !== 'undefined' ? self : this);
