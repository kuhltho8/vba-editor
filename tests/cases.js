/*
 * cases.js – Testprogramme für den Interpreter.
 * Jeder Fall: { name, src, out?: [...Zeilen], inputs?: [...], msg?: [...MsgBox-Texte],
 *              error?: RegExp, line?: Fehlerzeile, entry?, maxSteps? }
 * Wird von tests/run-tests.js (Node) und tests/tests.html (Browser) benutzt.
 */
(function (root) {
  const L = (s) => s.replace(/^\n/, '');
  const FORM1 = {
    caption: 'Rechner', width: 480, height: 320, controls: [
      { kind: 'textbox', name: 'E_Zahl1', caption: 'Zahl 1', x: 100, y: 20, w: 120, h: 24 },
      { kind: 'textbox', name: 'E_Zahl2', caption: 'Zahl 2', x: 100, y: 60, w: 120, h: 24 },
      { kind: 'textbox', name: 'A_Ergebnis', caption: 'Ergebnis', x: 100, y: 100, w: 120, h: 24 },
      { kind: 'label', name: 'lblInfo', caption: '', x: 20, y: 140, w: 200, h: 24 },
      { kind: 'button', name: 'btnBerechnen', caption: 'Berechnen', x: 260, y: 20, w: 120, h: 30 },
      { kind: 'optiongroup', name: 'E_Art', caption: 'Art', x: 260, y: 70, w: 160, h: 90, options: [{ caption: 'Normal', value: 1 }, { caption: 'Rabatt', value: 2 }] }
    ]
  };
  const cases = [
    // ---------- Ausgabe und Ausdrücke ----------
    { name: 'Debug.Print mit ; und ,', src: L(`
Sub Main()
  Debug.Print "a"; "b"
  Debug.Print 5; -3; "x"
  Debug.Print "A", "B"
  Debug.Print 1.5
  Debug.Print True; False
  Debug.Print "ohne Umbruch ";
  Debug.Print "weiter"
  Debug.Print
End Sub`), out: ['ab', ' 5 -3 x', 'A             B', ' 1,5 ', 'WahrFalsch', 'ohne Umbruch weiter', ''] },

    { name: 'Rangfolge der Operatoren', src: L(`
Sub Main()
  Debug.Print 2 + 3 * 4; (2 + 3) * 4; -2 ^ 2; 2 ^ 3 ^ 2
  Debug.Print 10 / 4; 10 \\ 4; 10 Mod 4; -7 Mod 3
  Debug.Print 5 > 3 And 2 > 1; Not 5 > 3; 3 Xor 5
End Sub`), out: [' 14  20 -4  64 ', ' 2,5  2  2 -1 ', 'WahrFalsch 6 '] },

    { name: '+ verkettet Strings, & immer', src: L(`
Sub Main()
  Dim a As String, b As String, x As Integer
  a = "5": b = "3": x = 2
  Debug.Print a + b
  Debug.Print a + x
  Debug.Print a & x & b
  Debug.Print "Summe: " & 1.5 + 1
End Sub`), out: ['53', ' 7 ', '523', 'Summe: 2,5'] },

    { name: 'Groß-/Kleinschreibung egal', src: L(`
sub main()
  DIM Zahl as INTEGER
  zahl = 4
  DEBUG.PRINT ZAHL * 2
  debug.print ucase("abc"); LEN("Hallo")
END SUB`), out: [' 8 ', 'ABC 5 '] },

    { name: 'Kommentare, Rem und Zeilenfortsetzung', src: L(`
' Kommentar
Rem auch ein Kommentar
Sub Main() ' Kommentar am Zeilenende
  Dim s As String
  s = "Hallo " & _
      "Welt"   ' weiter
  Debug.Print s: Rem nach Doppelpunkt
End Sub`), out: ['Hallo Welt'] },

    // ---------- Variablen und Typen ----------
    { name: 'Typumwandlung bei Zuweisung', src: L(`
Sub Main()
  Dim i As Integer, l As Long, d As Double, s As String, b As Boolean, c As Currency
  i = 2.5: Debug.Print i
  i = 3.5: Debug.Print i
  i = "42": Debug.Print i
  d = "1,25": Debug.Print d
  s = 3.75: Debug.Print s
  b = 1: Debug.Print b
  c = 1.23456: Debug.Print c
  l = 100000: Debug.Print l
End Sub`), out: [' 2 ', ' 4 ', ' 42 ', ' 1,25 ', '3,75', 'Wahr', ' 1,2346 ', ' 100000 '] },

    { name: 'Standardwerte und Variant', src: L(`
Sub Main()
  Dim i As Integer, s As String, b As Boolean, v
  Debug.Print i; "["; s; "]"; b; IsEmpty(v)
  v = 5: Debug.Print TypeName(v)
  v = "x": Debug.Print TypeName(v)
  Debug.Print TypeName(i); TypeName(s)
End Sub`), out: [' 0 [];Falsch;Wahr'.replace(/;/g, ''), 'Integer', 'String', 'IntegerString'] },

    { name: 'Const und Option Explicit', src: L(`
Option Explicit
Const MWST As Double = 0.19
Sub Main()
  Const NETTO = 100
  Dim brutto As Double
  brutto = NETTO * (1 + MWST)
  Debug.Print brutto
End Sub`), out: [' 119 '] },

    { name: 'Ohne Option Explicit: implizite Variablen', src: L(`
Sub Main()
  x = 5
  y = x * 2
  Debug.Print y; IsEmpty(z)
End Sub`), out: [' 10 Wahr'] },

    // ---------- Verzweigungen ----------
    { name: 'If / ElseIf / Else', src: L(`
Sub Main()
  Dim n As Integer
  For n = 1 To 4
    If n = 1 Then
      Debug.Print "eins"
    ElseIf n = 2 Then
      Debug.Print "zwei"
    ElseIf n = 3 Then
      Debug.Print "drei"
    Else
      Debug.Print "viel"
    End If
  Next
End Sub`), out: ['eins', 'zwei', 'drei', 'viel'] },

    { name: 'Einzeiliges If mit Else und Doppelpunkt', src: L(`
Sub Main()
  Dim a As Integer
  a = 7
  If a > 5 Then Debug.Print "groß": Debug.Print "!" Else Debug.Print "klein"
  If a < 5 Then Debug.Print "klein" Else Debug.Print "nicht klein"
  If a = 7 Then Debug.Print "sieben"
End Sub`), out: ['groß', '!', 'nicht klein', 'sieben'] },

    { name: 'Select Case mit Is, To und Listen', src: L(`
Sub Main()
  Dim n As Variant
  For Each n In Array(1, 4, 7, 12, 20, "x")
    Select Case n
      Case 1, 2, 3
        Debug.Print n; "klein"
      Case 4 To 9
        Debug.Print n; "mittel"
      Case "x"
        Debug.Print "Text"
      Case Is >= 20
        Debug.Print n; "riesig"
      Case Else
        Debug.Print n; "groß"
    End Select
  Next
End Sub`), out: [' 1 klein', ' 4 mittel', ' 7 mittel', ' 12 groß', ' 20 riesig', 'Text'] },

    { name: 'Select Case mit Strings', src: L(`
Sub Main()
  Dim note As String
  note = "B"
  Select Case note
    Case "A": Debug.Print "sehr gut"
    Case "B", "C": Debug.Print "gut"
    Case Else: Debug.Print "?"
  End Select
End Sub`), out: ['gut'] },

    // ---------- Schleifen ----------
    { name: 'For mit Step, auch rückwärts', src: L(`
Sub Main()
  Dim i As Integer, s As String
  For i = 10 To 1 Step -3
    s = s & i & " "
  Next i
  Debug.Print s; i
  For i = 1 To 0
    Debug.Print "nie"
  Next
  Dim x As Double
  For x = 0 To 1 Step 0.5
    Debug.Print x;
  Next
  Debug.Print
End Sub`), out: ['10 7 4 1 -2 ', ' 0  0,5  1 '] },

    { name: 'Do-Schleifen (alle Varianten) und While', src: L(`
Sub Main()
  Dim i As Integer
  i = 0
  Do While i < 3: i = i + 1: Loop
  Debug.Print i;
  Do Until i = 0: i = i - 1: Loop
  Debug.Print i;
  Do: i = i + 2: Loop While i < 5
  Debug.Print i;
  Do: i = i - 1: Loop Until i <= 2
  Debug.Print i;
  While i < 10: i = i * 2: Wend
  Debug.Print i
  Do
    i = i + 1
    If i > 20 Then Exit Do
  Loop
  Debug.Print i
End Sub`), out: [' 3  0  6  2  16 ', ' 21 '] },

    { name: 'Exit For in verschachtelter Schleife', src: L(`
Sub Main()
  Dim i As Integer, j As Integer, n As Integer
  For i = 1 To 3
    For j = 1 To 10
      If j > i Then Exit For
      n = n + 1
    Next j
  Next i
  Debug.Print n
End Sub`), out: [' 6 '] },

    // ---------- Prozeduren ----------
    { name: 'Sub mit und ohne Call, Function mit Rückgabe', src: L(`
Sub Main()
  Gruss "Anna"
  Call Gruss("Ben")
  Gruss ("Cem")
  Debug.Print Quadrat(7)
  Debug.Print Summe(1, 2) + Summe(3, 4)
End Sub
Sub Gruss(name As String)
  Debug.Print "Hallo " & name
End Sub
Function Quadrat(x As Double) As Double
  Quadrat = x * x
End Function
Function Summe(a, b)
  Summe = a + b
End Function`), out: ['Hallo Anna', 'Hallo Ben', 'Hallo Cem', ' 49 ', ' 10 '] },

    { name: 'ByRef (Standard) und ByVal', src: L(`
Sub Main()
  Dim a As Integer, b As Integer
  a = 1: b = 1
  Erhoehe a
  ErhoeheWert b
  Debug.Print a; b
  Erhoehe (a)
  Debug.Print a
  Dim t(1 To 2) As Integer
  Erhoehe t(2)
  Debug.Print t(2)
End Sub
Sub Erhoehe(x As Integer)
  x = x + 1
End Sub
Sub ErhoeheWert(ByVal x As Integer)
  x = x + 1
End Sub`), out: [' 2  1 ', ' 2 ', ' 1 '] },

    { name: 'Tauschen per ByRef', src: L(`
Sub Tausche(ByRef a As Variant, ByRef b As Variant)
  Dim h
  h = a: a = b: b = h
End Sub
Sub Main()
  Dim x, y
  x = "links": y = "rechts"
  Tausche x, y
  Debug.Print x; " "; y
End Sub`), out: ['rechts links'] },

    { name: 'Optional-Parameter und IsMissing', src: L(`
Function Begruessung(name As String, Optional anrede As String = "Hallo") As String
  Begruessung = anrede & " " & name
End Function
Function Info(Optional x)
  If IsMissing(x) Then Info = "fehlt" Else Info = "da: " & x
End Function
Sub Main()
  Debug.Print Begruessung("Tom")
  Debug.Print Begruessung("Tom", "Moin")
  Debug.Print Info(); " / "; Info(3)
End Sub`), out: ['Hallo Tom', 'Moin Tom', 'fehlt / da: 3'] },

    { name: 'Rekursion (Fakultät, Fibonacci)', src: L(`
Function Fak(n As Long) As Long
  If n <= 1 Then
    Fak = 1
  Else
    Fak = n * Fak(n - 1)
  End If
End Function
Function Fib(n)
  If n < 2 Then Fib = n Else Fib = Fib(n - 1) + Fib(n - 2)
End Function
Sub Main()
  Debug.Print Fak(10); Fib(15)
End Sub`), out: [' 3628800  610 '] },

    { name: 'Exit Sub und Exit Function', src: L(`
Function Pruefe(x) As String
  Pruefe = "positiv"
  If x < 0 Then Pruefe = "negativ": Exit Function
  If x = 0 Then Pruefe = "null"
End Function
Sub Main()
  Debug.Print Pruefe(-1); Pruefe(0); Pruefe(5)
  Exit Sub
  Debug.Print "nie"
End Sub`), out: ['negativnullpositiv'] },

    { name: 'Static-Variable und modulweite Variable', src: L(`
Dim zaehler As Integer
Sub Zaehle()
  Static n As Integer
  n = n + 1
  zaehler = zaehler + 10
  Debug.Print n; zaehler
End Sub
Sub Main()
  Zaehle
  Zaehle
End Sub`), out: [' 1  10 ', ' 2  20 '] },

    { name: 'Start-Sub wählbar', entry: 'Zweite', src: L(`
Sub Erste()
  Debug.Print "eins"
End Sub
Sub Zweite()
  Debug.Print "zwei"
End Sub`), out: ['zwei'] },

    { name: 'End beendet das Programm', src: L(`
Sub Main()
  Debug.Print "vorher"
  Ende
  Debug.Print "nie"
End Sub
Sub Ende()
  End
End Sub`), out: ['vorher'] },

    // ---------- Arrays ----------
    { name: 'Arrays mit Dim, LBound, UBound', src: L(`
Sub Main()
  Dim a(1 To 5) As Integer, i As Integer, s As Long
  For i = LBound(a) To UBound(a)
    a(i) = i * i
  Next
  For i = 1 To 5: s = s + a(i): Next
  Debug.Print s; LBound(a); UBound(a)
  Dim b(3) As String
  Debug.Print LBound(b); UBound(b)
End Sub`), out: [' 55  1  5 ', ' 0  3 '] },

    { name: 'ReDim und ReDim Preserve', src: L(`
Sub Main()
  Dim a() As Integer, i As Integer
  ReDim a(1 To 3)
  For i = 1 To 3: a(i) = i: Next
  ReDim Preserve a(1 To 5)
  a(5) = 50
  For i = 1 To 5: Debug.Print a(i);: Next
  Debug.Print
  ReDim a(2)
  Debug.Print a(0); UBound(a)
End Sub`), out: [' 1  2  3  0  50 ', ' 0  2 '] },

    { name: 'Zweidimensionales Array', src: L(`
Sub Main()
  Dim m(1 To 2, 1 To 3) As Integer, z As Integer, s As Integer
  For z = 1 To 2
    For s = 1 To 3
      m(z, s) = z * 10 + s
    Next
  Next
  Debug.Print m(2, 3); UBound(m, 2)
End Sub`), out: [' 23  3 '] },

    { name: 'Array, Split, Join, For Each', src: L(`
Sub Main()
  Dim teile, t, summe
  teile = Split("3;4;5", ";")
  For Each t In teile
    summe = summe + CInt(t)
  Next
  Debug.Print summe; UBound(teile)
  Debug.Print Join(Array("a", "b", "c"), "-")
End Sub`), out: [' 12  2 ', 'a-b-c'] },

    { name: 'Array an Prozedur übergeben', src: L(`
Function Summe(werte() As Integer) As Integer
  Dim i As Integer
  For i = LBound(werte) To UBound(werte)
    Summe = Summe + werte(i)
  Next
End Function
Sub Main()
  Dim z(1 To 3) As Integer
  z(1) = 1: z(2) = 2: z(3) = 3
  Debug.Print Summe(z)
End Sub`), out: [' 6 '] },

    // ---------- Eingebaute Funktionen ----------
    { name: 'String-Funktionen', src: L(`
Sub Main()
  Dim s As String
  s = "  Hallo Welt  "
  Debug.Print "[" & Trim(s) & "]"; Len(s)
  Debug.Print Left("Hallo", 2); Right("Hallo", 3); Mid("Hallo", 2, 3); Mid("Hallo", 3)
  Debug.Print InStr("Hallo", "l"); InStr(4, "Hallo", "l"); InStr("Hallo", "x")
  Debug.Print UCase("abc"); LCase("ÄBC"); Replace("a-b-c", "-", "+")
  Debug.Print Chr(65); Asc("a"); StrReverse("abc"); Space(2) & "|"; String(3, "*")
End Sub`), out: ['[Hallo Welt] 14 ', 'Halloallllo', ' 3  4  0 ', 'ABCäbca+b+c', 'A 97 cba  |***'] },

    { name: 'Umwandlung und Prüfung', src: L(`
Sub Main()
  Debug.Print Val("12.5abc"); Val("1,5"); Val("abc")
  Debug.Print CInt("7") + 1; CDbl("2,5") * 2; CStr(3.5) & "!"
  Debug.Print IsNumeric("12"); IsNumeric("12a"); IsNumeric("1,5")
  Debug.Print CBool("Wahr"); CBool(0); CLng(2.5); CLng(3.5)
End Sub`), out: [' 12,5  1  0 ', ' 8  5 3,5!', 'WahrFalschWahr', 'WahrFalsch 2  4 '] },

    { name: 'Mathematische Funktionen', src: L(`
Sub Main()
  Debug.Print Round(2.5); Round(3.5); Round(2.345, 2); Round(-2.5)
  Debug.Print Int(-3.7); Fix(-3.7); Int(3.7); Abs(-4); Sqr(16); Sgn(-2)
  Debug.Print 7 Mod 2 = 1; 2 ^ 10
End Sub`), out: [' 2  4  2,34 -2 ', '-4 -3  3  4  4 -1 ', 'Wahr 1024 '] },

    { name: 'Format', src: L(`
Sub Main()
  Debug.Print Format(1234.5, "0.00"); " | "; Format(1234.5, "#,##0.00"); " | "; Format(0.125, "0.0%")
  Debug.Print Format(7, "000"); " | "; Format(1234.5, "Currency"); " | "; Format(3.14159, "Fixed")
  Debug.Print Format(#12/24/2026#, "dd.mm.yyyy"); " | "; Format(#12/24/2026#, "dddd"); " | "; Format(#3/5/2026 14:07:09#, "hh:nn:ss")
  Debug.Print Format(True, "Yes/No"); " | "; Format("abc", ">")
End Sub`), out: ['1234,50 | 1.234,50 | 12,5%', '007 | 1.234,50 € | 3,14', '24.12.2026 | Donnerstag | 14:07:09', 'Ja | ABC'] },

    { name: 'Datumsfunktionen', src: L(`
Sub Main()
  Dim d As Date
  d = #11/26/2026#
  Debug.Print d; Year(d); Month(d); Day(d); Weekday(d)
  Debug.Print d + 7; DateAdd("m", 3, d); DateDiff("d", d, #12/24/2026#)
  Debug.Print DateSerial(2026, 2, 30); CDate("01.03.2026") - CDate("28.02.2026")
  Debug.Print IsDate("31.12.2026"); IsDate("32.12.2026"); Year(Now) >= 2024
End Sub`), out: ['26.11.2026 2026  11  26  5 ', '03.12.202626.02.2027 28 ', '02.03.2026 1 ', 'WahrFalschWahr'] },

    { name: 'Rnd ohne Randomize liefert feste Folge', src: L(`
Sub Main()
  Dim i As Integer, ok As Boolean
  ok = True
  For i = 1 To 100
    Dim z As Integer
    z = Int(Rnd * 6) + 1
    If z < 1 Or z > 6 Then ok = False
  Next
  Debug.Print ok
  Randomize
  Debug.Print Rnd() < 1
End Sub`), out: ['Wahr', 'Wahr'] },

    { name: 'IIf, Nz, Like, Konstanten', src: L(`
Sub Main()
  Debug.Print IIf(3 > 2, "ja", "nein"); Nz(Null, 0); "Meier" Like "M*r"; "A1" Like "[A-Z]#"
  Debug.Print Len(vbCrLf); vbYes; vbOKCancel
End Sub`), out: ['ja 0 WahrWahr', ' 2  6  1 '] },

    // ---------- MsgBox / InputBox ----------
    { name: 'InputBox und MsgBox', inputs: ['5', '7'], msgRet: [1, 6], src: L(`
Sub Main()
  Dim a As Integer, b As Integer
  a = InputBox("Erste Zahl:")
  b = InputBox("Zweite Zahl:", "Eingabe", "0")
  MsgBox "Summe: " & (a + b)
  If MsgBox("Nochmal?", vbYesNo) = vbYes Then Debug.Print "ja"
End Sub`), out: ['ja'], msg: ['Summe: 12', 'Nochmal?'] },

    { name: 'InputBox liefert Text: + verkettet!', inputs: ['2', '3'], src: L(`
Sub Main()
  Dim a, b
  a = InputBox("a")
  b = InputBox("b")
  Debug.Print a + b; " "; Val(a) + Val(b)
End Sub`), out: ['23  5 '] },

    { name: 'InputBox abbrechen liefert ""', inputs: [null], src: L(`
Sub Main()
  Dim s As String
  s = InputBox("Name?")
  Debug.Print "[" & s & "]"; Len(s)
End Sub`), out: ['[] 0 '] },

    // ---------- Syntaxfehler ----------
    { name: 'Fehler: Next ohne For', access: 'Next ohne For', src: L(`
Sub Main()
  Dim i
  i = 1
  Next
End Sub`), error: /Next ohne For/, line: 4 },

    { name: 'Fehler: For ohne Next', access: 'For ohne Next', src: L(`
Sub Main()
  Dim i
  For i = 1 To 3
    Debug.Print i
End Sub`), error: /For ohne Next/, line: 3 },

    { name: 'Fehler: If ohne End If', access: 'Block If ohne End If', src: L(`
Sub Main()
  If 1 = 1 Then
    Debug.Print "x"
End Sub`), error: /If-Block ohne End If/, line: 2 },

    { name: 'Fehler: End If ohne If', access: 'End If ohne Block If', src: L(`
Sub Main()
  Debug.Print "x"
  End If
End Sub`), error: /End If ohne If/, line: 3 },

    { name: 'Fehler: Loop ohne Do', access: 'Loop ohne Do', src: L(`
Sub Main()
  Loop
End Sub`), error: /Loop ohne Do/, line: 2 },

    { name: 'Fehler: Sub ohne End Sub', access: 'Erwartet: End Sub', src: L(`
Sub Main()
  Debug.Print 1
Sub Zwei()
End Sub`), error: /Sub ohne End Sub/, line: 1 },

    { name: 'Fehler: Option Explicit, Variable nicht deklariert', access: 'Variable nicht definiert', src: L(`
Option Explicit
Sub Main()
  Dim summe As Integer
  summe = 1
  sume = summe + 1
End Sub`), error: /Variable 'sume' wurde nicht deklariert/, line: 5 },

    { name: 'Fehler: unbekannte Funktion', access: 'Sub oder Function nicht definiert', src: L(`
Sub Main()
  Debug.Print Lenn("abc")
End Sub`), error: /Sub oder Function 'Lenn' ist nicht definiert/, line: 2 },

    { name: 'Fehler: Klammern beim Sub-Aufruf mit 2 Argumenten', access: 'Erwartet: =', src: L(`
Sub Main()
  Ausgabe(1, 2)
End Sub
Sub Ausgabe(a, b)
End Sub`), error: /ohne Klammern/, line: 2 },

    { name: 'Fehler: fehlendes Then', access: 'Erwartet: Then oder GoTo', src: L(`
Sub Main()
  If 1 = 1
    Debug.Print 1
  End If
End Sub`), error: /Erwartet: 'Then'/, line: 2 },

    { name: 'Fehler: offene Zeichenkette', access: 'Syntaxfehler', src: L(`
Sub Main()
  Debug.Print "Hallo
End Sub`), error: /Zeichenkette ist nicht abgeschlossen/, line: 2 },

    { name: 'Fehler: Anweisung außerhalb einer Prozedur', access: 'Ungültig außerhalb einer Prozedur', src: L(`
Debug.Print "x"`), error: /außerhalb einer Prozedur/, line: 1 },

    { name: 'Fehler: Zuweisung an Konstante', access: 'Zuweisung zu einer Konstanten nicht zulässig', src: L(`
Sub Main()
  Const PI = 3.14
  PI = 3
End Sub`), error: /Konstanten/, line: 3 },

    { name: 'Fehler: Argument nicht optional', access: 'Argument ist nicht optional', src: L(`
Sub Main()
  Debug.Print Doppelt()
End Sub
Function Doppelt(x)
  Doppelt = 2 * x
End Function`), error: /Argument nicht optional/, line: 2 },

    { name: 'Fehler: Exit For außerhalb', access: 'Exit For nicht innerhalb For...Next', src: L(`
Sub Main()
  Exit For
End Sub`), error: /Exit For ist nur innerhalb/, line: 2 },

    { name: 'Fehler: Doppelte Deklaration', access: 'Doppelte Deklaration im aktuellen Gültigkeitsbereich', src: L(`
Sub Main()
  Dim a As Integer
  Dim a As String
End Sub`), error: /Doppelte Deklaration/, line: 3 },

    { name: 'Fehler: Me. wird nicht benötigt', access: 'Ungültige Verwendung des Me-Schlüsselwortes', src: L(`
Sub Main()
  Me.Text1 = 5
End Sub`), error: /Me\./, line: 2 },

    // ---------- Laufzeitfehler ----------
    { name: 'Laufzeit: Typen unverträglich', access: 'Typen unverträglich', src: L(`
Sub Main()
  Dim x As Integer
  x = "abc"
End Sub`), error: /Typen unverträglich \| "abc" ist keine Zahl/, line: 3 },

    { name: 'Laufzeit: Überlauf bei Integer', access: 'Überlauf', src: L(`
Sub Main()
  Dim x As Integer
  x = 30000
  x = x + 5000
End Sub`), error: /Überlauf/, line: 4 },

    { name: 'Laufzeit: Division durch Null', access: 'Division durch Null', src: L(`
Sub Main()
  Dim a, b
  a = 5
  Debug.Print a / b
End Sub`), error: /Division durch Null/, line: 4 },

    { name: 'Laufzeit: Index außerhalb', access: 'Index außerhalb des gültigen Bereichs', src: L(`
Sub Main()
  Dim a(1 To 3) As Integer
  a(4) = 1
End Sub`), error: /Index 4 liegt nicht zwischen 1 und 3/, line: 3 },

    { name: 'Laufzeit: Fehler in aufgerufener Function meldet deren Zeile', access: 'Division durch Null', src: L(`
Sub Main()
  Debug.Print Teile(1, 0)
End Sub
Function Teile(a, b)
  Teile = a / b
End Function`), error: /Division durch Null/, line: 5 },

    { name: 'Laufzeit: ByRef-Typ unverträglich', access: 'ByRef-Argumenttyp unverträglich', src: L(`
Sub Main()
  Dim z As Long
  Verdopple z
End Sub
Sub Verdopple(x As Integer)
  x = x * 2
End Sub`), error: /ByRef-Argumenttyp unverträglich/, line: 3 },

    { name: 'Laufzeit: Endlosschleife wird abgebrochen', access: 'Ausführung des Codes wurde unterbrochen', maxSteps: 50000, src: L(`
Sub Main()
  Dim i As Integer
  Do While i < 10
    Debug.Print "ups";
  Loop
End Sub`), error: /Endlosschleife/, line: 4 },

    { name: 'Laufzeit: leere Endlosschleife wird abgebrochen', maxSteps: 50000, src: L(`
Sub Main()
  Do
  Loop
End Sub`), error: /Endlosschleife/ },

    { name: 'Laufzeit: Rekursion ohne Ende', access: 'Nicht genügend Stapelspeicher', src: L(`
Sub Main()
  Endlos 1
End Sub
Sub Endlos(n)
  Endlos n + 1
End Sub`), error: /Stapelspeicher/ },

    { name: 'Laufzeit: Mid mit Start 0', access: 'Ungültiger Prozeduraufruf oder ungültiges Argument', src: L(`
Sub Main()
  Debug.Print Mid("abc", 0, 1)
End Sub`), error: /Startposition bei Mid/, line: 2 },

    // ---------- Formulare (Phase 2) ----------
    { name: 'Formular: Textfelder lesen und schreiben, Button-Klick', form: FORM1,
      actions: [{ input: ['E_Zahl1', '2'] }, { input: ['E_Zahl2', '3,5'] }, { click: 'btnBerechnen' }],
      controls: { A_Ergebnis: '5,5', lblInfo: 'Fertig' }, src: L(`
Option Explicit
Private Sub btnBerechnen_Click()
    Dim zahl1 As Double, zahl2 As Double
    zahl1 = E_Zahl1
    zahl2 = E_Zahl2
    A_Ergebnis = zahl1 + zahl2
    lblInfo.Caption = "Fertig"
End Sub`) },

    { name: 'Formular: Textfeld-Werte sind Text ("2" + "3" = "23")', form: FORM1,
      actions: [{ input: ['E_Zahl1', '2'] }, { input: ['E_Zahl2', '3'] }, { click: 'btnBerechnen' }],
      controls: { A_Ergebnis: '23' }, src: L(`
Private Sub btnBerechnen_Click()
    A_Ergebnis = E_Zahl1 + E_Zahl2
End Sub`) },

    { name: 'Formular: leeres Textfeld ist Null', form: FORM1,
      actions: [{ click: 'btnBerechnen' }], out: ['Wahr', ' 0 '], src: L(`
Private Sub btnBerechnen_Click()
    Debug.Print IsNull(E_Zahl1)
    Debug.Print Nz(E_Zahl1, 0)
End Sub`) },

    { name: 'Formular: leeres Feld an Integer zuweisen → Laufzeitfehler 94', form: FORM1,
      actions: [{ click: 'btnBerechnen' }], error: /Unzulässige Verwendung von Null/, access: 'Unzulässige Verwendung von Null', line: 3, src: L(`
Private Sub btnBerechnen_Click()
    Dim x As Integer
    x = E_Zahl1
End Sub`) },

    { name: 'Formular: Modulvariablen bleiben zwischen Klicks erhalten, Form_Load', form: FORM1,
      actions: [{ click: 'btnBerechnen' }, { click: 'btnBerechnen' }, { click: 'btnBerechnen' }],
      controls: { A_Ergebnis: '13' }, src: L(`
Dim zaehler As Integer
Private Sub Form_Load()
    zaehler = 10
End Sub
Private Sub btnBerechnen_Click()
    zaehler = zaehler + 1
    A_Ergebnis = zaehler
End Sub`) },

    { name: 'Formular: Optionsgruppe', form: FORM1,
      actions: [{ select: ['E_Art', 2] }, { click: 'btnBerechnen' }], out: ['Rabatt'], src: L(`
Private Sub btnBerechnen_Click()
    Select Case E_Art
        Case 1: Debug.Print "Normal"
        Case 2: Debug.Print "Rabatt"
    End Select
    E_Art = 1
    If E_Art.Value <> 1 Then Debug.Print "falsch"
End Sub`) },

    { name: 'Formular: Eigenschaften Visible, Enabled, Me.', form: FORM1,
      actions: [{ click: 'btnBerechnen' }], out: ['Falsch Wahr x'], src: L(`
Private Sub btnBerechnen_Click()
    A_Ergebnis.Visible = False
    Me.A_Ergebnis = "x"
    btnBerechnen.Enabled = True
    Debug.Print A_Ergebnis.Visible & " " & btnBerechnen.Enabled & " " & Me.A_Ergebnis.Value
End Sub`) },

    { name: 'Formular: Zuweisung an Bezeichnungsfeld → Laufzeitfehler 438', form: FORM1,
      actions: [{ click: 'btnBerechnen' }], error: /Caption/, access: 'Objekt unterstützt diese Eigenschaft oder Methode nicht', line: 2, src: L(`
Private Sub btnBerechnen_Click()
    lblInfo = "Hallo"
End Sub`) },

    { name: 'Formular: unbekannte Eigenschaft', form: FORM1, error: /Value/, access: 'Methode oder Datenobjekt nicht gefunden', line: 2, src: L(`
Private Sub btnBerechnen_Click()
    lblInfo.Value = 1
End Sub`) },

    { name: 'Formular: Option Explicit kennt die Steuerelemente', form: FORM1,
      actions: [{ input: ['E_Zahl1', '4'] }, { click: 'btnBerechnen' }], controls: { A_Ergebnis: '16' }, src: L(`
Option Explicit
Private Sub btnBerechnen_Click()
    A_Ergebnis = Quadrat(E_Zahl1)
End Sub
Function Quadrat(ByVal x As Double) As Double
    Quadrat = x * x
End Function`) },

    { name: 'Formular: Variable mit Steuerelementnamen → Mehrdeutiger Name', form: FORM1,
      error: /Mehrdeutiger Name/, access: 'Mehrdeutiger Name erkannt: E_Zahl1', line: 1, src: L(`
Dim E_Zahl1 As Integer
Private Sub btnBerechnen_Click()
End Sub`) },
    { name: 'Keine startbare Sub', src: L(`
Function F()
  F = 1
End Function`), error: /Keine startbare Sub/ },
  ];

  /** Führt einen Fall aus und liefert { ok, details }. */
  async function runCase(VBA, c) {
    const out = [];
    const msgs = [];
    const inputs = (c.inputs || []).slice();
    const msgRet = (c.msgRet || []).slice();
    const host = {
      print: (l) => out.push(l),
      msgBox: async (p) => { msgs.push(p); return msgRet.length ? msgRet.shift() : 1; },
      inputBox: async () => (inputs.length ? inputs.shift() : '')
    };
    let error = null;
    let runtime = null;
    try {
      if (c.form) {
        runtime = new VBA.FormRuntime(c.form);
        const prog = VBA.compile(c.src, { controls: c.form.controls });
        const session = VBA.createRun(prog, host, { maxSteps: c.maxSteps, controls: runtime.adapters() });
        session.open();
        if (session.has('Form_Load')) await session.invoke('Form_Load');
        for (const a of c.actions || []) {
          if (a.input) runtime.userInput(a.input[0], a.input[1]);
          if (a.select) runtime.userSelect(a.select[0], a.select[1]);
          if (a.click) await session.invoke(a.click + '_Click');
        }
      } else {
        const prog = VBA.compile(c.src);
        await VBA.createRun(prog, host, { maxSteps: c.maxSteps }).start(c.entry);
      }
    } catch (e) {
      error = e;
    }
    const problems = [];
    if (c.error) {
      if (!error) problems.push('Fehler erwartet, aber keiner aufgetreten');
      else {
        const full = error.message + ' | ' + (error.hint || '');
        if (!c.error.test(full)) problems.push(`Fehlermeldung passt nicht: ${full}`);
        if (c.access && error.accessMessage !== c.access) problems.push(`Access-Text "${error.accessMessage}" statt "${c.access}"`);
        if (c.line !== undefined && error.line !== c.line) problems.push(`Zeile ${error.line} statt ${c.line}`);
      }
    } else if (error) {
      problems.push(`Unerwarteter Fehler: ${error.message}`);
    }
    if (c.out) {
      const a = JSON.stringify(out), b = JSON.stringify(c.out);
      if (a !== b) problems.push(`Ausgabe:\n    erhalten: ${a}\n    erwartet: ${b}`);
    }
    if (c.controls && runtime) {
      for (const [n, want] of Object.entries(c.controls)) {
        const st = runtime.get(n);
        const got = st.def.kind === 'label' || st.def.kind === 'button' ? st.caption : runtime.displayText(n);
        if (got !== want) problems.push(`Steuerelement ${n}: "${got}" statt "${want}"`);
      }
    }
    if (c.msg) {
      const a = JSON.stringify(msgs), b = JSON.stringify(c.msg);
      if (a !== b) problems.push(`MsgBox:\n    erhalten: ${a}\n    erwartet: ${b}`);
    }
    return { ok: problems.length === 0, details: problems.join('\n  ') };
  }

  const api = { cases, runCase };
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.VBATests = api;
})(typeof self !== 'undefined' ? self : this);
