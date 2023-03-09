VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Red, NKonto, BA21, Q29, S29
Private Sub Workbook_Open()


    With Application
        .ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
        .CommandBars("Status Bar").Visible = True
        .DisplayFormulaBar = True
        .DisplayScrollBars = True
        .Caption = "KNJIGOVODSTVENI PROGRAM FATALITI"
                .ActiveWindow.DisplayWorkbookTabs = True
    End With

End Sub
Sub Otvaranje(ByVal control As IRibbonControl)
    Sheets("MATRICA").Select
    NKonto = Range("C1")
    Sheets("GLKNJIGA").Select
    Range("C1").Select
    Red = Selection.Row
    Konto = Range("C1")
2   If Konto = "" Then
1       If Red = 1 Then
            Smesti
        Else
            Selection.End(xlUp).Select
            Red = Selection.Row + 6
            Smesti
        End If
        End
    End If
    If NKonto = Konto Then
        MsgBox ("Konto  " & NKonto & "  vec postoji")
        End
    End If
    If NKonto > Konto Then
        Red = Red + 3 + Cells(Red + 1, 8)
        Cells(Red, 3).Select
        Selection.End(xlDown).Select
        Red = Selection.Row
        Konto = Selection
        GoTo 2
    Else
        Range(Cells(Red, 1), Cells(Red + 8, 8)).Select
        Application.CutCopyMode = False
        Selection.EntireRow.Insert
        Smesti
    End If
       ' Application.CutCopyMode = False
End Sub
Sub Smesti()
    Kolona = Red + 5 + Cells(Red + 1, 8)
    Adresa = "=GLKNJIGA!$A$" & Red & ":$H$" & Kolona
    Sheets("MATRICA").Select
    Range("A1:H4").Select
    Selection.Copy
    Sheets("GLKNJIGA").Select
    Cells(Red, 1).Select
    ActiveSheet.Paste
    Names.Add "_KK" & NKonto, Adresa
End Sub
Sub Dodavanje_fin_nalog(ByVal control As IRibbonControl)
    Sheets("NALOG").Select
    m1 = Range("M1") + 1
    novired = Application.InputBox("Koliko novih redova zelite?", "DODAVANJE REDOVA", Type:=1)
    If novired = False Then End
    Range(Cells(m1, 1), Cells(m1 + 5, 23)).Select
    Selection.Cut
    Range(Cells(m1 + novired, 1), Cells(m1 + novired, 1)).Select
    ActiveSheet.Paste
    Range(Cells(m1 - 1, 1), Cells(m1 - 1, 23)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Cells(m1, 1), Cells(m1 + novired - 1, 23)).Select
    ActiveSheet.Paste
    Range("M1") = m1 - 1 + novired
    Application.CutCopyMode = False
    Range("A1").Select
End Sub
Sub Oduzimanje_fin_nalog(ByVal control As IRibbonControl)
    Sheets("NALOG").Select
    m1 = Range("M1") + 1
    If m1 < 5 Then MsgBox ("Ne mozete oduzimati redove.")
    If m1 < 5 Then End
    novired = Application.InputBox("Koliko redova zelite da izbrisete?", "ODUZIMANJE REDOVA", Type:=1)
    If novired = False Then End
    If novired > m1 - 4 Then MsgBox ("Ne mozete oduzeti toliko redova.")
    If novired > m1 - 4 Then End
    Range(Cells(m1, 1), Cells(m1 + 105, 23)).Select
    Selection.Cut
    Range(Cells(m1 - novired, 1), Cells(m1 - novired, 1)).Select
    ActiveSheet.Paste
    Range("M1") = m1 - novired - 1
    Application.CutCopyMode = False
    Range("A1").Select
    
End Sub
Sub Knjizenje_fin(ByVal control As IRibbonControl)
    With Application
        .Calculation = xlAutomatic
    End With
    Provera
    Range("F1") = "KNJIZEN"
    oldStatusBar = Application.DisplayStatusBar
    Application.DisplayStatusBar = True
    For BA20 = 3 To BA21
        Sheets("NALOG").Select
        Konto = "_KK" & Cells(BA20, 4)
        Application.StatusBar = "Knjizenje na kartici   " & Konto
        Range(Cells(BA20, 15), Cells(BA20, 22)).Select
        Selection.Copy
        Sheets("GLKNJIGA").Select
        Red = 5 + Range(Konto).Cells(2, 8)
        Range(Konto).Cells(Red, 1).Select
        ERed = Selection.Row
        Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
        Selection.PasteSpecial Paste:=xlFormats, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
        Range(Konto).Cells(Red + 1, 1).Select
        Application.CutCopyMode = False
        Selection.EntireRow.Insert
        If Red = 5 Then
            Range(Konto).Cells(Red, 7).Formula = "=E" & ERed & "-F" & ERed
        Else
            Range(Konto).Cells(Red, 7).Formula = "=G" & ERed - 1 & "+E" & ERed & "-F" & ERed
        End If
        Range(Konto).Cells(2, 8) = Range(Konto).Cells(2, 8) + 1
        Range(Konto).Select
        Adresa = "=GLKNJIGA!" & Selection.Address
        Names.Add Konto, Adresa

    Next BA20
    Dnev
    Application.StatusBar = False
    Application.DisplayStatusBar = oldStatusBar
    With Application
        .Calculation = xlAutomatic
    End With
    
    Sheets("NALOG").Select
    Range("A1").Select
End Sub
Sub Provera()
    If Abs(Range("N1")) > 0.0001 Then
        MsgBox ("NALOG NIJE U RAVNOTEZI")
        End
    End If
    On Error Resume Next
    If Range("E1") < 30000 Then
        Err = 0
        MsgBox ("DATUM NIJE DOBRO UPISAN")
        End
    End If
    On Error GoTo 0
    If Range("F1") = "KNJIZEN" Then
        MsgBox ("NALOG JE VEC KNJIZEN")
        End
    End If
    With Application
        .Calculation = xlManual
    End With
    
    BA20 = 2
    oldStatusBar = Application.DisplayStatusBar
    Application.DisplayStatusBar = True
1   BA20 = BA20 + 1
    If BA20 > Range("M1") Then GoTo 2
    If Cells(BA20, 4) = "" Then GoTo 2
    BC20 = Cells(BA20, 4)
    Konto = "_KK" & BC20
    Application.StatusBar = "Provera kartice   " & Konto
    On Error Resume Next
    Proba = Range(Konto).Cells(1, 3)
    If Err Then
        Err = 0
        MsgBox (" NE POSTOJI KARTICA  " & BC20)
        Application.StatusBar = False
        Application.DisplayStatusBar = oldStatusBar
        End
    End If
    On Error GoTo 0
    BA21 = BA20
    GoTo 1
2   BC20 = ""
   
    Application.StatusBar = False
    Application.DisplayStatusBar = oldStatusBar
  
End Sub
Sub Dnev()
        Application.StatusBar = "Knjizenje u dnevniku"
        Sheets("DNEVNIK").Select
        Red = 5 + Cells(2, 8)
        Range(Cells(Red + 1, 1), Cells(Red + BA21 - 2, 1)).Select
        Application.CutCopyMode = False
        Selection.EntireRow.Insert
        Sheets("NALOG").Select
        Range(Cells(3, 15), Cells(BA21, 23)).Select
        Selection.Copy
        Sheets("DNEVNIK").Select
        Cells(Red, 1).Select
        Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
        Selection.PasteSpecial Paste:=xlFormats, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
        Cells(Red, 7).Formula = "=G" & Red - 1 & "+E" & Red & "-F" & Red
        Cells(Red, 7).Select
        Selection.Copy
        Range(Cells(Red + 1, 7), Cells(Red + BA21 - 3, 7)).Select
        ActiveSheet.Paste
        If Red = 5 Then Cells(5, 7).Formula = "=E5-F5"
        Cells(2, 8) = Cells(2, 8) + BA21 - 2
End Sub
Sub Novinalog_fin(ByVal control As IRibbonControl)
    Sheets("NALOG").Select
    Cells.Select
    Selection.Clear
    Sheets("MATRICA").Select
    Range("AA1:AW20").Select
    Selection.Copy
    Sheets("NALOG").Select
    Range("A1").Select
    ActiveSheet.Paste
    Sheets("MATRICA").Select
    Range("A1").Select
    Sheets("NALOG").Select
    Application.CutCopyMode = False
    Range("A1").Select
End Sub
Sub Kontrola()
    m1 = Range("M1")
    Range(Cells(1, 1), Cells(m1 + 3, 6)).Select
    Application.DataEntryMode = xlOn
    
End Sub
Sub fiksiranje()
'
' fiksiranje Macro
' fiksira datume
'
' Keyboard Shortcut: Ctrl+f
'
    Sheets("ODLIV").Select
    Range("AX10:BC54").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("AX10:AZ54").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("ULIV").Select
    Range("C10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("ODLIV").Select
    Range("BA12:BC54").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C12").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Sheets("ODLIV").Select
End Sub
Sub Stavke_fin(ByVal control As IRibbonControl)
    RACUN = Application.InputBox("UPISITE KONTO KOMITENTA", "OTVORENE STAVKE KOMITENATA", Type:=2)
    If RACUN = False Then End
    If RACUN = "" Then End
    Konto = "_KK" & RACUN
    Application.GoTo Reference:=Konto
    Selection.Copy
    Sheets("DUGOVI").Select
    Range("A1").Select
    ActiveSheet.Paste
    Range("g4").ClearContents
    If Range("G2") = 0 Then End
    If Range("H2") = 1 Then End
    AI5 = Range("H2")
    Range(Cells(5, 1), Cells(4 + AI5, 8)).Select
    Selection.Sort Key1:=Range("H5"), Order1:=xlAscending, Key2:=Range("B5") _
        , Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom
    Cells(5 + AI5, 8) = "****"
    AG1 = 0
    AG2 = 5
    Ag3 = 5
    AH1 = Range("H5")
2   Ag3 = Ag3 + 1
    If Ag3 > 5 + AI5 Then GoTo 1
    AH2 = Cells(Ag3, 8)
    If AH1 = AH2 Then GoTo 2
    Ag3 = Ag3 - 1
    AH3 = Cells(Ag3, 7)
    If Abs(AG1 - AH3) < 0.0001 Then GoTo 3
    Cells(Ag3, 9) = AH3 - AG1
    Cells(Ag3, 9).Select
    Selection.NumberFormat = "#,##0.00_);(#,##0.00)"
    AG1 = AH3
    AG2 = Ag3 + 1
    AH1 = Cells(Ag3 + 1, 8)
    GoTo 2
3   Range(Cells(Ag3 + 1, 1), Cells(100 + AI5, 8)).Select
    Application.CutCopyMode = False
    Selection.Cut
    Cells(AG2, 1).Select
    ActiveSheet.Paste
    AI5 = AI5 - (Ag3 - AG2 + 1)
    Range("H2") = AI5
    Cells(5, 7).Formula = "=G4+E5-F5"
    Cells(5, 7).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Cells(5, 7), Cells(5 + AI5, 7)).Select
    ActiveSheet.Paste
    AH1 = AH2
    Ag3 = AG2
    GoTo 2
1   Range("E2").Formula = "=sum(E5:E" & 5 + AI5 & ")"
    Range("F2").Formula = "=sum(F5:F" & 5 + AI5 & ")"
    Range("I4") = "Dug"
    
End Sub
Sub Bruto_fin(ByVal control As IRibbonControl)
    Sheets("BRUTO").Select
    oldStatusBar = Application.DisplayStatusBar
    Application.DisplayStatusBar = True
    Application.StatusBar = "Bruto"
    Range("Q14") = "Sifra"
    Range("P18") = "Sifra"
    Range("Q18") = "Sifra"
    Range("Q15") = "_KK*"
    Range("Q19").Select
    Selection.ListNames
    Selection.End(xlDown).Select
    Red = Selection.Row
    Range(Cells(18, 17), Cells(Red, 17)).AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range _
        ("Q14:Q15"), CopyToRange:=Range("P18"), Unique:=False
    Range("P18").Select
    Selection.End(xlDown).Select
    Red = Selection.Row
    If Red > 50000 Then Range("A1").Select
    If Red > 50000 Then MsgBox ("Greska.")
    If Red > 50000 Then End
    Range(Cells(21, 1), Cells(21 + 5, 7)).Select
    Selection.Cut
    Cells(Red + 1, 1).Select
    ActiveSheet.Paste
    Range(Cells(20, 1), Cells(20, 7)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Cells(20, 1), Cells(Red, 7)).Select
    ActiveSheet.Paste
    For Q29 = 19 To Red
        Application.StatusBar = "Cita  " & Q29
        Konto = Cells(Q29, 16)
        Cells(Q29, 4) = Range(Konto).Cells(2, 5)
        Cells(Q29, 5) = Range(Konto).Cells(2, 6)
    Next Q29
    Application.StatusBar = "Bruto"
    Range(Cells(19, 1), Cells(Red + 2, 7)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
    For AZ1 = 1 To 10
        Range("BB26") = "" & AZ1 - 1 & "*"
        Range(Cells(7, 1), Cells(Red, 5)).AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range( _
            "BB25:BB26"), CopyToRange:=Range("BB28:BD28"), Unique:=False
        Cells(AZ1 + 7, 4) = Range("BC26")
        Cells(AZ1 + 7, 5) = Range("BD26")
    Next AZ1

    Q29 = 18
2   Q29 = Q29 + 1
3   If Q29 > Red Then GoTo 1
    Application.StatusBar = "Sredjuje  " & Q29
    If Cells(Q29 + 1, 1) = Cells(Q29, 1) Then
        Sabiraj
        GoTo 3
    End If
    GoTo 2
    
1   Application.StatusBar = "Bruto"
    Q29 = 19
4   Ubacivanje

    If Q29 <= Red Then GoTo 4
End Sub
Sub Sabiraj()
    q1 = Q29
2   q1 = q1 + 1
    If q1 > Red Then GoTo 1
    If Cells(q1 + 1, 1) = Cells(Q29, 1) Then GoTo 2
1   Cells(Q29, 4) = WorksheetFunction.Sum(Range(Cells(Q29, 4), Cells(q1, 4)))
    Cells(Q29, 5) = WorksheetFunction.Sum(Range(Cells(Q29, 5), Cells(q1, 5)))
    Cells(Q29, 6) = WorksheetFunction.Sum(Range(Cells(Q29, 6), Cells(q1, 6)))
    Cells(Q29, 7) = WorksheetFunction.Sum(Range(Cells(Q29, 7), Cells(q1, 7)))
    Range(Cells(Q29 + 1, 1), Cells(q1, 1)).Select
    Selection.EntireRow.Delete
    Red = Red - (q1 - Q29)
End Sub
Sub Ubacivanje()
    Cells(Q29, 1).Select
    Range("S1") = Cells(Q29, 1)
    Application.CutCopyMode = False
    Selection.EntireRow.Insert
    Cells(Q29, 1) = Range("S2")
    
    Range(Cells(Q29, 1), Cells(Q29, 7)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    With Selection.Interior
        .ColorIndex = 6
        .Pattern = xlSolid
    End With
    Ime = "PLAN" & Range("T2")
    Cells(Q29, 2) = WorksheetFunction.VLookup(Cells(Q29, 1), Range(Ime), 2, False)
    S29 = Q29
    Red = Red + 1
    Zbroji
    
End Sub
Sub Zbroji()
2   Q29 = Q29 + 1
    If Q29 > Red Then GoTo 1
    Range("S1") = Cells(Q29, 1)
    If Cells(S29, 1) = Range("S2") Then
        Ime = "PLAN" & Range("T2")
    On Error Resume Next
       Cells(Q29, 2) = WorksheetFunction.VLookup(Cells(Q29, 1), Range(Ime), 2, False)
    On Error GoTo 0
        GoTo 2
    End If
    
1   Cells(S29, 4) = WorksheetFunction.Sum(Range(Cells(S29 + 1, 4), Cells(Q29 - 1, 4)))
    Cells(S29, 5) = WorksheetFunction.Sum(Range(Cells(S29 + 1, 5), Cells(Q29 - 1, 5)))
    Cells(S29, 6) = WorksheetFunction.Sum(Range(Cells(S29 + 1, 6), Cells(Q29 - 1, 6)))
    Cells(S29, 7) = WorksheetFunction.Sum(Range(Cells(S29 + 1, 7), Cells(Q29 - 1, 7)))
    Range(Cells(S29, 4), Cells(S29, 7)).Select
    Selection.NumberFormat = "#,##0.00_);(#,##0.00)"
    Selection.Font.Bold = True
    
End Sub
Sub Dodavanje_uliv(ByVal control As IRibbonControl)
    Sheets("ULIV").Select
    AM1 = Range("AM1") + 1
    If AM1 < 16 Then End
    novired = Application.InputBox("KOLIKO NOVIH REDOVA", "DODAVANJE REDOVA", Type:=1)
    If novired = False Then End
    Range(Cells(AM1, 1), Cells(AM1 + 5, 40)).Select
    Selection.Cut
    Range(Cells(AM1 + novired, 1), Cells(AM1 + novired, 1)).Select
    ActiveSheet.Paste
    Range(Cells(AM1 - 1, 1), Cells(AM1 - 1, 40)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Cells(AM1, 1), Cells(AM1 + novired - 1, 40)).Select
    ActiveSheet.Paste
    Range("AM1") = AM1 - 1 + novired
End Sub
Sub Oduzimanje_uliv(ByVal control As IRibbonControl)
    Sheets("ULIV").Select
    AM1 = Range("AM1") + 1
    If AM1 < 16 Then End
    Range(Cells(AM1, 1), Cells(AM1 + 5, 40)).Select
    Selection.Cut
    Range(Cells(AM1 - 1, 1), Cells(AM1 - 1, 1)).Select
    ActiveSheet.Paste
    Range("AM1") = AM1 - 2
End Sub
Sub Dodavanje_odliv(ByVal control As IRibbonControl)
    Sheets("ODLIV").Select
    AM1 = Range("AM1") + 1
    If AM1 < 16 Then End
    novired = Application.InputBox("KOLIKO NOVIH REDOVA", "DODAVANJE REDOVA", Type:=1)
    If novired = False Then End
    Range(Cells(AM1, 1), Cells(AM1 + 5, 40)).Select
    Selection.Cut
    Range(Cells(AM1 + novired, 1), Cells(AM1 + novired, 1)).Select
    ActiveSheet.Paste
    Range(Cells(AM1 - 1, 1), Cells(AM1 - 1, 40)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Cells(AM1, 1), Cells(AM1 + novired - 1, 40)).Select
    ActiveSheet.Paste
    Range("AM1") = AM1 - 1 + novired
End Sub
Sub Oduzimanje_odliv(ByVal control As IRibbonControl)
    Sheets("ODLIV").Select
    AM1 = Range("AM1") + 1
    If AM1 < 16 Then End
    Range(Cells(AM1, 1), Cells(AM1 + 5, 40)).Select
    Selection.Cut
    Range(Cells(AM1 - 1, 1), Cells(AM1 - 1, 1)).Select
    ActiveSheet.Paste
    Range("AM1") = AM1 - 2
End Sub
Sub FIN_nalog(ByVal control As IRibbonControl)
Sheets("NALOG").Select
End Sub
Sub FIN_matricu(ByVal control As IRibbonControl)
Sheets("MATRICA").Select
Range("C1").Select
End Sub
Sub FIN_dnevnik(ByVal control As IRibbonControl)
Sheets("DNEVNIK").Select
Range("C1").Select
End Sub
Sub FIN_glavnuknj(ByVal control As IRibbonControl)
Sheets("GLKNJIGA").Select
Range("C1").Select
End Sub
Sub FIN_dugove(ByVal control As IRibbonControl)
Sheets("DUGOVI").Select
Range("C1").Select
End Sub
Sub FIN_komiten(ByVal control As IRibbonControl)
Sheets("KOMITE").Select
Range("C1").Select
End Sub
Sub FIN_bruto(ByVal control As IRibbonControl)
Sheets("BRUTO").Select
Range("C1").Select
End Sub
Sub FIN_konplan(ByVal control As IRibbonControl)
Sheets("KONPLAN").Select
End Sub
Sub FIN_ios(ByVal control As IRibbonControl)
Sheets("IOS").Select
Range("C1").Select
End Sub
Sub FIN_odliv(ByVal control As IRibbonControl)
Sheets("ODLIV").Select
Range("C1").Select
End Sub
Sub Graf_lik(ByVal control As IRibbonControl)
Sheets("OCODLIV").Select
End Sub
Sub FIN_uliv(ByVal control As IRibbonControl)
Sheets("ULIV").Select
Range("C1").Select
End Sub
Sub Graf_pril(ByVal control As IRibbonControl)
Sheets("OCPRIL").Select
End Sub

Sub NoviBruto(ByVal control As IRibbonControl)
  Sheets("BRUTOM").Select
    Cells.Select
    Selection.Copy
    Sheets("BRUTO").Select
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("C1").Select
End Sub
Sub Uputstvo(ByVal control As IRibbonControl)

Dim myShell As Object
Set myShell = CreateObject("WScript.Shell")
myShell.Run "C:\FATALITI\Uputstvo.doc"

End Sub

