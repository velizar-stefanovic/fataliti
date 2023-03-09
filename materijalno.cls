VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Red, NKonto, BA21
Private Sub Workbook_Open()
'Worksheets("Ulaz").Activate
    With Application
        .ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
        .CommandBars("Status Bar").Visible = True
        .DisplayFormulaBar = True
        .DisplayScrollBars = True
        .Caption = "KNJIGOVODSTVENI PROGRAM FATALITI"
        .ActiveWindow.DisplayHeadings = True
        .ActiveWindow.DisplayWorkbookTabs = True
    End With

End Sub
Sub Otvaranjem(ByVal control As IRibbonControl)
Attribute Otvaranjem.VB_ProcData.VB_Invoke_Func = "o\n14"
    Sheets("MATRICA").Select
    NKonto = Range("C1")
    Sheets("POKNJIGA").Select
    Range("C1").Select
    Red = Selection.Row
    Konto = Range("C1")
2   If Konto = "" Then
1       If Red = 1 Then
            Smestim
        Else
            Selection.End(xlUp).Select
            Red = Selection.Row + 5
            Smestim
        End If
        End
    End If
    If NKonto = Konto Then
        MsgBox ("Sifra  " & NKonto & "  vec postoji")
        End
    End If
    If NKonto > Konto Then
        Red = Red + 5 + Cells(Red + 1, 11)
        Cells(Red, 3).Select
        Selection.End(xlDown).Select
        Red = Selection.Row
        Konto = Selection
        GoTo 2
    Else
        Range(Cells(Red, 1), Cells(Red + 8, 11)).Select
        Application.CutCopyMode = False
        Selection.EntireRow.Insert
        Smestim
    End If
        'Application.CutCopyMode = False
End Sub
Sub Smestim()
    Kolona = Red + 6 + Cells(Red + 1, 11)
    Adresa = "=POKNJIGA!$A$" & Red & ":$K$" & Kolona
    Sheets("MATRICA").Select
    Range("A1:K5").Select
    Selection.Copy
    Sheets("POKNJIGA").Select
    Cells(Red, 1).Select
    ActiveSheet.Paste
    Names.Add "_MK" & NKonto, Adresa
End Sub
Sub Dodavanjem(ByVal control As IRibbonControl)
    Sheets("NALOG").Select
    m1 = Range("M1") + 1
    'If M1 < 13 Then End
    novired = Application.InputBox("Koliko novih redova zelite?", "DODAVANJE REDOVA", Type:=1)
    If novired = False Then End
    Range(Cells(m1, 1), Cells(m1 + 5, 25)).Select
    Selection.Cut
    Range(Cells(m1 + novired, 1), Cells(m1 + novired, 1)).Select
    ActiveSheet.Paste
    Range(Cells(m1 - 1, 1), Cells(m1 - 1, 25)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Cells(m1, 1), Cells(m1 + novired - 1, 25)).Select
    ActiveSheet.Paste
    Range("M1") = m1 - 1 + novired
    Application.CutCopyMode = False
    Range("A1").Select
End Sub
Sub Oduzimanjem(ByVal control As IRibbonControl)
    Sheets("NALOG").Select
    m1 = Range("M1") + 1
    If m1 < 5 Then MsgBox ("Ne mozete oduzimati redove.")
    If m1 < 5 Then End
    novired = Application.InputBox("Koliko redova zelite da izbrisete?", "ODUZIMANJE REDOVA", Type:=1)
    If novired = False Then End
    If novired > m1 - 4 Then MsgBox ("Ne mozete oduzeti toliko redova.")
    If novired > m1 - 4 Then End
    Range(Cells(m1, 1), Cells(m1 + 105, 25)).Select
    Selection.Cut
    Range(Cells(m1 - novired, 1), Cells(m1 - novired, 1)).Select
    ActiveSheet.Paste
    Range("M1") = m1 - novired - 1
    Application.CutCopyMode = False
    Range("A1").Select
End Sub
Sub Knjizenjem(ByVal control As IRibbonControl)
    With Application
        .Calculation = xlAutomatic
    End With
    Provera_ma
    Range("G1") = "KNJIZEN"
    oldStatusBar = Application.DisplayStatusBar
    Application.DisplayStatusBar = True
    For BA20 = 3 To BA21
        Sheets("NALOG").Select
        Konto = "_MK" & Cells(BA20, 2)
        Application.StatusBar = "Knjizenje na kartici   " & Konto
        Range(Cells(BA20, 15), Cells(BA20, 25)).Select
        Selection.Copy
        Sheets("POKNJIGA").Select
        Red = 6 + Range(Konto).Cells(2, 11)
        Range(Konto).Cells(Red, 1).Select
        Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
        Selection.PasteSpecial Paste:=xlFormats, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
        Range(Konto).Cells(Red + 1, 1).Select
        Application.CutCopyMode = False
        LRed = Selection.Row - 1
        Application.CutCopyMode = False
        Selection.EntireRow.Insert
        If Red = 6 Then
            Range(Konto).Cells(Red, 7).Formula = "=E" & LRed & "-F" & LRed
            Range(Konto).Cells(Red, 11).Formula = "=I" & LRed & "-J" & LRed
        Else
            Range(Konto).Cells(Red, 7).Formula = "=G" & LRed - 1 & "+E" & LRed & "-F" & LRed
            Range(Konto).Cells(Red, 11).Formula = "=K" & LRed - 1 & "+I" & LRed & "-J" & LRed
        End If
        Range(Konto).Cells(Red, 9).Formula = "=E" & LRed & "*H" & LRed
        Range(Konto).Cells(Red, 10).Formula = "=F" & LRed & "*H" & LRed
        Range(Konto).Cells(2, 11) = Range(Konto).Cells(2, 11) + 1
        Range(Konto).Select
        Adresa = "=POKNJIGA!" & Selection.Address
        Names.Add Konto, Adresa

    Next BA20
    Application.StatusBar = False
    Application.DisplayStatusBar = oldStatusBar
    With Application
        .Calculation = xlAutomatic
    End With
    
    Sheets("NALOG").Select
    Range("A1").Select
End Sub
Sub Provera_ma()
    Sheets("NALOG").Select
    On Error Resume Next
    If Range("E1") < 30000 Then
        Err = 0
        MsgBox ("DATUM NIJE DOBRO UPISAN")
        End
    End If
    On Error GoTo 0
    If Range("G1") = "KNJIZEN" Then
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
    If Cells(BA20, 2) = "" Then GoTo 2
    BC20 = Cells(BA20, 2)
    Konto = "_MK" & BC20
    Application.StatusBar = "Provera kartice   " & Konto
    On Error Resume Next
    proba = Range(Konto).Cells(3, 7)
    If Err Then
        Err = 0
        MsgBox (" NE POSTOJI KARTICA  " & BC20)
        Application.StatusBar = False
        Application.DisplayStatusBar = oldStatusBar
        End
    End If
    On Error GoTo 0
    If proba - Cells(BA20, 4) < 0 Then
        MsgBox ("KOLICINA CE OTICI U MINUS NA KARTICI  " & BC20)
        Application.StatusBar = False
        Application.DisplayStatusBar = oldStatusBar
        End
    End If
    BA21 = BA20
    GoTo 1
2   BC20 = ""
   
    Application.StatusBar = False
    Application.DisplayStatusBar = oldStatusBar
  
End Sub
Sub NoviNalogm(ByVal control As IRibbonControl)
Attribute NoviNalogm.VB_ProcData.VB_Invoke_Func = "r\n14"
    Sheets("NALOG").Select
    Cells.Select
    Selection.Clear
    Sheets("MATRICA").Select
    Range("AA1:AY20").Select
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
Sub Kontrola_ma()
    m1 = Range("M1")
    Range(Cells(1, 1), Cells(m1 + 3, 6)).Select
    Application.DataEntryMode = xlOn
    
End Sub
Sub listam(ByVal control As IRibbonControl)
    Sheets("LISTA").Select
    Kljuc = Application.InputBox("UPISITE KLJUC ZA IZBOR KARTICA", "KLJUC ZA LISTANJE", Type:=2)
    If Kljuc = False Then End
    If Kljuc = "" Then End
    Range("Z4") = "Sifra"
    Range("Z9") = "Sifra"
    Range("Y9") = "Sifra"
    Range("Z5") = "_MK" & Kljuc
    Range("Z10").Select
    Selection.ListNames
    Selection.End(xlDown).Select
    Red = Selection.Row
    Range(Cells(9, 26), Cells(Red, 26)).AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range _
        ("Z4:Z5"), CopyToRange:=Range("Y9"), Unique:=False
    Range("Y9").Select
    Selection.End(xlDown).Select
    Red = Selection.Row
    'If Red > 50000 Then End
    If Red > 50000 Then Range("A1").Select
    If Red > 50000 Then MsgBox ("Greska.")
    If Red > 50000 Then End
    For i = 10 To Red
        Cells(i, 1).Select
        ActiveCell.FormulaR1C1 = "=INDEX(" & Cells(i, 25) & ",3,8)"
        Selection.NumberFormat = "#,##0.00_);(#,##0.00)"
        Cells(i, 2).Select
        ActiveCell.FormulaR1C1 = "=MID(RC[23],4,LEN(RC[23]))"
        Cells(i, 3).Select
        ActiveCell.FormulaR1C1 = "=INDEX(" & Cells(i, 25) & ",2,3)"
        Cells(i, 4).Select
        ActiveCell.FormulaR1C1 = "=INDEX(" & Cells(i, 25) & ",3,3)"
        Cells(i, 5).Select
        ActiveCell.FormulaR1C1 = "=INDEX(" & Cells(i, 25) & ",3,7)"
        Cells(i, 6).Select
        ActiveCell.FormulaR1C1 = "=INDEX(" & Cells(i, 25) & ",3,11)"
        Selection.NumberFormat = "#,##0.00_);(#,##0.00)"
    Next i
    Range("F7").Formula = "=SUM(F10:F" & Red & ")"
         
End Sub
Sub Izdvajanjem(ByVal control As IRibbonControl)
    Sheets("LISTA").Select
    Range("F9").Select
    Selection.End(xlDown).Select
    Red = Selection.Row
    'If Red > 50000 Then End
    If Red > 50000 Then Range("A1").Select
    If Red > 50000 Then MsgBox ("Greska.")
    If Red > 50000 Then End
    Range(Cells(9, 1), Cells(Red, 6)).AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range( _
        "A3:E4"), CopyToRange:=Range("K10:P10"), Unique:=False
    Range("K8:P8").Select
End Sub
Sub Popism(ByVal control As IRibbonControl)
    Sheets("LISTA").Select
    Range("L10").Select
    Selection.End(xlDown).Select
    Red = Selection.Row
    'If Red > 50000 Then End
    If Red > 50000 Then Range("A1").Select
    If Red > 50000 Then MsgBox ("Greska.")
    If Red > 50000 Then End
    Range("Q2:V5").Select
    Selection.Copy
    Range("Q8").Select
    ActiveSheet.Paste
    Range("Q11:V11").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Cells(12, 17), Cells(Red, 17)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("Q11").Select
End Sub
Sub Preuzmim(ByVal control As IRibbonControl)
    Sheets("NIVEL").Select
    Range("B10").Select
    Selection.End(xlDown).Select
    Red = Selection.Row
    For i = 11 To Red
        Konto = "_MK" & Cells(i, 2)
        On Error Resume Next
        Cells(i, 3) = Range(Konto).Cells(2, 3)
        If Err Then
            Err = 0
            MsgBox (" NE POSTOJI KARTICA  " & Konto)
            End
        End If
        On Error GoTo 0
        Cells(i, 4) = Range(Konto).Cells(3, 3)
        Cells(i, 5) = Range(Konto).Cells(3, 7)
        Cells(i, 6) = Range(Konto).Cells(3, 8)
    Next i
                
End Sub
Sub Nivelacijam(ByVal control As IRibbonControl)
    With Application
        .Calculation = xlAutomatic
    End With
    Provera_niv
    Range("G1") = "KNJIZEN"
    oldStatusBar = Application.DisplayStatusBar
    Application.DisplayStatusBar = True
    For BA20 = 11 To BA21
        Sheets("NIVEL").Select
        Konto = "_MK" & Cells(BA20, 2)
        Application.StatusBar = "Knjizenje na kartici   " & Konto
        Range(Cells(BA20, 15), Cells(BA20, 25)).Select
        Selection.Copy
        Sheets("POKNJIGA").Select
        Red = 6 + Range(Konto).Cells(2, 11)
        Range(Konto).Cells(Red, 1).Select
        Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
        Selection.PasteSpecial Paste:=xlFormats, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
        Range(Konto).Cells(Red + 1, 1).Select
        LRed = Selection.Row - 1
        Application.CutCopyMode = False
        Selection.EntireRow.Insert
        Range(Konto).Cells(Red, 7).Formula = "=G" & LRed - 1 & "+E" & LRed & "-F" & LRed
        Range(Konto).Cells(Red, 9).Formula = "=E" & LRed & "*H" & LRed
        Range(Konto).Cells(Red, 10).Formula = "=F" & LRed & "*H" & LRed
        Range(Konto).Cells(Red, 11).Formula = "=K" & LRed - 1 & "+I" & LRed & "-J" & LRed
        Range(Konto).Cells(2, 11) = Range(Konto).Cells(2, 11) + 1
        
        Sheets("NIVEL").Select
        Range(Cells(10, 15), Cells(10, 25)).Select
        Selection.Copy
        Sheets("POKNJIGA").Select
        Red = 6 + Range(Konto).Cells(2, 11)
        Range(Konto).Cells(Red, 1).Select
        Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
        Selection.PasteSpecial Paste:=xlFormats, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
        Range(Konto).Cells(Red + 1, 1).Select
        Application.CutCopyMode = False
        Selection.EntireRow.Insert
        Range(Konto).Cells(Red, 7).Formula = "=G" & LRed & "+E" & LRed + 1 & "-F" & LRed + 1
        Range(Konto).Cells(Red, 9).Formula = "=E" & LRed + 1 & "*H" & LRed + 1
        Range(Konto).Cells(Red, 10).Formula = "=F" & LRed + 1 & "*H" & LRed + 1
        Range(Konto).Cells(Red, 11).Formula = "=K" & LRed & "+I" & LRed + 1 & "-J" & LRed + 1
        Range(Konto).Cells(2, 11) = Range(Konto).Cells(2, 11) + 1
        
        Range(Konto).Select
        Adresa = "=POKNJIGA!" & Selection.Address
        Names.Add Konto, Adresa

    Next BA20
    Application.StatusBar = False
    Application.DisplayStatusBar = oldStatusBar
    With Application
        .Calculation = xlAutomatic
    End With
    
    Sheets("NIVEL").Select
    Range("A1").Select
End Sub
Sub Provera_niv()
    Sheets("NIVEL").Select
    On Error Resume Next
    If Range("B5") < 30000 Then
        Err = 0
        MsgBox ("DATUM NIJE DOBRO UPISAN")
        End
    End If
    On Error GoTo 0
    If Range("G1") = "KNJIZEN" Then
        MsgBox ("NALOG JE VEC KNJIZEN")
        End
    End If
    With Application
        .Calculation = xlManual
    End With
    
    BA20 = 10
    oldStatusBar = Application.DisplayStatusBar
    Application.DisplayStatusBar = True
1   BA20 = BA20 + 1
    If BA20 > Range("M1") Then GoTo 2
    If Cells(BA20, 2) = "" Then GoTo 2
    BC20 = Cells(BA20, 2)
    Konto = "_MK" & BC20
    Application.StatusBar = "Provera kartice   " & Konto
    On Error Resume Next
    proba = Range(Konto).Cells(3, 7)
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
Sub NovaNivel(ByVal control As IRibbonControl)
    Sheets("NIVEL").Select
    Cells.Select
    Selection.Clear
    Sheets("MATRICA").Select
    Range("BA1:BY50").Select
    Selection.Copy
    Sheets("NIVEL").Select
    Range("A1").Select
    ActiveSheet.Paste
    Sheets("MATRICA").Select
    Range("A1").Select
    Sheets("NIVEL").Select
    Range("A11").Select
    
End Sub
Sub Poknjiga(ByVal control As IRibbonControl)
Sheets("POKNJIGA").Select
Range("C1").Select
End Sub
Sub Listnivelacija(ByVal control As IRibbonControl)
Sheets("NIVEL").Select
Range("C1").Select
End Sub
Sub Matrica(ByVal control As IRibbonControl)
Sheets("MATRICA").Select
Range("C1").Select
End Sub
Sub Uputstvo(ByVal control As IRibbonControl)
Dim myShell As Object
Set myShell = CreateObject("WScript.Shell")
myShell.Run "C:\FATALITI\Uputstvo.doc"

End Sub

