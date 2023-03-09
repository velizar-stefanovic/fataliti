VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Red, NKonto, BrReda, PrRed, PoRed
Private Sub Workbook_Open()


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
Sub Dodavanje_redova(ByVal CONTROL As IRibbonControl)
    m1 = Range("M1") + 1
    novired = Application.InputBox("Koliko novih redova zelite?", "DODAVANJE REDOVA", Type:=1)
    If novired = False Then End
    Range(Cells(m1, 1), Cells(m1 + 10, 100)).Select
    Selection.Cut
    Range(Cells(m1 + novired, 1), Cells(m1 + novired, 1)).Select
    ActiveSheet.Paste
    Range(Cells(m1 - 1, 1), Cells(m1 - 1, 100)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Cells(m1, 1), Cells(m1 + novired - 1, 100)).Select
    ActiveSheet.Paste
    Range("M1") = m1 - 1 + novired
    Application.CutCopyMode = False
    Range("A32").Select
End Sub
Sub Oduzimanje_reda(ByVal CONTROL As IRibbonControl)
    m1 = Range("M1") + 1
    If m1 < 34 Then MsgBox ("Ne mozete oduzimati redove.")
    If m1 < 34 Then End
    novired = Application.InputBox("Koliko redova zelite da izbrisete?", "ODUZIMANJE REDOVA", Type:=1)
    If novired = False Then End
    If novired > m1 - 33 Then MsgBox ("Ne mozete oduzeti toliko redova.")
    If novired > m1 - 33 Then End
    Range(Cells(m1, 1), Cells(m1 + 105, 100)).Select
    Selection.Cut
    Range(Cells(m1 - novired, 1), Cells(m1 - novired, 1)).Select
    ActiveSheet.Paste
    Range("M1") = m1 - novired - 1
    Application.CutCopyMode = False
    Range("A32").Select
End Sub
Sub Akontacija(ByVal CONTROL As IRibbonControl)
    Tekuci = ActiveSheet.Name
    AKONTA = Application.InputBox("UPISITE NAZIV RADNOG LISTA", "TABELA SA AKONTACIJOM", Type:=2)
        If AKONTA = False Then End
        If AKONTA = "" Then End
    On Error Resume Next
    Sheets(AKONTA).Select
    If Err Then
        Err = 0
        MsgBox ("NE POSTOJI RADNI LIST  " & AKONTA)
        End
    End If
    On Error GoTo 0
    m1 = Range("M1")
    Range(Cells(32, 63), Cells(m1, 63)).Select
    Selection.Copy
    Sheets(Tekuci).Select
    Range("BL32").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
    Sheets(AKONTA).Select
    Range("DD8:DD56").Select
    Selection.Copy
    Sheets(Tekuci).Select
    Range("DE8").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
End Sub
Sub Listici(ByVal CONTROL As IRibbonControl)
    PrRed = Application.InputBox("UPISITE POCETNI REDNI BROJ", "STAMPANJE LISTICA ZARADA", Type:=1)
        If PrRed = False Then End
    PoRed = Application.InputBox("UPISITE KRAJNJI REDNI BROJ", "STAMPANJE LISTICA ZARADA", Type:=1)
        If PoRed = False Then End
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    ActiveSheet.PageSetup.PrintArea = "$DJ$1:$DQS$61"
    With ActiveSheet.PageSetup
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    For i = PrRed + 31 To PoRed + 31 Step 2
        Range(Cells(i, 1), Cells(i + 1, 70)).Select
        Application.CutCopyMode = False
        Selection.Copy
        Cells(20, 1).Select
        Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
        ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
    Next i
    
End Sub
Sub Otvaranje_zar()
    Sheets("KARTON").Select
    NKonto = Range("D2")
    Sheets("PLATE").Select
    Range("D2").Select
    Red = Selection.Row
    KONTO = Range("D2")
2   If KONTO = "" Then
        Smestiz
        GoTo 1
    End If
    If NKonto = KONTO Then
        MsgBox ("Konto  " & NKonto & "  vec postoji")
        End
    End If
    If NKonto > KONTO Then
        Red = Red + 60
        KONTO = Cells(Red, 4)
        GoTo 2
    Else
        Range(Cells(Red - 1, 1), Cells(Red + 58, 7)).Select
        Application.CutCopyMode = False
        Selection.EntireRow.Insert
        Smestiz
    End If
        
1   End Sub
Sub Smestiz()
    Adresa = "=PLATE!$A$" & Red - 1 & ":$G$" & Red + 55
    Sheets("KARTON").Select
    Range("A1:G56").Select
    Selection.Copy
    Sheets("PLATE").Select
    Cells(Red - 1, 1).Select
    ActiveSheet.Paste
    Names.Add "_KZ" & NKonto, Adresa
End Sub
Sub Preuzima(ByVal CONTROL As IRibbonControl)
'
'   Program za otvaranje kartona iz tabele zarada
'
    TABELA = Application.InputBox("UPISITE NAZIV RADNOG LISTA", "TABELA ZA AZURIRANJE", Type:=2)
        If TABELA = False Then End
        If TABELA = "" Then End
    On Error Resume Next
    Sheets(TABELA).Select
    If Err Then
        Err = 0
        MsgBox ("NE POSTOJI RADNI LIST  " & TABELA)
        End
    End If
    On Error GoTo 0
    PrRed = Application.InputBox("REDNI BROJ PRVOG ZAPISA", "ODREDJIVANJE OPSEGA", Type:=1)
    If PrRed = False Then End
    PoRed = Application.InputBox("REDNI BROJ POSLEDNJEG ZAPISA", "ODREDJIVANJE OPSEGA", Type:=1)
    If PoRed = False Then End
    If PoRed < PrRed Then End
    For BrReda = PrRed + 31 To PoRed + 31
        Sheets("KARTON").Select
        'Range("D2") = Worksheets(TABELA).Cells(BrReda, 3)
        'Range("D3") = Worksheets(TABELA).Cells(BrReda, 6)
        'Range("D5") = Worksheets(TABELA).Cells(BrReda, 2)
        'Range("F2") = "'" & Worksheets(TABELA).Cells(BrReda, 5)
        'Range("F3") = Worksheets(TABELA).Cells(BrReda, 4)
        'Range("F2") = Worksheets(TABELA).Cells(BrReda, 4)
        'Range("F3") = Worksheets(TABELA).Cells(BrReda, 9)
        
        
        Range("D2") = Worksheets(TABELA).Cells(BrReda, 3)
        Range("D3") = Worksheets(TABELA).Cells(BrReda, 6)
        Range("D5") = Worksheets(TABELA).Cells(BrReda, 2)
        Range("F2") = Worksheets(TABELA).Cells(BrReda, 5)
        Range("F3") = Worksheets(TABELA).Cells(BrReda, 4)
        
        Otvaranje_zar
    Next BrReda
    
End Sub
Sub Puniz(ByVal CONTROL As IRibbonControl)
    TABELA = Application.InputBox("UPISITE NAZIV RADNOG LISTA", "TABELA SA PODACIMA", Type:=2)
        If TABELA = False Then End
        If TABELA = "" Then End
    On Error Resume Next
    Sheets(TABELA).Select
    If Err Then
        Err = 0
        MsgBox ("NE POSTOJI RADNI LIST  " & TABELA)
        End
    End If
    On Error GoTo 0
    PrRed = Application.InputBox("REDNI BROJ PRVOG ZAPISA", "ODREDJIVANJE OPSEGA", Type:=1)
    If PrRed = False Then End
    PoRed = Application.InputBox("REDNI BROJ POSLEDNJEG ZAPISA", "ODREDJIVANJE OPSEGA", Type:=1)
    If PoRed = False Then End
    If PoRed < PrRed Then End
    Proveraz
    For BrReda = PrRed + 31 To PoRed + 31
        KONTO = "_KZ" & Worksheets(TABELA).Cells(BrReda, 3)
        Sheets("PLATE").Select
        Range(KONTO).Cells(1, 1).Select
        Red = Selection.Row
        Kolona = 7 + Range(KONTO).Cells(4, 4)
        Range(Cells(Red, Kolona), Cells(Red + 48, Kolona)).Select
        Application.CutCopyMode = False
        Selection.Cut
        Range(Cells(Red, Kolona + 1), Cells(Red + 48, Kolona + 1)).Select
        ActiveSheet.Paste
        Sheets(TABELA).Select
        Range(Cells(BrReda, 1), Cells(BrReda, 65)).Select
        Application.CutCopyMode = False
        Selection.Copy
        Cells(20, 1).Select
        Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
        Range("GA6:GA46").Select
        Selection.Copy
        Sheets("PLATE").Select
        Cells(Red + 5, Kolona - 1).Select
        Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
        Selection.PasteSpecial Paste:=xlFormats, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
        Cells(Red + 4, Kolona - 1) = Worksheets(TABELA).Cells(27, 7)
        Range(KONTO).Cells(4, 4) = Kolona - 6
        
    Next BrReda
    Application.StatusBar = False
    Application.DisplayStatusBar = oldStatusBar
    
End Sub
Sub Proveraz()
    oldStatusBar = Application.DisplayStatusBar
    Application.DisplayStatusBar = True
    For BrReda = PrRed + 31 To PoRed + 31
        BC20 = Cells(BrReda, 3)
        KONTO = "_KZ" & BC20
        Application.StatusBar = "Provera kartice   " & KONTO
        On Error Resume Next
        Proba = Range(KONTO).Cells(32, 4)
        If Err Then
            Err = 0
            MsgBox (" NE POSTOJI KARTICA  " & BC20)
            Application.StatusBar = False
            Application.DisplayStatusBar = oldStatusBar
            End
        End If
        On Error GoTo 0
    Next BrReda
    
End Sub
Sub Otvaranje(ByVal CONTROL As IRibbonControl)
    Sifra = Application.InputBox("Upisite naziv novog radnog lista:", "OTVARANJE NOVE TABELE", Type:=2)
    If Sifra = False Then End
    If Sifra = "" Then End
    Sheets("MATRICA").Copy Before:=Sheets("KARTON")
    Sheets("MATRICA (2)").Select
    Sheets("MATRICA (2)").Name = Sifra
End Sub

Sub uputstvo(ByVal CONTROL As IRibbonControl)
Dim myShell As Object
Set myShell = CreateObject("WScript.Shell")
myShell.Run "C:\FATALITI\Uputstvo.doc"

End Sub



