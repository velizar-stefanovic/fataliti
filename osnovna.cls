VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
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

Sub Otvaranje(ByVal CONTROL As IRibbonControl)
    Sifra = Application.InputBox("UPISITE SIFRU GRUPE SREDSTAVA", "OTVARANJE NOVE TABELE", Type:=2)
    If Sifra = False Then End
    If Sifra = "" Then End
    Sheets("MATRICA").Copy Before:=Sheets("PRAZNA")
    Sheets("MATRICA (2)").Select
    Sheets("MATRICA (2)").Name = Sifra
End Sub
Sub Dodavanje_redova_oss(ByVal CONTROL As IRibbonControl)
    m1 = Range("M1") + 1
    novired = Application.InputBox("KOLIKO NOVIH REDOVA", "DODAVANJE REDOVA", Type:=1)
    If novired = False Then End
    Range(Cells(m1, 1), Cells(m1 + 5, 20)).Select
    Selection.Cut
    Range(Cells(m1 + novired, 1), Cells(m1 + novired, 1)).Select
    ActiveSheet.Paste
    Range(Cells(m1 - 1, 1), Cells(m1 - 1, 20)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Cells(m1, 1), Cells(m1 + novired - 1, 20)).Select
    ActiveSheet.Paste
    Range("M1") = m1 - 1 + novired
    Application.CutCopyMode = False
    Range("A11").Select
    'Cells(11 + novired, 1).Select
    
End Sub
Sub Oduzimanje_reda_oss(ByVal CONTROL As IRibbonControl)
    m1 = Range("M1") + 1
    If m1 < 13 Then MsgBox ("Ne mozete oduzimati redove.")
    If m1 < 13 Then End
    novired = Application.InputBox("Koliko redova zelite da izbrisete?", "ODUZIMANJE REDOVA", Type:=1)
    If novired = False Then End
    If novired > m1 - 12 Then MsgBox ("Ne mozete oduzeti toliko redova.")
    If novired > m1 - 12 Then End
    'If M1 < 13 Then End
    Range(Cells(m1, 1), Cells(m1 + 2000, 20)).Select
    Selection.Cut
    Range(Cells(m1 - novired, 1), Cells(m1 - novired, 1)).Select
    ActiveSheet.Paste
    Range("M1") = m1 - novired - 1
    Application.CutCopyMode = False
    Range("A11").Select
End Sub
Sub Azuriranje(ByVal CONTROL As IRibbonControl)
    m1 = Range("M1")
    Range(Cells(11, 17), Cells(m1, 17)).Select
    Selection.Copy
    Range("X11").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
    Range(Cells(11, 19), Cells(m1, 19)).Select
    Selection.Copy
    Range("J11").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
    Range(Cells(11, 24), Cells(m1, 24)).Select
    Selection.Copy
    Range("I11").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
    Range(Cells(11, 17), Cells(m1, 17)).Select
    Selection.ClearContents
    Range(Cells(11, 19), Cells(m1, 19)).Select
    Selection.ClearContents
End Sub
Sub Nova_godina(ByVal CONTROL As IRibbonControl)
    m1 = Range("M1")
    Range(Cells(11, 13), Cells(m1, 13)).Select
    Selection.Copy
    Range("J11").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
    Application.CutCopyMode = False
    Range("J11").Select
    
End Sub
Sub uputstvo(ByVal CONTROL As IRibbonControl)
Dim myShell As Object
Set myShell = CreateObject("WScript.Shell")
myShell.Run "C:\FATALITI\Uputstvo.doc"

End Sub


