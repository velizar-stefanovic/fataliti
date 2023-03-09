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
Sub Otvaranje_matrica(ByVal CONTROL As IRibbonControl)
    Sifra = Application.InputBox("Unesite naziv novog radnog lista:", "OTVARANJE NOVOG RADNOG LISTA", Type:=2)
    If Sifra = False Then End
    If Sifra = "" Then End
    Sheets("MATRICA").Copy Before:=Sheets("KALNAB")
    Sheets("MATRICA (2)").Select
    Sheets("MATRICA (2)").Name = Sifra
End Sub
Sub Otvaranje_nabavka(ByVal CONTROL As IRibbonControl)
    Sifra = Application.InputBox("Unesite naziv novog radnog lista:", "OTVARANJE NOVOG RADNOG LISTA", Type:=2)
    If Sifra = False Then End
    If Sifra = "" Then End
    Sheets("KALNAB").Copy Before:=Sheets("KALPROD")
    Sheets("KALNAB (2)").Select
    Sheets("KALNAB (2)").Name = Sifra
End Sub
Sub Otvaranje_prodaja(ByVal CONTROL As IRibbonControl)
    Sifra = Application.InputBox("Unesite naziv novog radnog lista:", "OTVARANJE NOVOG RADNOG LISTA", Type:=2)
    If Sifra = False Then End
    If Sifra = "" Then End
    Sheets("KALPROD").Copy After:=Sheets(Sheets.Count)
    Sheets("KALPROD (2)").Select
    Sheets("KALPROD (2)").Name = Sifra
End Sub
Sub Dodavanje_redova_nab(ByVal CONTROL As IRibbonControl)
Attribute Dodavanje_redova_nab.VB_ProcData.VB_Invoke_Func = "d\n14"
    m1 = Range("M1") + 1
    novired = Application.InputBox("Koliko novih redova zelite?", "DODAVANJE REDOVA", Type:=1)
    If novired = False Then End
    Range(Cells(m1, 1), Cells(m1 + 5, 11)).Select
    Selection.Cut
    Range(Cells(m1 + novired, 1), Cells(m1 + novired, 1)).Select
    ActiveSheet.Paste
    Range(Cells(m1 - 1, 1), Cells(m1 - 1, 11)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Cells(m1, 1), Cells(m1 + novired - 1, 11)).Select
    ActiveSheet.Paste
    Range("M1") = m1 - 1 + novired
    Application.CutCopyMode = False
    Range("A12").Select
End Sub
Sub Oduzimanje_redova_nab(ByVal CONTROL As IRibbonControl)
Attribute Oduzimanje_redova_nab.VB_ProcData.VB_Invoke_Func = "g\n14"
    m1 = Range("M1") + 1
    If m1 < 14 Then MsgBox ("Ne mozete oduzimati redove.")
    If m1 < 14 Then End
    novired = Application.InputBox("Koliko redova zelite da izbrisete?", "ODUZIMANJE REDOVA", Type:=1)
    If novired = False Then End
    If novired > m1 - 13 Then MsgBox ("Ne mozete oduzeti toliko redova.")
    If novired > m1 - 13 Then End
    Range(Cells(m1, 1), Cells(m1 + 105, 26)).Select
    Selection.Cut
    Range(Cells(m1 - novired, 1), Cells(m1 - novired, 1)).Select
    ActiveSheet.Paste
    Range("M1") = m1 - novired - 1
    Application.CutCopyMode = False
    Range("A12").Select
End Sub
Sub Dodavanje_redova_prod(ByVal CONTROL As IRibbonControl)
    m1 = Range("M1") + 1
    novired = Application.InputBox("Koliko novih redova zelite?", "DODAVANJE REDOVA", Type:=1)
    If novired = False Then End
    Range(Cells(m1, 1), Cells(m1 + 5, 26)).Select
    Selection.Cut
    Range(Cells(m1 + novired, 1), Cells(m1 + novired, 1)).Select
    ActiveSheet.Paste
    Range(Cells(m1 - 1, 1), Cells(m1 - 1, 26)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Cells(m1, 1), Cells(m1 + novired - 1, 26)).Select
    ActiveSheet.Paste
    Range("M1") = m1 - 1 + novired
    Application.CutCopyMode = False
    Range("E12").Select
End Sub
Sub Oduzimanje_reda_prod(ByVal CONTROL As IRibbonControl)
    m1 = Range("M1") + 1
    If m1 < 14 Then MsgBox ("Ne mozete oduzimati redove.")
    If m1 < 14 Then End
    novired = Application.InputBox("Koliko redova zelite da izbrisete?", "ODUZIMANJE REDOVA", Type:=1)
    If novired = False Then End
    If novired > m1 - 13 Then MsgBox ("Ne mozete oduzeti toliko redova.")
    If novired > m1 - 13 Then End
    Range(Cells(m1, 1), Cells(m1 + 105, 26)).Select
    Selection.Cut
    Range(Cells(m1 - novired, 1), Cells(m1 - novired, 1)).Select
    ActiveSheet.Paste
    Range("M1") = m1 - novired - 1
    Application.CutCopyMode = False
    Range("E12").Select
End Sub
Sub Dodavanje_redova_matrica(ByVal CONTROL As IRibbonControl)
    m1 = Range("M1") + 1
    novired = Application.InputBox("Koliko novih redova zelite?", "DODAVANJE REDOVA", Type:=1)
    If novired = False Then End
    Range(Cells(m1, 1), Cells(m1 + 10, 9)).Select
    Selection.Cut
    Range(Cells(m1 + novired, 1), Cells(m1 + novired, 1)).Select
    ActiveSheet.Paste
    Range(Cells(m1 - 1, 1), Cells(m1 - 1, 9)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Cells(m1, 1), Cells(m1 + novired - 1, 9)).Select
    ActiveSheet.Paste
    Range("M1") = m1 - 1 + novired
    Application.CutCopyMode = False
    Range("A11").Select
End Sub
Sub Oduzimanje_reda_matrica(ByVal CONTROL As IRibbonControl)
    m1 = Range("M1") + 1
    If m1 < 13 Then MsgBox ("Ne mozete oduzimati redove.")
    If m1 < 13 Then End
    novired = Application.InputBox("Koliko redova zelite da izbrisete?", "ODUZIMANJE REDOVA", Type:=1)
    If novired = False Then End
    If novired > m1 - 12 Then MsgBox ("Ne mozete oduzeti toliko redova.")
    If novired > m1 - 12 Then End
    Range(Cells(m1, 1), Cells(m1 + 105, 9)).Select
    Selection.Cut
    Range(Cells(m1 - novired, 1), Cells(m1 - novired, 1)).Select
    ActiveSheet.Paste
    Range("M1") = m1 - novired - 1
    Application.CutCopyMode = False
    Range("A11").Select
End Sub
Sub Kalck(ByVal CONTROL As IRibbonControl)
Sheets("MATRICA").Select
Range("A1").Select
End Sub
Sub Kalnc(ByVal CONTROL As IRibbonControl)
Sheets("KALNAB").Select
Range("A1").Select
End Sub
Sub Kalpc(ByVal CONTROL As IRibbonControl)
Sheets("KALPROD").Select
Range("A1").Select
End Sub
Sub uputstvo(ByVal CONTROL As IRibbonControl)
Dim myShell As Object
Set myShell = CreateObject("WScript.Shell")
myShell.Run "C:\FATALITI\Uputstvo.doc"

End Sub

