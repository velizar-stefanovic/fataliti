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
Worksheets("Ulaz").Activate

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


Private Sub Workbook_BeforeClose(Cancel As Boolean)

    With Application
        .ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
        .CommandBars("Status Bar").Visible = True
        .DisplayFormulaBar = True
        .DisplayScrollBars = True
        .ActiveWindow.DisplayHeadings = True
        .ActiveWindow.DisplayWorkbookTabs = True
   End With

End Sub
Sub Ulaz_u_gl_meni()
    Sheets("MENI").Select
End Sub
Sub Finansijsko_meni()
    Sheets("FINANSIJSKO").Select
End Sub
Sub Materijalno_meni()
    Sheets("MATERIJALNO").Select
End Sub
Sub Osnovna_meni()
    Sheets("OSNOVNA").Select
End Sub
Sub Zarade_meni()
    Sheets("ZARADE").Select
End Sub
Sub Kalkulacije_meni()
    Sheets("KALKULACIJE").Select
End Sub
Sub Izlaz_iz_glavnog_menija()
    Sheets("ULAZ").Select
End Sub
Sub Izlaz_iz_programa()
odgovor = MsgBox("Da li ste sigurni da zelite da izadjete?", vbOKCancel, "SNIMITI")
    If odgovor = 1 Then
        ThisWorkbook.Close True
    Else
        Range("A1").Select
    End If
    
   'odgovor = MsgBox("Da li ste sigurni da zelite da izadjete?", vbOKCancel, "SNIMITI")
'If odgovor = 1 Then
'Application.DisplayAlerts = False
    
    'Application.DisplayAlerts = True
    'Application.Quit
    
'Else
'Range("A1").Select
'End If
    
End Sub

Sub Uputstvod()

Dim myShell As Object
Set myShell = CreateObject("WScript.Shell")
myShell.Run "C:\FATALITI\Uputstvo.doc"

End Sub


