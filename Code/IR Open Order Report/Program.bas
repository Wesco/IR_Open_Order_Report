Attribute VB_Name = "Program"
Option Explicit
Public Const VersionNumber As String = "1.0.0"

Sub Main()

    'Import IR Open Order Report
    UserImportFile Sheets("IR OOR").Range("A1"), False

    'Import 117 Open Order Report
    Import117

    'Import Master Part List
    ImportMaster
    
    'Import GAPS inventory file
    ImportGaps

    'Import Previous Combined Open Order Report
    'ImportPrevCOOR

End Sub

Sub Clean()
    Dim PrevDispAlert As Boolean
    Dim PrevScrnUpdat As Boolean
    Dim PrevActivWkbk As Workbook
    Dim s As Worksheet

    PrevDispAlert = Application.DisplayAlerts
    PrevScrnUpdat = Application.ScreenUpdating
    Set PrevActivWkbk = ActiveWorkbook

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    ThisWorkbook.Activate

    For Each s In ThisWorkbook.Sheets
        If s.Name <> "Macro" Then
            s.Select
            Cells.Delete
            Range("A1").Select
        End If
    Next

    Sheets("Macro").Select
    Range("C7").Select

    PrevActivWkbk.Activate
    Application.DisplayAlerts = PrevDispAlert
    Application.ScreenUpdating = PrevScrnUpdat
End Sub
