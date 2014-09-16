Attribute VB_Name = "Program"
Option Explicit
Public Const VersionNumber As String = "1.0.2"
Public Const RepositoryName As String = "IR_Open_Order_Report"

Sub Main()
    Application.ScreenUpdating = False

    On Error GoTo Import_Error

    'Import IR Open Order Report
    UserImportFile Sheets("IR OOR").Range("A1"), False

    'Import 117 Open Order Report
    Import117 Crit:=AllOrders, _
              Seq:=ByInsideSalesperson, _
              RepDate:=Now, _
              SeqRng:=One, _
              SeqData:="24", _
              Branch:="3615", _
              Detail:=True, _
              Destination:=Sheets("117 OOR").Range("A1")

    'Import Master Part List
    ImportMaster

    'Import GAPS inventory file
    ImportGaps

    'Import Previous Open Order Report
    ImportPrevOOR

    On Error GoTo 0

    'Move descriptions to the first column and clean them up
    FormatMaster

    'Clean up 117 report and add UID column
    Format117

    'Format IR Open Order Report
    FormatIROOR

    'Create Wesco's Open Order Report
    CreateOOR

    'Format Wesco's Open Order Report
    FormatOOR

    'Export Wesco's Open Order Report to the network
    ExportOOR

    'Remove all data from the macro workbook
    Clean

    Application.ScreenUpdating = True

    'Notify user that the macro finished
    MsgBox "Complete!", vbOKOnly, "Macro"

    Exit Sub

Main_Error:
    If Err.Source = "ImportPrevOOR" And Err.Number = Errors.FILE_NOT_FOUND Then
        MsgBox "The previous OOR could not be found."
        Resume Next
    Else
        MsgBox Prompt:="Error " & Err.Number & " (" & Err.Description & ") occurred in " & Err.Source & ".", _
               Title:="Oops!"
    End If
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
            s.AutoFilterMode = False
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
