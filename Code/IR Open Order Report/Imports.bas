Attribute VB_Name = "Imports"
Option Explicit

Sub ImportPrevOOR()
    Dim PrevDispAlert As Boolean
    Dim FileName As String
    Dim FilePath As String
    Dim dt As Date
    Dim i As Long
    

    'Look back up to 30 days for the combined open order report
    For i = 1 To 30
        dt = Date - i
        FileName = "OOR " & Format(dt, "yyyy-mm-dd") & ".xlsx"
        FilePath = "\\7938-HP02\Shared\IR-Davidson-Mox\Open Order Report\" & Format(dt, "yyyy") & "\" & Format(dt, "mmm") & "\"
        
        If FileExists(FilePath & FileName) Then
            Exit For
        End If
    Next

    'If the 117 open order report was found, import it
    If FileExists(FilePath & FileName) Then
        PrevDispAlert = Application.DisplayAlerts
        Application.DisplayAlerts = False

        Workbooks.Open FilePath & FileName
        ActiveSheet.AutoFilterMode = False
        ActiveSheet.Columns.Hidden = False
        ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Sheets("Prev OOR").Range("A1")
        ActiveWorkbook.Close

        Application.DisplayAlerts = PrevDispAlert
    Else
        Err.Raise Errors.FILE_NOT_FOUND, "ImportPrevOOR", "Previous OOR Report not found."
    End If
End Sub

Sub ImportMaster()
    Dim PrevDispAlert As Boolean
    Dim FileName As String
    Dim FilePath As String
    
    FileName = "IR Master " & Format(Date, "yyyy") & ".xlsx"
    FilePath = "\\br3615gaps\gaps\IR\Master\"

    If FileExists(FilePath & FileName) Then
        PrevDispAlert = Application.DisplayAlerts
        Application.DisplayAlerts = False
        
        Workbooks.Open FilePath & FileName
        ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Sheets("Master").Range("A1")
        ActiveWorkbook.Close
        
        Application.DisplayAlerts = PrevDispAlert
    Else
        Err.Raise Errors.FILE_NOT_FOUND, "ImportMaster", "IR Master not found."
    End If
End Sub
