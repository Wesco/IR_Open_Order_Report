Attribute VB_Name = "Imports"
Option Explicit

Sub Import_IR_OOR()

    On Error GoTo IMPORT_ERR
    UserImportFile DestRange:=ThisWorkbook.Sheets("IR DLC").Range("A1"), _
                   DelFile:=False, _
                   ShowAllData:=True
    On Error GoTo 0

    Sheets("IR DLC").Select
    
    On Error GoTo COL_NOT_FOUND
    FindColumn ("PO Rel #")
    On Error GoTo 0

    If Range("A1").Value = "" Then
        UserImportFile DestRange:=ThisWorkbook.Sheets("IR DLC").Range("A1"), _
                       DelFile:=False, _
                       ShowAllData:=True
    Else
        UserImportFile DestRange:=ThisWorkbook.Sheets("IR Mox").Range("A1"), _
                       DelFile:=False, _
                       ShowAllData:=True
    End If

    Exit Sub

COL_NOT_FOUND:
    ActiveSheet.UsedRange.Cut Destination:=Sheets("IR Mox").Range("A1")
    Resume Next

IMPORT_ERR:
    Debug.Print Err.Number
    Debug.Print Err.Description
    Exit Sub
End Sub
