Attribute VB_Name = "Imports"
Option Explicit

Sub Import_IR_OOR()

    On Error GoTo Import_Err
    UserImportFile DestRange:=ThisWorkbook.Sheets("OOR1").Range("A1"), _
                   DelFile:=False, _
                   ShowAllData:=True
    On Error GoTo 0

    Sheets("OOR1").Select
    If FindColumn("PO Rel #") = 0 Then ActiveSheet.UsedRange.Cut Destination:=Sheets("OOR2").Range("A1")

    If Range("A1").Value = "" Then
        UserImportFile DestRange:=ThisWorkbook.Sheets("OOR1").Range("A1"), _
                       DelFile:=False, _
                       ShowAllData:=True
    Else
        UserImportFile DestRange:=ThisWorkbook.Sheets("OOR2").Range("A1"), _
                       DelFile:=False, _
                       ShowAllData:=True
    End If

    Exit Sub

Import_Err:
    Debug.Print ERR.Number
    Debug.Print ERR.Description
End Sub
