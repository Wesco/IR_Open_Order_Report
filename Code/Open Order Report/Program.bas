Attribute VB_Name = "Program"
Option Explicit

Sub Main()
    On Error GoTo ErrHandler
    Clean
    Import_IR_OOR
    Import117
    FixHeaders "IR DLC"
    FixHeaders "IR Mox"
    CreateReport
    On Error GoTo 0
    Exit Sub

ErrHandler:
    Select Case Err.Number
        Case CustErr.COLNOTFOUND:
            MsgBox Prompt:="Column """ & Err.Description & """ not found!", _
                   Title:="Error " & CustErr.COLNOTFOUND
        Case 53:
            Exit Sub
        
        Case Else:
            On Error GoTo 0
            Resume
    End Select
End Sub

Sub Clean()
    Dim w As Variant

    For Each w In ThisWorkbook.Sheets
        If Not w.Name = "Macro" Then
            w.Cells.Delete
        End If
    Next
End Sub
