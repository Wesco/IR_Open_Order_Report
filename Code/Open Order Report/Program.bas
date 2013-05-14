Attribute VB_Name = "Program"
Option Explicit

Sub Main()
    On Error GoTo ErrHandler
    Import_IR_OOR
    FixHeaders "IR DLC"
    FixHeaders "IR Mox"
    CopyReport
    On Error GoTo 0
    Exit Sub

ErrHandler:
    Select Case Err.Number
        Case CustErr.COLNOTFOUND:
            MsgBox Prompt:="Column """ & Err.Description & """ not found!", _
                   Title:="Error " & CustErr.COLNOTFOUND
    End Select
End Sub
