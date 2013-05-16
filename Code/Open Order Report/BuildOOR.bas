Attribute VB_Name = "BuildOOR"
Option Explicit

Sub CreateReport()
    Dim Col As Integer
    Dim ColList As Variant
    Dim TotalRows As Long
    Dim OORTotalRows As Long
    Dim i As Integer

    Sheets("IR DLC").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count

    'The column these will be copied to is respective to its position in the array
    ColList = Array("PO #", "PO Rel #", "PO Line #", "Item Number", "Item Description", "Need By Date", "PO Qty", "Open PO Qty")

    For i = 0 To UBound(ColList)
        Col = FindColumn(ColList(i))
        Range(Cells(2, Col), Cells(TotalRows, Col)).Copy Destination:=Sheets("OOR").Cells(2, i + 1)
    Next

    Sheets("OOR").Select
    'Set column headers
    Range("A1:H1") = Array("PO", "Rel", "Line", "Part", "Description", "Need By Date", "PO Qty", "Open Qty")
    'Get the number of rows
    OORTotalRows = ActiveSheet.UsedRange.Rows.Count + 1

    Sheets("IR Mox").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count

    For i = 0 To UBound(ColList)
        On Error GoTo COL_NOT_FOUND
        Col = FindColumn(ColList(i))
        Range(Cells(2, Col), Cells(TotalRows, Col)).Copy Destination:=Sheets("OOR").Cells(OORTotalRows, i + 1)
        On Error GoTo 0
    Next

    Sheets("OOR").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count

    Columns("A:A").Insert
    [A1].Value = "UID"
    With Range(Cells(2, 1), Cells(TotalRows, 1))
        .Formula = "=" & Cells(2, FindColumn("PO")).Address(False, False) & "&" & Cells(2, FindColumn("Line")).Address(False, False)
        .Value = .Value
    End With

    Exit Sub

COL_NOT_FOUND:
    If ColList(i) = "PO Rel #" Then
        Col = 0
        Resume Next
    Else
        MsgBox Prompt:="Column """ & Err.Description & """ not found!", _
               Title:="Error " & CustErr.COLNOTFOUND
    End If
End Sub
