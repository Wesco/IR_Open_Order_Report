Attribute VB_Name = "BuildOOR"
Option Explicit

Sub CopyReport()
    Dim Col As Integer
    Dim ColList As Variant
    Dim OpenQty_Col As Integer
    Dim TotalRows As Long
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
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    Range("A1:H1") = Array("PO", "Rel", "Line", "Part", "Description", "Need By Date", "PO Qty", "Open Qty")

End Sub
