Attribute VB_Name = "Module1"
Option Explicit

Sub CreateOOR()
    Sheets("OOR").Select
    Range("A1:H1") = Array("PO", "Rel", "Line", "Part", "Description", "Need By Date", "PO Qty", "Open Qty")
End Sub

Sub CopyReport()
    Dim Col As Integer
    Dim OpenQty_Col As Integer
    Dim TotalRows As Long

    Sheets("OOR1").Select
    
    Col = FindColumn("PO #")
    CopyColumn Col, Sheets("OOR").Range("A2")
    
    Col = FindColumn("PO Rel #")
    CopyColumn Col, Sheets("OOR").Range("B2")
    
    Col = FindColumn("PO Line #")
    CopyColumn Col, Sheets("OOR").Range("C2")
    
    Col = FindColumn("Item Number")
    CopyColumn Col, Sheets("OOR").Range("D2")
    
    Col = FindColumn("Item Description")
    CopyColumn Col, Sheets("OOR").Range("E2")
    
    Col = FindColumn("Need By Date")
    CopyColumn Col, Sheets("OOR").Range("F2")
    
    Col = FindColumn("PO Qty")
    CopyColumn Col, Sheets("OOR").Range("G2")
    
    Col = FindColumn("Open PO Qty")
    CopyColumn Col, Sheets("OOR").Range("H2")
    
    TotalRows = ActiveSheet.UsedRange.Rows.Count
End Sub

Sub CopyColumn(Column As Integer, Destination As Range)
    If Column > 0 Then Range(Cells(2, Column), Cells(ActiveSheet.UsedRange.Rows.Count, Column)).Copy Destination:=Destination
End Sub
