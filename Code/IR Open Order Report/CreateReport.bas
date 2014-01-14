Attribute VB_Name = "CreateReport"
Option Explicit

Sub CreateOOR()
    Dim ColHeaders As Variant
    Dim TotalCols As Integer
    Dim TotalRows As Long
    Dim i As Long

    Sheets("IR OOR").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    TotalCols = ActiveSheet.UsedRange.Columns.Count

    'Subtotal by UID
    ActiveSheet.UsedRange.Subtotal GroupBy:=1, _
                                   Function:=xlSum, _
                                   TotalList:=Array(11, 12, 13, 17, 18), _
                                   Replace:=True, _
                                   PageBreaks:=False, _
                                   SummaryBelowData:=True

    'Copy subtotals and paste as values
    Cells.Copy
    Range("A1").PasteSpecial xlPasteValues

    'Remove subtotal formatting
    ActiveSheet.UsedRange.RemoveSubtotal

    'Filter for subtotals
    ActiveSheet.UsedRange.AutoFilter Field:=1, Criteria1:="=* Total"

    'Copy subtotals to the open order report
    ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Sheets("Open Order Report").Range("A1")

    'Store column headers
    ColHeaders = Range(Cells(1, 1), Cells(1, TotalCols)).Value

    'Remove subtotals
    Cells.Delete

    'Reinsert column headers
    Rows(1).Insert
    Range(Cells(1, 1), Cells(1, TotalCols)).Value = ColHeaders

    Sheets("Open Order Report").Select

    'Subtract 1 because "Grand Total" will be deleted
    TotalRows = ActiveSheet.UsedRange.Rows.Count - 1

    'Remove Grand Total
    Rows(ActiveSheet.UsedRange.Rows.Count).Delete

    'Fix UID
    Columns(1).Insert
    Range("A1").Value = "UID"
    Range("A2:A" & TotalRows).Formula = "=""=""""""&SUBSTITUTE(B2,"" Total"","""")&"""""""""
    Range("A2:A" & TotalRows).Value = Range("A2:A" & TotalRows).Value
    Columns(2).Delete

    'PO Number
    FillColumn Range("D2:D" & TotalRows), "=IFERROR(IF(VLOOKUP(A2,'IR OOR'!A:D,4,FALSE)=0,"""",VLOOKUP(A2,'IR OOR'!A:D,4,FALSE)),"""")"

    'Line Number
    FillColumn Range("E2:E" & TotalRows), "=IFERROR(IF(VLOOKUP(A2,'IR OOR'!A:E,5,FALSE)=0,"""",VLOOKUP(A2,'IR OOR'!A:E,5,FALSE)),"""")"

    'PO Release #
    FillColumn Range("F2:F" & TotalRows), "=IFERROR(IF(VLOOKUP(A2,'IR OOR'!A:F,6,FALSE)=0,"""",VLOOKUP(A2,'IR OOR'!A:F,6,FALSE)),"""")"

    'IR Part Number
    FillColumn Range("G2:G" & TotalRows), "=IFERROR(IF(VLOOKUP(A2,'IR OOR'!A:G,7,FALSE)=0,"""",VLOOKUP(A2,'IR OOR'!A:G,7,FALSE)),"""")"

    'IR Part Description
    FillColumn Range("H2:H" & TotalRows), "=IFERROR(IF(VLOOKUP(A2,'IR OOR'!A:H,8,FALSE)=0,"""",VLOOKUP(A2,'IR OOR'!A:H,8,FALSE)),"""")"

    'Due Date
    FillColumn Range("O2:O" & TotalRows), "=TEXT(IFERROR(IF(VLOOKUP(A2,'IR OOR'!A:O,15,FALSE)=0,"""",VLOOKUP(A2,'IR OOR'!A:O,15,FALSE)),""""),""mmm dd, yyyy"")", "mmm dd, yyyy"

    'Remove unused columns
    Columns("P:R").Delete
    Columns("N").Delete
    Columns("I:J").Delete
    Columns("B:C").Delete

    'Remove POs older than 60 days
    RemoveData "<" & Format(Date - 60, "mm/dd/yyyy"), 10

    'Remove PO# 341236
    RemoveData "=341236", 2

    TotalRows = Rows(Rows.Count).End(xlUp).Row

    'WESCO PO
    AddColumn "WESCO PO", "=IFERROR(IF(VLOOKUP(A2,'117 OOR'!A:L,12,FALSE)=0,"""",TRIM(VLOOKUP(A2,'117 OOR'!A:L,12,FALSE))),"""")", "@"

    'SUPPLIER
    AddColumn "SUPPLIER", "=IFERROR(IF(VLOOKUP(A2,'117 OOR'!A:N,14,FALSE)=0,"""",TRIM(VLOOKUP(A2,'117 OOR'!A:N,14,FALSE))),"""")", "@"

    'PROMISE DATE
    AddColumn "PROMISE DATE", "=IFERROR(IF(VLOOKUP(A2,'117 OOR'!A:M,13,FALSE)=0,"""",TEXT(VLOOKUP(A2,'117 OOR'!A:M,13,FALSE), ""mmm dd, yyyy"")),"""")", "mmm dd, yyyy"

    'BO
    AddColumn "BO", "=IFERROR(VLOOKUP(A2,'117 OOR'!A:J,10,FALSE),0)"

    'RTS
    AddColumn "RTS", "=IFERROR(VLOOKUP(A2,'117 OOR'!A:I,9,FALSE),0)"

    'SHIPPED
    AddColumn "SHIPPED", "=IFERROR(VLOOKUP(A2,'117 OOR'!A:K,11,FALSE),0)"

    'OLD STATUS
    AddColumn "OLD STATUS", "=IFERROR(IF(VLOOKUP(A2,'Prev OOR'!A:Z," & Sheets("Prev OOR").UsedRange.Columns.Count - 1 & ",FALSE)=0,"""",VLOOKUP(A2,'Prev OOR'!A:Z," & Sheets("Prev OOR").UsedRange.Columns.Count - 1 & ",FALSE)),"""")"

    'STATUS - This must always be the second to last column
    AddColumn "STATUS", "=IF(NOT(IFERROR(VLOOKUP(A2,'117 OOR'!A:A,1,FALSE),"""")="""")=TRUE,IF(IFERROR(VLOOKUP(A2,'117 OOR'!A:J,10,FALSE),0)>0,""B/O"",IF(G2=IFERROR(VLOOKUP(A2,'117 OOR'!A:I,9,FALSE),0),""RTS"",IF(IFERROR(VLOOKUP(A2,'117 OOR'!A:K,11,FALSE),0)=G2,""SHIPPED"",""CHECK""))),""NOO"")"

    'NOTES - This must always be the last column
    AddColumn "NOTES", "=IFERROR(IF(VLOOKUP(A2,'Prev OOR'!A:R," & Sheets("Prev OOR").UsedRange.Columns.Count & ",FALSE)=0,"""",VLOOKUP(A2,'Prev OOR'!A:R," & Sheets("Prev OOR").UsedRange.Columns.Count & ",FALSE)),"""")"
End Sub

Private Sub FillColumn(Rng As Range, Formula As String, Optional NumberFormat As String = "General")
    With Rng
        .NumberFormat = NumberFormat
        .Formula = Formula
        .Value = .Value
    End With
End Sub

Private Sub AddColumn(Header As String, Formula As String, Optional NumberFormat As String = "General")
    Dim TotalRows As Long
    Dim TotalCols As Integer

    TotalRows = ActiveSheet.UsedRange.Rows.Count
    TotalCols = ActiveSheet.UsedRange.Columns.Count + 1

    Cells(1, TotalCols).Value = Header

    With Range(Cells(2, TotalCols), Cells(TotalRows, TotalCols))
        .Formula = Formula
        .NumberFormat = NumberFormat
        .Value = .Value
    End With
End Sub

Private Sub RemoveData(Criteria As String, Field As Integer)
    Dim ColHeaders As Variant
    Dim TotalCols As Integer
    Dim TotalRows As Long

    TotalCols = Columns(Columns.Count).End(xlToLeft).Column
    TotalRows = Rows(Rows.Count).End(xlUp).Row
    ColHeaders = Range(Cells(1, 1), Cells(1, TotalCols)).Value

    'Filter the active sheet
    Range(Cells(1, 1), Cells(TotalRows, TotalCols)).AutoFilter Field:=Field, Criteria1:=Criteria, Operator:=xlAnd

    'Remove the filtered data
    Cells.Delete

    'Reinsert the column headers
    Rows(1).Insert
    Range(Cells(1, 1), Cells(1, TotalCols)).Value = ColHeaders
End Sub
