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
    With Range("D2:D" & TotalRows)
        .NumberFormat = "General"
        .Formula = "=IFERROR(IF(VLOOKUP(A2,'IR OOR'!A:D,4,FALSE)=0,"""",VLOOKUP(A2,'IR OOR'!A:D,4,FALSE)),"""")"
        .Value = .Value
    End With

    'Line Number
    With Range("E2:E" & TotalRows)
        .NumberFormat = "General"
        .Formula = "=IFERROR(IF(VLOOKUP(A2,'IR OOR'!A:E,5,FALSE)=0,"""",VLOOKUP(A2,'IR OOR'!A:E,5,FALSE)),"""")"
        .Value = .Value
    End With

    'PO Release #
    With Range("F2:F" & TotalRows)
        .NumberFormat = "General"
        .Formula = "=IFERROR(IF(VLOOKUP(A2,'IR OOR'!A:F,6,FALSE)=0,"""",VLOOKUP(A2,'IR OOR'!A:F,6,FALSE)),"""")"
        .Value = .Value
    End With

    'IR Part Number
    With Range("G2:G" & TotalRows)
        .NumberFormat = "General"
        .Formula = "=IFERROR(IF(VLOOKUP(A2,'IR OOR'!A:G,7,FALSE)=0,"""",VLOOKUP(A2,'IR OOR'!A:G,7,FALSE)),"""")"
        .Value = .Value
    End With

    'IR Part Description
    With Range("H2:H" & TotalRows)
        .NumberFormat = "General"
        .Formula = "=IFERROR(IF(VLOOKUP(A2,'IR OOR'!A:H,8,FALSE)=0,"""",VLOOKUP(A2,'IR OOR'!A:H,8,FALSE)),"""")"
        .Value = .Value
    End With

    'Due Date
    With Range("O2:O" & TotalRows)
        .NumberFormat = "mmm dd, yyyy"
        .Formula = "=TEXT(IFERROR(IF(VLOOKUP(A2,'IR OOR'!A:O,15,FALSE)=0,"""",VLOOKUP(A2,'IR OOR'!A:O,15,FALSE)),""""),""mmm dd, yyyy"")"
        .Value = .Value
    End With

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

    'FOUND
    Range("K1").Value = "FOUND"
    Range("K2:K" & TotalRows).Formula = "=NOT(IFERROR(VLOOKUP(A2,'117 OOR'!A:A,1,FALSE),"""")="""")"
    Range("K2:K" & TotalRows).Value = Range("K2:K" & TotalRows).Value

    'ON BO
    Range("L1").Value = "ON BO"
    Range("L2:L" & TotalRows).Formula = "=IF(IFERROR(VLOOKUP(A2,'117 OOR'!A:J,10,FALSE),0)>0,TRUE,FALSE)"
    Range("L2:L" & TotalRows).Value = Range("L2:L" & TotalRows).Value

    'BO
    Range("M1").Value = "BO"
    Range("M2:M" & TotalRows).Formula = "=IFERROR(VLOOKUP(A2,'117 OOR'!A:J,10,FALSE),0)"
    Range("M2:M" & TotalRows).Value = Range("M2:M" & TotalRows).Value

    'RTS
    Range("N1").Value = "RTS"
    Range("N2:N" & TotalRows).Formula = "=IFERROR(VLOOKUP(A2,'117 OOR'!A:I,9,FALSE),0)"
    Range("N2:N" & TotalRows).Value = Range("N2:N" & TotalRows).Value

    'SHIPPED
    Range("O1").Value = "SHIPPED"
    Range("O2:O" & TotalRows).Formula = "=IFERROR(VLOOKUP(A2,'117 OOR'!A:K,11,FALSE),0)"
    Range("O2:O" & TotalRows).Value = Range("O2:O" & TotalRows).Value

    'STATUS
    Range("P1").Value = "STATUS"
    Range("P2:P" & TotalRows).Formula = "=IF(K2=TRUE,IF(M2>0,""B/O"",IF(G2=N2,""RTS"",IF(O2=G2,""SHIPPED"",""CHECK""))),""NOO"")"
    Range("P2:P" & TotalRows).Value = Range("P2:P" & TotalRows).Value

    'OLD STATUS
    Range("Q1").Value = "OLD STATUS"
    Range("Q2:Q" & TotalRows).Formula = "=IFERROR(IF(VLOOKUP(A2,'Prev OOR'!A:P,16,FALSE)=0,"""",VLOOKUP(A2,'Prev OOR'!A:P,16,FALSE)),"""")"
    Range("Q2:Q" & TotalRows).Value = Range("Q2:Q" & TotalRows).Value

    'NOTES
    Range("R1").Value = "NOTES"
    Range("R2:R" & TotalRows).Formula = "=IFERROR(IF(VLOOKUP(A2,'Prev OOR'!A:R,18,FALSE)=0,"""",VLOOKUP(A2,'Prev OOR'!A:R,18,FALSE)),"""")"
    Range("R2:R" & TotalRows).Value = Range("R2:R" & TotalRows).Value
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
