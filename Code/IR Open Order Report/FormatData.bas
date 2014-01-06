Attribute VB_Name = "FormatData"
Option Explicit

Sub FormatGaps()
    Dim TotalRows As Long

    Sheets("Gaps").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count

    'Remove SIM numbers stored as numbers
    Columns(1).Delete

    'Insert column for SIMs
    Columns(1).Insert

    'Store SIMs as text
    Range("A1").Value = "SIM"
    Range("A2:A" & TotalRows).Formula = "=""=""""""&C2&D2&"""""""""
    Range("A2:A" & TotalRows).Value = Range("A2:A" & TotalRows).Value
End Sub

Sub FormatMaster()
    Dim TotalRows As Long

    Sheets("Master").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count

    'Insert a column for descriptions
    Columns(1).Insert

    'Lookup descriptions on GAPS, if they are not found, use the description listed on the master
    Range("A1").Value = "Description"
    Range("A2:A" & TotalRows).Formula = "=TRIM(IFERROR(VLOOKUP(D2,Gaps!A:F,6,FALSE),G2))"
    Range("A2:A" & TotalRows).Value = Range("A2:A" & TotalRows).Value
End Sub

Sub Format117()
    Dim ColHeaders As Variant
    Dim TotalCols As Integer
    Dim TotalRows As Long
    Dim PartList As Variant
    Dim DescList As Variant
    Dim Result As Variant
    Dim i As Long
    Dim j As Long

    Sheets("117 OOR").Select
    TotalCols = ActiveSheet.UsedRange.Columns.Count

    'Remove report footer
    Rows(ActiveSheet.UsedRange.Rows.Count).Delete

    'Remove report header
    Rows(1).Delete

    'Remove all unneeded columns
    For i = TotalCols To 1 Step -1
        If Cells(1, i).Value <> "ORDER NO" And _
           Cells(1, i).Value <> "CUSTOMER REFERENCE NO" And _
           Cells(1, i).Value <> "CUSTOMER PART NUMBER" And _
           Cells(1, i).Value <> "LINE NO" And _
           Cells(1, i).Value <> "ITEM DESCRIPTION" And _
           Cells(1, i).Value <> "ORDER QTY" And _
           Cells(1, i).Value <> "AVAILABLE QTY" And _
           Cells(1, i).Value <> "QTY TO SHIP" And _
           Cells(1, i).Value <> "BO QTY" And _
           Cells(1, i).Value <> "QTY SHIPPED" Then
            Columns(i).Delete
        End If
    Next

    'Load the correct column order into an array
    ColHeaders = Array("ORDER NO", _
                       "CUSTOMER REFERENCE NO", _
                       "CUSTOMER PART NUMBER", _
                       "LINE NO", _
                       "ITEM DESCRIPTION", _
                       "ORDER QTY", _
                       "AVAILABLE QTY", _
                       "QTY TO SHIP", _
                       "BO QTY", _
                       "QTY SHIPPED")

    'Compare the correct column order to the actual column order
    For i = 0 To UBound(ColHeaders)
        If Cells(1, i + 1).Value <> ColHeaders(i) Then
            Err.Raise CustErr.INVALID_COLUMN_ORDER, "Format117", "The column order on the 117 report is incorrect."
        End If
    Next

    TotalCols = ActiveSheet.UsedRange.Columns.Count
    TotalRows = ActiveSheet.UsedRange.Rows.Count

    'Remove spaces from "CUSTOMER REFERENCE NO" & "CUSTOMER PART NUMBER"
    Range("B2:C" & TotalRows).Replace "=""", ""
    Range("B2:C" & TotalRows).Replace """", ""
    Range("B2:C" & TotalRows).Replace " ", ""

    'Remove extra spaces from part descriptions
    Columns(5).Insert
    Range("E1").Value = "ITEM DESCRIPTION"
    Range("E2:E" & TotalRows).Formula = "=TRIM(F2)"
    Range("E2:E" & TotalRows).Value = Range("E2:E" & TotalRows).Value
    Columns(6).Delete

    'Lookup missing part numbers by description on Master
    Columns(3).Insert
    Range("C1").Value = "CUSTOMER PART NUMBER"
    'Lookup part on master, if it is not found and part number is not blank, keep the part number listed on the OOR
    'if the part number was not found and is blank, lookup part on master using the item description
    Range("C2:C" & TotalRows).Formula = _
    "=IFERROR(IF(NOT(IFERROR(VLOOKUP(D2,Master!B:B,1,FALSE),"""")=""""),VLOOKUP(D2,Master!B:B,1,FALSE),VLOOKUP(F2,Master!A:B,2,FALSE)),"""")"
    Range("C2:C" & TotalRows).Value = Range("C2:C" & TotalRows).Value

    'If a part number could not be found on the master use the part number listed from the system
    For i = 2 To TotalRows
        If Cells(i, 3).Value = "" And Cells(i, 4).Value <> "" Then Cells(i, 3).Value = Cells(i, 4).Value
    Next

    Columns(4).Delete

    'Load part numbers into an array
    Sheets("Master").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    PartList = Range("B2:B" & TotalRows)

    'Load item descriptions into an array
    Sheets("117 OOR").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    DescList = Range("F2:F" & TotalRows)

    'Find part numbers in item descriptions
    For i = 1 To UBound(DescList)
        'If CUSTOMER PART NUMBER is blank
        If Cells(i + 1, 3).Value = "" Then
            'See if any part numbers are in the item description
            For j = 1 To UBound(PartList)
                Result = InStr(1, DescList(i, 1), PartList(j, 1))
                If Result <> 0 Then
                    Cells(i + 1, 3).Value = PartList(j, 1)
                    Exit For
                End If
            Next
        End If
    Next

    'Create UID column
    Columns(1).Insert
    Range("A1").Value = "UID"
    Range("A2:A" & TotalRows).Formula = "=""="""""" & C2 & D2 & """""""""
    Range("A2:A" & TotalRows).Value = Range("A2:A" & TotalRows).Value
End Sub

Sub FormatIROOR()
    Dim ColHeaders As Variant
    Dim TotalRows As Long
    Dim TotalCols As Integer
    Dim i As Long

    Sheets("IR OOR").Select
    ActiveSheet.UsedRange.UnMerge

    'Remove report header
    Rows(1).Delete

    'Store correct column header order
    ColHeaders = Array("Supplier Code", _
                       "Supplier Name", _
                       "Location Name", _
                       "PO Number", _
                       "Line Number", _
                       "PO Releases", _
                       "IR Part Number", _
                       "IR Part Description", _
                       "Supplier Part Number", _
                       "Order Date", _
                       "Ordered Quantity", _
                       "Quantity Received", _
                       "Open Quantity", _
                       "Performance Date", _
                       "Actual PO Due Date", _
                       "Currency Code", _
                       "PO Price", _
                       "Extended PO Price")

    'Compare correct column headers to actual column headers
    For i = 0 To UBound(ColHeaders)
        If Cells(1, i + 1).Value <> ColHeaders(i) Then
            Err.Raise CustErr.INVALID_COLUMN_ORDER
        End If
    Next

    TotalRows = ActiveSheet.UsedRange.Rows.Count
    TotalCols = ActiveSheet.UsedRange.Columns.Count

    'Copy PO numbers into blank cells
    For i = 2 To TotalRows
        If Cells(i, 4).Value = "" Then Cells(i, 4).Value = Cells(i - 1, 4).Value
    Next

    'Replace 0's in release number with blank string
    For i = 2 To TotalRows
        If Cells(i, 6).Value = "0" Then Cells(i, 6).Value = ""
    Next

    'Remove "USF HOLLAND" from PO Releases
    For i = 2 To TotalRows
        If Cells(i, 6).Value = "USF HOLLAND" Then
            Cells(i, 6).ClearContents
        End If
    Next

    'Create UID column
    Range("A1").Value = "UID"
    Range("A2:A" & TotalRows).ClearContents
    Range("A2:A" & TotalRows).NumberFormat = "General"
    Range("A2:A" & TotalRows).Formula = "=""=""""""&IF(NOT(F2=""""),D2&""-""&F2&G2,D2&G2)&"""""""""
    Range("A2:A" & TotalRows).Value = Range("A2:A" & TotalRows).Value

    'Remove irrelevant lines
    RemoveData "=*1L*", 4
    RemoveData "=", 1

    TotalRows = ActiveSheet.UsedRange.Rows.Count

    'Autofilter data
    ActiveSheet.UsedRange.AutoFilter

    'Sort A-Z
    With ActiveSheet.AutoFilter.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("A1:A" & TotalRows), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub FormatOOR()
    Sheets("Open Order Report").Select
    With ActiveSheet.UsedRange
        .WrapText = False
        .Interior.Pattern = xlNone
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        .Font.Name = "Calibri"
        .Font.Size = "11"
        .Columns.AutoFit
    End With
End Sub

Private Sub RemoveData(Criteria As String, Field As Integer)
    Dim ColHeaders As Variant
    Dim TotalCols As Integer
    Dim TotalRows As Long

    TotalCols = ActiveSheet.UsedRange.Columns.Count
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    ColHeaders = Range(Cells(1, 1), Cells(1, TotalCols)).Value

    'Filter the active sheet
    Range(Cells(1, 1), Cells(TotalRows, TotalCols)).AutoFilter Field:=Field, Criteria1:=Criteria, Operator:=xlAnd

    'Remove the filtered data
    Cells.Delete

    'Reinsert the column headers
    Rows(1).Insert
    Range(Cells(1, 1), Cells(1, TotalCols)).Value = ColHeaders
End Sub
