Attribute VB_Name = "FormatData"
Option Explicit

Sub FixHeaders(SheetName As String)
    Dim Col As Integer
    Dim PO_List As Variant
    Dim LINE_List As Variant
    Dim PART_List As Variant
    Dim DESC_List As Variant
    Dim NEEDBY_List As Variant
    Dim QTY_List As Variant
    Dim OPEN_List As Variant
    Dim PrevSheet As Worksheet
    Dim i As Integer
    Dim MaxListSize As Integer

    Set PrevSheet = ActiveSheet
    Sheets(SheetName).Select

    PO_List = Array("PO#", "PO Number", "PO")
    LINE_List = Array("Line", "Line Number", "Line Num", "Line #", "Line #")
    PART_List = Array("Part", "Part #", "Part#", "Part Number", "Item #", "Item#", "Item")
    DESC_List = Array("Description", "Item Description", "Part Description")
    NEEDBY_List = Array("Due Date", "Due")
    QTY_List = Array("Qty")
    OPEN_List = Array("PO Open Qty", "Open Qty", "Open")

    MaxListSize = WorksheetFunction.Max(UBound(LINE_List), UBound(PART_List), UBound(DESC_List), UBound(PO_List))

    For i = 0 To MaxListSize
        On Error GoTo COL_NOT_FOUND
        Col = FindColumn(PO_List(i))
        If Not Col = 0 Then Cells(1, Col).Value = "PO #"

        Col = FindColumn(LINE_List(i))
        If Not Col = 0 Then Cells(1, Col).Value = "PO Line #"

        Col = FindColumn(PART_List(i))
        If Not Col = 0 Then Cells(1, Col).Value = "Item Number"

        Col = FindColumn(DESC_List(i))
        If Not Col = 0 Then Cells(1, Col).Value = "Item Description"

        Col = FindColumn(NEEDBY_List(i))
        If Not Col = 0 Then Cells(1, Col).Value = "Need By Date"

        Col = FindColumn(QTY_List(i))
        If Not Col = 0 Then Cells(1, Col).Value = "PO Qty"

        Col = FindColumn(OPEN_List(i))
        If Not Col = 0 Then Cells(1, Col).Value = "Open PO Qty"
        On Error GoTo 0
    Next

    PrevSheet.Select
    Exit Sub

COL_NOT_FOUND:
    Col = 0
    Resume Next
End Sub
