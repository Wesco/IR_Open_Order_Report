Attribute VB_Name = "FormatData"
Option Explicit

Sub FixHeaders(SheetName As String)
    Dim Col As Integer
    Dim PO_Variant_List As Variant
    Dim PrevSheet As Worksheet
    Dim i As Integer

    Set PrevSheet = ActiveSheet
    Sheets(SheetName).Select
    PO_Variant_List = Array("PO#", "PO Number", "PO", "PO Numbers")

    For i = 0 To UBound(PO_Variant_List)
        On Error GoTo COL_NOT_FOUND
        Col = FindColumn(PO_Variant_List(i))
        On Error GoTo 0
        If Not Col = 0 Then
            Cells(1, Col).Value = "PO #"
            Exit For
        End If
    Next
    PrevSheet.Select
    Exit Sub

COL_NOT_FOUND:
    Col = 0
    Resume Next
End Sub
