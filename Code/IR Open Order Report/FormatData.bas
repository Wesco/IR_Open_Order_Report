Attribute VB_Name = "FormatData"
Option Explicit

Sub FormatMaster()
    
End Sub

Sub Format117()
    Dim TotalCols As Integer
    Dim TotalRows As Long
    Dim i As Long

    Sheets("117 OOR").Select
    TotalCols = ActiveSheet.UsedRange.Columns.Count

    'Remove report footer
    Rows(ActiveSheet.UsedRange.Rows.Count).Delete

    'Remove report header
    Rows(1).Delete

    'Remove all unneeded columns
    For i = TotalCols To 1 Step -1
        If Cells(1, i).Value <> "CUSTOMER REFERENCE NO" And _
           Cells(1, i).Value <> "CUSTOMER PART NUMBER" And _
           Cells(1, i).Value <> "ITEM DESCRIPTION" And _
           Cells(1, i).Value <> "ORDER QTY" And _
           Cells(1, i).Value <> "AVAILABLE QTY" And _
           Cells(1, i).Value <> "QTY TO SHIP" And _
           Cells(1, i).Value <> "BO QTY" And _
           Cells(1, i).Value <> "QTY SHIPPED" Then
            Columns(i).Delete
        End If
    Next

    TotalCols = ActiveSheet.UsedRange.Columns.Count
    TotalRows = ActiveSheet.UsedRange.Rows.Count

    'Remove spaces from "CUSTOMER REFERENCE NO" & "CUSTOMER PART NUMBER"
    Range("A2:B" & TotalRows).Replace "=""", ""
    Range("A2:B" & TotalRows).Replace """", ""
    Range("A2:B" & TotalRows).Replace " ", ""

    'Remove extra spaces from part descriptions
    Columns(3).Insert
    Range("C1").Value = "ITEM DESCRIPTION"
    Range("C2:C" & TotalRows).Formula = "=TRIM(D2)"
    Range("C2:C" & TotalRows).Value = Range("C2:C" & TotalRows).Value
    Columns(4).Delete

    'Create Part Number Column
    Columns(1).Insert
    Range("A1").Value = "PART NUMBER"
    For i = 2 To TotalRows
        'Search the item description for the customer part number

    Next

    'Create UID column
    Columns(1).Insert
    Range("A1").Value = "UID"
    Range("A2:A" & TotalRows).Formula = ""

End Sub

Sub FindPart()
    Dim TotalRows As Long
    Dim PartList As Variant
    Dim DescList As Variant
    Dim Result As Variant
    Dim i As Long
    Dim j As Long

    'Load part numbers into an array
    Sheets("Master").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    PartList = Range("B2:B" & TotalRows)

    'Load item descriptions into an array
    Sheets("117 OOR").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    DescList = Range("D2:D" & TotalRows)


    For i = 1 To UBound(DescList)
        If Cells(i + 1, 3).Value = "" Then
            For j = 1 To UBound(PartList)
                Result = InStr(1, DescList(i, 1), PartList(j, 1))

                If Result <> 0 Then
                    Cells(i + 1, 1).Value = PartList(j, 1)
                    Exit For
                End If
            Next
        End If
    Next
End Sub


















