Attribute VB_Name = "Imports"
Option Explicit

Sub Import_IR_OOR()

    On Error GoTo IMPORT_ERR
    UserImportFile DestRange:=ThisWorkbook.Sheets("IR DLC").Range("A1"), _
                   DelFile:=False, _
                   ShowAllData:=True
    On Error GoTo 0

    Sheets("IR DLC").Select

    On Error GoTo COL_NOT_FOUND
    FindColumn ("PO Rel #")
    On Error GoTo 0

    If Range("A1").Value = "" Then
        UserImportFile DestRange:=ThisWorkbook.Sheets("IR DLC").Range("A1"), _
                       DelFile:=False, _
                       ShowAllData:=True
    Else
        UserImportFile DestRange:=ThisWorkbook.Sheets("IR Mox").Range("A1"), _
                       DelFile:=False, _
                       ShowAllData:=True
    End If

    Exit Sub

COL_NOT_FOUND:
    ActiveSheet.UsedRange.Cut Destination:=Sheets("IR Mox").Range("A1")
    Resume Next

IMPORT_ERR:
    Debug.Print Err.Number
    Debug.Print Err.Description
    Exit Sub
End Sub

Sub Import117()
    Dim PrevSheet As Worksheet
    Dim TotalRows As Long
    Dim TotalCols As Integer
    Dim CustPartCol As String
    Dim ItemDescCol As String
    Dim Col As Integer

    Set PrevSheet = ActiveSheet
    Import117byISN ALL, Sheets("117").Range("A1"), "24"

    Sheets("117").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    TotalCols = ActiveSheet.UsedRange.Columns.Count

    If [A1].Value = "" Then
        Err.Raise 53
    Else
        Rows(TotalRows).Delete
        Rows(1).Delete

        With Range(Cells(2, 1), Cells(TotalRows, TotalCols))
            .Replace "=""", "", xlPart, xlByRows, False
            .Replace """", "", xlPart, xlByRows, False
            .Replace " ", "", xlPart, xlByRows, False
        End With

        Columns("A:A").Insert
        [A1].Value = "UID"
        With Range(Cells(2, 1), Cells(TotalRows, 1))
            .Formula = "=IF(N2="""",M2&BK2,M2&N2)"
            .Value = .Value
        End With
        Col = FindColumn("CUSTOMER PART NUMBER") + 1
        Columns(Col).Insert
        Cells(1, Col).Value = "CUSTOMER PART NUMBER"

        CustPartCol = Cells(2, FindColumn("CUSTOMER PART NUMBER")).Address(False, False)
        ItemDescCol = Cells(2, FindColumn("ITEM DESCRIPTION")).Address(False, False)

        With Range(Cells(2, Col), Cells(TotalRows, Col))
            .Formula = _
            "=IF(" & CustPartCol & "="""",IFERROR(MID(" & ItemDescCol & ",FIND(""***""," & ItemDescCol & ")+3,8),MID(" & _
                       ItemDescCol & ",FIND(""""," & ItemDescCol & ")-0,9))," & CustPartCol & ")"

            .Value = .Value
            .HorizontalAlignment = xlLeft
        End With
        Columns(Col - 1).Delete
    End If

End Sub
