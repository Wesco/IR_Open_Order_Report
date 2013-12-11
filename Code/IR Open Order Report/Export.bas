Attribute VB_Name = "Export"
Option Explicit

Sub ExportOOR()
    Dim FilePath As String
    Dim FileName As String
    Dim FileExt As String

    FilePath = "\\7938-HP02\Shared\IR-Davidson-Mox\Open Order Report\" & Format(Date, "yyyy") & "\" & Format(Date, "mmm") & "\"
    FileName = "OOR " & Format(Date, "yyyy-mm-dd")
    FileExt = ".xlsx"

    If Not FolderExists(FilePath) Then RecMkDir Path
    
    Sheets("Open Order Report").Copy
    ActiveWorkbook.SaveAs FilePath & FileName & FileExt, xlOpenXMLWorkbook
    ActiveWorkbook.Close
End Sub
