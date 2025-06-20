'Copied from Sheet1

Private Sub GetFilePath()
    Dim FilePath As String
    FilePath = ThisWorkbook.FullName
    'MsgBox FilePath
    Sheets("Sheet2").Range("A1").Value = FilePath
    
End Sub

