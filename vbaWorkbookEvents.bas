' Code originally from ThisWorkbook object

Option Explicit

Private Sub Workbook_Open()
    ' Check if Sheet1 cell B2 says Job Number meaning it's a new sheet
    If Sheets("Contract Review").Range("B2").Value = "" Then
        ' Call GetFilePath if b2 is blank or says job number
        Call GetFilePath
    Else
    End If
    
End Sub


Private Sub GetFilePath()
    Dim FilePath As String
    Dim Substring1 As String
    Dim Substring2 As String
    Dim StartIndex1 As Integer
    Dim EndIndex1 As Integer
    Dim StartIndex2 As Integer
    Dim EndIndex2 As Integer

    Application.ScreenUpdating = False

    FilePath = ThisWorkbook.FullName
    'MsgBox FilePath
    Sheets("Sheet2").Range("A1").Value = FilePath
    
    ' Find the position of "JOBS\" in the FilePath string
    StartIndex1 = InStr(1, FilePath, "CURRENT JOBS\") + 13
    
    ' Find the position of "\" in the FilePath string, starting from the position after "NT JOBS\"
    EndIndex1 = InStr(StartIndex1, FilePath, "\")
    
    ' Find the position of "-" in the FilePath string
    StartIndex2 = InStr(1, FilePath, "-")
    
    ' Check if "JOBS\", "-", and "Contract Review" were found in the FilePath string
    If StartIndex1 > 0 And EndIndex1 > 0 And StartIndex2 > 0 Then
        ' Extract the characters between "CURRENT JOBS\" and "-"
        Substring1 = Mid(FilePath, StartIndex1, StartIndex2 - StartIndex1)
        
        ' Extract the characters between "-" and "Contract Review"
        Substring2 = Mid(FilePath, StartIndex2 + 1, EndIndex1 - StartIndex2 - 1)
        
        ' Display the extracted substrings
        MsgBox "Job Number: " & Substring1 & vbNewLine & "Job Name: " & Substring2
        
                ' Insert Substring1 into Sheet1 cell B2
        Sheets("Contract Review").Range("B2").Value = Substring1
        
        ' Insert Substring2 into Sheet1 cell B3
        Sheets("Contract Review").Range("B3").Value = Substring2
        
    Else
        ' Either "JOBS\", "-", or "\AISC" was not found in the FilePath string
        ' MsgBox "Substrings not found"
    End If
        
    Application.ScreenUpdating = True
        
End Sub




