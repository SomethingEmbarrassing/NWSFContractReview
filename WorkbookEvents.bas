' Code originally from ThisWorkbook object

Option Explicit

Private Sub Workbook_Open()
    ' Check if Sheet1 cell B2 is blank - indicates a new workbook
    If Sheets("Contract Review").Range("B2").Value = "" Then
        ' Populate job number and name from file path
        Call GetFilePath
    End If

    ' If PM (F1) or Ton (E2) are blank pull info from Job List
    If Sheets("Contract Review").Range("F1").Value = "" _
        Or Sheets("Contract Review").Range("E2").Value = "" Then
        Call FillFromJobList
    End If

End Sub

Private Sub FillFromJobList()
    Const SRC_PATH As String = "F:\JOB LIST\JOB LIST2.xlsx"
    Dim srcWB As Workbook
    Dim ws As Worksheet
    Dim jobNum As String
    Dim foundCell As Range
    Dim opened As Boolean

    Application.ScreenUpdating = False

    jobNum = Sheets("Contract Review").Range("B2").Value

    'Check if source workbook already open
    On Error Resume Next
    Set srcWB = Workbooks("JOB LIST2.xlsx")
    On Error GoTo 0

    If srcWB Is Nothing Then
        'Open read-only if locked
        On Error Resume Next
        Set srcWB = Workbooks.Open(SRC_PATH, ReadOnly:=True)
        On Error GoTo 0
        opened = True
    End If

    If srcWB Is Nothing Then
        Application.ScreenUpdating = True
        MsgBox "Unable to open Job List workbook." & vbCrLf & SRC_PATH, vbExclamation
        Exit Sub
    End If

    Set ws = srcWB.Sheets("Add Jobs Here")
    Set foundCell = ws.Columns("C").Find(What:=jobNum, LookIn:=xlValues, LookAt:=xlWhole)

    If Not foundCell Is Nothing Then
        Sheets("Contract Review").Range("E1").Value = ws.Cells(foundCell.Row, "A").Value
        Sheets("Contract Review").Range("E2").Value = ws.Cells(foundCell.Row, "J").Value
    Else
        MsgBox "Job number " & jobNum & " not found in Job List.", vbInformation
    End If

    If opened Then srcWB.Close SaveChanges:=False

    Application.ScreenUpdating = True
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




