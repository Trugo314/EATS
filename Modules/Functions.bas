Attribute VB_Name = "Functions"
Option Explicit

Function ConvertColumnToDay(ColumnNumber As Integer) As String
    Select Case ColumnNumber
    Case 4
        ConvertColumnToDay = "Monday"
    Case 5
        ConvertColumnToDay = "Tuesday"
    Case 6
        ConvertColumnToDay = "Wednesday"
    Case 7
        ConvertColumnToDay = "Thursday"
    Case 8
        ConvertColumnToDay = "Friday"
    Case 9
        ConvertColumnToDay = "Saturday"
    Case 10
        ConvertColumnToDay = "Sunday"
    Case Else
        ErrWrite "Error with select case statement. No case matches " & ColumnNumber
        End
    End Select
End Function

Function ConvertDayToColumn(Day As String) As Integer
    Select Case Day
    Case "Monday"
        ConvertDayToColumn = 4
    Case "Tuesday"
        ConvertDayToColumn = 5
    Case "Wednesday"
        ConvertDayToColumn = 6
    Case "Thursday"
        ConvertDayToColumn = 7
    Case "Friday"
        ConvertDayToColumn = 8
    Case "Saturday"
        ConvertDayToColumn = 9
    Case "Sunday"
        ConvertDayToColumn = 10
    Case Else
        ErrWrite "Error with select case statement. No case matches " & Day
        End
    End Select
End Function

Function ConvertTxtFileToStringArray(ByVal FilePath As String, ByVal Delimiter As String) As String()
    'Declare integers
    Dim fileNum As Integer
    
    'Declare strings
    Dim fileContents As String
    
    fileNum = FreeFile
    
    Open FilePath For Binary Access Read As #fileNum
    fileContents = String$(LOF(fileNum), Chr$(0))
    Get #fileNum, , fileContents
    Close #fileNum
    
    ConvertTxtFileToStringArray = Split(fileContents, Delimiter)
End Function

Function FileSelector(Optional FilterList As String = "") As String
    'Declare variants
    Dim FilePath As Variant
    
    FilePath = Application.GetOpenFilename(FilterList)
    
    If FilePath = False Then
        FileSelector = ""
    Else
        FileSelector = FilePath
    End If
End Function

Public Function GetJobCount(SheetName As String) As Integer
    'Declare integers
    Dim Jobs As Integer
    
    'Declare ranges
    Dim CheckCell As Range
    
    Set CheckCell = Sheets(SheetName).Cells(9, 3)
    
    Do
        If ([CheckCell].Value <> stEmpty) Then
            Incr Jobs
        End If
        
        Set CheckCell = CheckCell.Offset(1, 0)
    Loop Until ([CheckCell].Value = stEmpty)
    
    GetJobCount = Jobs
End Function

Public Function GetLastDataRow(SheetName As String, ColumnNum As Integer) As Integer
    GetLastDataRow = Sheets(SheetName).Cells(Rows.Count, ColumnNum).End(xlUp).Row
End Function

Public Function GetUser()
    GetUser = Environ$("UserName")
End Function

Public Function IntDiv(IntNum1 As Integer, IntNum2 As Integer) As Double
    'Declare doubles
    Dim DblNum1 As Double, DblNum2 As Double
    
    DblNum1 = CDbl(IntNum1)
    DblNum2 = CDbl(IntNum2)
    
    IntDiv = (DblNum1 / DblNum2)
End Function

Public Function IsValidInteger(Candidate As Variant) As Boolean
    'Declare integers
    Dim DecimalLoc As Integer
    
    'Declare strings
    Dim stCandidate As String
    
    stCandidate = CStr(Candidate)
    DecimalLoc = InStr(1, stCandidate, ".")
    
    If (DecimalLoc > 0) Then
        IsValidInteger = False
    Else
        IsValidInteger = True
    End If
End Function

Public Function ReadFile(Optional FilterList As String) As String()
    'Declare examples
    'Example filter list: "Text documents(.txt*) (*.txt*), *.txt*"

    'Declare strings
    Dim lines() As String
    Dim path As String

    path = FileSelector(FilterList)
    lines = ConvertTxtFileToStringArray(path, vbCrLf)
    
    ReadFile = lines
End Function

Public Function UIAlphaEntry(Optional header As String = "", Optional Message As String = "", Optional InitString As String = "") As String
    'Declare strings
    Dim entry As String
    
    Do
        entry = Application.InputBox(Message, header, InitString, Type:=2)
        
        If (entry = "") Then
            ErrWrite "Please provide at least one character!"
        ElseIf (entry = "False") Then
            'User pressed cancel
            End
        End If
    Loop Until (entry <> "")
    
    UIAlphaEntry = entry
End Function

Public Function UINumEntry(MinValue As Variant, MaxValue As Variant, Optional header As String = "", Optional Message As String = "", Optional UseBounds As Boolean = False, Optional AsInteger As Boolean = False, Optional DefaultNum As Integer = 0) As Variant
    'Declare booleans
    Dim entryValid As Boolean
    
    'Declare variants
    Dim entry As Variant
    
    If (UseBounds = True) Then
        Message = Message & vbCrLf & _
            "Min value: " & CStr(MinValue)
            
        Message = Message & vbCrLf & _
            "Max value: " & CStr(MaxValue)
    End If
    
    If (AsInteger = True) Then
        Message = Message & vbCrLf & _
            "As integer: " & CStr(AsInteger)
    End If
    
    Do
        If (DefaultNum = 0) Then
            entry = Application.InputBox(Message, header, Type:=1)
        Else
            entry = Application.InputBox(Message, header, DefaultNum, Type:=1)
        End If
        
        'Check for empty entry
        'TODO - If a value of 0 is entered, then due to the variant type entry also equals False
        'and the procedure ends
        If (entry = "") Then
            ErrWrite "Please provide a number!"
            entryValid = False
        ElseIf (entry = False) Then
            'User pressed cancel
            End
        Else
            entryValid = True
        End If
        
        'Check for boundary
        If (UseBounds = True) Then
            If (entry < MinValue) Then
                ErrWrite "Entered value is less than specified minimum!"
                entryValid = False
            ElseIf (entry > MaxValue) Then
                ErrWrite "Entered value is greater than specified maximum!"
                entryValid = False
            Else
                entryValid = True
            End If
        End If
        
        'Check for integer
        If ((AsInteger = True) And (entryValid = True)) Then
            entryValid = IIf((IsValidInteger(entry) = True), True, False)
        End If
    Loop Until (entryValid = True)
    
    UINumEntry = entry
End Function


Public Function UITimeEntry(Optional header As String = "", Optional Message As String = "") As String
    'Declare integers
    Dim entryLength As Integer
    Dim hourDigits As Integer
    
    'Declare strings
    Dim entry As String
    
    Do
        entry = Application.InputBox(Message, header, Type:=2)
        
        If (entry = "") Then
            ErrWrite "Please enter a time!"
        ElseIf (entry = "False") Then
            'User pressed cancel
            End
        ElseIf (IsDate(entry) = False) Then
            ErrWrite "Please enter a time in a correct 24-hour format!"
        End If
    Loop Until (IsDate(entry))
    
    UITimeEntry = entry
End Function

Public Function ValidateDayName(TestString As String) As Boolean
    ValidateDayName = ((TestString = "Monday") Or _
        (TestString = "Tuesday") Or _
        (TestString = "Wednesday") Or _
        (TestString = "Thursday") Or _
        (TestString = "Friday") Or _
        (TestString = "Saturday") Or _
        (TestString = "Sunday"))
End Function
