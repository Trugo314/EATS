Attribute VB_Name = "Procedures"
Option Explicit

Sub CenterAlign(Sheet As String, Optional ColumnRange As String = "", Optional cell As Object = vbNull)
    If (ColumnRange <> "") Then
        Sheets(Sheet).Columns(ColumnRange).HorizontalAlignment = xlCenter
        Sheets(Sheet).Columns(ColumnRange).AutoFit
    ElseIf (cell <> vbNull) Then
        cell.HorizontalAlignment = xlCenter
    Else
        ErrWrite "In subprocedure CenterAlign(): Argument ""ColumnRange"" or ""Cell"" is required!"
        End
    End If
End Sub

Sub Clear(ByRef Value As Variant)
    If ((VarType(Value) <> vbInteger) And _
    (VarType(Value) <> vbLong) And _
    (VarType(Value) <> vbSingle) And _
    (VarType(Value) <> vbDouble) And _
    (VarType(Value) <> vbDecimal)) Then
        ErrWrite "Error in subprocedure Clear()" & vbCrLf & "Variable type " & CStr(VarType(Value)) & " is not a numeric type."
    Else
        Value = 0
    End If
End Sub

Sub Decr(ByRef Value As Integer)
    Value = Value - 1
End Sub

Sub ErrLog(ErrName As String, DoNotContinue As Boolean)
    'For logging pre-defined errors
End Sub

Sub ErrWrite(ByVal Message As String)
    'For writing custom errors in-line
    
    resOKOnly = MsgBox(Message, vbCritical + vbOKOnly, stError)
End Sub

Sub Incr(ByRef Value As Integer)
    Value = Value + 1
End Sub

Sub UpdateSheet(ColumnNumber As Integer, EndTime As String)
    'Declare integers
    Dim JobCount As Integer
    Dim JobIndex As Integer
    Dim JobOffset As Integer
    
    'Declare ranges
    Dim CurrentJobCell As Range
    Dim EndTimeCell As Range
    Dim FirstJobCell As Range
    Dim HoursWorkedCell As Range
    Dim MealTimeCell As Range
    Dim StartTimeCell As Range
    Dim TotalHoursCell As Range
    
    'Declare strings
    Const stDefaultProjectNumber As String = "LXC-xxx"
    Const stNoJobsError As String = "No jobs present to track hours for! At least one job number must be added."
    
    With Sheets("Current Week")
        Set StartTimeCell = .Cells(3, ColumnNumber)
        Set MealTimeCell = .Cells(4, ColumnNumber)
        Set EndTimeCell = .Cells(5, ColumnNumber)
        Set TotalHoursCell = .Cells(6, ColumnNumber)
        Set HoursWorkedCell = .Cells(7, ColumnNumber)
        Set FirstJobCell = .Cells(9, 3)
    End With
    
    JobCount = GetJobCount("Current Week")
    
    If (JobCount = 0) Then
        resYesNo = MsgBox("No jobs detected, would you like to add one?", vbQuestion + vbYesNo, "Data Entry")
        
        If (resYesNo = vbYes) Then
            [FirstJobCell].Value = UIAlphaEntry("Job Entry", "Enter a project number", stDefaultProjectNumber)
            JobIndex = 1
            Set CurrentJobCell = FirstJobCell
        Else
            ErrWrite stNoJobsError
            End
        End If
    ElseIf (JobCount = 1) Then
        JobIndex = 1
        Set CurrentJobCell = FirstJobCell
    Else
        Do
            JobIndex = UINumEntry(1, JobCount, "Job Index Entry", "Multiple jobs detected. Enter the index to update.", True, True)
            
            If (JobIndex > JobCount) Then
                ErrWrite "The entered job index exceeds the current job count!"
                End
            End If
            
            JobOffset = (JobIndex - 1) 'JobOffset of 0 equals first job, JobOffset of 1 equals second job, etc.
            Set CurrentJobCell = FirstJobCell.Offset(JobOffset, 0)
            
            
        Loop Until (bAlwaysFalse)
    End If
End Sub

Sub ValidatePreviousDays()
    
End Sub

Sub WriteLineToTxtFile(FilePath As String, ByVal Message As String)
    'Declare integers
    Dim fileNum As Integer
    
    fileNum = FreeFile
    
    Open FilePath For Append Access Write As #fileNum
    Print #fileNum, Message
    Close #fileNum
End Sub
