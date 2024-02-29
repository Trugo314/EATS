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

Sub CheckSheet()
     'Declare integers
     Dim ColumnNumber As Integer
     Dim WeekNumber As Integer
     
     Dim LastUpdateCell As Range
     
     'Declare strings
     Dim stDate As String
     Dim stEndTime As String
     Dim stToday As String
     Dim stUpdateTime As String

    'Assignments
    stToday = Format(Date, "dddd")
    stDate = Format(Date, "yyyy/mm/dd")
    stEndTime = Format(Time, "hh:mm")
    stUpdateTime = Format(Time, "hh:mm")
    WeekNumber = WorksheetFunction.IsoWeekNum(Now)
    
    With Sheets("Current Week")
        Set LastUpdateCell = .Cells(8, 2)
    End With
    
    ColumnNumber = ConvertDayToColumn(stToday)
    
    'Check if previous data needs updated
    'TODO - Probably need to differentiate between validating and updating previous data
    'When updating for the first time/day, I want to validate all previous days in the week as a sanity check
    'I also want to ask to update the previous work day end time even if all sanity checks pass
    If (stDate <> Left([LastUpdateCell].Text, 10)) Then
        ValidatePreviousDays ColumnNumber
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

Sub FixPreviousDays(Row As Integer, Column As Integer)
    'Declare strings
    Dim stData As String
    Dim stDay As String
    
    Select Case Row
    Case 3
        stData = "start time"
    Case 4
        stData = "meal duration"
    Case 5
        stData = "end time"
    Case Else
        ErrWrite "Error with select case statement. Row number invalid."
        End
    End Select
    
    stDay = ConvertColumnToDay(Column)
    
    Sheets("Current Week").Cells(Row, Column).Value = UITimeEntry("Fix Data!", "Data for " & stData & " on " & stDay & " is missing!" & vbCrLf & "Please enter the data to proceed!")
End Sub

Sub Incr(ByRef Value As Integer)
    Value = Value + 1
End Sub

Sub InitData()
    
End Sub

Sub UpdateSheet(ColumnNumber As Integer, EndTime As String)
    'Declare doubles
    Dim PreviousWorkHours As Double
    Dim WorkingHours As Double

    'Declare integers
    Dim JobCount As Integer
    Dim JobIndex As Integer
    Dim JobOffset As Integer
    Dim LastJobRow As Integer
    Dim LunchMinutes As Integer
    Dim TotalMinutes As Integer
    
    'Declare ranges
    Dim CurrentJobCell As Range
    Dim EndTimeCell As Range
    Dim FirstJobCell As Range
    Dim HourEntryCell As Range
    Dim HoursWorkedCell As Range
    Dim JobHoursRange As Range
    Dim MealTimeCell As Range
    Dim StartTimeCell As Range
    Dim TotalHoursCell As Range
    
    'Declare strings
    Const stDefaultProjectNumber As String = "LXC-xxx"
    Const stNoJobsError As String = "No jobs present to track hours for! At least one job number must be added."
    Dim stStartTime As String
    Dim stTotalHours As String
    Dim stTotalHoursSplit() As String
    
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
        Loop Until (JobIndex <= JobCount)
    End If
    
    'Determine which cell is receiving hours update
    Set HourEntryCell = Sheets("Current Week").Cells([CurrentJobCell].Row, ColumnNumber)
    
    'Update start time
    If ([StartTimeCell].Value = stEmpty) Then
        stStartTime = UITimeEntry("Start Time", "Enter the time at which you started working" & vbCrLf & "Format: ""hh:mm""")
    End If
    
    'Update meal duration
    If ([MealTimeCell].Value <> stEmpty) Then
        LunchMinutes = ([MealTimeCell].Value * 60) 'Conversion from decimal time to minutes
        'TODO - Fix ElseIf expression below. TimeValue argument on right side of operand should not be hard-coded
    ElseIf (TimeValue(EndTime) > TimeValue("12:00")) Then
        resYesNo = MsgBox("Would you like to enter lunch time?", vbQuestion + vbYesNo, "Lunch Time Entry")
        
        If (resYesNo = vbYes) Then
            LunchMinutes = UINumEntry(0, 60, "Lunch Time Entry", "Enter the time taken (in minutes) for lunch.", True, True, 30)
            [MealTimeCell].Value = (LunchMinutes / 60) 'Conversion from minutes to decimal time
        Else
            LunchMinutes = 0
            
            resYesNo = MsgBox("Are you taking lunch today?", vbQuestion + vbYesNo, "Lunch Time Entry")
            
            If (resYesNo = vbNo) Then
                'Not taking lunch for the day. Enter 0 to prevent checking on future updates
                [MealTimeCell].Value = LunchMinutes
            End If
        End If
    End If
    
    'Update end time
    [EndTimeCell].Value = EndTime
    
    'Grab total hours for evaluation
    stTotalHours = [TotalHoursCell].Text
    
    'Convert to decimal time
    stTotalHoursSplit = Split(stTotalHours, ":") 'Index 0 = hours, index 1 = minutes
    TotalMinutes = (stTotalHoursSplit(0) * 60)
    TotalMinutes = (TotalMinutes + stTotalHoursSplit(1))
    TotalMinutes = (TotalMinutes - LunchMinutes)
    
    WorkingHours = (TotalMinutes / 60)
    [HoursWorkedCell].Value = WorkingHours
    
    'Determine delta between last update and now
    LastJobRow = GetLastDataRow("Current Week", 3)
    Set JobHoursRange = Sheets("Current Week").Range(Cells([FirstJobCell].Row, ColumnNumber), Cells(LastJobRow, ColumnNumber))
    PreviousWorkHours = WorksheetFunction.Sum(JobHoursRange)
    PreviousWorkHours = (PreviousWorkHours - [HourEntryCell].Value)
    [HourEntryCell].Value = WorksheetFunction.Round((WorkingHours - PreviousWorkHours), 2)
End Sub

Sub ValidatePreviousDays(CurrentColumn As Integer)
    'Declare integers
    Dim i As Integer
    Dim j As Integer

    'Declare ranges
    Dim CheckCell As Range
    
    With Sheets("Current Week")
        Set CheckCell = .Cells(3, 4)
        
        For i = CheckCell.Column To (CurrentColumn - 1)
            For j = CheckCell.Row To 5
                If (.Cells(j, i).Value = stEmpty) Then
                    FixPreviousDays j, i
                End If
            Next j
        Next i
    End With
End Sub

Sub WriteLineToTxtFile(FilePath As String, ByVal Message As String)
    'Declare integers
    Dim fileNum As Integer
    
    fileNum = FreeFile
    
    Open FilePath For Append Access Write As #fileNum
    Print #fileNum, Message
    Close #fileNum
End Sub
