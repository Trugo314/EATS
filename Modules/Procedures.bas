Attribute VB_Name = "Procedures"
Option Explicit
'Declare booleans
Dim bUpdatingPreviousData As Boolean

'Declare integers
Dim intCurrentColumnNumber As Integer 'Integer that will always hold the column value of the current day
Dim intJobCount As Integer
Dim intLastJobRow As Integer
Dim intWeekNum As Integer

'Declare ranges
Dim LastUpdateCell As Range
Dim WeekNumberCell As Range

'Declare strings
Const stLogFileName As String = "programlog.txt"
Const stSupportFilesDir As String = "Spreadsheets\EATS\Support Files\"
Dim stEndTime As String
Dim stDate As String
Dim stDocumentsDir As String
Dim stLastUpdateDate As String
Dim stLogFilePath As String
Dim stOptionEndTime As String
Dim stStartTime As String
Dim stToday As String
Dim stUpdateTime As String

'****************************
'****************************
'Sheet Sequencer Section
'****************************
'****************************

Sub SheetSequencer()
    ' 1.)Get info for today
    ' 2.)Check previous days for valid data
    ' 3.)Ask to update previous day
    ' 4.)Shift data if necessary
    ' 5.)Update today
    
    InitDay
    ValidatePreviousDays
    UpdateSheet
End Sub

Sub InitDay()
    'Description: Initializes all data to be used for the current day.
    'Also needed to reference previous days
    
    stDocumentsDir = GetDocumentsDir
    stLogFilePath = stDocumentsDir & stSupportFilesDir & stLogFileName
    
    'Get information for the current day
    stToday = Format(Date, "dddd")
    stDate = Format(Date, "yyyy/mm/dd")
    stEndTime = Format(Time, "hh:mm") 'This will be overridden if the clock out time option is not empty
    stUpdateTime = Format(Time, "hh:mm")
    intWeekNum = WorksheetFunction.IsoWeekNum(Now)
    
    intCurrentColumnNumber = ConvertDayToColumn(stToday)
    intLastJobRow = GetLastDataRow("Current Week", 3)
    
    With Sheets("Current Week")
        Set LastUpdateCell = .Cells(8, 2)
        Set WeekNumberCell = .Cells(5, 2)
    End With
End Sub

'TODO - This needs reworked
Sub ValidatePreviousDays()
    'Description: Checks all previous days in the week to ensure data is present
    
    'Declare booleans
    Dim bStartTimePass As Boolean
    Dim bMealDurationPass As Boolean
    Dim bEndTimePass As Boolean
    Dim bHoursPass As Boolean
    
    'Declare integers
    'Dim intDaysFixed As Integer
    Dim intWeeksToShift As Integer
    
    'Declare ranges
    Dim CheckCells As Range
    
    'Declare strings
    Dim stInvalidDay As String
    
    'intDaysFixed = 0
    'TODO - New year will break this
    If (intWeekNum <> [WeekNumberCell].Value) Then
        intWeeksToShift = intWeekNum - [WeekNumberCell].Value
        ShiftData intWeeksToShift
        Exit Sub 'Temporary fix for checking after shifting?
    End If
    
    stLastUpdateDate = Left([LastUpdateCell].Text, 10)
    
    If ((stLastUpdateDate <> stDate) And (stToday <> "Monday")) Then
        For i = 4 To intCurrentColumnNumber - 1
            Set CheckCells = Sheets("Current Week").Cells(3, i)
        
            bStartTimePass = IIf([CheckCells].Text <> "", True, False)
            
            Set CheckCells = CheckCells.Offset(1, 0)
            bMealDurationPass = IIf([CheckCells].Text <> "", True, False)
            
            Set CheckCells = CheckCells.Offset(1, 0)
            bEndTimePass = IIf([CheckCells].Text <> "", True, False)
            
            Set CheckCells = Sheets("Current Week").Range(Cells(9, i), Cells(intLastJobRow, i))
            bHoursPass = IIf(WorksheetFunction.Sum(CheckCells) <> 0, True, False)
            
            If ((bStartTimePass = False) Or (bMealDurationPass = False) Or (bEndTimePass = False) Or (bHoursPass = False)) Then
                bUpdatingPreviousData = True
                stInvalidDay = ConvertColumnToDay(i)
                
                resOKOnly = MsgBox(stInvalidDay & " is missing data!" & vbCrLf & _
                    "EATS will update this data before continuing with today!", vbExclamation + vbOKOnly, "Previous Data Missing")
                    
                UpdateSheet i
                'Incr intDaysFixed
            End If
        Next i
        
        bUpdatingPreviousData = False
        
        resYesNo = MsgBox("EATS has processed all previous entries for the week. Do you need to update your end time from yesterday?", vbQuestion + vbYesNo, "Data Validation")
        
        If (resYesNo = vbYes) Then
            UpdateSheet intCurrentColumnNumber - 1
        End If
    End If
End Sub

Sub UpdateSheet(Optional ColumnNumber As Integer = -1)
    'Description: Updates hours for the current day
    
    'Declare doubles
    Dim dblPreviousWorkTime As Double
    Dim dblTotalHours As Double
    
    'Declare integers
    Dim intJobIndex As Integer
    Dim intJobRow As Integer
    Dim intLunchMinutes As Integer
    Dim intPresentHours As Integer
    Dim intPresentMinutes As Integer
    Dim intTotalMinutes As Integer
    
    'Declare ranges
    Dim EndTimeCell As Range 'Cell containing end time
    Dim HoursPresentCell As Range 'Cell containing total hours present that day
    Dim HoursWorkedTodayCell As Range 'Cell containing total hours worked that day
    Dim JobHoursCells As Range 'Cells containing individual hours for each job
    Dim JobHoursUpdateCell As Range 'Cell that is receiving new data
    Dim MealDurationCell As Range 'Cell containing the meal duration
    Dim StartTimeCell As Range 'Cell containing start time
    
    'Declare strings
    Dim stTimePresent() As String
    
    'Set cell references
    With Sheets("Current Week")
        If (ColumnNumber = -1) Then
            Set StartTimeCell = .Cells(3, intCurrentColumnNumber)
            Set MealDurationCell = .Cells(4, intCurrentColumnNumber)
            Set EndTimeCell = .Cells(5, intCurrentColumnNumber)
            Set HoursPresentCell = .Cells(6, intCurrentColumnNumber)
            Set HoursWorkedTodayCell = .Cells(7, intCurrentColumnNumber)
            Set JobHoursCells = .Range(Cells(9, intCurrentColumnNumber), Cells(intLastJobRow, intCurrentColumnNumber))
        Else
            Set StartTimeCell = .Cells(3, ColumnNumber)
            Set MealDurationCell = .Cells(4, ColumnNumber)
            Set EndTimeCell = .Cells(5, ColumnNumber)
            Set HoursPresentCell = .Cells(6, ColumnNumber)
            Set HoursWorkedTodayCell = .Cells(7, ColumnNumber)
            Set JobHoursCells = .Range(Cells(9, ColumnNumber), Cells(intLastJobRow, ColumnNumber))
        End If
    End With
    
    'Update start time cell
    If ([StartTimeCell].Value = "") Then
        If (ColumnNumber = -1) Then
            stStartTime = UITimeEntry("Start Time", "What time did you clock in today?")
        Else
            stStartTime = UITimeEntry("Start Time", "What time did you clock in on " & ConvertColumnToDay(ColumnNumber) & "?")
        End If
        
        'TODO
        'If (work spans two dates) Then
        '    Enter time and date
        'Else
        '    Enter time only
        'End If
        
        [StartTimeCell].Value = stStartTime
    End If
    
    'Update meal duration cell
    If ([MealDurationCell].Value <> "") Then
        intLunchMinutes = ([MealDurationCell].Value * 60)
    ElseIf ((TimeValue(stUpdateTime) > TimeValue("12:00")) Or (ColumnNumber <> -1)) Then
        If (ColumnNumber = -1) Then
            resYesNo = MsgBox("Would you like to enter lunch time?", vbQuestion + vbYesNo, "Lunch Time Entry")
        Else
            resYesNo = MsgBox("Would you like to enter lunch time for " & ConvertColumnToDay(ColumnNumber) & "?", vbQuestion + vbYesNo, "Lunch Time Entry")
        End If
        
        If (resYesNo = vbYes) Then
            intLunchMinutes = UINumEntry(0, 60, "Lunch Time Entry", "Enter the time taken (in minutes) for lunch.", True, True, 30)
            [MealDurationCell].Value = (intLunchMinutes / 60)
        Else
            intLunchMinutes = 0
            
            If (ColumnNumber = -1) Then
                resYesNo = MsgBox("Are you taking lunch today?", vbQuestion + vbYesNo, "Lunch Time Entry")
            Else
                resYesNo = MsgBox("Did you take lunch on " & ConvertColumnToDay(ColumnNumber) & "?", vbQuestion + vbYesNo, "Lunch Time Entry")
            End If
            
            If (resYesNo = vbNo) Then
                [MealDurationCell].Value = intLunchMinutes
            End If
        End If
    End If
    
    'Update end time cell
    If (ColumnNumber <> -1) Then
        'stEndTime = UITimeEntry("End Time", "What time did you clock out on " & ConvertColumnToDay(ColumnNumber) & "?")
        [EndTimeCell].Value = UITimeEntry("End Time", "What time did you clock out on " & ConvertColumnToDay(ColumnNumber) & "?")
    Else
        [EndTimeCell].Value = stEndTime
    End If
    
    'Update total hours worked today cell
    stTimePresent = Split([HoursPresentCell].Text, ":")
    intPresentHours = CInt(stTimePresent(0))
    intPresentMinutes = CInt(stTimePresent(1))
    
    intTotalMinutes = intPresentHours * 60
    intTotalMinutes = intTotalMinutes + intPresentMinutes
    intTotalMinutes = intTotalMinutes - intLunchMinutes
    
    dblTotalHours = intTotalMinutes / 60
    [HoursWorkedTodayCell].Value = dblTotalHours
    
    'Update job hours cell(s)
    'TODO - Add option to split missing time across all jobs
    intJobCount = GetJobCount("Current Week")
    
    If (intJobCount = 0) Then
        resYesNo = MsgBox("No jobs detected, would you like to add one?", vbQuestion + vbYesNo, "Data Entry")
        
        If (resYesNo = vbYes) Then
            AddJob True
            intJobIndex = 1 'Not needed for logic, just for user information
            intJobRow = 9
        Else
            ErrWrite "At least one job must be present before continuing!"
            End
        End If
    ElseIf (intJobCount = 1) Then
        intJobIndex = 1 'Not needed for logic, just for user information
        intJobRow = 9
    Else
        intJobIndex = UINumEntry(1, intJobCount, "Job Index Entry", "Multiple jobs detected. Please enter the index to update.", True, True)
        intJobRow = 9 + (intJobIndex - 1)
    End If
    
    If (ColumnNumber = -1) Then
        Set JobHoursUpdateCell = Sheets("Current Week").Cells(intJobRow, intCurrentColumnNumber)
    Else
        Set JobHoursUpdateCell = Sheets("Current Week").Cells(intJobRow, ColumnNumber)
    End If
    
    dblPreviousWorkTime = WorksheetFunction.Sum(JobHoursCells)
    dblPreviousWorkTime = dblPreviousWorkTime - [JobHoursUpdateCell].Value
    [JobHoursUpdateCell].Value = WorksheetFunction.Round((dblTotalHours - dblPreviousWorkTime), 2)
    
    'Set sheet data
    [LastUpdateCell].Value = stDate + " " + stUpdateTime
    [WeekNumberCell].Value = intWeekNum
    
    'Inform user and ask to save
    resYesNo = MsgBox("Job index " & CStr(intJobIndex) & " updated!" & vbCrLf & "Would you like to save?", vbQuestion + vbYesNo, "Update Complete")
        
    If (resYesNo = vbYes) Then
        ActiveWorkbook.Save
    End If
End Sub

Sub ShiftData(WeeksToShift As Integer)
    If (WeeksToShift >= 3) Then
        'Remove all data
        ClearAllSheets
    ElseIf (WeeksToShift = 2) Then
        ClearDataFromSheet "2 Weeks Ago"
        ClearDataFromSheet "1 Week Ago"
        CopyDataToSheet "Current Week", "2 Weeks Ago"
    ElseIf (WeeksToShift = 1) Then
        ClearDataFromSheet "2 Weeks Ago"
        CopyDataToSheet "1 Week Ago", "2 Weeks Ago"
        CopyDataToSheet "Current Week", "1 Week Ago"
        ClearDataFromSheet "Current Week"
    Else
        
    End If
End Sub

Sub ClearAllSheets()
    ClearDataFromSheet "Current Week"
    ClearDataFromSheet "1 Week Ago"
    ClearDataFromSheet "2 Weeks Ago"
End Sub

Sub ClearDataFromSheet(SheetName As String)
    'Declare objects
    Dim DataLocation As Object
    
    'Declare strings
    Const ShiftRange1 As String = "D3:J5" 'Start time, meal duration, end time
    Const ShiftRange2 As String = "D7:J7" 'Hours worked
    Dim ShiftRange3 As String 'Jobs and hours
    
    ShiftRange3 = "C9:K" & CStr(intLastJobRow)

    With Sheets(SheetName)
        For Each DataLocation In .Range(ShiftRange1)
            If ([DataLocation].Value <> "") Then
                [DataLocation].Value = ""
            End If
        Next
        
        For Each DataLocation In .Range(ShiftRange2)
            If ([DataLocation].Value <> "") Then
                [DataLocation].Value = ""
            End If
        Next
        
        For Each DataLocation In .Range(ShiftRange3)
            If ([DataLocation].Value <> "") Then
                [DataLocation].Value = ""
            End If
        Next
        
        .Range(ShiftRange3).Style = "Normal"
    End With
End Sub

Sub CopyDataToSheet(FromSheetName As String, ToSheetName As String)
    'Declare objects
    Dim DataLocation As Object
    
    'Declare strings
    Const ShiftRange1 As String = "D3:J5" 'Start time, meal duration, end time
    Const ShiftRange2 As String = "D7:J7" 'Hours worked
    Dim ShiftRange3 As String 'Jobs and hours
    
    ShiftRange3 = "C9:K" & CStr(intLastJobRow)
    
    Sheets(FromSheetName).Range(ShiftRange1).Copy Destination:=Sheets(ToSheetName).Range(ShiftRange1)
    Sheets(FromSheetName).Range(ShiftRange2).Copy Destination:=Sheets(ToSheetName).Range(ShiftRange2)
    Sheets(FromSheetName).Range(ShiftRange3).Copy Destination:=Sheets(ToSheetName).Range(ShiftRange3)
End Sub


'****************************
'****************************
'Support Procedure Section
'****************************
'****************************

Sub AddJob(Optional InhibitUpdate As Boolean = False)
    'Description: Adds a job the sheet
    
    'Declare integers
    Dim lastJobRow As Integer
    
    'Declare ranges
    Dim cellToUpdate As Range 'Cell that will take the new job
    Dim DataLocation As Range 'Used in the for each loop
    Dim jobRow As Range 'Contains job name and hours for all days
    
    lastJobRow = GetLastDataRow("Current Week", 3)
    
    Set cellToUpdate = Sheets("Current Week").Cells(lastJobRow + 1, 3)
    Set jobRow = Sheets("Current Week").Range(Cells(lastJobRow + 1, 3), Cells(lastJobRow + 1, 11))
    [cellToUpdate].Value = UIAlphaEntry("Job Entry", "Enter a job number", "LXC-xxx")
    
    For Each DataLocation In jobRow
        With DataLocation
            .Style = "Good"
            .BorderAround LineStyle:=xlContinuous, Weight:=xlThin
            .NumberFormat = "0.00"
            If (.Column = 11) Then
                .Formula = "=SUM(D" & CStr(cellToUpdate.Row) & ":J" & CStr(cellToUpdate.Row) & ")"
            End If
        End With
    Next DataLocation
    
    CenterAlign "Current Week", ColumnRange:="C:K"
    
    If ((bUpdatingPreviousData = False) And (InhibitUpdate = False)) Then
        resYesNo = MsgBox([cellToUpdate].Value & " has been added to the job list!" & vbCrLf & _
            "Would you like to update?", vbQuestion + vbYesNo, "Job Addition")
            
        If (resYesNo = vbYes) Then
            UpdateSheet
        End If
    End If
End Sub

Sub CenterAlign(Sheet As String, Optional ColumnRange As String = "", Optional Cell As Object = vbNull)
    'Declare ranges
    Dim CellsToFormat As Range

    If (ColumnRange <> "") Then
        With Sheets(Sheet)
            Set CellsToFormat = .Range(ColumnRange)
            CellsToFormat.HorizontalAlignment = xlCenter
            CellsToFormat.Columns.AutoFit
        End With
    ElseIf (Cell <> vbNull) Then
        Cell.HorizontalAlignment = xlCenter
        Cell.AutoFit
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
    'Description: Decrements an integer by 1
    
    Value = Value - 1
End Sub

Sub ErrWrite(ByVal Message As String)
    'Description: Used to show a custom error message
    
    resOKOnly = MsgBox(Message, vbCritical + vbOKOnly, stError)
End Sub

Sub Incr(ByRef Value As Integer)
    'Description: Increments an integer by 1
    
    Value = Value + 1
End Sub

Sub WriteLineToTxtFile(FilePath As String, ByVal Message As String)
    'Declare integers
    Dim fileNum As Integer
    
    If (Dir(FilePath) = "") Then
        resYesNo = MsgBox("The specified text file could not be found." & vbCrLf & _
            "Do you want to continue by creating the file below?" & vbCrLf & _
            FilePath, vbExclamation + vbYesNo, "File Not Found")
            
        If (resYesNo = vbNo) Then
            End
        End If
    End If
    
    fileNum = FreeFile
    
    Open FilePath For Append Access Write As #fileNum
    Print #fileNum, Message
    Close #fileNum
End Sub

'****************************
'****************************
'Testing Section
'****************************
'****************************

Sub ReadXML()
    Dim XDoc As Object, root As Object
    Dim NodeObject As Object
    Dim FieldNode As Object
    
    Set XDoc = CreateObject("MSXML2.DOMDocument")
    XDoc.async = False: XDoc.validateOnParse = False
    XDoc.Load ("C:\Users\GrantMumaugh\Documents\Spreadsheets\EATS\Support Files\Options.xml")
    Set root = XDoc.DocumentElement
    
    For Each NodeObject In root.ChildNodes
        
    Next NodeObject
End Sub
