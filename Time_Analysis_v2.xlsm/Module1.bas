Attribute VB_Name = "Module1"



Public Sub RefreshMondayData()
Dim outputRange As Range
    Set outputRange = Application.Run("VBAUtils.xlsm!HTTPDownloadFile", _
                "http://172.22.237.138/datafiles/Monday/6666786972.txt", _
                ActiveWorkbook, _
                "", "", 0, "start-of-day", "Monday Data", False, 1)
                
      Application.Run "VBAUtils.xlsm!sortRange", outputRange.Worksheet, outputRange, 6
                    
End Sub




Public Function ParseDuration(durationString As String) As String
    If InStr(durationString, "minutes") > 0 Then
        ParseDuration = Split(durationString, " ")(0)
    ElseIf InStr(durationString, "hour") > 0 Then
        ParseDuration = CInt(Split(durationString, " ")(0)) * 60
    ElseIf InStr(durationString, "days") > 0 Then
        ParseDuration = 0
    End If
End Function

Public Function ParseLocation(locationString As String) As String
    Dim markerInt As Integer
    Dim stringParts As Variant
    Dim newLocationString As String
    ParseLocation = "notset"
    
    ' take out zoom strings
    newLocationString = ""

    If (InStr(locationString, "id=")) Then
    'If (InStr(locationString, "Zoom")) Then
        'stringParts = Split(locationString, " ")
        
        stringParts = Split(locationString, "id=")
        
        For i = UBound(stringParts) To UBound(stringParts)
    
            newLocationString = newLocationString & stringParts(i)
        Next i
        
        Debug.Print newLocationString
    End If
        
    locationString = newLocationString

    ParseLocation = newLocationString

    
    
    'markerInt = InStr(locationString, "=")
    'If markerInt > 0 Then
    '    ParseLocation = Right(locationString, Len(locationString) - markerInt)
    'End If
End Function


Public Sub AddRecord(ByRef dataArray As Variant, recordNum As Integer, subject As String, startDateString As String, duration As String, _
        Category As String, recurFlag As String, startDate As Date, activityString As String, startHour As Integer, startMinute As Integer, _
        endHour As Integer, endMinute As Integer, locationString As String, statusString As String, monthYearString As String, _
        recipientString As String, organizerString As String)

    Dim durationHours As Integer
    Dim durationMinutes As Integer
    
    durationMinutes = CInt(duration)
    durationHours = durationMinutes / 60
    
    If subject = "DELETE" Or Category = "Delete" Then
        durationMinutes = duration * -1
        durationHours = durationHours * -1
    ElseIf Category = "" Then
        Category = "NOTSET"
    End If
    
    dataArray(recordNum, 1) = subject
    dataArray(recordNum, 2) = startDateString
    dataArray(recordNum, 3) = durationMinutes
    dataArray(recordNum, 4) = Category
    dataArray(recordNum, 5) = recurFlag
    dataArray(recordNum, 6) = Application.WeekNum(startDate)
    dataArray(recordNum, 7) = durationHours
    dataArray(recordNum, 8) = activityString
    dataArray(recordNum, 9) = Application.Weekday(startDate)
    dataArray(recordNum, 10) = startHour
    dataArray(recordNum, 11) = startMinute
    dataArray(recordNum, 12) = endHour
    dataArray(recordNum, 13) = endMinute
    dataArray(recordNum, 14) = locationString
    dataArray(recordNum, 15) = statusString
    dataArray(recordNum, 16) = monthYearString
    dataArray(recordNum, 17) = recipientString
    dataArray(recordNum, 18) = organizerString
    

End Sub

Public Sub parse()


    Dim tmpWorkbook As Workbook
    Dim dataWorksheet As Worksheet
    Dim resultWorksheet As Worksheet
    Dim dataRange As Range
    Dim recurrentRange As Range
    Dim categoriesRange As Range
    Dim locationRange As Range
    Dim durationRange As Range
    Dim subjectRange As Range
    Dim startRange As Range
    Dim endRange As Range
    Dim resultRange As Range
    Dim startHourRange As Range
    Dim endHourRange As Range
    Dim startMinuteRange As Range
    Dim endMinuteRange As Range
    Dim startYearRange As Range
    Dim statusRange As Range
    Dim recipientRange As Range
    Dim organizerRange As Range
    
    Dim startDate As Date
    Dim endDate As Date
    Dim nextDate As Date
    Dim todayDate As Date
    
    Dim dateString As String, monthYearString As String
    Dim subjectString As String
    
    Dim dataArray() As Variant
    
    Dim i As Double
    Dim j As Integer
    Dim k As Integer
    
    Dim diffWeeks As Long
    
    ReDim dataArray(1 To 3000, 1 To 18)
  
    Set tmpWorkbook = ActiveWorkbook
    'Workbooks("MYTIME v010.xlsm")
    tmpWorkbook.Activate
    
    Set resultWorksheet = tmpWorkbook.Sheets("ResultCalendar")
    Set dataWorksheet = tmpWorkbook.Sheets("RawData")
    dataWorksheet.Activate
    
    Set dataRange = dataWorksheet.Range("DATA")
    Set recurrentRange = dataWorksheet.Range("RECURRENCE")
    Set locationRange = dataWorksheet.Range("LOCATION")
    Set subjectRange = dataWorksheet.Range("SUBJECT")
    
    Set startRange = dataWorksheet.Range("START")
    Set endRange = dataWorksheet.Range("END")
    Set categoriesRange = dataWorksheet.Range("CATEGORIES")
    Set durationRange = dataWorksheet.Range("DURATION")
    
    Set startHourRange = dataWorksheet.Range("STARTHOUR")
    Set endHourRange = dataWorksheet.Range("STARTMINUTE")
    Set startMinuteRange = dataWorksheet.Range("ENDHOUR")
    Set endMinuteRange = dataWorksheet.Range("ENDMINUTE")
    Set startYearRange = dataWorksheet.Range("STARTYEAR")
    Set statusRange = dataWorksheet.Range("STATUS")

    Set recipientRange = dataWorksheet.Range("RECIPIENT")
    Set organizerRange = dataWorksheet.Range("ORGANIZER")
    
    resultWorksheet.Activate
    Set resultRange = resultWorksheet.Range(Cells(2, 1), Cells(3001, 18))
    resultRange.Clear
    
    k = 1
    For i = 2 To dataRange.Rows.Count
        
        'On Error GoTo err
        
        subjectString = subjectRange.Rows(i)

        If startYearRange.Rows(i) <> Year(Now()) Then
            'Debug.Print "ignoring " & subjectString & " as doesnt start in " & CStr(Year(Now()))
            GoTo donothing
        End If
        
        If statusRange.Rows(i) = "Cancelled" Then
            'Debug.Print "ignoring " & subjectString & " as is canceled "
            GoTo donothing
        End If
        
        dateString = endRange.Rows(i)
        If dateString = "" Then
            GoTo done
            
        End If
        

        'dateString = Right(dateString, Len(dateString) - 4)
        endDate = CDate(dateString)
        
        dateString = startRange.Rows(i)
        'dateString = Right(dateString, Len(dateString) - 4)
        startDate = CDate(dateString)
        
        'monthYearString = Year(endDate) & Month(endDate)
        
        diffWeeks = DateDiff("w", startDate, endDate)
        If diffWeeks > 104 Then
            diffWeeks = 104
        End If
        
            
        If subjectString = "Capsule cleanup / sales stats" Then
            Debug.Print
        End If
            
            
        If recurrentRange.Rows(i) = "Weekly" Then
        
            nextDate = startDate
            'For j = Application.WeekNum(startDate) - 1 To Application.WeekNum(endDate) + 1
            For j = Application.WeekNum(startDate) - 1 To Application.WeekNum(startDate) + diffWeeks
                
                    
                AddRecord dataArray, k, subjectString, Format(nextDate, "yyyy/mm/dd"), durationRange.Rows(i), _
                    categoriesRange.Rows(i), True, nextDate, ParseLocation(locationRange.Rows(i)), _
                    startHourRange.Rows(i), startMinuteRange.Rows(i), endHourRange.Rows(i), endMinuteRange.Rows(i), _
                    locationRange.Rows(i), statusRange.Rows(i), Month(nextDate), recipientRange.Rows(i), organizerRange.Rows(i)
                nextDate = DateAdd("ww", 1, nextDate)
                k = k + 1
                
            Next j
        ElseIf recurrentRange.Rows(i) = "Daily" Then
            nextDate = startDate
            'For j = (7 * Application.WeekNum(startDate)) - 1 To (7 * Application.WeekNum(endDate)) + 1
            For j = (7 * Application.WeekNum(startDate)) - 1 To (7 * (Application.WeekNum(startDate) + diffWeeks))
                nextDate = DateAdd("d", 1, nextDate)
                
                If Weekday(nextDate) > 1 And Weekday(nextDate) < 7 Then
                    AddRecord dataArray, k, subjectString, Format(nextDate, "yyyy/mm/dd"), durationRange.Rows(i), _
                        categoriesRange.Rows(i), True, nextDate, ParseLocation(locationRange.Rows(i)), _
                        startHourRange.Rows(i), startMinuteRange.Rows(i), endHourRange.Rows(i), endMinuteRange.Rows(i), _
                        locationRange.Rows(i), statusRange.Rows(i), Month(nextDate), recipientRange.Rows(i), organizerRange.Rows(i)
                        
                    k = k + 1
                End If
            Next j
        Else
                AddRecord dataArray, k, subjectString, Format(startDate, "yyyy/mm/dd"), durationRange.Rows(i), _
                    categoriesRange.Rows(i), False, startDate, ParseLocation(locationRange.Rows(i)), _
                    startHourRange.Rows(i), startMinuteRange.Rows(i), endHourRange.Rows(i), endMinuteRange.Rows(i), _
                    locationRange.Rows(i), statusRange.Rows(i), Month(startDate), recipientRange.Rows(i), organizerRange.Rows(i)
                k = k + 1
            End If
donothing:
    Next


done:
    resultWorksheet.Activate
    resultRange = dataArray
    
    'Exit Sub
    
    ActiveWorkbook.Worksheets("ResultCalendar").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("ResultCalendar").Sort.SortFields.Add2 Key:=Range( _
        "J2:J1658"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("ResultCalendar").Sort.SortFields.Add2 Key:=Range( _
        "L2:L1658"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("ResultCalendar").Sort
        .SetRange Range("A1:R2658")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Exit Sub
err:
        Debug.Print "error:" & subjectString

End Sub





Function GetRecurString(myrecurrence As OlRecurrenceType) As String

    If myrecurrence = olRecursDaily Then
        GetRecurString = "Daily"
    ElseIf myrecurrence = olRecursWeekly Then
        GetRecurString = "Weekly"
    Else
        Debug.Print
    End If

End Function

Sub UpdateOutputRange(ByRef myArr As Variant, startRow As Integer, lastEndRow As Integer, NextRow As Long, ws As Worksheet)
Dim outputRange As Range

    
    startRow = startRow + lastEndRow
    'lastEndRow = lastEndRow + NextRow - 1
    lastEndRow = lastEndRow + NextRow + startRow - 1
    ws.Activate
    Set outputRange = ws.Range(Cells(startRow, 1), Cells(lastEndRow, 16))
    
    outputRange.ClearContents
    outputRange.Value = WorksheetFunction.Transpose(myArr)
    
    Set outputRange = Nothing
End Sub
Public Sub ListAppointments()

    Application.ScreenUpdating = False

    Const olFolderCalendar As Byte = 9
    Const SchemaPropTag As String = "https://schemas.microsoft.com/mapi/proptag/"

    Dim olApp           As Outlook.Application: Set olApp = CreateObject("Outlook.Application")
    Dim olNS            As Outlook.Namespace: Set olNS = olApp.GetNamespace("MAPI")
    Dim olFolder        As Outlook.Folder
    Dim olColItems      As Outlook.Items, olColRestrictedItems As Outlook.Items
    Dim olApt           As Outlook.AppointmentItem
    Dim objOwner        As Outlook.Recipient

    Dim myRecurrPatt    As Object
    Dim myRecurrType    As Outlook.OlRecurrenceType

    Dim olMail          As Outlook.MailItem
    Dim oPA             As PropertyAccessor
    
    Dim NextRow As Long, numItem As Long, maxItemCount As Long, itemCount As Long
    Dim userCell As Range
    
    Dim ws              As Worksheet: Set ws = ThisWorkbook.Sheets("RawData")
    Dim wsResults       As Worksheet: Set wsResults = ThisWorkbook.Sheets("ResultStats")
    Dim wsReference     As Worksheet: Set wsReference = ThisWorkbook.Sheets("Reference")
    
    Dim outputRange             As Range, resultRange As Range
    Dim runDateRange            As Range: Set runDateRange = ws.Range("RUNDATE")
    Dim mailboxOwnerRange       As Range: Set mailboxOwnerRange = ws.Range("MAILBOX_OWNER")
    Dim categoryDefnRange       As Range: Set categoryDefnRange = wsReference.Range("CATEGORY_DEFN")
    Dim queryStartDateRange     As Range: Set queryStartDateRange = ws.Range("QUERYSTARTDATE")
    Dim queryEndDateRange       As Range: Set queryEndDateRange = ws.Range("QUERYENDDATE")
    Dim queryNumResultsRange    As Range: Set queryNumResultsRange = ws.Range("QUERYNUMRESULTS")
        
    Dim aptDate     As Date, datStartUTC As Date, datEndUTC As Date
    Dim i           As Integer, startRow As Integer, lastEndRow As Integer
     
    startRow = 3
    lastEndRow = 0
    maxItemCount = 10000000
    
    Set objOwner = olNS.CreateRecipient(mailboxOwnerRange.Value)
    objOwner.Resolve

    If objOwner.Resolved Then
        Set olFolder = olNS.GetSharedDefaultFolder(objOwner, olFolderCalendar)
    End If

    Set oMail = olApp.CreateItem(olMailItem)
    Set oPA = oMail.PropertyAccessor
    
    'datStartUTC = oPA.LocalTimeToUTC(DateAdd("d", -10, queryStartDateRange.Value))
    'datEndUTC = oPA.LocalTimeToUTC(DateAdd("d", 1, queryEndDateRange.Value))
    datStartUTC = oPA.LocalTimeToUTC(queryStartDateRange.Value)
    datEndUTC = oPA.LocalTimeToUTC(queryEndDateRange.Value)

    'Ensure there at least 1 item to continue
    If olFolder.Items.Count = 0 Then Exit Sub

    'This filter uses https://schemas.microsoft.com/mapi/proptag
    strFilter = AddQuotes("urn:schemas:calendar:dtend") _
    & " > '" & datStartUTC & "' AND " _
    & AddQuotes("urn:schemas:calendar:dtend") _
    & " < '" & datEndUTC & "'"
    
    
    Debug.Print strFilter
    
    'Count of items in Inbox
    Set olColItems = olFolder.Items
    Debug.Print (olColItems.Count)

    'This call succeeds with @SQL prefix
    Set olColRestrictedItems = olColItems.Restrict("@SQL=" & strFilter)
    'Set olColRestrictedItems = olColItems

    'Create an array large enough to hold all records
    Dim myArr() As Variant: ReDim myArr(0 To 17, 0 To olColRestrictedItems.Count - 1)

    On Error Resume Next
    
    olColRestrictedItems.IncludeRecurrences = True

    NextRow = 0
    itemCount = 1
    
    'clear current result set =A3:R44
    Set resultRange = ws.Range(Cells(3, 1), Cells(queryNumResultsRange.Value, 18))
    resultRange.ClearContents
        
    'write num of rows returned to the sheet
    queryNumResultsRange.Value = olColRestrictedItems.Count
    
    ' Draw column headers
    ws.Range("A2:R2").Value2 = Array("Subject", "Categories", "Duration", "Recurrence", "Location", "Start", "End", "Start Hour", _
            "Start Minute", "End Hour", "End Minute", "Start Year", "Status", "Recipients", "Organizer", "Item ID", "New Category", "New Monday ID")
            
    Application.StatusBar = CStr(itemCount) & "/" & CStr(olColRestrictedItems.Count)
    For Each olApt In olColRestrictedItems
        If NextRow > maxItemCount Then
            UpdateOutputRange myArr, startRow, lastEndRow, NextRow, ws
            GoTo cleanExit
        End If
    
        aptDate = olApt.Start
    
        If olApt.Categories <> "IGNORE" Then
            myArr(0, NextRow) = olApt.subject
            myArr(1, NextRow) = olApt.Categories
            myArr(2, NextRow) = olApt.duration
            myArr(3, NextRow) = "(None)"
            myArr(4, NextRow) = olApt.Location
            myArr(5, NextRow) = olApt.Start
            myArr(6, NextRow) = olApt.End
            myArr(7, NextRow) = Hour(olApt.Start)
            myArr(8, NextRow) = Minute(olApt.Start)
            myArr(9, NextRow) = Hour(olApt.End)
            myArr(10, NextRow) = Minute(olApt.End)
            myArr(11, NextRow) = Year(olApt.Start)
            
            For i = 1 To olApt.Recipients.Count
                If i = 1 Then
                    myArr(13, NextRow) = olApt.Recipients.Item(i).Name
                Else
                    myArr(13, NextRow) = myArr(13, NextRow) & "," & olApt.Recipients.Item(i).Name
                End If
            Next i
            
            myArr(14, NextRow) = olApt.Organizer
            myArr(15, NextRow) = olApt.EntryID
                            
            Set myRecurrPatt = olApt.GetRecurrencePattern
            
            If olApt.MeetingStatus = 5 Or olApt.MeetingStatus = 7 Then
                myArr(12, NextRow) = "Cancelled"
            End If
            
            If olApt.subject = "Competitor Analysis" Then
                Set myRecurrPatt = olApt.GetRecurrencePattern
            End If
            
            myArr(6, NextRow) = olApt.End
            
            If olApt.IsRecurring Then
            
                Set myRecurrPatt = olApt.GetRecurrencePattern
                myArr(5, NextRow) = myRecurrPatt.PatternStartDate
                myArr(6, NextRow) = myRecurrPatt.PatternEndDate
    
                If InStr(olApt.subject, "Daily") > 0 Then
                    myRecurrType = olRecursDaily
                Else
                    myRecurrType = myRecurrPatt.RecurrenceType
                End If
    
                myArr(3, NextRow) = GetRecurString(myRecurrType)
    
            Else
                'Debug.Print
            End If
            
            NextRow = NextRow + 1
        End If
        
        Application.StatusBar = CStr(itemCount) & "/" & CStr(olColRestrictedItems.Count)
        itemCount = itemCount + 1
    Next
    
    ReDim Preserve myArr(0 To 15, 0 To NextRow - 1)
    UpdateOutputRange myArr, startRow, lastEndRow, NextRow, ws
        
    On Error GoTo 0
        
cleanExit:

    Set resultRange = ws.Range(Cells(startRow - 1, 1), Cells(lastEndRow, 18))
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add2 Key:=resultRange.Columns(6), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ws.Sort
        .SetRange resultRange
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
        
    Set olApp = Nothing
    Set olNS = Nothing
    Set ws = Nothing
    Set wsResults = Nothing
    Set wsReference = Nothing
    Set objOwner = Nothing
    Set olColItems = Nothing
    Set myRecurrPatt = Nothing
    Set outputRange = Nothing
    Erase myArr

    Application.ScreenUpdating = True

    Exit Sub

ErrHand:
    'Add error handler
    Resume cleanExit
End Sub

Public Function AddQuotes(ByVal SchemaName As String) As String
    On Error Resume Next
    AddQuotes = Chr(34) & SchemaName & Chr(34)
End Function
