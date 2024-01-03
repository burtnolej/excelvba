Attribute VB_Name = "Module2"



Sub CreateRanges()

    Application.Run "VBAUtils.xlsm!CreateRefNamedRanges", "Reference", _
        "Z6:AB11", "2:2", "RawData", 2
End Sub


Sub TestAddMeeting()
Dim wsRawData As Worksheet: Set wsRawData = ThisWorkbook.Sheets("RawData")
Dim startDateRange As Range:  Set startDateRange = wsRawData.Range("ADD_START_DATE")
Dim startTimeRange As Range:  Set startTimeRange = wsRawData.Range("ADD_START_TIME")
Dim durationRange As Range:  Set durationRange = wsRawData.Range("ADD_DURATION")
Dim locationRange As Range:  Set locationRange = wsRawData.Range("ADD_LOCATION")
Dim subjectRange As Range:  Set subjectRange = wsRawData.Range("ADD_SUBJECT")
Dim categoryRange As Range:  Set categoryRange = wsRawData.Range("ADD_CATEGORY")

    AddMeetingExec startDateRange.value, startTimeRange.value, durationRange.value, "id=" & locationRange.value, categoryRange.value, subjectRange.value

End Sub

Sub TestAddMeetingFromSheet()



    AddMeetingExec "10/1/2023", "11:00:00", 90, "id=453345345345", "Tools", "this is it"
End Sub
Sub AddMeetingExec(startDay As String, startTime As String, duration As Long, location As String, category As String, subject As String)

Dim olApp       As Outlook.Application: Set olApp = CreateObject("Outlook.Application")
Dim olNS        As Outlook.Namespace: Set olNS = olApp.GetNamespace("MAPI")
Dim olApt       As Outlook.AppointmentItem
Dim olCategories As Variant
Dim statusMsg As String

Dim startDate As Date

    Set olApt = olApp.CreateItem(olAppointmentItem)
    olApt.subject = subject
    olApt.location = location
    olApt.Start = CDate(DateValue(startDay) & " " & TimeValue(startTime))
    olApt.duration = duration
    olApt.Categories = category

    olApt.Save

End Sub

Sub TestcolorCodeCategories()

    colorCodeCategories ActiveWorkbook.Sheets("RawData").Range("B3:B149")
    
End Sub
Sub AddColorCodingExec(Optional tmpRange As Range)
Dim myCell As Range
Dim categoryLookupRange As Range, categoryDefnRange As Range
Dim myArea As Range
Dim category As String
Dim categoryLookupIndex As Long, redValue As Long, greenValue As Long, blueValue As Long
Dim wsRef As Worksheet

    If tmpRange Is Nothing Then
        Set tmpRange = ActiveWorkbook.Sheets("RawData").Range("B3:B200,Q3:Q200")
    End If
    
    Application.Run "VBAUtils.xlsm!SetEventsOff"
    
    Set wsRef = ThisWorkbook.Sheets("Reference")

    Set categoryLookupRange = wsRef.Range("CATEGORY_LOOKUP")
    Set categoryDefnRange = wsRef.Range("CATEGORY_DEFN")

    For Each myCell In tmpRange.Cells
        Debug.Print myCell.Address
        category = myCell.value
        If category <> "" Then
            categoryLookupIndex = WorksheetFunction.Match(category, categoryLookupRange, 0)
            redValue = WorksheetFunction.index(categoryDefnRange, categoryLookupIndex, 5)
            greenValue = WorksheetFunction.index(categoryDefnRange, categoryLookupIndex, 6)
            blueValue = WorksheetFunction.index(categoryDefnRange, categoryLookupIndex, 7)
            myCell.Interior.Color = RGB(redValue, greenValue, blueValue)
        Else
            myCell.Interior.Color = RGB(255, 255, 255)
        End If
    Next myCell
    Application.Run "VBAUtils.xlsm!SetEventsOn"
    
End Sub
Sub UpdateItemsExec()

Dim wsRawData As Worksheet: Set wsRawData = ThisWorkbook.Sheets("RawData")
Dim wsRef As Worksheet: Set wsRef = ThisWorkbook.Sheets("Reference")

Dim updateItemLocationRange As Range: Set updateItemLocationRange = wsRawData.Range("RAWDATA_UPDATE_ITEM_LOCATION")
Dim updateItemIDRange As Range: Set updateItemIDRange = wsRawData.Range("RAWDATA_UPDATE_ITEM_ID")
Dim updateItemCategoryRange As Range:  Set updateItemCategoryRange = wsRawData.Range("RAWDATA_UPDATE_ITEM_CATEGORY")

Dim categoryLookupRange As Range:  Set categoryLookupRange = wsRef.Range("CATEGORY_LOOKUP")
Dim categoryDefnRange As Range:  Set categoryDefnRange = wsRef.Range("CATEGORY_DEFN")

Dim subjectRange As Range:  Set subjectRange = wsRawData.Range("RAWDATA_SUBJECT")

Dim currentItemCategoryRange As Range:  Set currentItemCategoryRange = wsRawData.Range("RAWDATA_CURRENT_ITEM_CATEGORY")
Dim currentItemLocationRange As Range: Set currentItemLocationRange = wsRawData.Range("RAWDATA_CURRENT_ITEM_LOCATION")
Dim currentItemLocation As Range, currentItemCategory As Range

Dim categoryLookupIndex As Long, redValue As Long, greenValue As Long, blueValue As Long

Dim statusMsg As String
Dim updateItem As Range
Dim newCategory As String, newLocation As String, currentSubject As String


    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    

    For Each updateItem In updateItemIDRange
        newCategory = updateItemCategoryRange.Rows(updateItem.Row - 2).value
        newLocation = updateItemLocationRange.Rows(updateItem.Row - 2).value
        currentSubject = subjectRange.Rows(updateItem.Row - 2).value
        Set currentItemCategory = currentItemCategoryRange.Rows(updateItem.Row - 2)
        
        
        Set currentItemLocation = currentItemLocationRange.Rows(updateItem.Row - 2)
        

        If newCategory <> "" Then
            If currentItemCategory <> newCategory Or currentItemLocation <> newLocation Then
                currentItemLocation.Select
                UpdateAppointmentCategory updateItem.value, newCategory, newLocation, currentItemCategory & ":" & currentItemLocation
                currentItemCategory = newCategory
                currentItemLocation = newLocation
    
                categoryLookupIndex = WorksheetFunction.Match(newCategory, categoryLookupRange, 0)
            
                redValue = WorksheetFunction.index(categoryDefnRange, categoryLookupIndex, 5)
                greenValue = WorksheetFunction.index(categoryDefnRange, categoryLookupIndex, 6)
                blueValue = WorksheetFunction.index(categoryDefnRange, categoryLookupIndex, 7)
                
                currentItemCategory.Interior.Color = RGB(redValue, greenValue, blueValue)
                updateItemCategoryRange.Rows(updateItem.Row - 2).Interior.Color = RGB(redValue, greenValue, blueValue)
    
            Else
                statusMsg = "nothing to change : " & currentSubject
                Application.StatusBar = statusMsg
            End If
        End If
        
    Next updateItem

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True



End Sub
Sub testUpdateAppointmentCategory()
Dim itemId As String

'itemId = "00000000A6C727E4353C314FBD93409EF639422707000D702909B30966478A168622D1D80E2400000000010D00000D702909B30966478A168622D1D80E2400031ED2FC4B0000"
itemId = "00000000A6C727E4353C314FBD93409EF639422707000D702909B30966478A168622D1D80E2400000000010D00000D702909B30966478A168622D1D80E240002534C93790000"

    UpdateAppointmentCategory itemId, "Client", "4234234234"
End Sub
Sub UpdateAppointmentCategory(itemId As String, newCategory As String, mondayTag As String, Optional oldValues As String = "")
Dim olApp       As Outlook.Application: Set olApp = CreateObject("Outlook.Application")
Dim olNS        As Outlook.Namespace: Set olNS = olApp.GetNamespace("MAPI")
Dim olApt       As Outlook.AppointmentItem
Dim olCategories As Variant
Dim statusMsg As String

    Set olApt = olNS.GetItemFromID(itemId)
    
    statusMsg = "updating: " & olApt.subject & ",category=" & newCategory & ",monday=" & mondayTag & "(" & oldValues & ")"
    Application.StatusBar = statusMsg
    Debug.Print statusMsg
    
    olApt.Categories = newCategory
    olApt.location = "id:" & mondayTag
    olApt.Save

    
End Sub


Sub GetCategories()
Dim olApp       As Outlook.Application: Set olApp = CreateObject("Outlook.Application")
Dim olNS        As Outlook.Namespace: Set olNS = olApp.GetNamespace("MAPI")
Dim olApt       As Outlook.AppointmentItem
Dim olCategory As Outlook.category
Dim categoryColor As OlCategoryColor

    
    
    For Each olCategory In olNS.Categories
        Debug.Print olCategory.Name
        categoryColor = olCategory.Color
        Debug.Print categoryColor
        Next olCategory

End Sub
