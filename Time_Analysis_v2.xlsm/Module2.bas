Attribute VB_Name = "Module2"



Sub CreateRanges()

    Application.Run "VBAUtils.xlsm!CreateRefNamedRanges", "Reference", _
        "Z6:AB11", "2:2", "RawData", 2
End Sub
Sub updateOutlookItemValues()

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
        newCategory = updateItemCategoryRange.Rows(updateItem.Row - 2).Value
        newLocation = updateItemLocationRange.Rows(updateItem.Row - 2).Value
        currentSubject = subjectRange.Rows(updateItem.Row - 2).Value
        Set currentItemCategory = currentItemCategoryRange.Rows(updateItem.Row - 2)
        
        
        Set currentItemLocation = currentItemLocationRange.Rows(updateItem.Row - 2)
        

        If newCategory <> "" Then
            If currentItemCategory <> newCategory Or currentItemLocation <> newLocation Then
                currentItemLocation.Select
                UpdateAppointmentCategory updateItem.Value, newCategory, newLocation, currentItemCategory & ":" & currentItemLocation
                currentItemCategory = newCategory
                currentItemLocation = newLocation
    
                categoryLookupIndex = WorksheetFunction.Match(newCategory, categoryLookupRange, 0)
            
                redValue = WorksheetFunction.Index(categoryDefnRange, categoryLookupIndex, 5)
                greenValue = WorksheetFunction.Index(categoryDefnRange, categoryLookupIndex, 6)
                blueValue = WorksheetFunction.Index(categoryDefnRange, categoryLookupIndex, 7)
                
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
    olApt.Location = "id:" & mondayTag
    olApt.Save

    
End Sub


Sub GetCategories()
Dim olApp       As Outlook.Application: Set olApp = CreateObject("Outlook.Application")
Dim olNS        As Outlook.Namespace: Set olNS = olApp.GetNamespace("MAPI")
Dim olApt       As Outlook.AppointmentItem
Dim olCategory As Outlook.Category
Dim categoryColor As OlCategoryColor

    
    
    For Each olCategory In olNS.Categories
        Debug.Print olCategory.Name
        categoryColor = olCategory.Color
        Debug.Print categoryColor
        Next olCategory

End Sub
