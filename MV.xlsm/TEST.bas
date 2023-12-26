Attribute VB_Name = "TEST"

Const STATUS_COMPLETED = "1"
Const STATUS_WORKING = "0"
Const STATUS_NOT_STARTED = "5"

Const ADMIN_BOARD_ID = "1140656959"
Const CONTENT_BOARD_ID = "2410623120"
Const CLIENTS_BOARD_ID = "2193345626"
Const MARKETING_BOARD_ID = "2259144314"

Const MARKETING_ITEM_ID = "2951551976"
Const ADMIN_ITEM_ID = "2951455425"
Const ADMIN_SUBITEM_ID = "2962223456"
Const CLIENTS_ITEM_ID = "2951573079"
Const CONTENT_ITEM_ID = "2951585128"

Const DDQ = """"

Dim boardIds() As Variant
Dim itemIds() As Variant




Public Function OpenFile(sPath As String, iRWFlag As Integer) As Object
Dim objFSO As Object
Dim oFile As Object

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set oFile = objFSO.OpenTextFile(sPath, 1)
    
    Set OpenFile = oFile
End Function

Sub TestGetRangeFromFile()
    RehydrateRangeFromFile "MV.xlsm", "Persist", "persistdata", "C:\Users\burtn\foo.csv"
End Sub

Sub TestPersistRangeToFile()
    PersistRangeToFile "MV.xlsm", "Persist", "persistdata", "C:\Users\burtn\foo.csv"
End Sub


Sub TestHTTPDownloadFile2()
Dim todayString As String, filterString As String

    todayString = Format(Now(), "YYYYMMDD")
    filterString = "nofilter"

    Application.Run "VBAUtils.xlsm!HTTPDownloadFile", _
                    "http://bumblebee/datafiles/1140656959.txt", _
                    "A", "REFERENCE", "start-of-day", "1140656959", True
    
    
End Sub


Sub TestHTTPDownloadFile(Optional newSheetName As String = "test")
Dim tmpWorkbook As Workbook
Dim tmpSheet As Worksheet
Dim tmpRange As Range
Dim fileLength As Long
Dim rowWidth As Long
Dim fileArray As Variant, lineArray As Variant

 
Set tmpWorkbook = ActiveWorkbook
Set tmpSheet = tmpWorkbook.Sheets.Add
tmpSheet.Name = newSheetName

Url = "http://bumblebee/datafiles/prod_data_clients_20230601.055501.txt"

Dim objHTTP As Object
Dim postData As String
Dim DDQ As String

DDQ = Chr(34)


Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
objHTTP.Open "GET", Url, False
objHTTP.setRequestHeader "Content-Type", "text/csv"
objHTTP.send


fileArray = Split(objHTTP.responseText, Chr(10))
fileLength = Len(fileArray)
For i = 1 To fileLength
    
    lineArray = Split(fileArray(i), "^")
    rowWidth = Len(lineArray)
    
    Set tmpRange = tmpSheet.Rows(i).Resize(rowWidth)
    tmpRange = lineArray
Next i


End Sub
Sub testCreateMondayFolderBatch()
Dim dataRange As Range, dataRow As Range
Dim itemId As String, itemContent As String, rs As String, rt As String, itemName As String

    Set dataRange = Selection
    For Each dataRow In dataRange.Rows
    
        itemContent = dataRow.Columns(2)
        itemName = dataRow.Columns(1)
        itemId = dataRow.Columns(3)
    
        Debug.Print "Processing : " & itemId & " " & itemName

        InitMondayFolder itemId, itemName, _
                        "C:\Users\burtn\Velox Financial Technology\Velox Shared Drive - Documents\General\Monday", _
                        ActiveWorkbook, _
                        itemContent, _
                        False
        Debug.Print
        Debug.Print
    Next dataRow
End Sub

Sub AddToMondayFiles()
Dim dataRange As Range, dataRow As Range
Dim itemId As String, itemLink As String, rs As String, rt As String, itemName As String, itemContent As String

    Set dataRange = Selection
    For Each dataRow In dataRange.Rows
        itemLink = dataRow.Columns(4)
        itemId = dataRow.Columns(3)
        itemContent = dataRow.Columns(2)
        itemName = dataRow.Columns(1)
        
        If itemLink <> "" Then
            AddToMondayFile itemId, _
                            "C:\Users\burtn\Velox Financial Technology\Velox Shared Drive - Documents\General\Monday", _
                            ActiveWorkbook, _
                            itemContent, _
                            itemLink, _
                            itemName, _
                            False
        End If
    Next dataRow

End Sub

Sub TestDeleteMondayItems()
Dim dataRange As Range, dataRow As Range
Dim itemId As String, itemLink As String, rs As String, rt As String, itemName As String, itemContent As String

    Set dataRange = Selection
    For Each dataRow In dataRange.Rows
        itemId = dataRow.Columns(1)
        

            DeleteMondayItem itemId, _
                            ActiveWorkbook, _
                            "C:\Users\burtn\Velox Financial Technology\Velox Shared Drive - Documents\General\Monday"
    Next dataRow
End Sub

Sub TestExportAllModules()

    ExportAllModules
    
End Function
Sub testAddMondayItemsBatch()
Dim dataRange As Range, dataRow As Range
Dim groupid As String, boardid As String, rs As String, rt As String, itemName As String
Dim status As String, owner As String, newItemName As String, newSubItemName As String, itemId As String
Dim newItemId As String, addedFlag As String, newItemUpdateMsg As String
Dim i As Integer

    Set dataRange = ActiveSheet.Range("NEWITEM_DATA")
    Set addedItemIdRange = ActiveSheet.Range("NEWITEM_ADDEDITEMID")

    'For Each dataRow In dataRange.Rows()
    For i = 1 To dataRange.Rows.Count
        groupid = ActiveSheet.Range("NEWITEM_GROUP_ID").Rows(i)
        boardid = ActiveSheet.Range("NEWITEM_BOARD_ID").Rows(i)
        itemName = ActiveSheet.Range("NEWITEM_ITEM_NAME").Rows(i)
        newItemName = ActiveSheet.Range("NEWITEM_NEWITEM_NAME").Rows(i)
        status = ActiveSheet.Range("NEWITEM_STATUS").Rows(i)
        owner = ActiveSheet.Range("NEWITEM_OWNER").Rows(i)
        itemId = ActiveSheet.Range("NEWITEM_ITEMID").Rows(i)
        newSubItemName = ActiveSheet.Range("NEWSUBITEM_NEWSUBITEM_NAME").Rows(i)
        newItemUpdateMsg = ActiveSheet.Range("NEWITEM_NEWITEM_UPDATE").Rows(i)
        
        addedFlag = addedItemIdRange.Rows(i)

        
        If addedFlag = "" Then
            If newItemName <> "" Then ' then its a new item
                If boardid <> "" Then
                    CreateMondayItem boardid, groupid, newItemName, status, owner, rs, rt
                    newItemId = getResponseItemid(rt, "create_item")
                    addedItemIdRange.Rows(i).value = newItemId

                    CreateMondaySubItem newItemId, newSubItemName, rs, rt
                    newItemId = getResponseItemid(rt, "create_subitem")
                    addedItemIdRange.Rows(i).value = newItemId
                    
                Else
                    Debug.Print "exiting from row #" & i & " as end of items to add"
                    Exit Sub
                End If
            ElseIf newSubItemName <> "" Then ' its a new sub item
                    CreateMondaySubItem itemId, newSubItemName, rs, rt
                    newItemId = getResponseItemid(rt, "create_subitem")
                    addedItemIdRange.Rows(i).value = newItemId
            End If
            
            ' then post the description field as an update
            PostUpdateMonday newItemId, newItemUpdateMsg, rs, rt
    
        Else
            Debug.Print "skipping row #" & i & " as already added"
        End If
        

        
    Next i
End Sub
Sub testDisplayGroups()
Dim tmpSheet As Worksheet

    Set tmpSheet = ActiveWorkbook.Sheets("Reference")
    tmpSheet.Activate
    SetEventsOff
    DisplayGroups tmpSheet, tmpSheet.Range("S2")

    SetEventsOn
    
End Sub

Sub testGetBoardColumns()
Dim rs As String, rt As String
Dim columnColl As Collection
Dim column As Variant
Dim userAcct As Dictionary

    Set columnColl = GetBoardColumns("1140656959", rs, rt)
    
    For Each column In columnColl
        Debug.Print column.item("title"),
        Debug.Print column.item("type")
    Next column
    
End Sub


Sub testMoveItemToBoard()
Dim columnMappings As String
Dim rs As String, rt As String
'3969481029
            
'2259144314  Marketing & Messaging   New Group   new_group28215
'1140656959  General Admin   TestGroup   new_group73661


columnMappings = GetColumnMapString("2259144314", "1140656959", "new_group73661", "3969481029")

MoveItemToBoard "2259144314", "1140656959", "3969481029", "new_group73661", columnMappings, rs, rt


End Sub

Function GetColumnMapString(oldBoardId As String, newBoardId As String, newGroupId As String, itemId As String) As String
Dim colName As Variant
Dim oldBoardRowNum As Long, newBoardRowNum As Long, colNum As Long
Dim columnMapBoardIdRange As Range, columnMapHeadersRange As Range, columnMapRange As Range
Dim Columns As Variant
Dim result As String, DDQ As String

DDQ = Chr(34)

Set columnMapRange = ActiveWorkbook.Sheets("REFERENCE").Range("COLUMN_MAP")
Set columnMapBoardIdRange = ActiveWorkbook.Sheets("REFERENCE").Range("COLUMN_MAP_BOARD_ID")
Set columnMapHeadersRange = ActiveWorkbook.Sheets("REFERENCE").Range("COLUMN_MAP_HEADERS")

Columns = Array("name", "status", "person", "subitem", "tags", "item_id", "created_by", "last_updated")

oldBoardRowNum = Application.Match(oldBoardId, columnMapBoardIdRange, 0)
newBoardRowNum = Application.Match(newBoardId, columnMapBoardIdRange, 0)
result = ""

For i = 0 To UBound(Columns)
    colName = Columns(i)
    colNum = Application.Match(colName, columnMapHeadersRange, 0)
    oldColName = columnMapRange.Cells(oldBoardRowNum, colNum).value
    newColName = columnMapRange.Cells(newBoardRowNum, colNum).value
    If i > 0 Then result = result + ","
    result = result & "{source:" & "\" & DDQ & oldColName & "\" & DDQ & ", target:" & "\" & DDQ & newColName & "\" & DDQ & "}"
    
Next i

GetColumnMapString = result
End Function

Sub TestMoveItem()
Dim rs As String, rt As String
Dim item As Collection, subitems As Collection, subitemsUpdates As Collection, itemUpdates As Collection
Dim itemId As String, targetBoardId As String, targetGroupId As String, newItemId As String, newSubItemId As String, origSubItemName As String
Dim origItemName As String, origItemStatus As Integer, origSubItemStatus As Integer
Dim columnValueString As Variant
Dim itemColumnValues As Collection, subItemColumnValues As Collection
Dim itemValue As Dictionary

    itemId = "4977366193"
    targetBoardId = "4977328922"
    targetGroupId = "topics"
    
    GetItemDetails itemId, item, subitems, subitemsUpdates, itemUpdates, rs, rt
    
    For i = 1 To item.Count
        origItemName = item(i)("name")
        Set itemColumnValues = item(i)("column_values")
        For j = 1 To itemColumnValues.Count
            column_value_title = itemColumnValues(j)("id")
            If Left(column_value_title, 6) = "status" Then
                columnValueString = itemColumnValues(j)("value")
                Set jsonObject = ParseJson(columnValueString)
                origItemStatus = jsonObject("index")
            ElseIf column_value_title = "status" Then
            End If
        Next j
        newItemId = CreateMondayItem(targetBoardId, targetGroupId, origItemName & "_copy", CStr(origItemStatus), rs, rt)
        
        Set itemUpdates = item(i)("updates")
        
        Set subitems = item(i)("subitems")
        For k = 1 To subitems.Count
            origSubItemName = subitems(k)("name")
            Set subItemColumnValues = subitems(k)("column_values")
            For l = 1 To subItemColumnValues.Count
                column_value_title = subItemColumnValues(l)("id")
                If Left(column_value_title, 6) = "status" Then
                    columnValueString = subItemColumnValues(l)("value")
                    Set jsonObject = ParseJson(columnValueString)
                    origSubItemStatus = jsonObject("index")
                ElseIf column_value_title = "status" Then
                End If
            Next l
            newSubItemId = CreateMondaySubItem(newItemId, origSubItemName & "_copy", CStr(origSubItemStatus), rs, rt)
            
            Set subItemUpdates = item(i)("updates")
        Next k
    Next i
    
    'newItemId = CreateMondayItem("4977328922", "topics", "new foo", "1", rs, rt)
    'newSubItemId = CreateMondaySubItem(newItemId, "new sub foo", "1", rs, rt)
    'newUpdateId = AddUpdate(newSubItemId, "this is an update", rs, rt)
    

    
    'need to add the ability to pass status, created, updated etc etc
    
    

    
    'For i = 1 To subitems.Count
    '    CreateMondaySubItem
    'Next i
    
    'For i = 1 To itemsUpdates.Count
    '    AddItemUpdates
    'Next i
    
    'For i = 1 To subitemsUpdates.Count
    '    AddSubItemUpdates
    'Next i
    
    'Delete
    
End Sub
Public Sub TestMondayAPI()
Dim i As Integer
Dim rs As String, rt As String
Dim jsonObject As Object
Dim dataDict As New Dictionary
Dim subItemBoardId As String, columnid As String


                    
    ReDim boardsIds(0 To 3)
    ReDim itemIds(0 To 3)
    
    AddTag "foofoof", rs, rt

    'TestGetTags rs, rt
    'TestGetUsers rs, rt


    UpdateTagsMonday "1140656959", "4699993069", "19676602,19045698", rs, rt
    
    UpdateOwnerMonday "2259144314", "4960032448", "Mike", rs, rt
    UpdateOwnerMonday "1140656959", "3969481029", "Chris", rs, rt
    'Debug.Print TestCreateMondaySubItem("3969481029", "foo", rs, rt)
    
    'Exit Sub

    'Admin Software& Services
    '{\"tag_ids\":[10165564,10166350]}"},
    
    '3969481029
    
    'UpdateItemAttributeMonday "1140656959", "3969481240", "name", "testnamesubitem2", rs, rt
    
    'UpdateItemAttributeMonday "1140656959", "3969481240", "people", "Alison Hood", rs, rt
    'UpdateItemAttributeMonday "1140656959", "3969481029", "person", "Alison Hood", rs, rt
    
    Set dataDict = GetSubitemColumns("1140656959", "3977248134", rs, rt)
    
    'columnid = CreateDropdownColumn("1140656959", "foobar3", "foo,bar", rs, rt)
    'SetDropdownColumnValues "1140656959", "3969481029", "dropdown6", "1 High, JB Nike", rs, rt
    'UpdateOwnerMonday "1140656959", "3969481029", "Alison Hood", rs, rt
    
    Exit Sub
    
    
    
    
    'TestCreateMondayItem "2410623120", "content", "test", rs, rt
    
    'TestGetGroupsForBoard CLIENTS_BOARD_ID, rs, rt
    'TestGetBoards rs, rt

    
    boardIds = Array(ADMIN_BOARD_ID, CONTENT_BOARD_ID, CLIENTS_BOARD_ID, MARKETING_BOARD_ID)
    itemIds = Array(ADMIN_ITEM_ID, CONTENT_ITEM_ID, CLIENTS_ITEM_ID, MARKETING_ITEM_ID)
    subitemIds = Array(ADMIN_SUBITEM_ID)
    
    ' items
    For i = 0 To UBound(boardIds)
        Debug.Print boardIds(i), itemIds(i),
        TestPostUpdate itemIds(i), rs, rt
        Debug.Print rs,
        Debug.Print rt
        TestUpdateStatus itemIds(i), boardIds(i), rs, rt
        Debug.Print rs,
        Debug.Print rt
    Next i
    
    ' subitems
    For i = 0 To UBound(subitemIds)
        subItemBoardId = TestGetBoardId(itemIds(i), rs, rt)
        Debug.Print rs,
        Debug.Print rt
        
        TestUpdateStatus itemIds(i), subItemBoardId, rs, rt
        Debug.Print rs,
        Debug.Print rt
    Next i

End Sub

Public Sub TestGetTags(ByRef rs As String, ByRef rt As String)

    
    Set tagColl = GetTags(rs, rt)
    For Each tag In tagColl
        Debug.Print tag.item("id"), tag.item("name")
        'Set userAcct = user("account")
        'Debug.Print userAcct.Item("name")
    Next tag
    
End Sub


Public Sub TestPostUpdate(itemId As Variant, ByRef rs As String, ByRef rt As String)
    PostUpdateMonday CStr(itemId), "1", rs, rt
   
End Sub

Public Sub TestUpdateStatus(itemId As Variant, boardid As Variant, ByRef rs As String, ByRef rt As String)
    UpdateStatusMonday CStr(boardid), CStr(itemId), "0", rs, rt
End Sub

Public Sub TestUpdateSubItemStatusMonday(itemId As Variant, boardid As Variant, ByRef rs As String, ByRef rt As String)
    UpdateSubItemStatusMonday CStr(boardid), CStr(itemId), "0", rs, rt
End Sub

Public Function TestGetBoardId(itemId As Variant, ByRef rs As String, ByRef rt As String) As Variant
    TestGetBoardId = GetBoardId(CStr(itemId), rs, rt)
End Function

Public Function TestGetBoards(ByRef rs As String, ByRef rt As String) As Variant

    Set TestGetBoards = GetBoards(rs, rt)
    
End Function

Public Function TestGetUsers(ByRef rs As String, ByRef rt As String) As Variant
Dim userColl As Collection
Dim User As Variant
Dim userAcct As Dictionary
    
    Set userColl = GetUsers(rs, rt)
    For Each User In userColl
        Debug.Print User.item("email")
        Debug.Print User.item("id")
        Debug.Print User.item("name")
        'Set userAcct = user("account")
        'Debug.Print userAcct.Item("name")
    Next User
    
End Function
Public Function TestGetGroupsForBoard(boardid As String, ByRef rs As String, ByRef rt As String) As Collection
Dim groupColl As Collection
Dim boardGroup As Variant

    Set groupColl = GetGroupsForBoard(boardid, rs, rt)
    For Each boardGroup In groupColl
        Debug.Print boardGroup("title")
    Next boardGroup

End Function

Public Sub TestCreateMondayItem(boardid As String, groupid As String, itemName As String, ByRef rs As String, ByRef rt As String)
    CreateMondayItem boardid, groupid, itemName, rs, rt
End Sub


Public Function TestCreateMondaySubItem(parentItemId As String, itemName As String, ByRef rs As String, ByRef rt As String) As String
    CreateMondaySubItem parentItemId, itemName, rs, rt
    TestCreateMondaySubItem = getResponseItemid(rt, "create_subitem")
End Function


