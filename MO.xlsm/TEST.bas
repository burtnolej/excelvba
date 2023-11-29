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


Dim boardIds() As Variant
Dim itemIds() As Variant

Public Sub TestExportAllModules()

    Application.Run "vbautils.xlsm!ExportAllModules"
End Sub
Sub testCreateMondayFolderBatch()
Dim dataRange As Range, dataRow As Range
Dim itemID As String, itemContent As String, rs As String, rt As String, itemName As String

    Set dataRange = Selection
    For Each dataRow In dataRange.Rows
    
        itemContent = dataRow.Columns(2)
        itemName = dataRow.Columns(1)
        itemID = dataRow.Columns(3)
    
        Debug.Print "Processing : " & itemID & " " & itemName

        InitMondayFolder itemID, itemName, _
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
Dim itemID As String, itemLink As String, rs As String, rt As String, itemName As String, itemContent As String

    Set dataRange = Selection
    For Each dataRow In dataRange.Rows
        itemLink = dataRow.Columns(4)
        itemID = dataRow.Columns(3)
        itemContent = dataRow.Columns(2)
        itemName = dataRow.Columns(1)
        
        If itemLink <> "" Then
            AddToMondayFile itemID, _
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
Dim itemID As String, itemLink As String, rs As String, rt As String, itemName As String, itemContent As String

    Set dataRange = Selection
    For Each dataRow In dataRange.Rows
        itemID = dataRow.Columns(1)
        

            DeleteMondayItem itemID, _
                            ActiveWorkbook, _
                            "C:\Users\burtn\Velox Financial Technology\Velox Shared Drive - Documents\General\Monday"
    Next dataRow
End Sub

Sub testGetTags()
    Set tagColl = GetTags(rs, rt)
    For Each tag In tagColl
        Debug.Print tag.Item("id"), tag.Item("name")
        'Set userAcct = user("account")
        'Debug.Print userAcct.Item("name")
    Next tag
End Sub
Sub testAddMondayItemsBatch()
Dim dataRange As Range, dataRow As Range
Dim groupid As String, boardid As String, rs As String, rt As String, itemName As String, tags_string As String, tag_name As String
Dim status As String, owner As String, newItemName As String, newSubItemName As String, itemID As String, tag_id As String, DDQ As String, createFolderFlag As String
Dim newItemId As String, addedFlag As String, newItemUpdateMsg As String
Dim i As Integer
Dim tag_index As Variant
Dim first_tag As Boolean
Dim tags_array As Variant
Dim tagsIdsRange As Range, tagsNamesRange As Range, addedItemURLRange As Range, addedItemFolderRange As Range
Dim tagsDict As Dictionary, statusDict As Dictionary
Dim tmpSheet As Worksheet

    

    Set tagsDict = New Dictionary
    Set statusDict = New Dictionary
    DDQ = """"
    
    Set tmpSheet = Workbooks("MO.xlsm").Sheets("AddNewItems")
    Set dataRange = tmpSheet.Range("NEWITEM_DATA")
    Set addedItemIdRange = tmpSheet.Range("NEWITEM_ADDEDITEMID")
    Set addedItemURLRange = tmpSheet.Range("NEWITEM_ADDEDITEMURL")
    Set addedItemFolderRange = tmpSheet.Range("NEWITEM_ADDEDITEMFOLDER")
    
    
    
    createFolderFlag = tmpSheet.Range("CREATE_FOLDER_FLAG").Value

    Application.Run "vbautils.xlsm!RangeToDict", Workbooks("MO.xlsm"), "Reference", "TAGS_DATA", tagsDict
    Application.Run "vbautils.xlsm!RangeToDict", Workbooks("MO.xlsm"), "Reference", "STATUS_DATA", statusDict
    
    For i = 1 To 1  ' just do the first row for now 6/24/23
        groupid = ActiveSheet.Range("NEWITEM_GROUP_ID").Rows(i)
        boardid = ActiveSheet.Range("NEWITEM_BOARD_ID").Rows(i)
        itemName = ActiveSheet.Range("NEWITEM_ITEM_NAME").Rows(i)
        newItemName = ActiveSheet.Range("NEWITEM_NEWITEM_NAME").Rows(i)
        status = ActiveSheet.Range("NEWITEM_STATUS").Rows(i)
        tags = ActiveSheet.Range("NEWITEM_TAG").Rows(i)
        owner = ActiveSheet.Range("NEWITEM_OWNER").Rows(i)
        itemID = ActiveSheet.Range("NEWITEM_ITEMID").Rows(i)
        newSubItemName = ActiveSheet.Range("NEWSUBITEM_NEWSUBITEM_NAME").Rows(i)
        newItemUpdateMsg = ActiveSheet.Range("NEWITEM_NEWITEM_UPDATE").Rows(i)
        
        
        
        
        addedFlag = addedItemIdRange.Rows(i)

        tags_array = Split(tags, "^")
        tags_string = ""
        first_tag = True
        
        
        status = statusDict.Item(status)
        
        For j = 0 To UBound(tags_array)
            tag_name = tags_array(j)

            If tagsDict.Exists(tag_name) = True Then
                tag_id = Str(tagsDict.Item(tag_name))
            Else
                tag_id = AddTag(tag_name, rs, rt)
            End If
            
            If first_tag = False Then
                tags_string = tags_string & "," & tag_id
            Else
                tags_string = tag_id
            End If
        
            If first_tag = True Then
                first_tag = False
            End If
                
        Next j
        
        If addedFlag = "" Then
            If newItemName <> "" Then ' then its a new item
                If boardid <> "" Then
                    If newItemName <> "N/A" Then  ' if N/A then its adding a sibitem anyway 06/24/23
                        CreateMondayItem boardid, groupid, newItemName, status, owner, tags_string, rs, rt
                        newItemId = getResponseItemid(rt, "create_item")
                    Else
                        newItemId = itemID
                    End If

                    If newSubItemName <> "" Then  ' if "" then its adding  a parent only 06/24/23
                        CreateMondaySubItem newItemId, newSubItemName, status, owner, tags_string, rs, rt
                        newItemId = getResponseItemid(rt, "create_subitem")
                        newItemName = newSubItemName
                    End If
                    
                Else
                    Debug.Print "exiting from row #" & i & " as end of items to add"
                    Exit Sub
                End If
            ElseIf newSubItemName <> "" Then ' its a new sub item
                    newItemName = newSubItemName
                    CreateMondaySubItem itemID, newItemName, status, owner, tags_string, rs, rt
                    newItemId = getResponseItemid(rt, "create_subitem")
            End If
            
            ' add the new item id back into the worksheet
           
            addedItemIdRange.Rows(i).Value = newItemId
            
            ' add the new URL into the worksheet
            addedItemURLRange.Rows(i).Formula = "=HYPERLINK(" & DDQ & "https://veloxfintech.monday.com/boards/" & boardid & "/pulses/" & newItemId & DDQ & ")"
            
            ' then post the description field as an update
            PostUpdateMonday newItemId, newItemUpdateMsg, rs, rt
            
            If createFolderFlag = "YES" Then
                folderPath = CreateSimpleMondayFolder("E:\Velox Financial Technology\Velox Shared Drive - Documents\General\Monday", newItemId, newItemName)
            
                addedItemFolderRange.Rows(i).Formula = "=HYPERLINK(" & DDQ & folderPath & DDQ & ")"
            End If
        
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

Sub testDisplayTags()
Dim tmpSheet As Worksheet

    Set tmpSheet = ActiveWorkbook.Sheets("Reference")
    tmpSheet.Activate
    SetEventsOff
    DisplayTags tmpSheet, tmpSheet.Range("AP2")

    SetEventsOn
    
End Sub


Public Sub TestMondayAPI()
Dim i As Integer
Dim rs As String, rt As String
Dim jsonObject As Object
Dim dataDict As New Dictionary
Dim subItemBoardId As String

    ReDim boardsIds(0 To 3)
    ReDim itemIds(0 To 3)

    'TestCreateMondayItem "4977328922", "topics", "test", "19676602,19045698", rs, rt
    TestCreateMondaySubItem "5013296680", "test", "1", "22121", "19676602,19045698", rs, rt
    
    Exit Sub
    
    
    'Debug.Print TestCreateMondaySubItem("2951573079", "foo", rs, rt)
    
    
    

    'Admin Software& Services
    '{\"tag_ids\":[10165564,10166350]}"},
    
    UpdateItemAttributeMonday CLIENTS_BOARD_ID, CLIENTS_ITEM_ID, "name", "foo", rs, rt
    
    
    'UpdateTagsMonday CLIENTS_BOARD_ID, CLIENTS_ITEM_ID, "10165564,10166350", rs, rt
    
    
    
    
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


Public Sub TestPostUpdate(itemID As Variant, ByRef rs As String, ByRef rt As String)
    PostUpdateMonday CStr(itemID), "1", rs, rt
   
End Sub

Public Sub TestUpdateStatus(itemID As Variant, boardid As Variant, ByRef rs As String, ByRef rt As String)
    UpdateStatusMonday CStr(boardid), CStr(itemID), "0", rs, rt
End Sub

Public Sub TestUpdateSubItemStatusMonday(itemID As Variant, boardid As Variant, ByRef rs As String, ByRef rt As String)
    UpdateSubItemStatusMonday CStr(boardid), CStr(itemID), "0", rs, rt
End Sub

Public Function TestGetBoardId(itemID As Variant, ByRef rs As String, ByRef rt As String) As Variant
    TestGetBoardId = GetBoardId(CStr(itemID), rs, rt)
End Function

Public Function TestGetBoards(ByRef rs As String, ByRef rt As String) As Variant

    Set TestGetBoards = GetBoards(rs, rt)
    
End Function
Public Function TestGetGroupsForBoard(boardid As String, ByRef rs As String, ByRef rt As String) As Collection
Dim groupColl As Collection
Dim boardGroup As Variant

    Set groupColl = GetGroupsForBoard(boardid, rs, rt)
    For Each boardGroup In groupColl
        Debug.Print boardGroup("title")
    Next boardGroup

End Function

Public Sub TestCreateMondayItem(boardid As String, groupid As String, itemName As String, tags As String, ByRef rs As String, ByRef rt As String)
    CreateMondayItem boardid, groupid, itemName, "1", "22121", tags, rs, rt
End Sub


Public Function TestCreateMondaySubItem(parentItemId As String, itemName As String, status As String, owner As String, tags As String, ByRef rs As String, ByRef rt As String) As String

    CreateMondaySubItem parentItemId, itemName, status, owner, tags, rs, rt
    TestCreateMondaySubItem = getResponseItemid(rt, "create_subitem")
End Function


