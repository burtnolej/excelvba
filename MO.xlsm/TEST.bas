Attribute VB_Name = "TEST"

'Public Sub TestExportAllModules()
'Sub testCreateMondayFolderBatch()
'Sub AddToMondayFiles()
'Sub TestDeleteMondayItems()
'Sub testGetTags()

'Public Sub TestMondayAPI()
'Public Sub TestPostUpdate(itemID As Variant, ByRef rs As String, ByRef rt As String)
'Public Sub TestUpdateStatus(itemID As Variant, boardid As Variant, ByRef rs As String, ByRef rt As String)
'Public Sub TestUpdateSubItemStatusMonday(itemID As Variant, boardid As Variant, ByRef rs As String, ByRef rt As String)
'Public Function TestGetBoardId(itemID As Variant, ByRef rs As String, ByRef rt As String) As Variant
'Public Function TestGetBoards(ByRef rs As String, ByRef rt As String) As Variant
'Public Function TestGetGroupsForBoard(boardid As String, ByRef rs As String, ByRef rt As String) As Collection
'Public Sub TestCreateMondayItem(boardid As String, groupid As String, itemName As String, tags As String, ByRef rs As String, ByRef rt As String)
'Public Function TestCreateMondaySubItem(parentItemId As String, itemName As String, status As String, owner As String, tags As String, ByRef rs As String, ByRef rt As String) As String

'Sub RefreshGroupsExec()
'Sub RefreshTagsExec()
'Sub RefreshItemsExec()
'Sub RefreshUsersExec()

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

Sub TestCreateSharepointMondayFolder()

    CreateSharepointMondayFolder "foo"

End Sub
Public Sub CreateSharepointMondayFolder(folderName As String)
Dim objShell As Object
Dim PSExe, PSScript As String
    
    Set objShell = VBA.CreateObject("Wscript.Shell")

    PythonExe = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
    PythonScript = "" & Environ("USERPROFILE") & "\Deploy\CreaterFolder-Nodep.ps1"

    objShell.Run PythonExe & " " & PythonScript & " " & """/sites/VeloxSharedDrive/Shared%20Documents/General/Monday""" & " " & folderName

End Sub
Public Sub TestExportAllModules()

    Application.Run "vbautils.xlsm!ExportAllModules"
End Sub
Sub testCreateMondayFolderBatch()
Dim dataRange As Range, dataRow As Range
Dim itemid As String, itemContent As String, rs As String, rt As String, itemName As String

    Set dataRange = Selection
    For Each dataRow In dataRange.Rows
    
        itemContent = dataRow.Columns(2)
        itemName = dataRow.Columns(1)
        itemid = dataRow.Columns(3)
    
        Debug.Print "Processing : " & itemid & " " & itemName

        InitMondayFolder itemid, itemName, _
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
Dim itemid As String, itemLink As String, rs As String, rt As String, itemName As String, itemContent As String

    Set dataRange = Selection
    For Each dataRow In dataRange.Rows
        itemLink = dataRow.Columns(4)
        itemid = dataRow.Columns(3)
        itemContent = dataRow.Columns(2)
        itemName = dataRow.Columns(1)
        
        If itemLink <> "" Then
            AddToMondayFile itemid, _
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
Dim itemid As String, itemLink As String, rs As String, rt As String, itemName As String, itemContent As String

    Set dataRange = Selection
    For Each dataRow In dataRange.Rows
        itemid = dataRow.Columns(1)
        

            DeleteMondayItem itemid, _
                            ActiveWorkbook, _
                            "C:\Users\burtn\Velox Financial Technology\Velox Shared Drive - Documents\General\Monday"
    Next dataRow
End Sub

Sub testGetTags()
    Set tagColl = GetTags(rs, rt)
    For Each tag In tagColl
        Debug.Print tag.Item("id"), tag.Item("name")
    Next tag
End Sub

Function GetTagsString(tagArray As Variant) As String
Dim first_tag As Boolean
Dim tag_id As String, tag_name As String, rs As String, rt As String
Dim tagsDict As New Dictionary
    'Set tagsDict = New Dictionary
    GetTagsString = ""
    first_tag = True
    For j = 0 To UBound(tagArray)
        tag_name = tagArray(j)
        
        If tag_name <> "SELECT_ONE" Then
    
            If tagsDict.Exists(tag_name) = True Then
                tag_id = Str(tagsDict.Item(tag_name))
            Else
                tag_id = AddTag(tag_name, rs, rt)
            End If
            
            If first_tag = False Then
                GetTagsString = GetTagsString & "," & tag_id
            Else
                GetTagsString = tag_id
            End If
        
            If first_tag = True Then
                first_tag = False
            End If
        End If
    Next j
End Function

Public Sub NewItemExec()
Dim tmpSheet As Worksheet

    Set tmpSheet = Workbooks("MO.xlsm").Sheets("AddNewItems")
    
    tmpSheet.Range("ADD_ITEM_ITEM_NAMES").value = "NEW_ITEM"

    tmpSheet.Range("ADDITEM_BOARD_NAME").value = "SELECT_ONE"
    tmpSheet.Range("ADD_ITEM_GROUP_NAMES").value = "SELECT_ONE"
    tmpSheet.Range("NEWITEM_NEWITEM_NAME").value = "INPUT_ONE"
    tmpSheet.Range("NEWSUBITEM_NEWSUBITEM_NAME").value = "SELECT_ONE"
    tmpSheet.Range("NEWITEM_TAG").value = "SELECT_ONE"
    tmpSheet.Range("NEWITEM_TAG2").value = "SELECT_ONE"
    
    tmpSheet.Range("NEWSUBITEM_TAG").value = "SELECT_ONE"
    tmpSheet.Range("NEWSUBITEM_TAG2").value = "SELECT_ONE"
    tmpSheet.Range("NEWITEM_OWNER").value = "SELECT_ONE"
    tmpSheet.Range("NEWSUBITEM_OWNER").value = "SELECT_ONE"
    
    tmpSheet.Range("NEWITEM_ADDEDITEMID").value = ""
    tmpSheet.Range("NEWITEM_ADDEDSUBITEMID").value = ""
    tmpSheet.Range("NEWITEM_ADDEDITEMURL").value = ""
    tmpSheet.Range("NEWITEM_ADDEDITEMFOLDER").value = ""
    
    tmpSheet.Range("NEWITEM_NEWITEM_UPDATE").value = "INPUT_ONE"
    tmpSheet.Range("NEWSUBITEM_NEWSUBITEM_UPDATE").value = "INPUT_ONE"
    
    
    
    tmpSheet.Range("NEWSUBITEM_STATUS").value = "SELECT_ONE"
    tmpSheet.Range("NEWITEM_STATUS").value = "SELECT_ONE"
    

Set tmpSheet = Nothing

End Sub
Public Sub AddItemExec(ByRef rs As String, ByRef rt As String, ByRef sirs As String, ByRef sirt As String)
Dim dataRange As Range, dataRow As Range
Dim groupid As String, boardid As String, itemName As String, tags_string As String, tag_name As String, subitemStatus As String, ownerEnum As String, subitemtags_string As String
Dim status As String, owner As String, newItemName As String, newSubItemName As String, itemid As String, tag_id As String, DDQ As String, createFolderFlag As String, subitem_statusenum As String
Dim newItemId As String, addedFlag As String, newItemUpdateMsg As String, newSubItemId As String, newSubItemUpdateMsg As String, subitemOwnerEnum As String, newFolderName As String
Dim i As Integer
Dim tag_index As Variant
Dim first_tag As Boolean
Dim tags_array As Variant
Dim tagsIdsRange As Range, tagsNamesRange As Range, addedItemURLRange As Range, addedItemFolderRange As Range, addedSubItemIdRange As Range
Dim tagsDict As Dictionary, statusDict As Dictionary
Dim tmpSheet As Worksheet

    SetEventsOff

    Set tagsDict = New Dictionary
    Set statusDict = New Dictionary
    DDQ = """"
    
    Set tmpSheet = Workbooks("MO.xlsm").Sheets("AddNewItems")
    Set dataRange = tmpSheet.Range("NEWITEM_DATA")
    Set addedItemIdRange = tmpSheet.Range("NEWITEM_ADDEDITEMID")
    Set addedSubItemIdRange = tmpSheet.Range("NEWITEM_ADDEDSUBITEMID")
    
    Set addedItemURLRange = tmpSheet.Range("NEWITEM_ADDEDITEMURL")
    Set addedItemFolderRange = tmpSheet.Range("NEWITEM_ADDEDITEMFOLDER")
    Set addedSubItemURLRange = tmpSheet.Range("NEWITEM_ADDEDSUBITEMURL")
    Set addedSubItemFolderRange = tmpSheet.Range("NEWITEM_ADDEDSUBITEMFOLDER")
    
    createFolderFlag = tmpSheet.Range("CREATE_FOLDER_FLAG").value
    createSubItemFolderFlag = tmpSheet.Range("CREATE_SUBITEM_FOLDER_FLAG").value

    Application.Run "vbautils.xlsm!RangeToDict", Workbooks("MO.xlsm"), "Reference", "TAGS_DATA", tagsDict
    Application.Run "vbautils.xlsm!RangeToDict", Workbooks("MO.xlsm"), "Reference", "STATUS_DATA", statusDict
    
    'For i = 1 To 1  ' just do the first row for now 6/24/23
    groupid = tmpSheet.Range("NEWITEM_GROUP_ID").value
    boardid = tmpSheet.Range("NEWITEM_BOARD_ID").value
    itemName = tmpSheet.Range("NEWITEM_ITEM_NAME").value
    newItemName = tmpSheet.Range("NEWITEM_NEWITEM_NAME").value
    status = tmpSheet.Range("NEWITEM_STATUS").value
    subitemStatus = tmpSheet.Range("NEWSUBITEM_STATUS").value
    tags = tmpSheet.Range("NEWITEM_TAG").value
    tags2 = tmpSheet.Range("NEWITEM_TAG2").value
    subitemtags = tmpSheet.Range("NEWSUBITEM_TAG").value
    subitemtags2 = tmpSheet.Range("NEWSUBITEM_TAG2").value
    owner = tmpSheet.Range("NEWITEM_OWNER").value
    On Error Resume Next ' 1/3/24 to account for adding subitem only
    ownerEnum = tmpSheet.Range("OWNERID").value
    On Error GoTo 0
    subitemOwnerEnum = tmpSheet.Range("SUBITEMOWNERID").value
    itemid = tmpSheet.Range("NEWITEM_ITEMID").value
    newSubItemName = tmpSheet.Range("NEWSUBITEM_NEWSUBITEM_NAME").value
    newItemUpdateMsg = tmpSheet.Range("NEWITEM_NEWITEM_UPDATE").value
    newSubItemUpdateMsg = tmpSheet.Range("NEWSUBITEM_NEWSUBITEM_UPDATE").value
    statusenum = tmpSheet.Range("STATUS_ENUM").value
    subitem_statusenum = tmpSheet.Range("SUBITEM_STATUS_ENUM").value
    
    addedFlag = addedItemIdRange.value

    tags_string = GetTagsString(Array(tags, tags2))
    subitemtags_string = GetTagsString(Array(subitemtags, subitemtags))
    
    'you need to know what the columns are called on each board to update them. so person on Test2 wont work need to pull in the json column definitions

    'need to also change the visible columns on the report so that the input works
    
    If addedFlag = "" Then
        If newItemName <> "" Then ' then its a new item
            If boardid <> "" Then
                If itemid = "" Then ' changes on 1/3/24 to better reflect new input form
                'If newItemName <> "N/A" Then  ' if N/A then its adding a sibitem anyway 06/24/23
                
                    If boardid = "4977328922" Then
                        ' test board
                        CreateMondayItem boardid, groupid, newItemName, CStr(statusenum), ownerEnum, tags_string, rs, rt, "people8"
                    Else
                        CreateMondayItem boardid, groupid, newItemName, CStr(statusenum), ownerEnum, tags_string, rs, rt
                    End If
                    
                    newItemId = getResponseItemid(rt, "create_item")
                    UpdateStatusMondayMultiVal boardid, newItemId, status, rs, rt
                Else
                    newItemId = itemid
                End If

                If newSubItemName <> "" Then  ' if "" then its adding  a parent only 06/24/23
                    If boardid = "4977328922" Then
                        CreateMondaySubItem newItemId, newSubItemName, CStr(subitem_statusenum), subitemOwnerEnum, subitemtags_string, sirs, sirt, "people5"
                    Else
                        CreateMondaySubItem newItemId, newSubItemName, CStr(subitem_statusenum), subitemOwnerEnum, subitemtags_string, sirs, sirt
                    End If
                    newSubItemId = getResponseItemid(sirt, "create_subitem")
                    
                    boardid = getResponseItemid(sirt, "create_subitem", "board")
                    UpdateStatusMondayMultiVal boardid, newSubItemId, subitemStatus, sirs, sirt
                     
                    newItemName = newSubItemName
                End If
                
            Else
                Debug.Print "exiting from row #" & i & " as end of items to add"
                Exit Sub
            End If
        ElseIf newSubItemName <> "" Then ' its a new sub item
                newItemName = newSubItemName
                CreateMondaySubItem itemid, newItemName, status, owner, tags_string, sirs, sirt
                newItemId = getResponseItemid(sirt, "create_subitem")
        End If
        
        ' then post the description field as an update
        PostUpdateMonday newItemId, newItemUpdateMsg, rs, rt
        
        ' then post the description field as an update
        PostUpdateMonday newSubItemId, newSubItemUpdateMsg, rs, rt
        
        ' add the new item id back into the worksheet
        addedItemIdRange.value = newItemId
        addedSubItemIdRange.value = newSubItemId
        
        If createFolderFlag = "YES" Then
            newFolderName = newItemId & "_" & newItemName
            newFolderPath = "https://veloxfintechcom.sharepoint.com/sites/VeloxSharedDrive/Shared%20Documents/General/Monday/" & newFolderName
            CreateSharepointMondayFolder newFolderName
            addedItemFolderRange.Formula = "=HYPERLINK(" & DDQ & newFolderPath & DDQ & "," & DDQ & newFolderName & DDQ & ")"
            addedItemURLRange.Formula = "=HYPERLINK(" & DDQ & "https://veloxfintech.monday.com/boards/" & boardid & "/pulses/" & newItemId & DDQ & "," & DDQ & newFolderName & DDQ & ")"
        End If
        
        If createSubItemFolderFlag = "YES" Then
            newFolderName = newSubItemId & "_" & newSubItemName
            newFolderPath = "https://veloxfintechcom.sharepoint.com/sites/VeloxSharedDrive/Shared%20Documents/General/Monday/" & newFolderName
            CreateSharepointMondayFolder newFolderName
            addedSubItemFolderRange.Formula = "=HYPERLINK(" & DDQ & newFolderPath & DDQ & "," & DDQ & newFolderName & DDQ & ")"
            addedSubItemURLRange.Formula = "=HYPERLINK(" & DDQ & "https://veloxfintech.monday.com/boards/" & boardid & "/pulses/" & newSubItemId & DDQ & "," & DDQ & newFolderName & DDQ & ")"
        End If
        
        
            'folderPath = CreateSimpleMondayFolder("E:\Velox Financial Technology\Velox Shared Drive - Documents\General\Monday", newItemId, newItemName)
        
            'addedItemFolderRange.Formula = "=HYPERLINK(" & DDQ & folderPath & DDQ & ")"
        'End If
        
        ' add the new URL into the worksheet
        

    Else
        Debug.Print "skipping row #" & i & " as already added"
    End If
        

        
    SetEventsOn
End Sub
Function RefreshGroupsExec() As Long
Dim tmpSheet As Worksheet

    Set tmpSheet = ActiveWorkbook.Sheets("Reference")
    tmpSheet.Activate
    SetEventsOff
    RefreshGroupsExec = DisplayGroups(tmpSheet, tmpSheet.Range("S2"))

    SetEventsOn
    
End Function

Function RefreshTagsExec() As Long
Dim tmpSheet As Worksheet

    Set tmpSheet = ActiveWorkbook.Sheets("Reference")
    tmpSheet.Activate
    SetEventsOff
    RefreshTagsExec = DisplayTags(tmpSheet, tmpSheet.Range("AP2"))

    SetEventsOn
    
End Function

Function RefreshUsersExec() As Long
Dim tmpSheet As Worksheet

    Set tmpSheet = ActiveWorkbook.Sheets("Reference")
    tmpSheet.Activate
    SetEventsOff
    RefreshUsersExec = DisplayUsers(tmpSheet, tmpSheet.Range("AS2"))

    SetEventsOn
    
End Function

Public Sub TestMondayAPI()
Dim i As Integer
Dim rs As String, rt As String
Dim jsonObject As Object
Dim dataDict As New Dictionary
Dim subItemBoardId As String

    ReDim boardsIds(0 To 3)
    ReDim itemIds(0 To 3)

    TestGetGroupsForBoard "4977328922", rs, rt
    TestGetBoards rs, rt
    Exit Sub
    
    'TestCreateMondayItem "4977328922", "topics", "test", "19676602,19045698", rs, rt
    TestCreateMondaySubItem "5013296680", "test", "1", "22121", "19676602,19045698", rs, rt
    
    Exit Sub
    
    
    'Debug.Print TestCreateMondaySubItem("2951573079", "foo", rs, rt)
    
    
    

    'Admin Software& Services
    '{\"tag_ids\":[10165564,10166350]}"},
    
    UpdateItemAttributeMonday CLIENTS_BOARD_ID, CLIENTS_ITEM_ID, "name", "foo", rs, rt
    
    
    'UpdateTagsMonday CLIENTS_BOARD_ID, CLIENTS_ITEM_ID, "10165564,10166350", rs, rt
    
    
    
    
    
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


Public Sub TestPostUpdate(itemid As Variant, ByRef rs As String, ByRef rt As String)
    PostUpdateMonday CStr(itemid), "1", rs, rt
   
End Sub

Public Sub TestUpdateStatus(itemid As Variant, boardid As Variant, ByRef rs As String, ByRef rt As String)
    UpdateStatusMonday CStr(boardid), CStr(itemid), "0", rs, rt
End Sub

Public Sub TestUpdateSubItemStatusMonday(itemid As Variant, boardid As Variant, ByRef rs As String, ByRef rt As String)
    UpdateSubItemStatusMonday CStr(boardid), CStr(itemid), "0", rs, rt
End Sub

Public Function TestGetBoardId(itemid As Variant, ByRef rs As String, ByRef rt As String) As Variant
    TestGetBoardId = GetBoardId(CStr(itemid), rs, rt)
End Function

Public Function TestGetBoards(ByRef rs As String, ByRef rt As String) As Variant
Dim boardColl As Collection
Dim board As Variant

    Set TestGetBoards = GetBoards(rs, rt)
    
    Set boardColl = GetBoards(rs, rt)
    For Each board In boardColl
        Debug.Print board("name"), board("id")
    Next board
    
    
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


