Attribute VB_Name = "FORCOPY"




Public Sub AddTag(tagname As String, ByRef responseStatus As String, ByRef responseText As String)
Dim DDQ As String, postData As String
    DDQ = Chr(34)
    postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "mutation {create_or_get_tag (tag_name: \" & DDQ & tagname & "\" & DDQ & ") {id}}" & DDQ & "}"
    WriteToMondayAPI postData, responseStatus, responseText
    Set jsonObject = ParseJson(responseText)
    If jsonObject.Exists("error_code") = True Then:
       responseText = jsonObject.item("error_code") & " : " & jsonObject.item("error_message")
       
End Sub

Public Sub UpdateStatusMonday(board_id As String, itemId As String, newStatus As String, ByRef responseStatus As String, ByRef responseText As String)
Dim DDQ As String, postData As String
    DDQ = Chr(34)
    postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "mutation { change_simple_column_value(board_id : " & board_id & ", item_id :" & itemId & ", column_id: \" & DDQ & "status" & "\" & DDQ & ", value: \" & DDQ & _
            newStatus & "\" & DDQ & ") {id}}" & DDQ & "}"

    LogIt postData, "UpdateStatusMonday", "postData"
    WriteToMondayAPI postData, responseStatus, responseText
    LogIt responseStatus & ":" & responseText, "UpdateStatusMonday", "response"

    WriteToMondayAPI postData, responseStatus, responseText
    
    Set jsonObject = ParseJson(responseText)
    If jsonObject.Exists("error_code") = True Then:
       responseText = jsonObject.item("error_code") & " : " & jsonObject.item("error_message")
       
   
    
End Sub


Sub MoveItemToBoard(old_board_id As String, board_id As String, item_id As String, new_group_id As String, columnMappings As String, _
         ByRef responseStatus As String, ByRef responseText As String)
Dim DDQ As String, postData As String
DDQ = Chr(34)

'mutation {
'  move_item_to_board (board_id:1234567890, group_id: "new_group", item_id:9876543210,

'columns_mapping: [{source:"status", target:"status2"}, {source:"person", target:"person"}, {source:"date", target:"date4"}]) {
'    ID
'  }
'}

     postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "mutation { move_item_to_board(" & _
            "item_id : " & item_id & _
            ", board_id :" & board_id & _
            ", group_id :" & "\" & DDQ & new_group_id & "\" & DDQ & _
            " ,columns_mapping: " & "[" & columnMappings & "]" & _
            ") {id}}" _
            & DDQ & "}"
            
    
    WriteToMondayAPI postData, responseStatus, responseText
    
    Set jsonObject = ParseJson(responseText)
    If jsonObject.Exists("error_code") = True Then:
       responseText = jsonObject.item("error_code") & " : " & jsonObject.item("error_message")
       
End Sub


Sub GetItemDetails(item_id As String, ByRef items As Collection, ByRef subitems As Collection, _
                    ByRef subitemsUpdates As Collection, ByRef itemsUpdates As Collection, _
                    ByRef responseStatus As String, ByRef responseText As String)
Dim DDQ As String, postData As String
DDQ = Chr(34)

  postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "query { items (ids: [" & item_id & "]) {name, id,  column_values { id,value },updates {creator_id,item_id ,text_body,assets{name,url,file_extension}}, " & _
            "subitems{name,parent_item {id},column_values{id, value} updates {creator_id,item_id ,text_body,assets{name,url,file_extension}}}" & _
            "}}" & DDQ & "}"
            
    
    WriteToMondayAPI postData, responseStatus, responseText

    Set jsonObject = ParseJson(responseText)
    
    Set items = jsonObject("data")("items")
    Set itemsUpdates = items(1)("updates")
    Set subitems = items(1)("subitems")
    Set subitemsUpdates = subitems(1)("updates")
    
    If jsonObject.Exists("error_code") = True Then
        Debug.Print
    End If
    
End Sub

Sub AddItem(board_id As String, group_id As String, ByRef items As Collection, ByRef subitems As Collection, _
                    ByRef subitemsUpdates As Collection, ByRef itemsUpdates As Collection, _
                    ByRef responseStatus As String, ByRef responseText As String)

End Sub

Public Sub UpdateTagsMonday(board_id As String, itemId As String, newTags As Variant, ByRef responseStatus As String, ByRef responseText As String)
Dim DDQ As String, postData As String
Dim jsonObject As Dictionary

    DDQ = Chr(34)


    newTags = "{\\\" & DDQ & "tags" & "\\\" & DDQ & ": " & "{" & "\\\" & DDQ & "tag_ids" & "\\\" & DDQ & ":[" & newTags & "]}}"
    
    postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "mutation { change_multiple_column_values(" & _
            "item_id : " & itemId & _
            ", board_id :" & board_id & _
            " ,column_values: " & "\" & DDQ & newTags & "\" & DDQ & _
            ") {id}}" _
            & DDQ & "}"
       
   WriteToMondayAPI postData, responseStatus, responseText
   
   Set jsonObject = ParseJson(responseText)
   If jsonObject.Exists("error_code") = True Then:
       responseText = jsonObject.item("error_code") & " : " & jsonObject.item("error_message")
    
    
End Sub




Public Sub UpdateOwnerMonday(board_id As String, itemId As String, userId As Variant, ByRef responseStatus As String, ByRef responseText As String)
Dim DDQ As String, postData As String
    DDQ = Chr(34)

    newOwner = "{\\\" & DDQ & "person" & "\\\" & DDQ & ":{\\\" & DDQ & "personsAndTeams" & "\\\" & DDQ & ":[{\\\" & DDQ & "id" & "\\\" & DDQ & ":" & Str(userId) & ",\\\" & DDQ & "kind" & "\\\" & DDQ & ":\\\" & DDQ & "person" & "\\\" & DDQ & "}]}}"


    postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "mutation { change_multiple_column_values(" & _
                "board_id : " & board_id & _
                ", item_id :" & itemId & _
                ", column_values: " & "\" & DDQ & newOwner & "\" & DDQ & _
                ") {id}}" _
                & DDQ & "}"


    WriteToMondayAPI postData, responseStatus, responseText

    Set jsonObject = ParseJson(responseText)
    If jsonObject.Exists("error_code") = True Then:
       responseText = jsonObject.item("error_code") & " : " & jsonObject.item("error_message")
End Sub




Public Sub UpdateItemAttributeMonday(board_id As String, itemId As String, itemType As String, newItemValue As String, ByRef responseStatus As String, ByRef responseText As String)
Dim DDQ As String, postData As String
    DDQ = Chr(34)
    postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "mutation { change_simple_column_value(board_id : " & board_id & ", item_id :" & itemId & ", column_id: \" & DDQ & itemType & "\" & DDQ & ", value: \" & DDQ & _
            newItemValue & "\" & DDQ & ") {id}}" & DDQ & "}"

    LogIt postData, "UpdateItemAttributeMonday", "postData"
    
    WriteToMondayAPI postData, responseStatus, responseText
    LogIt responseStatus & ":" & responseText, "UpdateItemAttributeMonday", "response"
End Sub




Public Sub PostUpdateMonday(itemId As String, msg As String, ByRef responseStatus As String, ByRef responseText As String)
Dim DDQ As String, postData As String
    DDQ = Chr(34)
    postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "mutation { create_update (item_id: " & itemId & ", body: " & "\" & DDQ & msg & "\" & DDQ & ") {id}}" & DDQ & "}"
    
    WriteToMondayAPI postData, responseStatus, responseText
End Sub

  
  

Public Function GetBoardId(itemId As String, ByRef responseStatus As String, ByRef responseText As String) As Variant
Dim DDQ As String, postData As String, GetBoardIdStr As String
Dim jsonObject As Object
    DDQ = Chr(34)

    postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "query {items (ids: [ " & itemId & "]) { board { id } } }" & DDQ & "}"
    LogIt postData, "GetBoardId", "postData"
    
    WriteToMondayAPI postData, responseStatus, responseText
    LogIt responseStatus & ":" & responseText, "GetBoardId", "response"
    
    Set jsonObject = ParseJson(responseText)
    GetBoardIdStr = jsonObject("data")("items")(1)("board")("id")
    LogIt GetBoardIdStr, "GetBoardId", "result"
    GetBoardId = GetBoardIdStr
        
End Function




Public Function GetBoards(ByRef responseStatus As String, ByRef responseText As String) As Collection
'query { boards (ids: 157244624) { groups (ids: status) { Title Color Position } }}
'query { boards () {name State board_folder_id items {ID name column_values { Text } } } }

Dim DDQ As String, postData As String
Dim jsonObject As Object
Dim boardsObjects As Collection
Dim itemDict As Variant

    DDQ = Chr(34)

    postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "query {boards () { id name permissions state} }" & DDQ & "}"
    WriteToMondayAPI postData, responseStatus, responseText
    
    Set jsonObject = ParseJson(responseText)
    Set GetBoards = jsonObject("data")("boards")

End Function




Public Function GetSubitemColumns(boardid As String, subitemid As String, ByRef responseStatus As String, ByRef responseText As String) As Collection

Dim DDQ As String, postData As String
Dim jsonObject As Object
Dim boardsObjects As Collection
Dim tmpDict As Dictionary
Dim itemDict As Variant

 
    DDQ = Chr(34)


    postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "query {items (ids:" & subitemid & "){ column_values  {id title value text }}}" & DDQ & "}"
    WriteToMondayAPI postData, responseStatus, responseText
    
    Set jsonObject = ParseJson(responseText)
    Set tmpDict = jsonObject("data")
    Set tmpDict = tmpDict.item("items")(1)
    Set GetSubitemColumns = tmpDict.item("column_values")
    
End Function




Public Function GetBoardColumns(boardid As String, ByRef responseStatus As String, ByRef responseText As String) As Collection
Dim DDQ As String, postData As String
Dim jsonObject As Object
Dim boardsObjects As Collection
Dim itemDict As Variant

 '{"query":"query {   boards (ids: 1234567)   {owner{ id }  columns {   title   type }}}"}'
 
    DDQ = Chr(34)

    postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "query {boards (ids:" & boardid & ")  {owner{ id }  columns { id  title   type }} }" & DDQ & "}"
    WriteToMondayAPI postData, responseStatus, responseText
    
    Set jsonObject = ParseJson(responseText)
    Set GetBoardColumns = jsonObject("data")("boards")(1)("columns")

     
End Function




Public Function GetUsers(ByRef responseStatus As String, ByRef responseText As String) As Collection
'query { users () { created_at email account { name id}}

Dim DDQ As String, postData As String
Dim jsonObject As Object
Dim boardsObjects As Collection
Dim itemDict As Variant

    DDQ = Chr(34)

    postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "query { users () { created_at email name id account { name id }}}" & DDQ & "}"
    
    WriteToMondayAPI postData, responseStatus, responseText
    
    Set jsonObject = ParseJson(responseText)
    Set GetUsers = jsonObject("data")("users")

End Function




Public Function GetTags(ByRef responseStatus As String, ByRef responseText As String) As Collection
Dim DDQ As String, postData As String
Dim jsonObject As Object
Dim boardsObjects As Collection
Dim itemDict As Variant

    DDQ = Chr(34)

    postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "query {tags () { id name } }" & DDQ & "}"
    WriteToMondayAPI postData, responseStatus, responseText
    
    Set jsonObject = ParseJson(responseText)
    Set GetTags = jsonObject("data")("tags")

End Function




Public Sub DeleteItem(itemId As String, ByRef responseStatus As String, ByRef responseText As String)
Dim DDQ As String, postData As String
'mutation { delete_item (item_id: 12345678) {ID}}

    DDQ = Chr(34)

    postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "mutation { delete_item (item_id : " & itemId & ") {id}}" & DDQ & "}"
    WriteToMondayAPI postData, responseStatus, responseText

End Sub




Public Function GetGroupsForBoard(boardid As String, ByRef responseStatus As String, ByRef responseText As String) As Collection
'query { boards (ids: 157244624) { groups (ids: status) { Title Color Position  } }}
Dim DDQ As String, postData As String
Dim jsonObject As Object
Dim boardsObjects As Collection
Dim itemDict As Variant

    DDQ = Chr(34)

    postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "query {boards (ids: " & boardid & ") { groups () {id title}} }" & DDQ & "}"
    WriteToMondayAPI postData, responseStatus, responseText
    
    Set jsonObject = ParseJson(responseText)
    

    Set boardsObjects = jsonObject("data")("boards")
    Set GetGroupsForBoard = boardsObjects(1)("groups")
    
End Function




Public Function CreateDropdownColumn(boardid As String, dropDownName As String, ByRef responseStatus As String, ByRef responseText As String) As String
Dim DDQ As String, postData As String
Dim jsonObject As Object

    DDQ = Chr(34)

    postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "mutation {create_column (board_id: " & boardid & ",title:" & "\" & DDQ & dropDownName & "\" & DDQ & _
                ",column_type:dropdown" & ") { id } } " & DDQ & "}"
 
    WriteToMondayAPI postData, responseStatus, responseText
    
    
    Set jsonObject = ParseJson(responseText)
    CreateDropdownColumn = jsonObject("data")("create_column")("id")
    
End Function




Public Sub SetDropdownColumnValues(boardid As String, itemId As String, dropDownName As String, dropDownValues As String, _
    ByRef responseStatus As String, ByRef responseText As String)
Dim DDQ As String, postData As String

    DDQ = Chr(34)

    postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "mutation {change_simple_column_value (item_id: " & itemId & ", board_id: " & boardid & _
                ",column_id:" & "\" & DDQ & dropDownName & "\" & DDQ & ",value:" & "\" & DDQ & dropDownValues & "\" & DDQ & ", create_labels_if_missing: true) { name id } } " & DDQ & "}"
 
    WriteToMondayAPI postData, responseStatus, responseText

    
    
End Sub




Public Function CreateMondayItem(boardid As String, groupid As String, itemName As String, status As String, _
        ByRef responseStatus As String, ByRef responseText As String) As String

Dim DDQ As String, postData As String
Dim jsonObject As Object
Dim boardsObjects As Collection
Dim itemDict As Variant

    DDQ = Chr(34)
 
    column_values = "{\\\" & DDQ & "status" & "\\\" & DDQ & ": " & "\\\" & DDQ & status & "\\\" & DDQ & "}"
    
    postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "mutation { create_item(" & _
                "board_id : " & boardid & _
                ", group_id :" & "\" & DDQ & groupid & "\" & DDQ & _
                ", item_name :" & "\" & DDQ & itemName & "\" & DDQ & _
                ", column_values: " & "\" & DDQ & column_values & "\" & DDQ & _
                ") {id}}" _
                & DDQ & "}"
                
                
    WriteToMondayAPI postData, responseStatus, responseText

    Set jsonObject = ParseJson(responseText)
    CreateMondayItem = jsonObject("data")("create_item")("id")
    
End Function


Public Function CreateMondaySubItem(parentId As String, itemName As String, status As String, ByRef responseStatus As String, ByRef responseText As String) As String

Dim DDQ As String, postData As String
Dim jsonObject As Object
Dim boardsObjects As Collection
Dim itemDict As Variant

    DDQ = Chr(34)
 
    column_values = "{\\\" & DDQ & "status" & "\\\" & DDQ & ": " & "\\\" & DDQ & status & "\\\" & DDQ & "}"
    
    postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "mutation { create_subitem(" & _
                "parent_item_id :" & "\" & DDQ & parentId & "\" & DDQ & _
                ", item_name :" & "\" & DDQ & itemName & "\" & DDQ & _
                ", column_values: " & "\" & DDQ & column_values & "\" & DDQ & _
                ") {id}}" _
                & DDQ & "}"
                
    WriteToMondayAPI postData, responseStatus, responseText
    

    Set jsonObject = ParseJson(responseText)
    CreateMondaySubItem = jsonObject("data")("create_subitem")("id")
    
    
End Function

Public Function AddUpdate(parentId As String, body As String, ByRef responseStatus As String, ByRef responseText As String) As String
Dim DDQ As String, postData As String

    DDQ = Chr(34)
'mutation {
'    create_update (item_id: 1234567890, body: "This update will be added to the item") {
'    ID
'  }
'}
    postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "mutation { create_update(" & _
                "item_id :" & "\" & DDQ & parentId & "\" & DDQ & _
                ", body :" & "\" & DDQ & body & "\" & DDQ & _
                ") {id}}" _
                & DDQ & "}"

    WriteToMondayAPI postData, responseStatus, responseText
    

    Set jsonObject = ParseJson(responseText)
    AddUpdate = jsonObject("data")("create_update")("id")
End Function
                
Public Sub AddFileToUpdate(updateIt As String, file As Variant, ByRef responseStatus As String, ByRef responseText As String)
Dim DDQ As String, postData As String

    DDQ = Chr(34)
    
'mutation {
'    add_file_to_update (update_id: 1234567890, file: YOUR_FILE) {
'        ID
'    }
'}

    postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "mutation { add_file_to_update(" & _
                "update_id :" & "\" & DDQ & parentId & "\" & DDQ & _
                ", file :" & "\" & DDQ & file & "\" & DDQ & _
                ") {id}}" _
                & DDQ & "}"

    WriteToMondayAPI postData, responseStatus, responseText

End Sub

Public Sub ArchiveItem()
'mutation {
'    archive_item (item_id: 1234567890) {
'        ID
'    }
'}
End Sub


Public Sub WriteToMondayAPI(postData As String, ByRef responseStatus As String, ByRef responseText As String)
Dim objHTTP As Object
Dim DDQ As String, apiKey As String, Url As String
    DDQ = Chr(34)
    
    apiKey = "eyJhbGciOiJIUzI1NiJ9.eyJ0aWQiOjExMTgyMDkwNiwidWlkIjoxNTE2MzEwNywiaWFkIjoiMjAyMS0wNS0zMFQxMTowMDo1OS4wMDBaIiwicGVyIjoibWU6d3JpdGUiLCJhY3RpZCI6NjY5MDk4NSwicmduIjoidXNlMSJ9.zIeOeoqeaZ2Q8NuKBPPw2LQFh2JRPvPwIkhhn4e5Q08"
    Url = "https://api.monday.com/v2"
    Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    objHTTP.Open "POST", Url, False
    objHTTP.setRequestHeader "Authorization", apiKey
    objHTTP.setRequestHeader "Content-Type", "application/json"
    objHTTP.setRequestHeader "API-Version", "2023-10"
    objHTTP.send postData
    
    responseStatus = objHTTP.status
    responseText = objHTTP.responseText
    
    'myParseJson objHTTP.responseText
    
End Sub




Public Sub LogIt(logMsg As String, logFunction As String, logMsgDescription As String)
Dim tmpSheet As Worksheet
Dim lastUsedRow As Long
Dim activityRange As Range

    Set tmpSheet = ActiveWorkbook.Sheets("Logs")
    Set activityRange = tmpSheet.Range("ACTIVITY_LOG")
    lastUsedRow = activityRange.SpecialCells(xlCellTypeLastCell).Row
    
    
    activityRange.Rows(lastUsedRow + 1).Columns(1).Value = Now()
    activityRange.Rows(lastUsedRow + 1).Columns(2).Value = logMsg
    activityRange.Rows(lastUsedRow + 1).Columns(3).Value = logFunction
    activityRange.Rows(lastUsedRow + 1).Columns(4).Value = logMsgDescription

endsub:
    Set activityRange = Nothing
    Set tmpSheet = Nothing
    
End Sub




Sub BatchUpdateMondayStatus()
Dim Target As Range, sourceColumn As Range
Dim tmpVal As String, itemId As String, boardid As String, responseStatus As String, responseText As String, newStatus As String, itemType As String, subitemid As String
Dim userId As Variant, tagId As Variant, tag As Variant
Dim userIdRange As Range, tagNameRange As Range

    Set sourceColumn = ActiveSheet.Range("COLUMN_UPDATES_MONDAY_STATUS")
    For Each Target In sourceColumn.offset(1, 0).Cells
        If Not Target.Value = vbNullString Then
            itemId = ActiveSheet.Range("COLUMN_ITEMID").Rows(Target.Row - 3).Value
            boardid = ActiveSheet.Range("COLUMN_BOARDID").Rows(Target.Row - 3).Value
            itemType = ActiveSheet.Range("COLUMN_TYPE").Rows(Target.Row - 3).Value
            If itemType = "subitem" Then boardid = GetBoardId(CStr(itemId), responseStatus, responseText)
            If Target.Value = "Working" Then newStatus = "0" Else If Target.Value = "Completed" Then newStatus = "1" Else If Target.Value = "Duplicate" Then newStatus = "6" Else If Target.Value = "Ongoing" Then newStatus = "7" Else newStatus = "7"
            UpdateStatusMonday boardid, itemId, newStatus, responseStatus, responseText
            If responseStatus = "200" Then
                Debug.Print "Successfullly updated [" & itemId & "] to " & Target.Value & "[" & responseText & "]"
           Else
                Debug.Print "Failed to update  [" & itemId & "] to " & Target.Value & "[" & responseText & "]"
           End If
       End If

    Next Target

End Sub




Sub BatchUpdateMondayOwner()
Dim Target As Range, sourceColumn As Range
Dim tmpVal As String, itemId As String, boardid As String, responseStatus As String, responseText As String, newStatus As String, itemType As String, subitemid As String
Dim userId As Variant, tagId As Variant, tag As Variant
Dim userIdRange As Range, tagNameRange As Range

    'Set sourceColumn = ActiveSheet.Range("COLUMN_UPDATES_MONDAY_OWNER")
    Set sourceColumn = Selection
    For Each Target In sourceColumn.offset(1, 0).Cells
         If Not Target.Value = vbNullString Then
            itemId = ActiveSheet.Range("COLUMN_ITEMID").Rows(Target.Row - 3).Value
            boardid = ActiveSheet.Range("COLUMN_BOARDID").Rows(Target.Row - 3).Value
            itemType = ActiveSheet.Range("COLUMN_TYPE").Rows(Target.Row - 3).Value
            If itemType = "subitem" Then boardid = GetBoardId(CStr(itemId), responseStatus, responseText)
            Set userIdRange = Worksheets("Reference").Range("DATA_USERNAME")
            If IsError(Application.Match(Target.Value, userIdRange, 0)) Then
                Debug.Print Target.Value & "not found"
            Else
                userRow = Application.Match(Target.Value, userIdRange, 0)
                userId = userIdRange.Rows(userRow).offset(, 3).Value
            End If
            UpdateOwnerMonday boardid, itemId, userId, responseStatus, responseText
            If responseStatus = "200" Then
                Debug.Print "Successfullly updated [" & itemId & "] to " & Target.Value & "[" & responseText & "]"
           Else
                Debug.Print "Failed to update  [" & itemId & "] to " & Target.Value & "[" & responseText & "]"
           End If
         End If
    Next Target
End Sub
