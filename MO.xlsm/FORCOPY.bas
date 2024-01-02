Attribute VB_Name = "FORCOPY"
'Public Function AddTag(tagname As String, ByRef responseStatus As String, ByRef responseText As String) As String
'Public Sub UpdateStatusMonday(board_id As String, itemID As String, newStatus As String, ByRef responseStatus As String, ByRef responseText As String)
'Public Sub UpdateTagsMonday(board_id As String, itemID As String, newTags As String, ByRef responseStatus As String, ByRef responseText As String)
'Public Sub UpdateItemAttributeMonday(board_id As String, itemID As String, itemType As String, newItemValue As String, ByRef responseStatus As String, ByRef responseText As String)
'Public Sub PostUpdateMonday(itemID As String, msg As String, ByRef responseStatus As String, ByRef responseText As String)
'Public Function GetBoardId(itemID As String, ByRef responseStatus As String, ByRef responseText As String) As Variant
'Public Function GetBoards(ByRef responseStatus As String, ByRef responseText As String) As Collection
'Public Function GetTags(ByRef responseStatus As String, ByRef responseText As String) As Collection
'Public Sub DeleteItem(itemID As String, ByRef responseStatus As String, ByRef responseText As String)
'Public Function GetGroupsForBoard(boardid As String, ByRef responseStatus As String, ByRef responseText As String) As Collection
'Public Sub CreateMondayItem(boardid As String, groupid As String, itemName As String, statusenum As String, owner As String, newTags As String, ByRef responseStatus As String, ByRef responseText As String)
'Public Sub CreateMondaySubItem(parentItemId As String, itemName As String, statusenum As String, owner As String, newTags As String, ByRef responseStatus As String, ByRef responseText As String)
'Public Sub WriteToMondayAPI(postData As String, ByRef responseStatus As String, ByRef responseText As String)


Public Function AddTag(tagname As String, ByRef responseStatus As String, ByRef responseText As String) As String
Dim DDQ As String, postData As String
    DDQ = Chr(34)
    postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "mutation {create_or_get_tag (tag_name: \" & DDQ & tagname & "\" & DDQ & ") {id}}" & DDQ & "}"
    WriteToMondayAPI postData, responseStatus, responseText
    
    Set jsonObject = ParseJson(responseText)
    'If jsonObject.Exists("error_code") = True Then:
    '   responseText = jsonObject.Item("error_code") & " : " & jsonObject.Item("error_message")

    Set jsonObject = ParseJson(responseText)
    AddTag = jsonObject("data")("create_or_get_tag")("id")
    
End Function
Public Sub UpdateStatusMonday(board_id As String, itemid As String, newStatus As String, ByRef responseStatus As String, ByRef responseText As String)
Dim DDQ As String, postData As String
    DDQ = Chr(34)
    postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "mutation { change_simple_column_value(board_id : " & board_id & ", item_id :" & itemid & ", column_id: \" & DDQ & "status" & "\" & DDQ & ", value: \" & DDQ & _
            newStatus & "\" & DDQ & ") {id}}" & DDQ & "}"

    Debug.Print postData

    WriteToMondayAPI postData, responseStatus, responseText
End Sub

Public Sub UpdateStatusMondayMultiVal(board_id As String, itemid As String, newStatus As String, ByRef responseStatus As String, ByRef responseText As String)
Dim DDQ As String, postData As String
    DDQ = Chr(34)

    
    status = "\" & DDQ & "{" & "\\\" & DDQ & "status" & "\\\" & DDQ & ":{" & "\\\" & DDQ & "label" & "\\\" & DDQ & ":" & "\\\" & DDQ & newStatus & "\\\" & DDQ & "}}" & "\" & DDQ
    postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "mutation { change_multiple_column_values(board_id: " & board_id & ",item_id: " & itemid & ",column_values: " & status & "){id}}" & DDQ & "}"
    
    Debug.Print postData
    WriteToMondayAPI postData, responseStatus, responseText

 'mutation {
 ' change_multiple_column_values(
 '   item_id:5764926352,
 '   board_id:4977328922,
 '   column_values: "{\"status\":{\"label\" : \"Working\"}}")

End Sub


Public Sub UpdateTagsMonday(board_id As String, itemid As String, newTags As String, ByRef responseStatus As String, ByRef responseText As String)
Dim DDQ As String, postData As String

'mutation { change_multiple_column_values(item_id:2951573079, board_id:2193345626, column_values: "{\"tags\" : {\"tag_ids\" : [10165564,10166350]}}") {ID}}
'mutation { change_multiple_column_values(board_id : 2193345626, item_id :2951573079, column_values: "{\"tags\": {\"tag_ids\":[10165564,10166350]}}") {id}}
 
    DDQ = Chr(34)
    postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "mutation { change_multiple_column_values(board_id : " & board_id & ", item_id :" & itemid & ", column_values: " _
            & DDQ & "{\" & DDQ & "tags" & "\" & DDQ & ": " & "{" & "\" & DDQ & "tag_ids" & "\" & DDQ & ":[" & newTags & "]}}" & DDQ & ") {id}}" & DDQ & "}"

    Debug.Print postData
    
    WriteToMondayAPI postData, responseStatus, responseText
End Sub

Public Sub UpdateItemAttributeMonday(board_id As String, itemid As String, itemType As String, newItemValue As String, ByRef responseStatus As String, ByRef responseText As String)
Dim DDQ As String, postData As String
    DDQ = Chr(34)
    postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "mutation { change_simple_column_value(board_id : " & board_id & ", item_id :" & itemid & ", column_id: \" & DDQ & itemType & "\" & DDQ & ", value: \" & DDQ & _
            newItemValue & "\" & DDQ & ") {id}}" & DDQ & "}"

    Debug.Print postData
    
    WriteToMondayAPI postData, responseStatus, responseText
End Sub


Public Sub PostUpdateMonday(itemid As String, msg As String, ByRef responseStatus As String, ByRef responseText As String)
Dim DDQ As String, postData As String
    DDQ = Chr(34)
    postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "mutation { create_update (item_id: " & itemid & ", body: " & "\" & DDQ & msg & "\" & DDQ & ") {id}}" & DDQ & "}"
    
    WriteToMondayAPI postData, responseStatus, responseText
End Sub


Public Function GetBoardId(itemid As String, ByRef responseStatus As String, ByRef responseText As String) As Variant
Dim DDQ As String, postData As String
Dim jsonObject As Object
    DDQ = Chr(34)

    postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "query {items (ids: [ " & itemid & "]) { board { id } } }" & DDQ & "}"
    WriteToMondayAPI postData, responseStatus, responseText
    
    Set jsonObject = ParseJson(responseText)
    GetBoardId = jsonObject("data")("items")(1)("board")("id")
        
End Function

Public Function GetBoards(ByRef responseStatus As String, ByRef responseText As String) As Collection
'query { boards (ids: 157244624) { groups (ids: status) { Title Color Position } }}
'query { boards () {name State board_folder_id items {ID name column_values { Text } } } }

Dim DDQ As String, postData As String
Dim jsonObject As Object
Dim boardsObjects As Collection
Dim itemDict As Variant

    DDQ = Chr(34)

    postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "query {boards (limit:100) { id name permissions state} }" & DDQ & "}"
    WriteToMondayAPI postData, responseStatus, responseText
    
    Set jsonObject = ParseJson(responseText)
    Set GetBoards = jsonObject("data")("boards")

End Function


Public Function GetUsers(ByRef responseStatus As String, ByRef responseText As String) As Collection
Dim DDQ As String, postData As String
Dim jsonObject As Object
Dim boardsObjects As Collection
Dim itemDict As Variant

    DDQ = Chr(34)

    postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "query {users () { id name } }" & DDQ & "}"
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


Public Sub DeleteItem(itemid As String, ByRef responseStatus As String, ByRef responseText As String)
Dim DDQ As String, postData As String
'mutation { delete_item (item_id: 12345678) {ID}}

    DDQ = Chr(34)

    postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "mutation { delete_item (item_id : " & itemid & ") {id}}" & DDQ & "}"
    Debug.Print postData
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

Public Sub CreateMondayItem(boardid As String, groupid As String, itemName As String, statusenum As String, owner As String, newTags As String, _
        ByRef responseStatus As String, ByRef responseText As String, Optional peopleFieldName As String = "people")

        
Dim DDQ As String, postData As String, itemid As String
Dim jsonObject As Object
Dim boardsObjects As Collection
Dim itemDict As Variant


    ' people8 for test
    
    DDQ = Chr(34)
    
    tags = "\\\" & DDQ & "tags" & "\\\" & DDQ & ": " & "{" & "\\\" & DDQ & "tag_ids" & "\\\" & DDQ & ":[" & newTags & "]}"
    
    person = "\\\" & DDQ & peopleFieldName & "\\\" & DDQ & ": {" & "\\\" & DDQ & "personsAndTeams" & "\\\" & DDQ & ":[{" & "\\\" & DDQ & "id" & "\\\" & DDQ & ":" & owner & "," & "\\\" & DDQ & "kind" & "\\\" & DDQ & ":" & "\\\" & DDQ & "person" & "\\\" & DDQ & "}]}"
    'person = "\\\" & DDQ & peopleFieldName & "\\\" & DDQ & ": {" & "\\\" & DDQ & "personsAndTeams" & "\\\" & DDQ & ":[{" & "\\\" & DDQ & "id" & "\\\" & DDQ & ":22027695" & "," & "\\\" & DDQ & "kind" & "\\\" & DDQ & ":" & "\\\" & DDQ & "person" & "\\\" & DDQ & "}]}"
    
    status = "\\\" & DDQ & "status" & "\\\" & DDQ & ":" & "\\\" & DDQ & statusenum & "\\\"
    
    column_values = "column_values:" & "\" & DDQ & "{" & status & DDQ & "," & person & "," & tags & "}" & "\" & DDQ
    'column_values = "column_values:" & "\" & DDQ & "{" & "\\\" & DDQ & "status" & "\\\" & DDQ & ":" & "\\\" & DDQ & statusenum & "\\\" & DDQ & "," & person & "," & tags & "}" & "\" & DDQ
    
    postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "mutation {create_item (board_id: " & boardid & ",group_id:" & "\" & DDQ & groupid & "\" & DDQ & ",item_name:" & "\" & DDQ & itemName & "\" & DDQ & "," _
                & column_values & ") { id } } " & DDQ & "}"
                
    'Debug.Print postData
     
    WriteToMondayAPI postData, responseStatus, responseText
    
    'Set jsonObject = ParseJson(responseText)
    'itemid = jsonObject("data")("create_item")("id")
    
    
    
End Sub

Public Sub CreateMondaySubItem(parentItemId As String, itemName As String, statusenum As String, owner As String, newTags As String, _
    ByRef responseStatus As String, ByRef responseText As String, Optional peopleFieldName As String = "people")

Dim DDQ As String, postData As String
Dim jsonObject As Object
Dim boardsObjects As Collection
Dim itemDict As Variant

    ' people 6 for test
    DDQ = Chr(34)

    tags = "\\\" & DDQ & "tags" & "\\\" & DDQ & ": " & "{" & "\\\" & DDQ & "tag_ids" & "\\\" & DDQ & ":[" & newTags & "]}"
    
    'person = "\\\" & DDQ & peopleFieldName & "\\\" & DDQ & ": {" & "\\\" & DDQ & "personsAndTeams" & "\\\" & DDQ & ":[{" & "\\\" & DDQ & "id" & "\\\" & DDQ & ":22027695" & "," & "\\\" & DDQ & "kind" & "\\\" & DDQ & ":" & "\\\" & DDQ & "person" & "\\\" & DDQ & "}]}"
    person = "\\\" & DDQ & peopleFieldName & "\\\" & DDQ & ": {" & "\\\" & DDQ & "personsAndTeams" & "\\\" & DDQ & ":[{" & "\\\" & DDQ & "id" & "\\\" & DDQ & ":" & owner & "," & "\\\" & DDQ & "kind" & "\\\" & DDQ & ":" & "\\\" & DDQ & "person" & "\\\" & DDQ & "}]}"
    
    column_values = "column_values:" & "\" & DDQ & "{" & "\\\" & DDQ & "status" & "\\\" & DDQ & ":" & "\\\" & DDQ & statusenum & "\\\" & DDQ & "," & person & "," & tags & "}" & "\" & DDQ
    
    'postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "mutation {create_subitem (parent_item_id: " & parentItemId & ",item_name:" & "\" & DDQ & itemName & "\" & DDQ & "," _
    '            & column_values & ") { id } } " & DDQ & "}"
    postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "mutation {create_subitem (parent_item_id: " & parentItemId & ",item_name:" & "\" & DDQ & itemName & "\" & DDQ & "," _
                & column_values & ") {id board {id}} } " & DDQ & "}"
                
            
            
    WriteToMondayAPI postData, responseStatus, responseText
    
    
End Sub




Public Sub WriteToMondayAPI(postData As String, ByRef responseStatus As String, ByRef responseText As String)
Dim objHTTP As Object
Dim DDQ As String, apiKey As String, Url As String
    apiKey = "eyJhbGciOiJIUzI1NiJ9.eyJ0aWQiOjExMTgyMDkwNiwidWlkIjoxNTE2MzEwNywiaWFkIjoiMjAyMS0wNS0zMFQxMTowMDo1OS4wMDBaIiwicGVyIjoibWU6d3JpdGUiLCJhY3RpZCI6NjY5MDk4NSwicmduIjoidXNlMSJ9.zIeOeoqeaZ2Q8NuKBPPw2LQFh2JRPvPwIkhhn4e5Q08"
    Url = "https://api.monday.com/v2"
    DDQ = Chr(34)
    Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    objHTTP.Open "POST", Url, False
    objHTTP.setRequestHeader "Authorization", apiKey
    objHTTP.setRequestHeader "Content-Type", "application/json"
    objHTTP.send postData
    
    responseStatus = objHTTP.status
    responseText = objHTTP.responseText
    
    'myParseJson objHTTP.responseText
    
End Sub


