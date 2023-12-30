Attribute VB_Name = "HTTPUtils"
'Public Function HTTPDownloadFile(url As String, tmpWorkbook As Workbook, _
'                     sheetNamePrefix As String, _
'                     configSheetName As String, _
'                     Optional startRow As Integer = 0, _
'                     Optional fileType As String = "start-of-day", _
'                     Optional newSheetName As String = "test", _
'                     Optional deleteSheet As Boolean = True, _
'                     Optional startRangeRow As Integer = 0) As Range

Sub GetRecordExec(lookupRecordId As String, recordType As String, ByRef rt As String, ByRef rs As String)
Dim recordDict As Dictionary, fieldDict As Dictionary
Dim fieldsCell As Range, valuesCell As Range, responseStatusRange As Range, responseTextRange As Range, postDataRange As Range
Dim key As Variant, value As Variant
Dim lookupClientId As String
Dim fieldsKey As Variant
Dim fieldColl As Collection
Dim fieldsDict As Dictionary

    Set tmpWorksheet = ActiveWorkbook.Sheets("Sheet2")
    Set fieldsCell = tmpWorksheet.Range("FIELDS").Rows(1)
    Set valuesCell = tmpWorksheet.Range("VALUES").Rows(1)
    Set responseStatusRange = tmpWorksheet.Range("RESPONSE_STATUS")
    Set responseTextRange = tmpWorksheet.Range("RESPONSE_TEXT")
    Set postDataRange = tmpWorksheet.Range("POST_DATA")
    'Set recordType = tmpWorksheet.Range("RECORD_TYPE")
    
    If recordType = "person" Then
        'Set lookupRecordId = tmpWorksheet.Range("LOOKUP_RECORD_ID")
        'Set recordDict = GetCapsuleRecord("https://api.capsulecrm.com/api/v2/parties/", lookupRecordId.value, rs, rt)
        Set recordDict = GetCapsuleRecord("https://api.capsulecrm.com/api/v2/parties/", lookupRecordId, rs, rt)
    ElseIf recordType = "organisation" Then
        'Set lookupRecordId = tmpWorksheet.Range("LOOKUP_CLIENT_ID")
        'Set recordDict = GetCapsuleRecord("https://api.capsulecrm.com/api/v2/parties/", lookupRecordId.value, rs, rt)
        Set recordDict = GetCapsuleRecord("https://api.capsulecrm.com/api/v2/parties/", lookupRecordId, rs, rt)
    ElseIf recordType = "opportunity" Then
        'Set lookupRecordId = tmpWorksheet.Range("LOOKUP_OPPORTUNITY_ID")
        'Set recordDict = GetCapsuleRecord("https://api.capsulecrm.com/api/v2/opportunities/", lookupRecordId.value, rs, rt)
        Set recordDict = GetCapsuleRecord("https://api.capsulecrm.com/api/v2/opportunities/", lookupRecordId, rs, rt, "opportunity")
        'Set recordDict = GetCapsuleRecord("https://api.capsulecrm.com/api/v2/opportunities/", lookupRecordId.value, rs, rt, "opportunity")
    End If
    
    For Each key In recordDict.Keys
        
        If key = "fields" Then
            Set fieldColl = recordDict(key)
            For Each fieldsDict In fieldColl
                fieldsCell.value = fieldsDict("definition")("name")
                valuesCell.value = fieldsDict("value")
                Set fieldsCell = fieldsCell.Offset(1)
                Set valuesCell = valuesCell.Offset(1)
            Next fieldsDict
        Else
            fieldsCell.value = key
        
            If TypeName(recordDict(key)) = "Dictionary" Then
                Set fieldDict = recordDict(key)
                value = fieldDict.Item("name")
            ElseIf TypeName(recordDict(key)) = "Collection" Then
            Else
                value = recordDict(key)
            End If
            valuesCell.value = value
            Set fieldsCell = fieldsCell.Offset(1)
            Set valuesCell = valuesCell.Offset(1)
        End If
    Next key
    
    'responseStatusRange.value = rt
    'responseTextRange.value = rs

exitsub:
    Set tmpWorksheet = Nothing
    Set fieldsCell = Nothing
    Set valuesCell = Nothing
    Set responseStatusRange = Nothing
    Set responseTextRange = Nothing
    Set postDataRange = Nothing
    'Set lookupRecordId = Nothing
    Set recordDict = Nothing
    Set fieldsCell = Nothing
    Set valuesCell = Nothing
    Set fieldColl = Nothing

End Sub

Sub UpdateRecordExec(lookupRecordId As String, recordType As String, ByRef rt As String, ByRef rs As String)
Dim recordDict As Dictionary, fieldDict As Dictionary
Dim fieldsCell As Range, valuesCell As Range, responseStatusRange As Range, responseTextRange As Range, postDataRange As Range
Dim url As String
Dim key As Variant, value As Variant
Dim objHTTP As Object
Dim tmpWorksheet As Worksheet

    Set tmpWorksheet = ActiveWorkbook.Sheets("Sheet2")
    
    Set responseStatusRange = tmpWorksheet.Range("RESPONSE_STATUS")
    Set responseTextRange = tmpWorksheet.Range("RESPONSE_TEXT")
    Set postDataRange = tmpWorksheet.Range("POST_DATA")
    Set lookupRecordIDRange = tmpWorksheet.Range("LOOKUP_RECORD_ID")
    
    'Set objHTTP = UpdateCapsuleRecordField("https://api.capsulecrm.com/api/v2/parties/", lookupRecordIDRange.value)
    
    url = GetUrl(recordType, "update", lookupRecordId)
    
    'Set objHTTP = UpdateCapsuleRecordField("https://api.capsulecrm.com/api/v2/parties/", lookupRecordId)
    Set objHTTP = UpdateCapsuleRecordField(url)
    
    Set jsonObject = ParseJson(objHTTP.responseText)
    
    responseStatusRange.value = objHTTP.Status
    rs = objHTTP.Status
    rt = objHTTP.responseText
    
exitsub:

    Set responseStatusRange = Nothing
    Set responseTextRange = Nothing
    Set postDataRange = Nothing
    Set objHTTP = Nothing
    Set jsonObject = Nothing
    
End Sub

'Function UpdateCapsuleRecordField(url As String, id As String) As Object
Function UpdateCapsuleRecordField(url As String) As Object
Dim resultFieldsRange As Range, resultValuesRange As Range, customfieldFieldsRange As Range, customfieldValuesRange As Range, customFieldsRange As Range, lookupsRange As Range, postData As Range, _
    responseStatusRange As Range, responseTextRange As Range, postDataRange As Range, optionFieldRange As Range, updatesRange As Range, fieldsRange As Range, valuesRange As Range, _
    nestedFieldsRange As Range, lookupRecordIDRange As Range
Dim tmpWorksheet As Worksheet
Dim datastr As String, responseStatus As String, responseText As String
Dim header As New Dictionary
Dim subdata As New Dictionary
Dim data As New Dictionary


    Set tmpWorksheet = ActiveWorkbook.Sheets("Sheet2")
    Set urlRange = tmpWorksheet.Range("URL")
    Set accessCodeRange = tmpWorksheet.Range("ACCESS_CODE")
    Set updatesRange = tmpWorksheet.Range("UPDATES")
    Set entityRange = tmpWorksheet.Range("ENTITY")
    Set fieldsRange = tmpWorksheet.Range("FIELDS")
    Set customFieldsRange = tmpWorksheet.Range("CUSTOMFIELD")
    Set valuesRange = tmpWorksheet.Range("VALUES")
    Set resultFieldsRange = tmpWorksheet.Range("RESULT_FIELDS")
    Set resultValuesRange = tmpWorksheet.Range("RESULT_VALUES")
    Set lookupsRange = tmpWorksheet.Range("LOOKUPS")
    Set responseStatusRange = tmpWorksheet.Range("RESPONSE_STATUS")
    Set responseTextRange = tmpWorksheet.Range("RESPONSE_TEXT")
    Set postDataRange = tmpWorksheet.Range("POST_DATA")
    Set optionFieldRange = tmpWorksheet.Range("OPTION_FIELD")

    

    For i = 1 To updatesRange.Count
        If updatesRange.Rows(i).value <> "" And customFieldsRange.Rows(i).value = False And lookupsRange.Rows(i) = False And fieldsRange.Rows(i).value <> "" Then
            If optionFieldRange.Rows(i) = "FALSE" Then
                MsgBox "invalid value " & fieldsRange.Rows(i).value & ":" & valuesRange.Rows(i).value
                GoTo exitsub
            ElseIf fieldsRange.Rows(i).value = "value" Then
                subdata.Add "amount", valuesRange.Rows(i).value
                subdata.Add "currency", "USD"
            Else
                subdata.Add fieldsRange.Rows(i).value, updatesRange.Rows(i).value
            End If
        End If
    Next i
    
    Set customFields = GetCustomFields(tmpWorksheet, fieldsRange, valuesRange)
    'Set customFields = GetCustomFields(tmpWorksheet, fieldsRange, valuesRange, customFieldsRange, optionFieldRange, updatesRange)
    subdata.Add "fields", customFields
    
    'GetLookupFields tmpWorksheet, fieldsRange, valuesRange, lookupsRange, subdata
    
    data.Add entityRange.value, subdata
    
    header.Add "Authorization", "Bearer " & accessCodeRange.value
    header.Add "Content-Type", "application/json"
    header.Add "Accept", "application/json"

    datastr = JsonConverter.ConvertToJson(data)
    postDataRange.value = datastr
    'url = url & id
    
    Set UpdateCapsuleRecordField = WriteToRESTAPI(datastr, url, "4rY0P12jmfi0iq41S0mtuDUi4yKEfnLefH260Ufkgnb8fE33xfdt/fb2dsqGeev7", header, "put", responseStatus, responseText)

exitsub:
    Set tmpWorksheet = Nothing
    Set accessCodeRange = Nothing
    Set updatesRange = Nothing
    Set entityRange = Nothing
    Set fieldsRange = Nothing
    Set customFieldsRange = Nothing
    Set valuesRange = Nothing
    Set resultFieldsRange = Nothing
    Set resultValuesRange = Nothing
    
    Set lookupsRange = Nothing
    Set responseStatusRange = Nothing
    Set responseTextRange = Nothing
    Set postDataRange = Nothing
    Set optionFieldRange = Nothing
    Set customFields = Nothing
    
End Function

Function GetCapsuleRecord(url As String, id As String, ByRef rt As String, ByRef rs As String, Optional entity As String = "party") As Dictionary
Dim access_code As String, responseStatus As String, responseText As String
Dim header As New Dictionary

    url = url & id & "?embed=fields"
    
    header.Add "Authorization", "Bearer " & "4rY0P12jmfi0iq41S0mtuDUi4yKEfnLefH260Ufkgnb8fE33xfdt/fb2dsqGeev7"
    header.Add "Content-Type", "application/json"
    header.Add "Accept", "application/json"
    
    Set objHTTP = GetFromRESTAPI(url, _
        "4rY0P12jmfi0iq41S0mtuDUi4yKEfnLefH260Ufkgnb8fE33xfdt/fb2dsqGeev7", header, responseStatus, responseText)
        
    Set jsonObject = ParseJson(objHTTP.responseText)
    
    'Set GetCapsuleRecord = jsonObject.Item("party")
    Set GetCapsuleRecord = jsonObject.Item(entity)
    
    
    rs = objHTTP.Status
    rt = objHTTP.responseText
        
End Function

Sub GetCapsuleDefinitions()
Dim optionHeaderRange As Range, optionHeaderCell As Range

    Set optionHeaderRange = ActiveSheet.Range("OPTIONS_HEADER")
    Set optionHeaderCell = optionHeaderRange.Cells(1, 1)
    
    Set optionHeaderCell = GetCapsuleDefinition("https://api.capsulecrm.com/api/v2/opportunities/fields/definitions", optionHeaderCell)
    GetCapsuleDefinition "https://api.capsulecrm.com/api/v2/parties/fields/definitions", optionHeaderCell
    

End Sub
Function GetCapsuleDefinition(url As String, optionHeaderCell As Range) As Range
Dim access_code As String, responseStatus As String, responseText As String, defnName As String, optionRangeName As String
Dim header As New Dictionary
Dim defnColl As Collection
Dim defnDict As Dictionary
Dim optionsColl As Collection
Dim optionValueCell As Range, optionValuesRange As Range



    header.Add "Authorization", "Bearer " & "4rY0P12jmfi0iq41S0mtuDUi4yKEfnLefH260Ufkgnb8fE33xfdt/fb2dsqGeev7"
    header.Add "Content-Type", "application/json"
    header.Add "Accept", "application/json"
    
    Set objHTTP = GetFromRESTAPI(url, _
        "4rY0P12jmfi0iq41S0mtuDUi4yKEfnLefH260Ufkgnb8fE33xfdt/fb2dsqGeev7", header, responseStatus, responseText)
        
    Set jsonObject = ParseJson(objHTTP.responseText)
    
    Set defnColl = jsonObject.Item("definitions")
    
    For i = 1 To defnColl.Count
        Set optionsColl = New Collection
        Set defnDict = defnColl(i)
        On Error Resume Next
        Set optionsColl = defnDict("options")
        On Error GoTo 0

        If optionsColl.Count > 0 Then
        
            optionHeaderCell.value = defnDict("name")
            optionRangeName = UCase(Replace(defnDict("name"), " ", "_") & "_OPTIONS")
            Set optionValueCell = optionHeaderCell.Offset(1)
            Set optionHeaderCell = optionHeaderCell.Offset(, 1)
        
            Set optionValuesRange = optionValueCell.Resize(optionsColl.Count)
            ActiveWorkbook.Names.Add optionRangeName, RefersTo:=optionValuesRange
            For j = 1 To optionsColl.Count
                On Error Resume Next
                optionValueCell.value = optionsColl(j)
                Set optionValueCell = optionValueCell.Offset(1)
                On Error GoTo 0
            Next j
        End If
    Next i
    
    Set GetCapsuleDefinition = optionHeaderCell.Offset(, 1)
    
    
End Function


Public Function GetFromRESTAPI(url As String, apiKey As String, header As Dictionary, _
        ByRef responseStatus As String, ByRef responseText As String) As Object
Dim objHTTP As Object
    
    Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    objHTTP.Open "get", url, False

    For Each key In header.Keys
        value = header.Item(key)
        objHTTP.setRequestHeader key, value
    Next key
    
    objHTTP.send
    
    Set GetFromRESTAPI = objHTTP

End Function


Function GetUrl(entity As String, action As String, id As String) As String
Dim url As String

    url = "https://api.capsulecrm.com/api/v2/"
    If entity = "opportunity" And action = "delete" Then
        url = url & "opportunities" & "/" & id
    ElseIf entity = "party" And action = "delete" Then
        url = url & "parties" & "/" & id
    ElseIf entity = "person" And action = "delete" Then
        url = url & "parties" & "/" & id
    ElseIf entity = "entry" And action = "delete" Then
        url = url & "entries" & "/" & id
    ElseIf entity = "person" And action = "update" Then
        url = url & "parties" & "/" & id
    ElseIf entity = "opportunity" And action = "update" Then
        url = url & "opportunities" & "/" & id
    End If
    
endsub:
    GetUrl = url
End Function
Sub DeleteRecordExec(lookupRecordId As String, recordType As String, ByRef responseStatus As String, ByRef responseText As String)
Dim header As New Dictionary
Dim objHTTP As Object
Dim url As String, access_code As String
Dim responseStatusRange As Range, responseTextRange As Range, postDataRange As Range
Dim currentTypeRange As Range
Dim currentId As Range

    Set tmpWorksheet = ActiveSheet
    
    Set responseStatusRange = tmpWorksheet.Range("RESPONSE_STATUS")
    Set responseTextRange = tmpWorksheet.Range("RESPONSE_TEXT")
    Set postDataRange = tmpWorksheet.Range("POST_DATA")
    
    'Set currentTypeRange = tmpWorksheet.Range("CURRENT_TYPE")
    'Set currentId = tmpWorksheet.Range("CURRENT_ID")
        
    'url = GetUrl(currentTypeRange.value, "delete", currentId.value)
    url = GetUrl(recordType, "delete", lookupRecordId)
    
    
    access_code = "4rY0P12jmfi0iq41S0mtuDUi4yKEfnLefH260Ufkgnb8fE33xfdt/fb2dsqGeev7"
    
    header.Add "Authorization", "Bearer " & access_code
    header.Add "Content-Type", "application/json"
    header.Add "Accept", "application/json"

    Set objHTTP = WriteToRESTAPI("", url, access_code, header, "DELETE", responseStatus, responseText)
    
    responseStatus = objHTTP.Status
    responseText = objHTTP.responseText
    datastr = datastr

exitsub:
    Set tmpWorksheet = Nothing
    Set responseStatusRange = Nothing
    Set responseTextRange = Nothing
    Set postDataRange = Nothing
    
    Set currentTypeRange = Nothing
    Set currentId = Nothing
    Set objHTTP = Nothing

End Sub


Sub SetupRecordDefaults(Optional recordType As String)
Dim recordFormRange As Range, recordDefaultsRange As Range, recordTypeRange As Range
Dim recordDefaultsTypeRange As Range, recordsDefaultsDefaultRange As Range, recordFormDefaultRange As Range, recordFormFieldNameRange As Range, _
    recordsDefaultsFieldNameRange As Range
Dim tmpWorksheet As Worksheet, refWorksheet As Worksheet
Dim fieldCount As Integer

    Set tmpWorksheet = ActiveWorkbook.Sheets("Sheet2")
    Set refWorksheet = ActiveWorkbook.Sheets("Reference")

    Set recordFormRange = tmpWorksheet.Range("RECORD_FORM")
    Set recordDefaultsRange = refWorksheet.Range("RECORD_DEFAULTS")
    If recordType <> "" Then
        'Set recordTypeRange.value = recordType
    Else
        recordType = tmpWorksheet.Range("RECORD_TYPE")
        'Set recordTypeRange = tmpWorksheet.Range("RECORD_TYPE")
    End If
    
    Set recordDefaultsTypeRange = recordDefaultsRange.Columns(1)
    Set recordsDefaultsDefaultRange = recordDefaultsRange.Columns(4)
    Set recordsDefaultsFieldNameRange = recordDefaultsRange.Columns(2)
    Set recordFormDefaultRange = recordFormRange.Columns(2)
    Set recordFormFieldNameRange = recordFormRange.Columns(1)
    
    fieldCount = 1
    
    recordFormRange.ClearContents
    
    For i = 1 To recordDefaultsRange.Rows.Count
        'If recordDefaultsTypeRange.Rows(i).value = recordTypeRange.value Then
        If recordDefaultsTypeRange.Rows(i).value = recordType Then
            recordFormDefaultRange.Rows(fieldCount).value = recordsDefaultsDefaultRange.Rows(i).value
            recordFormFieldNameRange.Rows(fieldCount).value = recordsDefaultsFieldNameRange.Rows(i).value
            fieldCount = fieldCount + 1
        End If
    
    Next i

exitsub:
    Set tmpWorksheet = Nothing
    Set refWorksheet = Nothing
    Set recordFormRange = Nothing
    Set recordDefaultsRange = Nothing
    Set recordTypeRange = Nothing
    Set recordDefaultsTypeRange = Nothing
    Set recordsDefaultsDefaultRange = Nothing
    Set recordsDefaultsFieldNameRange = Nothing
    Set recordFormDefaultRange = Nothing
    Set recordFormFieldNameRange = Nothing
End Sub
Sub GetLookupFields(tmpWorksheet As Worksheet, fieldsRange As Range, valuesRange As Range, lookupsRange As Range, ByRef data As Dictionary)
Dim lookupFieldsRange As Range, lookupValuesRange As Range
Dim lookupValue As Variant
'Dim customFields As New Collection
Dim lookupField As New Dictionary
    Set lookupFieldsRange = tmpWorksheet.Range("LOOKUP_FIELDS")
    Set lookupValuesRange = tmpWorksheet.Range("LOOKUP_VALUES")
    
    For i = 1 To fieldsRange.Count
        Set customField = New Dictionary
        lookupValue = lookupsRange.Rows(i).value
        FieldValue = fieldsRange.Rows(i)
        If lookupsRange.Rows(i).value <> False Then
            customField.Add "id", lookupsRange.Rows(i).value
            data.Add fieldsRange.Rows(i), customField
        End If
    Next i
        
End Sub
Function GetCustomFields(tmpWorksheet As Worksheet, fieldsRange As Range, valuesRange As Range) As Collection
Dim customfieldFieldsRange As Range, customfieldValuesRange As Range, nestedFieldsRange As Range
Dim customFields As New Collection
Dim subCustomField As New Dictionary, customField As New Dictionary
    Set customfieldFieldsRange = ActiveWorkbook.Sheets("Reference").Range("CUSTOMFIELD_FIELDS")
    Set customfieldValuesRange = ActiveWorkbook.Sheets("Reference").Range("CUSTOMFIELD_VALUES")
    Set customFieldsRange = ActiveWorkbook.Sheets("Sheet2").Range("CUSTOMFIELD")
    Set optionFieldRange = ActiveWorkbook.Sheets("Sheet2").Range("OPTION_FIELD")
    Set updatesRange = ActiveWorkbook.Sheets("Sheet2").Range("UPDATES")

    For i = 1 To fieldsRange.Count
        
        If customFieldsRange.Rows(i).value <> False And customFieldsRange.Rows(i).value <> "" Then
            If updatesRange Is Nothing Then
                If optionFieldRange.Rows(i) = False Then
                    MsgBox "invalid value " & fieldsRange.Rows(i).value & ":" & valuesRange.Rows(i).value
                    GoTo exitsub
                End If
                Set subCustomField = New Dictionary
                Set customField = New Dictionary
                subCustomField.Add "name", fieldsRange.Rows(i).value
                subCustomField.Add "id", customFieldsRange.Rows(i).value
                customField.Add "definition", subCustomField
                customField.Add "value", valuesRange.Rows(i).value
                customFields.Add customField
            Else
                If updatesRange.Rows(i).value <> "" Then
                    If optionFieldRange.Rows(i) = False Then
                        MsgBox "invalid value " & fieldsRange.Rows(i).value & ":" & updatesRange.Rows(i).value
                        GoTo exitsub
                    End If
                        
                    Set subCustomField = New Dictionary
                    Set customField = New Dictionary
                    subCustomField.Add "name", fieldsRange.Rows(i).value
                    subCustomField.Add "id", customFieldsRange.Rows(i).value
                    customField.Add "definition", subCustomField
                    customField.Add "value", updatesRange.Rows(i).value
                    customFields.Add customField

                End If
            End If
        End If
    
        Debug.Print customFieldsRange.Rows(i).Address
        Debug.Print customFieldsRange.Rows(i).value
        Next i
    
    Set GetCustomFields = customFields
    
exitsub:
    Set customfieldFieldsRange = Nothing
    Set customfieldValuesRange = Nothing
    Set subCustomField = Nothing
    Set customField = Nothing
End Function
Sub AddRecordExec(ByRef responseStatus As String, ByRef responseText As String, ByRef datastr As String)
Dim tmpWorksheet As Worksheet, refWorksheet As Worksheet
Dim urlRange As Range, accessCodeRange As Range, _
            fieldsRange As Range, valuesRange As Range, entityRange As Range
Dim field As Variant, key As Variant
Dim data As New Dictionary
Dim subdata As New Dictionary
Dim header As New Dictionary, listDict As New Dictionary
Dim listColl As New Collection
Dim i As Integer
Dim objHTTP As Object, jsonObject As Object
Dim capsuleRecord As Dictionary, capsuleSubRecord As Dictionary, tmpRecord As Dictionary
Dim outputFieldCount As Integer
Dim customFields As Collection
Dim resultFieldsRange As Range, resultValuesRange As Range, customfieldFieldsRange As Range, customfieldValuesRange As Range, customFieldsRange As Range, lookupsRange As Range, postData As Range, _
    responseStatusRange As Range, responseTextRange As Range, postDataRange As Range, optionFieldRange As Range

    Set tmpWorksheet = ActiveSheet
    Set refWorksheet = ActiveWorkbook.Sheets("Reference")

    Set urlRange = tmpWorksheet.Range("URL")
    Set accessCodeRange = tmpWorksheet.Range("ACCESS_CODE")
    Set entityRange = tmpWorksheet.Range("ENTITY")
    Set fieldsRange = tmpWorksheet.Range("FIELDS")
    Set customFieldsRange = tmpWorksheet.Range("CUSTOMFIELD")
    Set valuesRange = tmpWorksheet.Range("VALUES")
    Set resultFieldsRange = tmpWorksheet.Range("RESULT_FIELDS")
    Set resultValuesRange = tmpWorksheet.Range("RESULT_VALUES")
    Set lookupsRange = tmpWorksheet.Range("LOOKUPS")
    Set responseStatusRange = tmpWorksheet.Range("RESPONSE_STATUS")
    Set responseTextRange = tmpWorksheet.Range("RESPONSE_TEXT")
    Set postDataRange = tmpWorksheet.Range("POST_DATA")
    Set optionFieldRange = tmpWorksheet.Range("OPTION_FIELD")
    
    For i = 1 To fieldsRange.Count
        If customFieldsRange.Rows(i).value = False And lookupsRange.Rows(i) = False And fieldsRange.Rows(i).value <> "" Then
            If optionFieldRange.Rows(i) = "FALSE" Then
                MsgBox "invalid value " & fieldsRange.Rows(i).value & ":" & valuesRange.Rows(i).value
                GoTo exitsub
            ElseIf fieldsRange.Rows(i).value = "value" Then
                subdata.Add "amount", valuesRange.Rows(i).value
                subdata.Add "currency", "USD"
            Else
                If fieldsRange.Rows(i).value = "emailAddresses" Then
                
                    listDict.Add "type", "work"
                    listDict.Add "address", valuesRange.Rows(i).value
                    listColl.Add listDict
                    subdata.Add "emailAddresses", listColl
                    
                Else
                    subdata.Add fieldsRange.Rows(i).value, valuesRange.Rows(i).value
                End If
            End If
        End If
    Next i
    
    Set customFields = GetCustomFields(tmpWorksheet, fieldsRange, valuesRange, customFieldsRange, optionFieldRange)
    subdata.Add "fields", customFields
    
    GetLookupFields refWorksheet, fieldsRange, valuesRange, lookupsRange, subdata
    
    data.Add entityRange.value, subdata
    
    header.Add "Authorization", "Bearer " & accessCodeRange.value
    header.Add "Content-Type", "application/json"
    header.Add "Accept", "application/json"

    datastr = JsonConverter.ConvertToJson(data)
    
    Debug.Print datastr
    Set objHTTP = WriteToRESTAPI(datastr, urlRange.value, accessCodeRange.value, header, "POST", _
            responseStatus, responseText)
    
    Set jsonObject = ParseJson(objHTTP.responseText)
    
     
    If objHTTP.Status = "200" Or objHTTP.Status = "201" Then
        Set capsuleRecord = jsonObject.Item(entityRange.value)

        For Each Item In capsuleRecord.Keys
            If TypeName(capsuleRecord(Item)) = "Dictionary" Then
                Set capsuleSubRecord = capsuleRecord(Item)
                For Each subItem In capsuleSubRecord.Keys
                    outputStr = outputStr & subItem & "=" & capsuleSubRecord(subItem) & ","
                Next subItem
                'resultStr = Item
                resultFieldsRange.Rows(outputFieldCount).value = Item
                resultValuesRange.Rows(outputFieldCount).value = outputStr
            ElseIf TypeName(capsuleRecord(Item)) = "Collection" Then
                
            Else
                'resultStr = Item
                resultFieldsRange.Rows(outputFieldCount).value = Item
                resultValuesRange.Rows(outputFieldCount).value = capsuleRecord(Item)
                'outputStr = resultValuesRange.Rows(outputFieldCount).value
            End If
            outputFieldCount = outputFieldCount + 1
        Next Item
    End If
    
    responseStatus = objHTTP.Status
    responseText = objHTTP.responseText

    'postDataRange.value = datastr
    
exitsub:
    Set tmpWorksheet = Nothing
    Set urlRange = Nothing
    Set accessCodeRange = Nothing
    Set entityRange = Nothing
    Set fieldsRange = Nothing
    Set customFieldsRange = Nothing
    Set valuesRange = Nothing
    Set resultFieldsRange = Nothing
    Set resultValuesRange = Nothing
    Set lookupsRange = Nothing
    Set responseStatusRange = Nothing
    Set responseTextRange = Nothing
    Set postDataRange = Nothing
    Set optionFieldRange = Nothing
    Set capsuleRecord = Nothing
    Set jsonObject = Nothing
    Set objHTTP = Nothing
    Set customFields = Nothing
End Sub

Sub TestWriteToRESTAPI()
Dim data As New Dictionary
Dim subdata As New Dictionary
Dim header As New Dictionary
Dim url As String, access_code As String, datastr As String, responseStatus As String, responseText As String

    url = "https://api.capsulecrm.com/api/v2/parties"
    access_code = "4rY0P12jmfi0iq41S0mtuDUi4yKEfnLefH260Ufkgnb8fE33xfdt/fb2dsqGeev7"

    subdata.Add "firstName", "foo"
    subdata.Add "Contact Owner", "Jon"
    subdata.Add "jobTitle", "foobar"
    subdata.Add "lastName", "barfoo"
    subdata.Add "type", "person"
    subdata.Add "organisation", 188497262
    data.Add "party", subdata
    
    header.Add "Authorization", "Bearer " & access_code
    header.Add "Content-Type", "application/json"
    header.Add "Accept", "application/json"

    datastr = JsonConverter.ConvertToJson(data)
    

    WriteToRESTAPI datastr, url, access_code, header, "POST", responseStatus, responseText
    
End Sub

Public Function WriteToRESTAPI(postData As String, url As String, apiKey As String, header As Dictionary, _
                            restAction As String, _
                            ByRef responseStatus As String, ByRef responseText As String) As Object
Dim objHTTP As Object
Dim DDQ As String, value As String
    DDQ = Chr(34)
    
    Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    objHTTP.Open restAction, url, False
    
    For Each key In header.Keys
        value = header.Item(key)
        objHTTP.setRequestHeader key, value
    Next key

    objHTTP.send postData
    
    Set WriteToRESTAPI = objHTTP

    
End Function








Public Function HTTPDownloadFile(url As String, tmpWorkbook As Workbook, _
                     sheetNamePrefix As String, _
                     configSheetName As String, _
                     Optional startRow As Integer = 0, _
                     Optional fileType As String = "start-of-day", _
                     Optional newSheetName As String = "test", _
                     Optional deleteSheet As Boolean = True, _
                     Optional startRangeRow As Integer = 0) As Range
Dim tmpSheet As Worksheet
Dim tmpRange As Range, rowCountRange As Range, outputRange As Range
Dim fileLength As Long, rowWidth As Long, rowOffset As Long
Dim fileArray As Variant, lineArray As Variant
Dim objHTTP As Object
Dim rowCountRangeName As String
Dim origWorksheet As Worksheet

    Set origWorksheet = ActiveSheet

    On Error GoTo err
    If fileType <> "start-of-day" Then
        rowOffset = rowCountRange.value + startRangeRow
    Else
        rowOffset = startRangeRow
    End If
    
    'Application.ScreenUpdating = False
    'Application.EnableEvents = False
    'Application.Calculation = xlCalculationManual
    
    If fileType = "start-of-day" Then
        If deleteSheet = True Then
            On Error Resume Next
            Set tmpSheet = tmpWorkbook.Sheets(newSheetName)
            tmpSheet.Delete
            On Error GoTo 0
            
            Set tmpSheet = tmpWorkbook.Sheets.Add
            tmpSheet.Name = newSheetName
        Else
            Set tmpSheet = tmpWorkbook.Sheets(newSheetName)
            tmpSheet.Range("1:1048576").ClearContents
        End If
        
        
    Else
        Set tmpSheet = tmpWorkbook.Sheets(newSheetName)
        startRow = 1 ' dont need the headers
    End If
    
    Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    objHTTP.Open "GET", url, False
    objHTTP.setRequestHeader "Content-Type", "text/csv"
    objHTTP.send
    
    If objHTTP.Status = 404 Then
        Debug.Print objHTTP.StatusText
    Else
        fileArray = Split(objHTTP.responseText, Chr(10))
        fileLength = UBound(fileArray)
        
        For i = startRow To fileLength - 1
            'j = i - startRow
            lineArray = Split(fileArray(i), "^")
            rowWidth = UBound(lineArray) + 1
            
            If UBound(lineArray) > 0 Then
                Set tmpRange = tmpSheet.Rows(i + 1 + rowOffset).Resize(, rowWidth)
                tmpRange = lineArray
            End If
        Next i
    End If

    tmpSheet.Activate
    Set outputRange = tmpSheet.Range(Cells(1 + startRangeRow, 1), Cells(fileLength + startRangeRow, UBound(Split(fileArray(1), "^")) + 1))
    GoTo endsub
    
err:
    MsgBox "probably timedout"

endsub:
    origWorksheet.Activate
    Set HTTPDownloadFile = outputRange
    Set tmpWorkbook = Nothing
    Set tmpSheet = Nothing
    Set objHTTP = Nothing
    Set tmpRange = Nothing
    Set rowCountRange = Nothing
    Set origWorksheet = Nothing
    
    'Application.ScreenUpdating = True
    'Application.EnableEvents = True
    'Application.Calculation = xlCalculationAutomatic
End Function
