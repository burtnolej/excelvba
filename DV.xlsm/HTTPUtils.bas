Attribute VB_Name = "HTTPUtils"
'Public Function HTTPDownloadFile(url As String, tmpWorkbook As Workbook, _
'                     sheetNamePrefix As String, _
'                     configSheetName As String, _
'                     Optional startRow As Integer = 0, _
'                     Optional fileType As String = "start-of-day", _
'                     Optional newSheetName As String = "test", _
'                     Optional deleteSheet As Boolean = True, _
'                     Optional startRangeRow As Integer = 0) As Range


Function GetUrl(entity As String, action As String, id As Long) As String
Dim url As String

    url = "https://api.capsulecrm.com/api/v2/"
    If entity = "opportunity" And action = "delete" Then
        url = url & "opportunities" & "/" & CStr(id)
    ElseIf entity = "party" And action = "delete" Then
        url = url & "parties" & "/" & CStr(id)
    End If
        
endsub:
    GetUrl = url
End Function
Sub TestDeleteEntity()
Dim header As New Dictionary
Dim objHTTP As Object
Dim url As String, access_code As String, dataStr As String, responseStatus As String, responseText As String
Dim responseStatusRange As Range, responseTextRange As Range, postDataRange As Range
Dim currentTypeRange As Range
Dim currentId As Range

    Set tmpWorksheet = ActiveSheet
    
    Set responseStatusRange = tmpWorksheet.Range("RESPONSE_STATUS")
    Set responseTextRange = tmpWorksheet.Range("RESPONSE_TEXT")
    Set postDataRange = tmpWorksheet.Range("POST_DATA")
    
    Set currentTypeRange = tmpWorksheet.Range("CURRENT_TYPE")
    Set currentId = tmpWorksheet.Range("CURRENT_ID")
        
    url = GetUrl(currentTypeRange.value, "delete", currentId.value)
    
    access_code = "4rY0P12jmfi0iq41S0mtuDUi4yKEfnLefH260Ufkgnb8fE33xfdt/fb2dsqGeev7"
    
    header.Add "Authorization", "Bearer " & access_code
    header.Add "Content-Type", "application/json"
    header.Add "Accept", "application/json"

    Set objHTTP = WriteToRESTAPI("", url, access_code, header, "GET", responseStatus, responseText)
    
    responseStatusRange.value = objHTTP.Status
    responseTextRange.value = objHTTP.responseText
    postDataRange.value = dataStr

exitsub:
    Set tmpWorksheet = Nothing
    Set responseStatusRange = Nothing
    Set responseTextRange = Nothing
    Set postDataRange = Nothing
    
    Set currentTypeRange = Nothing
    Set currentId = Nothing
    Set objHTTP = Nothing

End Sub


Sub SetupRecordDefaults()
Dim recordFormRange As Range, recordDefaultsRange As Range, recordTypeRange As Range
Dim recordDefaultsTypeRange As Range, recordsDefaultsDefaultRange As Range, recordFormDefaultRange As Range, recordFormFieldNameRange As Range, _
    recordsDefaultsFieldNameRange As Range
Dim tmpWorksheet As Worksheet
Dim fieldCount As Integer

    Set tmpWorksheet = ActiveSheet

    Set recordFormRange = tmpWorksheet.Range("RECORD_FORM")
    Set recordDefaultsRange = tmpWorksheet.Range("RECORD_DEFAULTS")
    Set recordTypeRange = tmpWorksheet.Range("RECORD_TYPE")
    
    Set recordDefaultsTypeRange = recordDefaultsRange.Columns(1)
    Set recordsDefaultsDefaultRange = recordDefaultsRange.Columns(4)
    Set recordsDefaultsFieldNameRange = recordDefaultsRange.Columns(2)
    Set recordFormDefaultRange = recordFormRange.Columns(2)
    Set recordFormFieldNameRange = recordFormRange.Columns(1)
    
    fieldCount = 1
    
    recordFormRange.ClearContents
    
    For i = 1 To recordDefaultsRange.Rows.count
        If recordDefaultsTypeRange.Rows(i).value = recordTypeRange.value Then
            Debug.Print recordFormDefaultRange.Address
            Debug.Print recordsDefaultsDefaultRange.Rows(i).value
            recordFormDefaultRange.Rows(fieldCount).value = recordsDefaultsDefaultRange.Rows(i).value
            recordFormFieldNameRange.Rows(fieldCount).value = recordsDefaultsFieldNameRange.Rows(i).value
            fieldCount = fieldCount + 1
        End If
    
    Next i

End Sub
Sub GetLookupFields(tmpWorksheet As Worksheet, fieldsRange As Range, valuesRange As Range, lookupsRange As Range, ByRef data As Dictionary)
Dim lookupFieldsRange As Range, lookupValuesRange As Range
'Dim customFields As New Collection
Dim lookupField As New Dictionary
    Set lookupFieldsRange = tmpWorksheet.Range("LOOKUP_FIELDS")
    Set lookupValuesRange = tmpWorksheet.Range("LOOKUP_VALUES")
    
    For i = 1 To fieldsRange.count
        Set customField = New Dictionary
        If lookupsRange.Rows(i).value <> False Then
            customField.Add "id", lookupsRange.Rows(i).value
            data.Add fieldsRange.Rows(i), customField
        End If
    Next i
        
End Sub
Function GetCustomFields(tmpWorksheet As Worksheet, fieldsRange As Range, valuesRange As Range, customFieldsRange As Range) As Collection
Dim customfieldFieldsRange As Range, customfieldValuesRange As Range
Dim customFields As New Collection
Dim subCustomField As New Dictionary, customField As New Dictionary
    Set customfieldFieldsRange = tmpWorksheet.Range("CUSTOMFIELD_FIELDS")
    Set customfieldValuesRange = tmpWorksheet.Range("CUSTOMFIELD_VALUES")

    For i = 1 To fieldsRange.count
        If customFieldsRange.Rows(i).value <> False And customFieldsRange.Rows(i).value <> "" Then
            Set subCustomField = New Dictionary
            Set customField = New Dictionary
            subCustomField.Add "name", fieldsRange.Rows(i).value
            subCustomField.Add "id", customFieldsRange.Rows(i).value
            customField.Add "definition", subCustomField
            customField.Add "value", valuesRange.Rows(i).value
            customFields.Add customField
        End If
    Next i
    
    Set GetCustomFields = customFields
End Function
Sub TestWriteToRESTAPIFromSheet()
Dim tmpWorksheet As Worksheet
Dim urlRange As Range, accessCodeRange As Range, _
            fieldsRange As Range, valuesRange As Range, entityRange As Range
Dim responseStatus As String, responseText As String, dataStr As String, outputStr As String
Dim field As Variant, key As Variant
Dim data As New Dictionary
Dim subdata As New Dictionary
Dim header As New Dictionary
Dim i As Integer
Dim objHTTP As Object, jsonObject As Object
Dim capsuleRecord As Dictionary, capsuleSubRecord As Dictionary, tmpRecord As Dictionary
Dim outputFieldCount As Integer
Dim customFields As Collection
Dim resultFieldsRange As Range, resultValuesRange As Range, customfieldFieldsRange As Range, customfieldValuesRange As Range, customFieldsRange As Range, lookupsRange As Range, postData As Range, _
    responseStatusRange As Range, responseTextRange As Range, postDataRange As Range

    Set tmpWorksheet = ActiveSheet

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
    
    For i = 1 To fieldsRange.count
        If customFieldsRange.Rows(i).value = False And lookupsRange.Rows(i) = False And fieldsRange.Rows(i).value <> "" Then
            If fieldsRange.Rows(i).value = "value" Then
                Set tmpRecord = New Dictionary
                tmpRecord.Add "amount", valuesRange.Rows(i).value
                tmpRecord.Add "currency", "USD"
                subdata.Add "value", tmpRecord
            Else
                subdata.Add fieldsRange.Rows(i).value, valuesRange.Rows(i).value
            End If
        End If
    Next i
    
    Set customFields = GetCustomFields(tmpWorksheet, fieldsRange, valuesRange, customFieldsRange)
    subdata.Add "fields", customFields
    
    GetLookupFields tmpWorksheet, fieldsRange, valuesRange, lookupsRange, subdata
    
    data.Add entityRange.value, subdata
    
    header.Add "Authorization", "Bearer " & accessCodeRange.value
    header.Add "Content-Type", "application/json"
    header.Add "Accept", "application/json"

    dataStr = JsonConverter.ConvertToJson(data)
    
    Debug.Print dataStr
    Set objHTTP = WriteToRESTAPI(dataStr, urlRange.value, accessCodeRange.value, header, "POST", _
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
                resultFieldsRange.Rows(outputFieldCount).value = Item
                resultValuesRange.Rows(outputFieldCount).value = outputStr
            ElseIf TypeName(capsuleRecord(Item)) = "Collection" Then
                
            Else
                resultFieldsRange.Rows(outputFieldCount).value = Item
                resultValuesRange.Rows(outputFieldCount).value = capsuleRecord(Item)
            End If
            outputFieldCount = outputFieldCount + 1
        Next Item
    End If
    
    responseStatusRange.value = objHTTP.Status
    responseTextRange.value = objHTTP.responseText
    postDataRange.value = dataStr
    End Sub

Sub TestWriteToRESTAPI()
Dim data As New Dictionary
Dim subdata As New Dictionary
Dim header As New Dictionary
Dim url As String, access_code As String, dataStr As String, responseStatus As String, responseText As String

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

    dataStr = JsonConverter.ConvertToJson(data)
    

    WriteToRESTAPI dataStr, url, access_code, header, "POST", responseStatus, responseText
    
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
            If SheetExists(newSheetName, tmpWorkbook) = False Then
                Set tmpSheet = ActiveWorkbook.Sheets.Add
                tmpSheet.Name = newSheetName
            Else
                
                Set tmpSheet = tmpWorkbook.Sheets(newSheetName)
                tmpSheet.Range("1:1048576").ClearContents
            End If
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
        
        For i = 0 To fileLength - 1
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
