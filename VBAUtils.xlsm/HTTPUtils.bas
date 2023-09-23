Attribute VB_Name = "HTTPUtils"
'Public Function HTTPDownloadFile(url As String, tmpWorkbook As Workbook, _
'                     sheetNamePrefix As String, _
'                     configSheetName As String, _
'                     Optional startRow As Integer = 0, _
'                     Optional fileType As String = "start-of-day", _
'                     Optional newSheetName As String = "test", _
'                     Optional deleteSheet As Boolean = True, _
'                     Optional startRangeRow As Integer = 0) As Range






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

    
    If fileType <> "start-of-day" Then
        rowOffset = rowCountRange.Value + startRangeRow
    Else
        rowOffset = startRangeRow
    End If
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
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
            tmpSheet.Range("2:1048576").ClearContents
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
    Set outputRange = tmpSheet.Range(Cells(1 + startRangeRow, 1), Cells(fileLength - 1 + startRangeRow, UBound(Split(fileArray(1), "^")) + 1))
    

endsub:
    Set HTTPDownloadFile = outputRange
    Set tmpWorkbook = Nothing
    Set tmpSheet = Nothing
    Set objHTTP = Nothing
    Set tmpRange = Nothing
    Set rowCountRange = Nothing
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Function
