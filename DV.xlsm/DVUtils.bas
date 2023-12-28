Attribute VB_Name = "DVUtils"


Public Sub TestCustomSave(Optional param As Variant)
    ThisWorkbook.CustomSave ActiveWorkbook.Name
End Sub

Public Sub DVGetDataFile(filename)
Dim tmpSheet As Worksheet
Dim outputRange As Range
    sheetname = UCase(Split(filename, ".")(0))
    filename = filename & ".csv"
    url = RetrieveCheckEnvUrl()
    'If envUrl = "" Then
    '    url = "http://172.22.237.138/datafiles/"
    'Else
    ''    url = envUrl
    'End If
    

    Application.Run "DV.xlsm!SetEventsOff"
    
    Set outputRange = Application.Run("DV.xlsm!HTTPDownloadFile", url + "/" + filename, _
                ActiveWorkbook, _
                "", "", 0, "start-of-day", sheetname, False, 0)
    Application.Run "DV.xlsm!SetEventsOn"
    
    On Error Resume Next
    Set tmpSheet = ActiveWorkbook.Sheets(left(sheetname, 30))
    On Error GoTo 0

    tmpSheet.Names.Add left(UCase(sheetname), 30) & "_DATA", outputRange
    tmpSheet.Names.Add left(UCase(sheetname), 30) & "_DATA_HEADER", outputRange.Rows(1)
    
    DVCreateCustomNamedRanges outputRange, tmpSheet
    DVUpdateAvailableSheets left(sheetname, 30), outputRange
    
exitsub:
    Set tmpSheet = Nothing
    Set outputRange = Nothing
    
End Sub

Sub DVUpdateAvailableSheets(newSheetName As String, dataRange As Range)
Dim availableSheetsRange As Range, tmpCell As Range
    Set availableSheetsRange = ActiveWorkbook.Sheets("SHEETS").Range("AVAILABLE_SHEETS")

    For Each tmpCell In availableSheetsRange.Columns(1).Rows
        If tmpCell.value = newSheetName Then
            ' found it so updating
            GoTo endsub
        ElseIf tmpCell.value = "" Then
            ' not found so adding a row
            tmpCell.value = newSheetName
            GoTo endsub
        End If
    Next tmpCell
endsub:
    tmpCell.Offset(, 1).value = Now()
    tmpCell.Offset(, 2).value = dataRange.Rows.count
    tmpCell.Offset(, 3).value = dataRange.Columns.count
    Set availableSheetsRange = Nothing
End Sub
Sub DVCreateCustomNamedRanges(ByVal dataRows As Range, tmpSheet As Worksheet)
Dim tmpCell As Range, headerRow As Range
Dim rangeName As String, newHeaderName As String
Dim colCount As Integer

    Set headerRow = dataRows.Rows(1)
    Set dataRows = dataRows.Offset(1).Resize(dataRows.Rows.count - 1)
    
    For colCount = 1 To headerRow.Columns.count
        newHeaderName = Replace(UCase(headerRow.Columns(colCount).value), " ", "_")
        rangeName = tmpSheet.Name & "_" & newHeaderName
        tmpSheet.Names.Add rangeName, dataRows.Columns(colCount)
        headerRow.Columns(colCount).value = newHeaderName ' make the header rows spelling same as actual named range
    Next colCount

exitsub:
    Set headerRow = Nothing
    Set dataRows = Nothing

End Sub
