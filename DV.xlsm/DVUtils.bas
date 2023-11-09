Attribute VB_Name = "DVUtils"
Public Sub DVGetDataFile(filename)
Dim tmpSheet As Worksheet
Dim outputRange As Range
    sheetname = UCase(Split(filename, ".")(0))
    
    url = RetrieveCheckEnvUrl()
    'If envUrl = "" Then
    '    url = "http://172.22.237.138/datafiles/"
    'Else
    ''    url = envUrl
    'End If
    

    Application.Run "DV.xlsm!SetEventsOff"
    
    Set outputRange = Application.Run("DV.xlsm!HTTPDownloadFile", url + filename, _
                ActiveWorkbook, _
                "", "", 0, "start-of-day", sheetname, False, 0)
    Application.Run "DV.xlsm!SetEventsOn"
    
    Set tmpSheet = ActiveWorkbook.Sheets(sheetname)
    tmpSheet.Names.Add UCase(sheetname) & "_DATA", outputRange
    tmpSheet.Names.Add UCase(sheetname) & "_DATA_HEADER", outputRange.Rows(1)
    
    DVCreateCustomNamedRanges outputRange, tmpSheet
exitsub:
    Set tmpSheet = Nothing
    Set outputRange = Nothing
    
End Sub

Sub DVCreateCustomNamedRanges(dataRows As Range, tmpSheet As Worksheet)
Dim tmpCell As Range, headerRow As Range
Dim rangeName As String
Dim colCount As Integer

    Set headerRow = dataRows.Rows(1)
    Set dataRows = dataRows.Offset(1).Resize(dataRows.Rows.count - 1)
    
    For colCount = 1 To headerRow.Columns.count
        rangeName = tmpSheet.Name & "_" & Replace(UCase(headerRow.Columns(colCount).value), " ", "_")
        tmpSheet.Names.Add rangeName, dataRows.Columns(colCount)
    Next colCount

exitsub:
    Set headerRow = Nothing
    Set dataRows = Nothing

End Sub
