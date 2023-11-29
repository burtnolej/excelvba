Attribute VB_Name = "TestUtils"
'Public Sub TestRefreshCapsuleData()
'Public Sub TestExportModules()
'Public Sub TestRefreshMondayData()
'Sub TestCreateEmail()
'Sub TestIsInStr()
'Sub TestDumpFormats()
'Sub TestArrayToRange()
'Sub TestRangeToArray()
'Sub ArrayToRange(tmpWorkbook As Workbook, sheetNameStr As String, rangeNameStr As String, _
'    tmpArray() As Variant)
'Public Function RangeToArray(tmpWorkbook As Workbook, sheetNameStr As String, rangeNameStr As String, _
'    ByRef tmpArray() As Variant, Optional ByRef length As Long, Optional ByRef width As Long) As Variant
'Public Function RangeToDict(tmpWorkbook As Workbook, sheetNameStr As String, rangeNameStr As String, _
'    ByRef tmpDict As Dictionary, Optional ByRef length As Long, Optional ByRef width As Long) As Variant

Public Sub GenerateRibbon()
    LoadCustRibbon
End Sub

Public Sub OpenDesktop()
Dim tmpWorkbook As Workbook, tmpWindow As Window

    Set tmpWorkbook = Workbooks.Open("E:\Velox Financial Technology\Velox Shared Drive - Documents\General\Tools\Desktop.xlsm")
    
    Set tmpWindow = Windows("Desktop.xlsm")
    tmpWindow.WindowState = xlNormal
    tmpWindow.top = 0
    tmpWindow.left = 1300
    tmpWindow.width = 240
    tmpWindow.height = 800
    tmpWindow.Visible = True

exitsub:
    Set tmpWindow = Nothing
    Set tmpWorkbook = Nothing
    
End Sub


Public Sub TestRefreshCapsuleData()
Dim outputRange As Range
    Set outputRange = HTTPDownloadFile("http://172.22.237.138/datafiles/person.csv", _
                ActiveWorkbook, _
                "", "", 0, "start-of-day", "Sheet2", False, 1)
                
    Debug.Print outputRange.Address
    sortRange outputRange.Worksheet, outputRange, 14
                    
End Sub


Public Sub TestExportModules()

    ExportModules ActiveWorkbook, "E:\Velox Financial Technology\Velox Shared Drive - Documents\General\Tools\vba", _
        "Utils"
End Sub
Public Sub TestRefreshMondayData()
Dim outputRange As Range
    Set outputRange = HTTPDownloadFile("http://172.22.237.138/datafiles/Monday/6666786972.txt", _
                ActiveWorkbook, _
                "", "", 0, "start-of-day", "Sheet3", False, 1)
                
    sortRange outputRange.Worksheet, outputRange, 6
                    
End Sub


Sub TestCreateEmail()
Dim emailContent As String
    emailContent = readClipboard

    createEmail "MeetingSummaries@veloxfintech.com", "jon.butler@veloxfintech.com", _
            "foo", emailContent, True
End Sub


Sub TestIsInStr()

    Debug.Print IsInStr("foo", "barfooda")
    Debug.Print IsInStr("fodo", "barfooda")

    Debug.Print IsInStr("foo", "barfooda", True)
    Debug.Print IsInStr("fodo", "barfooda", True)
    
End Sub

Sub TestDumpFormats()
    DumpFormats
End Sub
Sub TestArrayToRange()
Dim testArray As Variant
Dim tmpArray() As Variant

    ReDim tmpArray(1 To 5, 1 To 1)
    
    tmpArray(1, 1) = "a"
    tmpArray(2, 1) = "b"
    tmpArray(3, 1) = "c"
    tmpArray(4, 1) = "d"
    tmpArray(5, 1) = "e"
    
    
    ArrayToRange ActiveWorkbook, "Sheet1", "TMPRANGE3", tmpArray
    
End Sub

Sub TestRangeToArray()
Dim testArray As Variant
Dim tmpArray() As Variant
Dim length As Long, width As Long

    RangeToArray ActiveWorkbook, "Sheet1", "TMPRANGE2", tmpArray, length, width
    
End Sub

Sub ArrayToRange(tmpWorkbook As Workbook, sheetNameStr As String, rangeNameStr As String, _
    tmpArray() As Variant)
Dim tmpRange As Range
Dim tmpSheet As Worksheet
Dim origSheet As Worksheet

    Set origSheet = ActiveSheet

    Set tmpSheet = tmpWorkbook.Sheets(sheetNameStr)
    
    tmpSheet.Activate
    Set tmpRange = tmpSheet.Range(rangeNameStr)
    origSheet.Activate
    
    tmpRange.value = tmpArray

    
endfunc:
    Set tmpRange = Nothing
    Set origSheet = Nothing
    Set tmpSheet = Nothing
    
    End Sub

Public Function RangeToArray(tmpWorkbook As Workbook, sheetNameStr As String, rangeNameStr As String, _
    ByRef tmpArray() As Variant, Optional ByRef length As Long, Optional ByRef width As Long) As Variant
Dim tmpRange As Range
Dim tmpSheet As Worksheet
Dim origSheet As Worksheet

    Set origSheet = ActiveSheet

    Set tmpSheet = tmpWorkbook.Sheets(sheetNameStr)
    
    tmpSheet.Activate
    Set tmpRange = tmpSheet.Range(rangeNameStr)
    origSheet.Activate
    
    tmpArray = tmpRange
    width = UBound(tmpArray, 2)
    length = UBound(tmpArray)
    
endfunc:

    RangeToArray = tmpArray
    Set tmpRange = Nothing
    Set origSheet = Nothing
    Set tmpSheet = Nothing
    
End Function
Public Function RangeToDict(tmpWorkbook As Workbook, sheetNameStr As String, rangeNameStr As String, _
    ByRef tmpDict As Dictionary, Optional ByRef length As Long, Optional ByRef width As Long) As Variant
Dim tmpRange As Range
Dim tmpSheet As Worksheet
Dim origSheet As Worksheet
Dim colvalues() As Variant

    'Set tmpDict = New Dictionary
    
    Set origSheet = ActiveSheet

    'Set tmpSheet = Workbooks("vbautils.xlsm").Sheets(sheetNameStr)
    
    tmpWorkbook.Sheets(sheetNameStr).Activate
    Set tmpRange = tmpWorkbook.Sheets(sheetNameStr).Range(rangeNameStr)
    
    ReDim colvalues(1 To tmpRange.Columns.count)
    
    For j = 1 To tmpRange.Rows.count
        If tmpDict.Exists(tmpRange(j, 1).value) = False Then
            For i = 2 To tmpRange.Columns.count
                colvalues(i - 1) = CStr(tmpRange(j, i).value)
            Next i
            tmpDict.Add tmpRange(j, 1).value, colvalues
        End If
        'Debug.Print tmpDict.Item(tmpRange(j, 2).Value)
    Next j
    
    origSheet.Activate
    width = 2
    length = tmpRange.Rows.count
    
endfunc:

    Set tmpRange = Nothing
    Set origSheet = Nothing
    Set tmpSheet = Nothing
    
End Function


#Const IsDebug = True

Sub testing()
#If IsDebug Then
    MsgBox "foo"
#End If
End Sub
