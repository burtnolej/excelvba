Attribute VB_Name = "TestMeetingMinutes"
'TestGetFolderFiles
'VerySimpleTableAdd
'TestIsInStr
'TestArrayToRange
'TestHyperlink

Sub TestGetFolderFiles()
    Application.Run "vbaUtils.xlsm!GetFolderFiles", _
        "E:\Velox Financial Technology\Velox Shared Drive - Documents\General\Sales Team\Meeting Summaries", _
        ActiveWorkbook.Sheets("MEETING_MINUTES")
End Sub

Sub VerySimpleTableAdd()
    Dim oTable As Table
    Set oTable = ActiveDocument.Tables.Add(Range:=Selection.Range, numRows:=3, NumColumns:=3)
End Sub

Sub TestGetOneDriveLink()

    Debug.Print GetOneDriveLink("E:\Velox Financial Technology\Velox Shared Drive - Documents\General\Monday\5095244215 - elias docs\Velox__Quantico__CEO__Meeting_1.pptx")
    
'https://veloxfintechcom.sharepoint.com/sites/VeloxSharedDrive/Shared%20Documents/General/Sales%20Cycle/In%20Sales%20Process/SANTANDER/082223

End Sub
Sub TestIsInStr()

    Debug.Print IsInStr("foo", "barfooda")
    Debug.Print IsInStr("fodo", "barfooda")

    Debug.Print IsInStr("foo", "barfooda", True)
    Debug.Print IsInStr("fodo", "barfooda", True)
    
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

#Const IsDebug = True
Sub TestRangeToArray()
Dim testArray As Variant
Dim tmpArray() As Variant
Dim length As Long, width As Long

    RangeToArray ActiveWorkbook, "Sheet1", "TMPRANGE2", tmpArray, length, width
    
End Sub
Sub testing()
#If IsDebug Then
    MsgBox "foo"
#End If
End Sub
Sub TestHyperlink()
    Dim wordTable As Table
    Dim wordApp As Word.Application
    Dim wordDoc As Word.Document
    Dim wordRange As Word.Range
    
    On Error Resume Next
    Set wordApp = GetObject(, "Word.Application")
    On Error GoTo 0
    If wordApp Is Nothing Then
        Set wordApp = CreateObject("Word.Application")
    End If
    
    Set wordDoc = wordApp.Documents.Add ' create a new document
    Set wordRange = wordDoc.Range
    Set wordTable = wordDoc.Tables.Add(wordRange, 3, 3, wdWord9TableBehavior, wdAutoFitFixed)
    wordTable.Style = "Table Grid"

    
    cellValue = "C:\Users\burtn\OneDrive\laptop\Documents\demo"
    wordTable.Cell(1, 1).Range.InsertAfter cellValue

    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    filesystemObjectExists = False
    On Error Resume Next
    filesystemObjectExists = FSO.folderExists(cellValue)
    If filesystemObjectExists = False Then
        filesystemObjectExists = FSO.fileExists(kcellValue)
    End If
    On Error GoTo 0
           
    If filesystemObjectExists = True Then
        wordDoc.Hyperlinks.Add wordTable.Cell(1, 1).Range, cellValue, "", , "foo"
    End If
    
exitsub:
    Set wordDoc = Nothing
    Set wordApp = Nothing
    Set wordRange = Nothing
    Set workTable = Nothing
    Set FSO = Nothing
    
End Sub








