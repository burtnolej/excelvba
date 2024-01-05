Attribute VB_Name = "Module1"
'Function boardIdArray()
'Function userNamesArray()
'Public Sub AddFormulas(sumaryWorkbook As Workbook, boardWorksheet As Worksheet, numRows As Integer, startContentRow As Integer)
'Public Sub AddUpdatesFormulas(targetSheet As Worksheet, numRows As Integer, mondayPrefix As String, mondaySuffix As String)
'Public Sub DeleteMondayItem(itemID As String, tmpWorkbook As Workbook, mondayFolder As String)
'Public Function GetFolderFromID(itemID As String, tmpWorkbook As Workbook) As String
'Public Sub AddToMondayFile(itemID As String, mondayFolderPath As String, tmpWorkbook As Workbook, initDocText As String, extraDocText As String, itemName As String, Optional dryRun As Boolean = False)
'Public Function CreateSimpleMondayFolder(mondayFolderPath As String, itemID As String, itemName As String) As String
'Public Sub InitMondayFolder(itemID As String, itemName As String, mondayFolderPath As String, tmpWorkbook As Workbook, Optional initDocText As String = "", Optional dryRun As Boolean = False)
'Public Sub AddFoldersFormulas(targetSheet As Worksheet, numRows As Integer, mondayPrefix As String, mondaySuffix As String)
'Public Function GetUpdates(updateFileDir As String, updatesSheet As Worksheet)
'Public Function GetSandboxFolder(sandboxDir As String, sandboxSheet As Worksheet)
'Public Sub ApplyFilter(filterSheet As Worksheet, colName As String, colValue As Variant)
'Public Sub ApplySort(sortSheet As Worksheet, sortColName As String)
'Public Sub LoadMondayDumpFile()
'Public Sub AddSearchTextColumn(maxRow As Integer, startRow As Integer, targetSheet As Worksheet, targetWorkbook As Workbook, sourceBook As Workbook)
'Public Function CreateViewerBook(genBook As Workbook, viewerBookname As String, viewerBookpath As String, outputFolder As String) As Workbook
'Public Function GetFolderSelection() As String
'Public Sub DisplayGroups(tmpSheet As Worksheet, topLeftCell As Range)
'Public Sub DisplayTags(tmpSheet As Worksheet, topLeftCell As Range)


Const UNDERLINE = "_"
Const SPACE = " "
Const DQ = """"
Const MONDAY_URL = "https://veloxfintech.monday.com/boards/"

Enum FilterColumn
    CreatedDate = 6
    UpdatedDate = 9
    AlisonHood = 13
    JonButler = 17
End Enum
Enum FilterCriteria
    DateLastWeek = 5
    DateToday = 1
End Enum
Function boardIdArray()
    boardIdArray = Array("2259144314", "1140656959", "2193345626", "4977328922", "2763786972", "4973959122", "4977328522", "4974012540", "4973204278")

End Function
Function userNamesArray()
    userNamesArray = Array("Alison Hood", "Ali Moosavi", "Christian Schuler", "Ross Lucas", "Jon Butler")
End Function


'''
' v1.1 fixed bug where it wasnt clearing contents of a previous folder run
' v1.1 added better error handling and IsDebug flag
' v1.2 added code to allow updates (posts) to be posted to Monday
' v1.3 completed post an update
' v1.4 part way through implementing update status - just need to do code gen
' v1.4 need to make the gened callback procedure work for posts and updates
' v1.4 just need to see why adding post now doesnt work
' v1.5 updates done need to fix screen split
' v1.6 fix screen split

#Const IsDebug = True

Public Sub AddFormulas(sumaryWorkbook As Workbook, boardWorksheet As Worksheet, numRows As Integer, startContentRow As Integer)
Dim tmpWorkbook As Workbook
Dim userName As Variant
Dim origUserName As String, hlinkString As String

Dim columnLink As Range, fillRange As Range, columnItemId As Range, ncolumnItemIdFirstCell As Range

    Set columnLink = boardWorksheet.Range("COLUMN_LINK")
    Set columnItemId = boardWorksheet.Range("COLUMN_ITEMID")
    Set columnItemIdFirstCell = columnItemId.Resize(1).Offset(1)
    Debug.Print columnItemIdFirstCell.Address
    
    '=HYPERLINK("https://veloxfintech.monday.com/boards/"&B2&"/pulses/"&H2)
    
    hlinkString = "=hyperlink(" & DQ & MONDAY_URL & DQ & "&B" & startContentRow & "&" & DQ & "/pulses/" & DQ & "&" & columnItemIdFirstCell.Address(0, 0) & ")"
    
    Debug.Print hlinkString
    
    columnLink.Cells(2, 1).Formula = hlinkString
        
    Set fillRange = columnLink.Resize(numRows).Offset(1)
    columnLink.Rows(2).Select
    Selection.AutoFill Destination:=fillRange
        
    For Each userName In userNamesArray
    
        origUserName = userName
        userName = UCase(userName)
        userName = Replace(userName, SPACE, UNDERLINE)
        
        Set columnLink = boardWorksheet.Range("COLUMN_" & userName)
        columnLink.Cells(2, 1).Formula = "=IF(AND(ISERROR(FIND(" & DQ & origUserName & DQ & ",$J" & startContentRow & ")),ISERROR(FIND(" & DQ & origUserName & DQ & ",$J" & startContentRow & "))),0,1)"
        Set fillRange = columnLink.Resize(numRows).Offset(1)
        columnLink.Rows(2).Select
        Debug.Print fillRange.Address
        Selection.AutoFill Destination:=fillRange
    Next userName

 
End Sub
Public Sub AddUpdatesFormulas(targetSheet As Worksheet, numRows As Integer, mondayPrefix As String, mondaySuffix As String)
Dim sectionUpdatesStartColID As Integer
Dim columnUpdateParentId As Range, columnUpdateItemId As Range, columnUpdateCreator As Range, columnUpdatesStart As Range
Dim columnItemId As Range, columnUpdateUpdatesTime As Range

    DDQ = Chr(34)
    ampersand = Chr(38)
    equals = Chr(61)
    
    ' need to make these more dynamic and not hardcoded
    sectionUpdatesStartColID = targetSheet.Range("COLUMN_UPDATES_START").Column
    Set columnItemId = targetSheet.Range("COLUMN_ITEMID").Columns ' reference to the parent item id to lookup by
    Set columnUpdateParentId = targetSheet.Range("COLUMN_UPDATES_PARENTID").Columns
    Set columnUpdateUpdatesTime = targetSheet.Range("COLUMN_UPDATES_UPDATETIME").Columns
    Set columnUpdateItemId = targetSheet.Range("COLUMN_UPDATES_ITEMID").Columns
    Set columnUpdateCreator = targetSheet.Range("COLUMN_UPDATES_CREATOR").Columns
    targetSheet.Activate
    Set columnUpdatesStart = targetSheet.Range("COLUMN_UPDATES_START")

    ' put the index of the update/post row, if one exists into the UPDATE START column
    columnUpdatesStart.Select
    columnUpdatesStart.Rows(2).Formula = "=IF(ISERROR(MATCH(" & columnItemId.Rows(2).Address(0, 0) & _
                             ",INDIRECT(" & columnUpdateParentId.Rows(1).Address(1, 0) & "),0))," & DQ & DQ & ",MATCH(" & columnItemId.Rows(2).Address(0, 0) & _
                             ",INDIRECT(" & columnUpdateParentId.Rows(1).Address(1, 0) & "),0))"

    Set fillRange = columnUpdatesStart.Offset(1).Resize(numRows)
    fillRange.Rows(1).Select
    Selection.AutoFill Destination:=fillRange, Type:=xlFillDefault
    
    ' fill in the update time column (will be a dupe of columnItemID), make the lookup col value dynamic so can fill right
    columnUpdateUpdatesTime.Rows(2).Formula = "=IF(ISERROR(INDEX(INDIRECT(" & columnUpdateUpdatesTime.Rows(1).Address(1, 0) & _
                            ")," & columnUpdatesStart.Rows(2).Address(0, 1) & _
                            "))," & DQ & DQ & ",INDEX(INDIRECT(" & columnUpdateUpdatesTime.Rows(1).Address(1, 0) & _
                            ")," & columnUpdatesStart.Rows(2).Address(0, 1) & _
                            "))"

    ' fill right across the update columns
    Set fillRange = columnUpdateUpdatesTime.Rows(2).Resize(, 5)
    fillRange.Cells(1, 1).Select
    Selection.AutoFill Destination:=fillRange, Type:=xlFillDefault

    ' fill down
    Set fillRange = fillRange.Resize(numRows)
    fillRange.Rows(1).Select
    Selection.AutoFill Destination:=fillRange, Type:=xlFillDefault

    
    
End Sub

Public Sub DeleteMondayItem(itemid As String, tmpWorkbook As Workbook, mondayFolder As String)
Dim rs As String, rt As String, mondayFolderName As String
    mondayFolderName = GetFolderFromID(itemid, tmpWorkbook)
    
    If mondayFolderName <> "-1" Then
        ' need to delete the folder
    End If
    
    DeleteItem itemid, rs, rt
End Sub
Public Function GetFolderFromID(itemid As String, tmpWorkbook As Workbook) As String
Dim folderNamesCol As Range, itemIdCell As Range

    Set folderNamesCol = tmpWorkbook.Sheets("Folders").Range("FOLDERS_COLUMNS").Columns(1)
    
    For Each itemIdCell In folderNamesCol.Cells
        If InStr(1, itemIdCell.value, itemid) <> 0 Then
            GetFolderFromID = itemIdCell.value
            Exit Function
        End If
    Next itemIdCell
    
    GetFolderFromID = -1

End Function
Public Sub AddToMondayFile(itemid As String, mondayFolderPath As String, tmpWorkbook As Workbook, initDocText As String, extraDocText As String, _
                itemName As String, Optional dryRun As Boolean = False)
Dim wordDoc As Word.Document
Dim fso As Scripting.FileSystemObject
Dim wordApp As Word.Application
Dim folderPath As String, mondayFolderName As String, folderName As String

    Set wordDoc = CreateObject("Word.Document")
    Set wordApp = CreateObject("Word.Application")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    mondayFolderName = GetFolderFromID(itemid, tmpWorkbook)
    
    wordDocName = itemid & "__description__" & ".docx"
    
    If mondayFolderName = "-1" Then
        mondayFolderName = itemid & " - " & itemName
        folderPath = fso.BuildPath(mondayFolderPath, mondayFolderName)
        'fso.CreateFolder folderPath
        Debug.Print "created folder : " & itemid & " " & itemName
    Else
        folderPath = fso.BuildPath(mondayFolderPath, mondayFolderName)
    End If

    'folderPath = fso.BuildPath(mondayFolderPath, mondayFolderName)
    docPath = fso.BuildPath(folderPath, wordDocName)
    
    If dryRun = False Then

        If Dir(docPath) = "" Then
            Set wordDoc = wordApp.Documents.Add
            wordDoc.content.InsertAfter text:=initDocText
            wordDoc.content.InsertAfter text:=extraDocText

            Debug.Print "created word doc : " & wordDocName
        Else
            Set wordDoc = wordApp.Documents.Open(docPath)
            
            If InStr(1, wordDoc.content.text, extraDocText) <> 0 Then
                Debug.Print "content allready exists : " & docPath
                GoTo exitsub
            End If
        
            wordDoc.content.InsertParagraphAfter
            wordDoc.content.InsertAfter text:=extraDocText
        End If

        Debug.Print "added to word doc : " & extraDocText
        wordDoc.SaveAs2 fso.BuildPath(folderPath, wordDocName)
        wordDoc.Close
            
    End If

exitsub:
    wordApp.Quit
    Set wordDoc = Nothing
    Set wordApp = Nothing
    Set fso = Nothing
End Sub

Public Function CreateSimpleMondayFolder(mondayFolderPath As String, itemid As String, itemName As String) As String
Dim folderPath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    folderPath = fso.BuildPath(mondayFolderPath, itemid & " " & itemName)
    If Dir(folderPath) = "" Then
        
        itemName = Replace(itemName, " ", "_")
        itemName = Replace(itemName, "/", "_")
        folderPath = fso.BuildPath(mondayFolderPath, itemid & " " & itemName)
        fso.CreateFolder folderPath
    End If
    
    CreateSimpleMondayFolder = folderPath
            
End Function
Public Sub InitMondayFolder(itemid As String, itemName As String, mondayFolderPath As String, tmpWorkbook As Workbook, _
    Optional initDocText As String = "", Optional dryRun As Boolean = False)
Dim fso As Scripting.FileSystemObject
Dim folderName As String, folderPath As String, wordDocName As String
Dim wordDoc As Word.Document
Dim wordApp As Word.Application
Dim objSelection As Variant
Dim folderNamesCol As Range, itemIdCell As Range

    Set wordDoc = CreateObject("Word.Document")
    Set wordApp = CreateObject("Word.Application")

    Set folderNamesCol = tmpWorkbook.Sheets("Folders").Range("FOLDERS_COLUMNS").Columns(1)
    
    For Each itemIdCell In folderNamesCol.Cells
        If InStr(1, itemIdCell.value, itemid) <> 0 Then
            Debug.Print "folder exists : " & itemIdCell.value
            GoTo exitsub
        End If
    Next itemIdCell
    wordDocName = itemid & "__description__" & ".docx"

    folderName = itemid & " - " & itemName
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    folderPath = fso.BuildPath(mondayFolderPath, folderName)
    If Dir(folderPath) = "" Then
    
        If dryRun = False Then
            fso.CreateFolder folderPath
            Set wordDoc = wordApp.Documents.Add
            wordDoc.content.InsertAfter text:=initDocText
            wordDoc.SaveAs2 fso.BuildPath(folderPath, wordDocName)
            wordDoc.Close
        End If
        
        Debug.Print "created folder : " & folderPath
        Debug.Print "created word doc : " & wordDocName
        
    End If
    
exitsub:

    wordApp.Quit
    Set wordDoc = Nothing
    Set wordApp = Nothing
    Set fso = Nothing
    

End Sub
Public Sub AddFoldersFormulas(targetSheet As Worksheet, numRows As Integer, mondayPrefix As String, mondaySuffix As String)
Dim startRange As Range, fillRange As Range, columnEmailRange As Range, columnEmailStartCell As Range, columnItemId As Range, columnItemId2 As Range
Dim columnBoardId As Range, columnBoardName As Range, boardIdStartCell As Range, boardNameStartCell As Range, folderLinkRange As Range
Dim emailLinkFormula As String, boardNameFormula As String
Dim doubleQuote As String
Dim sectionFolderStart As Integer
    
    doubleQuote = Chr(34)
    ampersand = Chr(38)
    equals = Chr(61)

    ' need to make these more dynamic and not hardcoded
    sectionFolderStart = targetSheet.Range("COLUMN_FOLDER_START").Column
    Set columnEmailRange = targetSheet.Range("COLUMN_EMAIL").Columns
    Set columnItemId = targetSheet.Range("COLUMN_ITEMID").Columns
    Set columnItemId2 = targetSheet.Range("COLUMN_ITEMID2").Columns ' with a prefix for matching
    Set columnBoardId = targetSheet.Range("COLUMN_BOARDID").Columns
    targetSheet.Activate
    Set folderStartRange = targetSheet.Range("COLUMN_FOLDER_START")
    
    folderStartRange.Cells(2, 1).Formula = "=IF(ISERROR(MATCH(" & columnItemId2.Rows(2).Address(0, 0) & ",Folders!I:I,0)),-1,MATCH(" & columnItemId2.Rows(2).Address(0, 0) & ",Folders!I:I,0))"
    Set fillRange = folderStartRange.Resize(numRows).Offset(1)
    folderStartRange.Rows(2).Select
    Selection.AutoFill Destination:=fillRange
    
    Set folderDataRange = folderStartRange.Offset(, 1).Resize(, 8)
    folderDataRange.Cells(2, 1).Formula = "=IF(" & folderStartRange.Cells(2, 1).Address(0, 1) & _
                    "<>-1,INDEX(FOLDER_COLUMNS," & folderStartRange.Cells(2, 1).Address(0, 1) & _
                    ",MATCH(" & folderDataRange.Cells(1, 1).Address(1, 0) & ",FOLDER_HEADERS,0))," & DQ & DQ & ")"
    
    ' fill across the row first
    Set fillRange = folderDataRange.Rows(1).Offset(1)
    fillRange.Cells(1, 1).Select
    Selection.AutoFill Destination:=fillRange, Type:=xlFillDefault
    
    ' fill down the row first
    Set fillRange = folderDataRange.Resize(numRows).Offset(1)
    fillRange.Rows(1).Select
    Selection.AutoFill Destination:=fillRange
    
    Set folderLinkRange = folderStartRange.Offset(, 9)
    folderLinkRange.Cells(2, 1).Formula = "=hyperlink(IF(" & folderStartRange.Cells(2, 1).Address(0, 1) & _
                    "<>-1,INDEX(FOLDER_COLUMNS," & folderStartRange.Cells(2, 1).Address(0, 1) & _
                    ",MATCH(" & folderLinkRange.Cells(1, 1).Address(1, 0) & ",FOLDER_HEADERS,0))," & DQ & DQ & "))"
    
    Set fillRange = folderLinkRange.Rows(2).Resize(numRows)
    fillRange.Rows(1).Select
    Selection.AutoFill Destination:=fillRange
    
    emailLinkFormula = equals & doubleQuote & mondayPrefix & doubleQuote & ampersand & _
                            columnItemId.Cells(2, 1).Address(0, 1) & ampersand & _
                            doubleQuote & mondaySuffix & doubleQuote
    Set columnEmailStartCell = columnEmailRange.Cells(2, 1)
    columnEmailStartCell.Formula = emailLinkFormula
    Set fillRange = columnEmailRange.Resize(numRows).Offset(1)
    columnEmailStartCell.Select
    Selection.AutoFill Destination:=fillRange
    

    Set columnBoardName = columnBoardId.Offset(, 1)

    Set boardIdStartCell = columnBoardId.Rows(2)
    'targetSheet.Activate
    columnBoardName.EntireColumn.Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    boardNameFormula = "=INDEX(DATA_BOARDNAMES,MATCH(" & boardIdStartCell.Address(0, 1) & ",DATA_BOARDID,0),1)"
    Set columnBoardName = columnBoardName.Offset(, -1)
    columnBoardName.Rows(1).value = "BOARD_NAME"
    Set boardNameStartCell = columnBoardName.Rows(2)
    boardNameStartCell.Select
    boardNameStartCell.Formula = boardNameFormula
    Set fillRange = columnBoardName.Resize(numRows).Offset(1)
    boardNameStartCell.Select
    Selection.AutoFill Destination:=fillRange

End Sub

Public Function GetUpdates(updateFileDir As String, updatesSheet As Worksheet)
'2022-06-17T08:20:37Z^2790791796^1557226394^Christian Schuler
'H:\My Drive\monday__dump\datafiles\20220617
Dim oFSO As FileSystemObject
Dim tmpWorkbook As Workbook

    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set tmpWorkbook = Workbooks.Open(fileName:=oFSO.BuildPath(updateFileDir, "updates.txt"), Format:=6, Delimiter:="^")
    tmpWorkbook.Sheets("updates").UsedRange.Select
    Selection.Copy
    updatesSheet.Activate
    updatesSheet.Cells(1, 1).Select
    updatesSheet.Paste
    tmpWorkbook.Close
    
End Function
Public Function GetSandboxFolder(sandboxDir As String, sandboxSheet As Worksheet)
 
Dim oFSO As Object, oFolders As Object, oFolder As Object, oFile As Object
Dim resultArray() As String, fileList As String
Dim outputRange As Range, columnLink As Range, fillRange As Range, itemLink As Range
Dim i As Integer

On Error GoTo err
    i = 0

    ' clear previous data but leave headers intact
    sandboxSheet.UsedRange.Offset(1).ClearContents

    ReDim resultArray(0 To 600, 0 To 8)
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolders = oFSO.GetFolder(sandboxDir).SubFolders
     
    For Each oFolder In oFolders
        resultArray(i, 0) = oFolder.Name
        resultArray(i, 1) = Format(CDate(oFolder.DateCreated), "YYYY/MM/DD")
        resultArray(i, 2) = Format(CDate(oFolder.DateLastModified), "YYYY/MM/DD")
        resultArray(i, 3) = oFolder.Path
        resultArray(i, 4) = oFolder.Size
        
        fileList = ""
        For Each oFile In oFolder.Files
            fileList = fileList & oFile.Name & ","
        Next oFile
        
        resultArray(i, 5) = oFolder.Files.Count
        resultArray(i, 6) = fileList
        resultArray(i, 8) = "a" & CStr(Left(oFolder.Name, 10))
        
        i = i + 1
     
    Next oFolder
    
    sandboxSheet.Activate
    With sandboxSheet
        Set outputRange = .Range(Cells(2, 1), Cells(oFolders.Count + 1, 9))
        outputRange = resultArray
        Set columnLink = .Range("FOLDER_COLUMN_LINK")
        columnLink.Cells(2, 1).Formula = "=hyperlink(D2)"
        Set fillRange = columnLink.Resize(oFolders.Count).Offset(1)
        columnLink.Rows(2).Select
        Selection.AutoFill Destination:=fillRange
    End With
    
    GoTo exitsub
    
err:
    MsgBox err.Number & ": " & err.Description, , ThisWorkbook.Name & ": GetSandboxFolder"
    #If IsDebug Then
        Stop            ' Used for troubleshooting - Then press F8 to step thru code
        Resume          ' Resume will take you to the line that errored out
    #Else
        Resume exitsub ' Exit procedure during normal running
    #End If
    
    
exitsub:
    
End Function


Public Sub ApplyFilter(filterSheet As Worksheet, colName As String, colValue As Variant)
Dim colFilterRange As Range, filterRange As Range
Dim colFilterNum As Integer
    
    Set colFilterRange = filterSheet.Range(colName)
    colFilterNum = colFilterRange.Column
    Set filterRange = filterSheet.AutoFilter.Range ' find the filter on the sheet

    With filterSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    filterRange.AutoFilter Field:=colFilterNum, Criteria1:=colValue, Operator:=xlFilterValues
    
End Sub

Public Sub ApplySort(sortSheet As Worksheet, sortColName As String)
Dim colSortRange As Range
Dim colSortNum As Integer

    Set colSortRange = sortSheet.Range(sortColName)
        
    sortSheet.AutoFilter.Sort.SortFields.Clear
    sortSheet.AutoFilter.Sort.SortFields.Add2 key:=colSortRange, SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With sortSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub


Public Sub LoadMondayDumpFile()
Dim dirnameString As String, fileNameSuffix As String, parentDirnameString As String, datafileDirname As String
Dim targetSheetName As String, targetFileName As String, dataFile As String, targetDirName As String
Dim inputUserName As String, inputDate As String, inputGDrive As String, inputFolder As String, folderSheetName As String
Dim fileExtension As String, inputFileExtension As String, outputFolderSheet As String, configSheetName As String, tmpPath As String
Dim mondayPrefix As String, mondaySuffix As String, updatesSheetName As String, outputBookname As String
Dim tmpWorkbook As Workbook, summaryWorkbook As Workbook, codeWorkbook As Workbook
Dim boardWorksheet As Worksheet, tmpWorksheet As Worksheet, sourceFolderSheet As Worksheet, targetFolderSheet As Worksheet, folderSheet As Worksheet
Dim tmpSheet As Worksheet, sourceUpdatesSheet As Worksheet, targetUpdatesSheet As Worksheet
Dim offsetFactor As Integer, i As Integer, numRows As Integer, sectionFolderStart As Integer
Dim columnLink As Range, startItemsCell As Range
Dim fs As Object

    On Error GoTo err
    
    SetEventsOff
    
    offsetFactor = 0
    numRows = 0
    fileExtension = ".xlsm"
    inputFileExtension = ".txt"
    viewerFileNameSuffix = "monday_viewer"
    targetSheetName = "Viewer"
    folderSheetName = "Folders"
    updatesSheetName = "Updates"
    configSheetName = "Config"
    outputBookname = "monday_viewer.xlsm"

    Set fs = CreateObject("Scripting.FileSystemObject")

    Set codeWorkbook = ThisWorkbook
        
    inputUserName = ActiveSheet.Range("INPUT_USER")
    inputDate = ActiveSheet.Range("INPUT_DATE")
    inputGDrive = ActiveSheet.Range("INPUT_GDRIVE")
    
    'If DirExist(fs.BuildPath(inputGDrive, "\")) = False Then
    '    MsgBox "Cannot find " & inputGDrive
    '    GoTo err
    'End If
    
    'tmpPath = fs.BuildPath(inputGDrive, "datafiles")
    'If DirExist(tmpPath) = False Then
    '    MsgBox "cannot find " & tmpPath
    '    GoTo err
    'End If
    
    inputFolder = ActiveSheet.Range("INPUT_FOLDER")
    outputFolder = ActiveSheet.Range("OUTPUT_FOLDER")
    inputOpenFlag = ActiveSheet.Range("INPUT_OPENFLAG")
    refreshUpdatesFlag = ActiveSheet.Range("REFRESH_UPDATE_FLAG")
    refreshFolderFlag = ActiveSheet.Range("REFRESH_FOLDER_FLAG")
    outputFolderSheet = ActiveSheet.Range("OUTPUT_FOLDER_SHEET")
    mondayPrefix = ActiveSheet.Range("MONDAY_PREFIX")
    mondaySuffix = ActiveSheet.Range("MONDAY_SUFFIX")

    ' build input absolute path and filename
    parentDirnameString = inputGDrive
    'datafileDirname = fs.BuildPath(parentDirnameString, "datafiles")
    'datafileDirname = fs.BuildPath(datafileDirname, inputDate)
    datafileDirname = parentDirnameString
    targetDirName = outputFolder
    targetFileName = fs.BuildPath(targetDirName, viewerFileNameSuffix & "_" & inputDate & "_" & Replace(inputUserName, " ", "") & fileExtension)
    
    ' existing reference sheets
    Set sourceFolderSheet = codeWorkbook.Sheets(folderSheetName)
    Set sourceUpdatesSheet = codeWorkbook.Sheets(updatesSheetName)

    ' create the output book with all its named ranges and sheets
    Set summaryWorkbook = CreateViewerBook(codeWorkbook, outputBookname, inputGDrive, targetDirName)
    
   
    Set boardWorksheet = summaryWorkbook.Sheets(targetSheetName) 'this is where the final report goes

    
    
    Set targetFolderSheet = summaryWorkbook.Sheets(outputFolderSheet) ' where the underlying folder data goes
    Set clipboardSheet = summaryWorkbook.Sheets.Add
    clipboardSheet.Name = "___tmp" ' temp store to collate all the input board data
    Set targetUpdatesSheet = summaryWorkbook.Sheets.Add ' where the underlying updates/posts data goes for reference fro final report sheet
    targetUpdatesSheet.Name = "updates"
    targetUpdatesSheet.Visible = xlSheetHidden
    
    If refreshFolderFlag = True Then
        GetSandboxFolder inputFolder, sourceFolderSheet
    End If
    
    If refreshUpdatesFlag = True Then
        GetUpdates datafileDirname, sourceUpdatesSheet
    End If
    
    offsetFactor = 4
    Set startItemsCell = boardWorksheet.Cells(offsetFactor + 1, 1)
    
    ' for each monday board input file
    For i = 0 To UBound(boardIdArray)
    
        Application.StatusBar = "Processing " & dataFile
                
        ' build the input file name and open it
        dataFile = fs.BuildPath(datafileDirname, boardIdArray(i) & inputFileExtension)
        Set tmpWorkbook = Workbooks.Open(fileName:=dataFile, Format:=6, Delimiter:="^")
        Set tmpWorksheet = tmpWorkbook.Sheets(boardIdArray(i))
        tmpWorksheet.Activate
        
        ' copy the items from the board
        ActiveSheet.UsedRange.Select
        Selection.Copy
        
        ' paste to temp store
        clipboardSheet.Activate
        clipboardSheet.Range("A1").Offset(offsetFactor).Select
        clipboardSheet.Paste

        offsetFactor = offsetFactor + Selection.Rows.Count

        Application.CutCopyMode = False
        tmpWorkbook.Close
    Next i

    ' copy from temp and paste into main output report, then delete the temp sheet
    Application.StatusBar = "Loading items to  " & boardWorksheet.Name
    clipboardSheet.Activate
    clipboardSheet.UsedRange.Select
    Selection.Copy
    boardWorksheet.Activate
    startItemsCell.Select
    startItemsCell.PasteSpecial Paste:=xlPasteValues
    clipboardSheet.Delete


        
    Application.StatusBar = "Adding formulas to  " & boardWorksheet.Name
    AddFormulas summaryWorkbook, boardWorksheet, offsetFactor, startItemsCell.Row
    startItemsCell.Select
    ActiveWindow.ScrollColumn = 1

    ' copy the folder data into the target  book
    sourceFolderSheet.UsedRange.Copy
    targetFolderSheet.Activate
    targetFolderSheet.Cells(1, 1).PasteSpecial Paste:=xlPasteValues
    
    summaryWorkbook.Names.Add "FOLDER_COLUMNS", RefersTo:=targetFolderSheet.UsedRange

    AddFoldersFormulas boardWorksheet, offsetFactor + 1, mondayPrefix, mondaySuffix
    
    ' copy the item updates  into the target  book
    sourceUpdatesSheet.UsedRange.Copy
    targetUpdatesSheet.Activate
    targetUpdatesSheet.Cells(1, 1).PasteSpecial Paste:=xlPasteValues

    summaryWorkbook.Names.Add "UPDATES_UPDATETIME", RefersTo:=targetUpdatesSheet.UsedRange.Columns(1)
    summaryWorkbook.Names.Add "UPDATES_PARENTID", RefersTo:=targetUpdatesSheet.UsedRange.Columns(2)
    summaryWorkbook.Names.Add "UPDATES_ITEMID", RefersTo:=targetUpdatesSheet.UsedRange.Columns(3)
    summaryWorkbook.Names.Add "UPDATES_CREATOR", RefersTo:=targetUpdatesSheet.UsedRange.Columns(4)
    summaryWorkbook.Names.Add "UPDATES_FIRSTLINE", RefersTo:=targetUpdatesSheet.UsedRange.Columns(5)
    
    AddUpdatesFormulas boardWorksheet, offsetFactor + 1, mondayPrefix, mondaySuffix
    
    AddSearchTextColumn 2000, 4, boardWorksheet, summaryWorkbook, codeWorkbook


    ApplySort boardWorksheet, "COLUMN_UPDATED_ON"
    
    ApplyFilter boardWorksheet, "COLUMN_OWNER", "Jon Butler"
    ApplyFilter boardWorksheet, "COLUMN_STATUS", Array("Working", "Not Started")
        
    summaryWorkbook.Activate
    summaryWorkbook.SaveAs FileFormat:=xlOpenXMLWorkbookMacroEnabled, fileName:=targetFileName
    
    GoTo exitsub
    
err:
    MsgBox err.Number & ": " & err.Description, , ThisWorkbook.Name & ": LoadMondayDumpFile"
    #If IsDebug Then
        Stop            ' Used for troubleshooting - Then press F8 to step thru code
        Resume          ' Resume will take you to the line that errored out
    #Else
        Resume exitsub ' Exit procedure during normal running
    #End If
    
    
exitsub:
        If Not summaryWorkbook Is Nothing Then
        summaryWorkbook.Close
        Set summaryWorkbook = Nothing
    End If
    
    codeWorkbook.Sheets(configSheetName).Activate
    
    Set fs = Nothing
    
    If inputOpenFlag = True Then
        Workbooks.Open targetFileName, UpdateLinks:=0
    End If
    
    SetEventsOn
        
End Sub

'Public Sub AddSearchTextColumn(maxRow As Integer, startRow As Integer, targetSheet As Worksheet, searchTerm As String, searchTermColIdx As Integer, targetWorkbook As Workbook)
Public Sub AddSearchTextColumn(maxRow As Integer, startRow As Integer, targetSheet As Worksheet, targetWorkbook As Workbook, sourceBook As Workbook)
Dim searchTermRange As Range, searchTermColIndexRange As Range, targetColumnRange As Range, colValueRange As Range, viewerColumnFilterRange As Range, targetColumnHeaderRange As Range
Dim searchTermRangeString As String, configSheetName As String, searchTermColIndexRangeString As String
Dim colValue As String
Dim viewerColumnFilterName As Name
Dim myWorkbooks As Workbooks
Dim sourceWorkbookNameIndex As Integer, targetWorkbookNameIndex As Integer
Dim searchTermArray As Variant, outputArray As Variant

    
    ReDim outputArray(0 To maxRow - 1, 1 To 1)
    Set myWorkbooks = Workbooks
    searchTermRangeString = "SEARCH_TEXT_COLUMN"
    searchTermColIndexRangeString = "SEARCH_TEXT_COLUMN_INDEX"
    
    configSheetName = "Reference"

    Set searchTermRange = sourceBook.Sheets(configSheetName).Range(searchTermRangeString)
    Set searchTermColIndexRange = sourceBook.Sheets(configSheetName).Range(searchTermColIndexRangeString)
    searchTermArray = Split(searchTermRange.value, ",")

    Set targetColumnRange = targetSheet.Columns(searchTermColIndexRange.value)
    
    For i = 1 To maxRow
        For j = LBound(searchTermArray) To UBound(searchTermArray)
            Set colValueRange = targetSheet.Range("COLUMN_" & searchTermArray(j))
            outputArray(i - 1, 1) = outputArray(i - 1, 1) & colValueRange.Rows(i + 1).value
        Next j
    Next i
    
    Set targetColumnRange = targetColumnRange.Resize(maxRow).Offset(startRow)
    targetColumnRange.value = outputArray
    
    Set viewerColumnFilterRange = targetSheet.Range("E1")
    Set viewerColumnFilterName = targetWorkbook.Names.Add("COLUMN_FILTER_SEARCHSTR", RefersTo:=viewerColumnFilterRange)
    viewerColumnFilterRange.Style = "input"
    
    ' this column is added separately and the AddFilterCode function assumes that column needs to be +1 so -1 is to compensate
    Set targetColumnRange = targetColumnRange.Offset(, -1)
    AddFilterCode targetWorkbook, targetSheet.Name, targetColumnRange, viewerColumnFilterName

End Sub


Public Function CreateViewerBook(genBook As Workbook, viewerBookname As String, viewerBookpath As String, outputFolder As String) As Workbook
Dim viewerSheet As Worksheet, viewerGenSheet As Worksheet, viewerRefSheet As Worksheet, viewerFoldersSheet As Worksheet, tmpSheet As Worksheet
Dim viewerColumnNamesRange As Range, viewerColumnNumRange As Range, viewerColumnRange As Range
Dim viewerGenColumnCell As Range, viewerGenColumnWidthRange As Range, formatCell As Range, viewerGenColumnFormat As Range
Dim viewerColumnFilterRange As Range, viewerGenColumnFilterInputRange As Range, typeCell As Range
Dim viewerGenColumnHeaderFormatRange As Range, headerFormatCell As Range
Dim viewerWorkbook As Workbook, viewerGenWorkbook As Workbook
Dim viewerGenColumnNameRange As Range, viewerGenDataBoardnameRange As Range, viewerGenColumnVisibleRange As Range
Dim viewerGenColumnRange As Range, viewerGenDataBoadrdidRange As Range, viewerGenColumnTypeRange As Range
Dim viewerColumnFilterName As Name
Dim startRow As Integer, foldersStartRow As Integer, maxRow As Integer

    startRow = 4
    foldersStartRow = 1
    maxRow = 2000

    'Set viewerGenWorkbook = Workbooks("monday_report_gen")
    Set viewerGenSheet = genBook.Sheets("Reference")
    
    Set viewerWorkbook = Workbooks.Add
    
    Set viewerFoldersSheet = ActiveWorkbook.Sheets.Add
    viewerFoldersSheet.Name = "Folders"
    viewerFoldersSheet.Visible = xlSheetHidden
    viewerWorkbook.Names.Add "FOLDER_HEADERS", RefersTo:=viewerFoldersSheet.Rows(foldersStartRow)
    
    Set viewerSheet = ActiveWorkbook.Sheets.Add
    viewerSheet.Name = "Viewer"
    Set viewerRefSheet = ActiveWorkbook.Sheets.Add
    viewerRefSheet.Name = "Reference"
    viewerRefSheet.Visible = xlSheetHidden
    
    'On Error Resume Next ' hide the default sheet if it exists
    Set tmpSheet = viewerWorkbook.Sheets("Sheet1")
    tmpSheet.Visible = xlSheetHidden
    'On Error GoTo 0
    
    Set viewerGenColumnNameRange = viewerGenSheet.Range("VIEWER_COLUMN_NAMES")
    Set viewerGenColumnWidthRange = viewerGenSheet.Range("VIEWER_COLUMN_WIDTH")
    Set viewerGenColumnFormat = viewerGenSheet.Range("VIEWER_COLUMN_FORMAT")
    Set viewerGenColumnVisibleRange = viewerGenSheet.Range("VIEWER_COLUMN_VISIBLE")
    Set viewerGenColumnRange = viewerGenSheet.Range("VIEWER_COLUMN")
    Set viewerGenDataBoardnameRange = viewerGenSheet.Range("VIEWER_DATA_BOARDNAMES")
    Set viewerGenDataBoadrdidRange = viewerGenSheet.Range("VIEWER_DATA_BOARDID")
    Set viewerGenColumnFilterInputRange = viewerGenSheet.Range("VIEWER_COLUMN_FILTER_INPUT")
    Set viewerGenColumnTypeRange = viewerGenSheet.Range("VIEWER_COLUMN_TYPE")
    Set viewerGenColumnHeaderFormatRange = viewerGenSheet.Range("VIEWER_COLUMN_HEADER_FORMAT")

    
    viewerSheet.Activate
        
    AddFilterCalbackSub viewerWorkbook, viewerSheet.Name
        
    For Each viewerGenColumnCell In viewerGenColumnNameRange
        Set viewerColumnRange = viewerSheet.Range(Cells(startRow, viewerGenColumnCell.Row() - 1), Cells(2000, viewerGenColumnCell.Row() - 1))
        viewerWorkbook.Names.Add "COLUMN_" & viewerGenColumnCell.value, RefersTo:=viewerColumnRange
        viewerColumnRange.Rows(1).value = viewerGenColumnCell.value
        
        ' set the column width
        viewerColumnRange.ColumnWidth = viewerGenColumnWidthRange.Rows(viewerGenColumnCell.Row() - 1)
        
        ' set visibility
        If viewerGenColumnVisibleRange.Rows(viewerGenColumnCell.Row() - 1) = 0 Then
            viewerColumnRange.EntireColumn.Hidden = True
        End If
        
        ' set formats
        Set formatCell = viewerGenColumnFormat.Rows(viewerGenColumnCell.Row() - 1)
        CopyCellFormat formatCell, viewerColumnRange

        ' set types
        Set typeCell = viewerGenColumnTypeRange.Rows(viewerGenColumnCell.Row() - 1)
        If typeCell <> -1 Then ' -1 means a formula
            viewerColumnRange.Select
            Selection.NumberFormat = typeCell.value
        End If
        
        ' set column header format
        Set headerFormatCell = viewerGenColumnHeaderFormatRange.Rows(viewerGenColumnCell.Row() - 1)
        CopyCellFormat headerFormatCell, viewerColumnRange.Rows(1)
        
        'add in the filter cell
        If viewerGenColumnFilterInputRange.Rows(viewerGenColumnCell.Row() - 1) = 1 Then

            Set viewerColumnFilterRange = viewerColumnRange.Resize(1).Offset(-2)
            Set viewerColumnFilterName = viewerWorkbook.Names.Add("COLUMN_FILTER_" & viewerGenColumnCell.value, RefersTo:=viewerColumnFilterRange)
            viewerColumnFilterRange.Style = "input"
            'viewerGenColumnFilterInputRange.AddComment = "> for multiple values enter a comma separated list i.e. word1,word2" & vbNewLine & "> for wildcards add a asterisk as suffix/prefix i.e. *word* " & vbNewLine & "> to reset to no filters enter an empty string " & vbNewLine & "> for nonblanks enter <>"
        
            AddFilterCode viewerWorkbook, viewerSheet.Name, viewerColumnFilterRange, viewerColumnFilterName
            
            ' add a tooltip
            viewerColumnFilterRange.Select
            viewerColumnFilterRange.AddCommentThreaded ( _
                "for multiple values enter a comma separated list i.e. word1,word2, " & Chr(10) & "" & Chr(10) & "for wildcards add a asterisk as suffix/prefix i.e. *word*" & Chr(10) & "" & Chr(10) & "to reset to no filters enter an empty string") & Chr(10) & "" & Chr(10) & "> for nonblanks enter <>"
        
        End If
    Next viewerGenColumnCell
    

    
    AddSendMondayUpdateCode genBook, viewerWorkbook, outputFolder & "\tmp.txt"
    
    viewerWorkbook.Names.Add "DATA_BOARDNAMES", RefersTo:=viewerRefSheet.Range(viewerGenDataBoardnameRange.Address)
    viewerWorkbook.Names.Add "DATA_BOARDID", RefersTo:=viewerRefSheet.Range(viewerGenDataBoadrdidRange.Address)
    
    viewerGenDataBoadrdidRange.Copy
    viewerRefSheet.Activate
    viewerRefSheet.Range(viewerGenDataBoadrdidRange.Address).Select
    viewerRefSheet.Paste

    viewerGenDataBoardnameRange.Copy
    viewerRefSheet.Range(viewerGenDataBoardnameRange.Address).Select
    viewerRefSheet.Paste

    
    viewerSheet.Activate
    Rows("4:4").Select
    Selection.AutoFilter
    
    Range("H5:H5").Select
    ActiveWindow.FreezePanes = True
    
    ActiveWindow.Zoom = 70

    Set CreateViewerBook = viewerWorkbook
    Set summaryWorkbook = Nothing
    Set viewerWorkbook = Nothing
    
End Function

Public Function GetFolderSelection() As String
Dim sFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then ' if OK is pressed
            sFolder = .SelectedItems(1)
        End If
    End With
    GetFolderSelection = sFolder
End Function


Public Function DisplayGroups(tmpSheet As Worksheet, topLeftCell As Range) As Long
Dim boardsColl As Collection, groupsColl As Collection
Dim rs As String, rt As String
Dim board As Variant, boardGroup As Variant, groupItem As Variant
Dim rowIndex As Integer
Dim nextCell As Range, namedRange As Range, lastCell As Range
Dim totalgroups As Long

    Set nextCell = topLeftCell
    rowIndex = 1
    Set boardsColl = GetBoards(rs, rt)
    
    For Each board In boardsColl
         
        Set groupsColl = GetGroupsForBoard(CStr(board("id")), rs, rt)
        
        If InStr(board("name"), "Subitems") = 1 Then
            
            GoTo nextiter
        End If
        
        If groupsColl.Count > 0 Then
            
            For i = 1 To groupsColl.Count
                nextCell.value = board("id")
                nextCell.Offset(, 1).value = board("name")
                nextCell.Offset(, 2).value = groupsColl(i)("title")
                nextCell.Offset(, 3).value = groupsColl(i)("id")
                Debug.Print "group:" & board("name") & "," & groupsColl(i)("title")
                Set nextCell = nextCell.Offset(1)
            Next i
            totalgroups = totalgroups + groupsColl.Count
            
        End If
nextiter:
    Next board
    
    Set lastCell = nextCell.Offset(-1)
    
    'Set lastCell = nextCell.Resize(-1)
    Set namedRange = tmpSheet.Range(topLeftCell, lastCell)
    ThisWorkbook.Names.Add Name:="BOARD_IDS", RefersTo:=namedRange

    Set namedRange = namedRange.Offset(, 1)
    ThisWorkbook.Names.Add Name:="BOARD_NAMES", RefersTo:=namedRange
    
    Set namedRange = namedRange.Offset(, 1)
    ThisWorkbook.Names.Add Name:="GROUP_NAMES", RefersTo:=namedRange
    
    Set namedRange = namedRange.Offset(, 1)
    ThisWorkbook.Names.Add Name:="GROUP_IDS", RefersTo:=namedRange
    
    DisplayGroups = totalgroups
End Function

Public Function DisplayTags(tmpSheet As Worksheet, topLeftCell As Range) As Long
Dim tagColl As Collection, tagsColl As Collection
Dim rs As String, rt As String
Dim tag As Variant
Dim rowIndex As Integer
Dim nextCell As Range
Dim numTags As Long

    Set nextCell = topLeftCell
    rowIndex = 1
    Set tagsColl = GetTags(rs, rt)

    nextCell.value = "SELECT_ONE"
    nextCell.Offset(, 1).value = "SELECT_ONE"
    Set nextCell = nextCell.Offset(1)
    For i = tagsColl.Count To 1 Step -1
        nextCell.value = tagsColl(i)("id")
        nextCell.Offset(, 1).value = tagsColl(i)("name")
        Set nextCell = nextCell.Offset(1)
        Debug.Print "tag:" & tagsColl(i)("name")
    Next i
    
    Set lastCell = nextCell.Offset(-1)

    Set namedRange = tmpSheet.Range(topLeftCell, lastCell)
    ThisWorkbook.Names.Add Name:="TAGS_ID", RefersTo:=namedRange
    
    Set namedRange = namedRange.Offset(, 1)
    ThisWorkbook.Names.Add Name:="TAGS_NAMES", RefersTo:=namedRange
    
    DisplayTags = tagsColl.Count
End Function

Public Function DisplayUsers(tmpSheet As Worksheet, topLeftCell As Range) As Long
Dim tagColl As Collection, tagsColl As Collection
Dim rs As String, rt As String
Dim tag As Variant
Dim rowIndex As Integer
Dim nextCell As Range
Dim numTags As Long

    Set nextCell = topLeftCell
    rowIndex = 1
    Set tagsColl = GetUsers(rs, rt)

    For i = tagsColl.Count To 1 Step -1
        nextCell.value = tagsColl(i)("id")
        nextCell.Offset(, 1).value = tagsColl(i)("name")
        Set nextCell = nextCell.Offset(1)
        Debug.Print "user:" & tagsColl(i)("name")
    Next i
    
    Set lastCell = nextCell.Offset(-1)

    Set namedRange = tmpSheet.Range(topLeftCell, lastCell)
    ThisWorkbook.Names.Add Name:="USERS_ID", RefersTo:=namedRange
    
    Set namedRange = namedRange.Offset(, 1)
    ThisWorkbook.Names.Add Name:="USERS_NAMES", RefersTo:=namedRange
    
    DisplayUsers = tagsColl.Count

End Function
