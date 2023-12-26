Attribute VB_Name = "Module1"

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
Function GetBoardIdArray(latest_flag As Boolean)
    'boardIdArray = Array("4495589558", "4504446525", "2763786972", "4486967796", "2410623120", "2193345626", "2259144314", "1140656959", "2872932195")
    'boardIdArray = Array("1140656959", "2193345626", "2259144314", "2763786972", "4973959122", "4974012540", "4909340518", "4973204278")
    
    If latest_flag = True Then
        GetBoardIdArray = Array("5555786972")
    Else
        GetBoardIdArray = Array("1140656959", "2193345626", "2259144314", "2763786972", "4973959122", "4974012540", "4909340518", "4973204278")
    End If
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

#Const IsDebug = False

Public Sub AddFormulas(sumaryWorkbook As Workbook, boardWorksheet As Worksheet, numRows As Integer, startContentRow As Integer)
Dim tmpWorkbook As Workbook
Dim userName As Variant
Dim origUserName As String, hlinkString As String

Dim columnLink As Range, fillRange As Range, columnItemId As Range, ncolumnItemIdFirstCell As Range

    Set columnLink = boardWorksheet.Range("COLUMN_LINK")
    Set columnItemId = boardWorksheet.Range("COLUMN_ITEMID")
    Set columnItemIdFirstCell = columnItemId.Resize(1).offset(1)
    
    hlinkString = "=hyperlink(" & DQ & MONDAY_URL & DQ & "&B" & startContentRow & "&" & DQ & "/pulses/" & DQ & "&" & columnItemIdFirstCell.Address(0, 0) & ")"
    
    columnLink.Cells(2, 1).Formula = hlinkString
        
    Set fillRange = columnLink.Resize(numRows).offset(1)
    columnLink.Rows(2).Select
    Selection.AutoFill destination:=fillRange
        
    For Each userName In userNamesArray
    
        origUserName = userName
        userName = UCase(userName)
        userName = Replace(userName, SPACE, UNDERLINE)
        
        Set columnLink = boardWorksheet.Range("COLUMN_" & userName)
        columnLink.Cells(2, 1).Formula = "=IF(AND(ISERROR(FIND(" & DQ & origUserName & DQ & ",$J" & startContentRow & ")),ISERROR(FIND(" & DQ & origUserName & DQ & ",$J" & startContentRow & "))),0,1)"
        Set fillRange = columnLink.Resize(numRows).offset(1)
        columnLink.Rows(2).Select
        Selection.AutoFill destination:=fillRange
    Next userName

exitsub:
 
    Set columnLink = Nothing
    Set columnItemId = Nothing
    Set columnItemIdFirstCell = Nothing
    Set fillRange = Nothing
    
 
End Sub
Public Sub AddTopicFilterFormulas(targetSheet As Worksheet, numRows As Integer, mondayPrefix As String, mondaySuffix As String)
Dim filterTopicRange As Range, refFilterTopicRange As Range
Dim filterType As String
Dim colIndex As Long, groupNameColIndex As Long, boardNameColIndex As Long

    Set refFilterTopicRange = ActiveWorkbook.Sheets("Reference").Range("DATA_DATATOPICFILTER")
    
    Set groupNameColIndex = targetSheet.Range("COLUMN_GROUP_NAME").column
    'Set boardNameColIndex = targetSheet.Range("COLUMN_BOARD_NAME").column   ' think need to create this reference
     
    For i = 1 To 10
        On Error GoTo 0
        Set filterTopicRange = targetSheet.Range("COLUMN_FILTER_TOPIC_FILTER" & Str(i))
        On Error Resume Next
        
        If filterTopicRange = Nothing Then
            GoTo exitsub
        Else
            colIndex = filterTopicRange.column
            
            filterType = refFilterTopicRange.Columns(i).Row(1)
            subItemFlag = refFilterTopicRange.Columns(i).Row(3)
            filterName = refFilterTopicRange.Columns(i).Row(4)
            
            '=IF(AND(D6=AU$3^OR(A6=AU$2^A6=AU$1))^1^0)
            'use Address() to give
            'Fill down
        End If

    Next i
    
    ' add in the master filter
exitsub:

End Sub
Public Sub AddUpdatesFormulas(targetSheet As Worksheet, numRows As Integer, mondayPrefix As String, mondaySuffix As String)
Dim sectionUpdatesStartColID As Integer
Dim columnUpdateParentId As Range, columnUpdateItemId As Range, columnUpdateCreator As Range, columnUpdatesStart As Range
Dim columnItemId As Range, columnUpdateUpdatesTime As Range

    DDQ = Chr(34)
    ampersand = Chr(38)
    equals = Chr(61)
    
    ' need to make these more dynamic and not hardcoded
    sectionUpdatesStartColID = targetSheet.Range("COLUMN_UPDATES_START").column
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

    Set fillRange = columnUpdatesStart.offset(1).Resize(numRows)
    fillRange.Rows(1).Select
    Selection.AutoFill destination:=fillRange, Type:=xlFillDefault
    
    ' fill in the update time column (will be a dupe of columnItemID), make the lookup col value dynamic so can fill right
    columnUpdateUpdatesTime.Rows(2).Formula = "=IF(ISERROR(INDEX(INDIRECT(" & columnUpdateUpdatesTime.Rows(1).Address(1, 0) & _
                            ")," & columnUpdatesStart.Rows(2).Address(0, 1) & _
                            "))," & DQ & DQ & ",INDEX(INDIRECT(" & columnUpdateUpdatesTime.Rows(1).Address(1, 0) & _
                            ")," & columnUpdatesStart.Rows(2).Address(0, 1) & _
                            "))"

    ' fill right across the update columns
    Set fillRange = columnUpdateUpdatesTime.Rows(2).Resize(, 5)
    fillRange.Cells(1, 1).Select
    Selection.AutoFill destination:=fillRange, Type:=xlFillDefault

    ' fill down
    Set fillRange = fillRange.Resize(numRows)
    fillRange.Rows(1).Select
    Selection.AutoFill destination:=fillRange, Type:=xlFillDefault

    
Exit Sub:
    Set columnItemId = Nothing
    Set columnUpdateParentId = Nothing
    Set columnUpdateUpdatesTime = Nothing
    Set columnUpdateItemId = Nothing
    Set columnUpdateCreator = Nothing
    Set fillRange = Nothing
    
End Sub

Public Sub DeleteMondayItem(itemId As String, tmpWorkbook As Workbook, mondayFolder As String)
Dim rs As String, rt As String, mondayFolderName As String
    mondayFolderName = GetFolderFromID(itemId, tmpWorkbook)
    
    If mondayFolderName <> "-1" Then
        ' need to delete the folder
    End If
    
    DeleteItem itemId, rs, rt
End Sub
Public Function GetFolderFromID(itemId As String, tmpWorkbook As Workbook) As String
Dim folderNamesCol As Range, itemIdCell As Range

    Set folderNamesCol = tmpWorkbook.Sheets("Folders").Range("FOLDERS_COLUMNS").Columns(1)
    
    For Each itemIdCell In folderNamesCol.Cells
        If InStr(1, itemIdCell.value, itemId) <> 0 Then
            GetFolderFromID = itemIdCell.value
            Exit Function
        End If
    Next itemIdCell
    
    GetFolderFromID = -1
exitfunction:
    Set folderNamesCol = Nothing

End Function
Public Sub AddToMondayFile(itemId As String, mondayFolderPath As String, tmpWorkbook As Workbook, initDocText As String, extraDocText As String, _
                itemName As String, Optional dryRun As Boolean = False)
Dim wordDoc As Word.Document
Dim fso As Scripting.FileSystemObject
Dim wordApp As Word.Application
Dim folderPath As String, mondayFolderName As String, foldername As String

    Set wordDoc = CreateObject("Word.Document")
    Set wordApp = CreateObject("Word.Application")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    mondayFolderName = GetFolderFromID(itemId, tmpWorkbook)
    
    wordDocName = itemId & "__description__" & ".docx"
    
    If mondayFolderName = "-1" Then
        mondayFolderName = itemId & " - " & itemName
        folderPath = fso.BuildPath(mondayFolderPath, mondayFolderName)
        Debug.Print "created folder : " & itemId & " " & itemName
    Else
        folderPath = fso.BuildPath(mondayFolderPath, mondayFolderName)
    End If

    docPath = fso.BuildPath(folderPath, wordDocName)
    
    If dryRun = False Then

        If Dir(docPath) = "" Then
            Set wordDoc = wordApp.Documents.Add
            wordDoc.Content.InsertAfter text:=initDocText
            wordDoc.Content.InsertAfter text:=extraDocText

            Debug.Print "created word doc : " & wordDocName
        Else
            Set wordDoc = wordApp.Documents.Open(docPath)
            
            If InStr(1, wordDoc.Content.text, extraDocText) <> 0 Then
                Debug.Print "content allready exists : " & docPath
                GoTo exitsub
            End If
        
            wordDoc.Content.InsertParagraphAfter
            wordDoc.Content.InsertAfter text:=extraDocText
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
Public Sub InitMondayFolder(itemId As String, itemName As String, mondayFolderPath As String, tmpWorkbook As Workbook, _
    Optional initDocText As String = "", Optional dryRun As Boolean = False)
Dim fso As Scripting.FileSystemObject
Dim foldername As String, folderPath As String, wordDocName As String
Dim wordDoc As Word.Document
Dim wordApp As Word.Application
Dim objSelection As Variant
Dim folderNamesCol As Range, itemIdCell As Range

    Set wordDoc = CreateObject("Word.Document")
    Set wordApp = CreateObject("Word.Application")

    Set folderNamesCol = tmpWorkbook.Sheets("Folders").Range("FOLDERS_COLUMNS").Columns(1)
    
    For Each itemIdCell In folderNamesCol.Cells
        If InStr(1, itemIdCell.value, itemId) <> 0 Then
            GoTo exitsub
        End If
    Next itemIdCell
    wordDocName = itemId & "__description__" & ".docx"

    foldername = itemId & " - " & itemName
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    folderPath = fso.BuildPath(mondayFolderPath, foldername)
    If Dir(folderPath) = "" Then
    
        If dryRun = False Then
            fso.CreateFolder folderPath
            Set wordDoc = wordApp.Documents.Add
            wordDoc.Content.InsertAfter text:=initDocText
            wordDoc.SaveAs2 fso.BuildPath(folderPath, wordDocName)
            wordDoc.Close
        End If
        
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
    sectionFolderStart = targetSheet.Range("COLUMN_FOLDER_START").column
    Set columnEmailRange = targetSheet.Range("COLUMN_EMAIL").Columns
    Set columnItemId = targetSheet.Range("COLUMN_ITEMID").Columns
    Set columnItemId2 = targetSheet.Range("COLUMN_ITEMID2").Columns ' with a prefix for matching
    Set columnBoardId = targetSheet.Range("COLUMN_BOARDID").Columns
    targetSheet.Activate
    Set folderStartRange = targetSheet.Range("COLUMN_FOLDER_START")
    
    folderStartRange.Cells(2, 1).Formula = "=IF(ISERROR(MATCH(" & columnItemId2.Rows(2).Address(0, 0) & ",Folders!I:I,0)),-1,MATCH(" & columnItemId2.Rows(2).Address(0, 0) & ",Folders!I:I,0))"
    Set fillRange = folderStartRange.Resize(numRows).offset(1)
    folderStartRange.Rows(2).Select
    Selection.AutoFill destination:=fillRange
    
    Set folderDataRange = folderStartRange.offset(, 1).Resize(, 8)
    folderDataRange.Cells(2, 1).Formula = "=IF(" & folderStartRange.Cells(2, 1).Address(0, 1) & _
                    "<>-1,INDEX(FOLDER_COLUMNS," & folderStartRange.Cells(2, 1).Address(0, 1) & _
                    ",MATCH(" & folderDataRange.Cells(1, 1).Address(1, 0) & ",FOLDER_HEADERS,0))," & DQ & DQ & ")"
    
    ' fill across the row first
    Set fillRange = folderDataRange.Rows(1).offset(1)
    fillRange.Cells(1, 1).Select
    Selection.AutoFill destination:=fillRange, Type:=xlFillDefault
    
    ' fill down the row first
    Set fillRange = folderDataRange.Resize(numRows).offset(1)
    fillRange.Rows(1).Select
    Selection.AutoFill destination:=fillRange
    
    Set folderLinkRange = folderStartRange.offset(, 9)
    folderLinkRange.Cells(2, 1).Formula = "=hyperlink(IF(" & folderStartRange.Cells(2, 1).Address(0, 1) & _
                    "<>-1,INDEX(FOLDER_COLUMNS," & folderStartRange.Cells(2, 1).Address(0, 1) & _
                    ",MATCH(" & folderLinkRange.Cells(1, 1).Address(1, 0) & ",FOLDER_HEADERS,0))," & DQ & DQ & "))"
    
    Set fillRange = folderLinkRange.Rows(2).Resize(numRows)
    fillRange.Rows(1).Select
    Selection.AutoFill destination:=fillRange
    
    emailLinkFormula = equals & doubleQuote & mondayPrefix & doubleQuote & ampersand & _
                            columnItemId.Cells(2, 1).Address(0, 1) & ampersand & _
                            doubleQuote & mondaySuffix & doubleQuote
    Set columnEmailStartCell = columnEmailRange.Cells(2, 1)
    columnEmailStartCell.Formula = emailLinkFormula
    Set fillRange = columnEmailRange.Resize(numRows).offset(1)
    columnEmailStartCell.Select
    Selection.AutoFill destination:=fillRange
    

    'Set columnBoardName = columnBoardId.offset(, 1)

    'Set boardIdStartCell = columnBoardId.Rows(2)
    'targetSheet.Activate
    'columnBoardName.EntireColumn.Select
    'Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    'boardNameFormula = "=INDEX(DATA_BOARDNAMES,MATCH(int(" & boardIdStartCell.Address(0, 1) & "),DATA_BOARDID,0),1)"
    'Set columnBoardName = columnBoardName.offset(, -1)
    'columnBoardName.Rows(1).Value = "BOARD_NAME"
    'Set boardNameStartCell = columnBoardName.Rows(2)
    'boardNameStartCell.Select
    'boardNameStartCell.Formula = boardNameFormula
    'Set fillRange = columnBoardName.Resize(numRows).offset(1)
    'boardNameStartCell.Select
    'Selection.AutoFill Destination:=fillRange

exitsub:

    Set columnEmailRange = Nothing
    Set columnItemId = Nothing
    Set columnItemId2 = Nothing
    Set columnBoardId = Nothing
    Set folderStartRange = Nothing
    Set fillRange = Nothing
    Set folderLinkRange = Nothing
    Set boardNameStartCell = Nothing
    Set columnBoardName = Nothing
    Set boardIdStartCell = Nothing
    Set columnEmailStartCell = Nothing
    
End Sub


Public Sub AddBoardNameColumn(targetSheet As Worksheet, numRows As Integer, mondayPrefix As String, mondaySuffix As String)
Dim startRange As Range, fillRange As Range, columnEmailRange As Range, columnEmailStartCell As Range, columnItemId As Range, columnItemId2 As Range
Dim columnBoardId As Range, columnBoardName As Range, boardIdStartCell As Range, boardNameStartCell As Range, folderLinkRange As Range
Dim emailLinkFormula As String, boardNameFormula As String
Dim doubleQuote As String
Dim sectionFolderStart As Integer
    
    doubleQuote = Chr(34)
    ampersand = Chr(38)
    equals = Chr(61)

    'need to put an arg to not create column if running from template
    'maybe the best way is to take the column out of the template
    
    Set columnBoardId = targetSheet.Range("COLUMN_BOARDID").Columns
    Set columnBoardName = columnBoardId.offset(, 1)

    Set boardIdStartCell = columnBoardId.Rows(2)
    targetSheet.Activate
    columnBoardName.EntireColumn.Select
    'columnBoardName.EntireColumn.offset(, 1).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    boardNameFormula = "=INDEX(DATA_BOARDNAMES,MATCH(int(" & boardIdStartCell.Address(0, 1) & "),DATA_BOARDID,0),1)"
    Set columnBoardName = columnBoardName.offset(, -1)
    columnBoardName.Rows(1).value = "BOARD_NAME"
    Set boardNameStartCell = columnBoardName.Rows(2)
    boardNameStartCell.Select
    boardNameStartCell.Formula = boardNameFormula
    Set fillRange = columnBoardName.Resize(numRows).offset(1)
    boardNameStartCell.Select
    Selection.AutoFill destination:=fillRange

exitsub:

    Set columnBoardId = Nothing
    Set fillRange = Nothing
    Set boardNameStartCell = Nothing
    Set columnBoardName = Nothing
    Set boardIdStartCell = Nothing
    
End Sub


Public Function GetUpdates(updateFileDir As String, ByVal updatesSheet As Worksheet)
'2022-06-17T08:20:37Z^2790791796^1557226394^Christian Schuler
'H:\My Drive\monday__dump\datafiles\20220617
Dim oFSO As FileSystemObject
Dim tmpWorkbook As Workbook

    dataurl = ActiveWorkbook.Sheets("Reference").Range("dataurl").value
    
    Application.Run "VBAUtils.xlsm!HTTPDownloadFile", _
            dataurl + "/Monday/updates.txt", _
            updatesSheet.Parent, _
            "", "REFERENCE", 1, "start-of-day", updatesSheet.Name, True
            
    'Application.Run "VBAUtils.xlsm!HTTPDownloadFile", _
    '        "http://172.22.237.138/datafiles/Monday/updates.txt", _
    '        updatesSheet.Parent, _
    '        "", "REFERENCE", 1, "start-of-day", updatesSheet.Name, True
    
exitfunction:
    Set oFSO = Nothing
    Set tmpWorkbook = Nothing
    
End Function
Public Function GetSandboxFolder(sandboxDir As String, sandboxSheet As Worksheet)
 
Dim oFSO As Object, oFolders As Object, oFolder As Object, oFile As Object
Dim folderArray() As String, fileList As String
Dim outputRange As Range, columnLink As Range, fillRange As Range, itemLink As Range
Dim i As Integer

'On Error GoTo err
    i = 0

    ' clear previous data but leave headers intact
    folderArray = Application.Run("vbautils.xlsm!GetMondayFolders", sandboxDir)
    sandboxSheet.UsedRange.offset(1).ClearContents

    sandboxSheet.Activate
    With sandboxSheet
        Set outputRange = .Range(Cells(2, 1), Cells(UBound(folderArray) + 1, 9))
        outputRange = folderArray
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
    Set oFSO = Nothing
    Set outputRange = Nothing
    Set columnLink = Nothing
    Set fillRange = Nothing
    Set oFolders = Nothing
    Erase folderArray
    
End Function


Public Sub ApplyFilter(filterSheet As Worksheet, colName As String, colValue As Variant)
Dim colFilterRange As Range, filterRange As Range
Dim colFilterNum As Integer
    
    Set colFilterRange = filterSheet.Range(colName)
    colFilterNum = colFilterRange.column
    Set filterRange = filterSheet.AutoFilter.Range ' find the filter on the sheet

    With filterSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    filterRange.AutoFilter Field:=colFilterNum, Criteria1:=colValue, Operator:=xlFilterValues
exitsub:
    Set colFilterRange = Nothing
    Set filterRange = Nothing
    
End Sub

Public Sub ApplySort(sortSheet As Worksheet, sortColName As String)
Dim colSortRange As Range
Dim colSortNum As Integer

    Set colSortRange = sortSheet.Range(sortColName)
        
    sortSheet.AutoFilter.Sort.SortFields.Clear
    sortSheet.AutoFilter.Sort.SortFields.Add2 Key:=colSortRange, SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With sortSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
exitsub:
     Set colSortRange = Nothing
End Sub

Function GetConfig(RV, configName) As String
    GetConfig = CallByName(RV, configName, VbGet)
    On Error Resume Next
    GetConfig = Replace(Split(GetConfig, "__")(1), "_", " ")
    On Error GoTo 0
End Function
Public Sub GenerateReport(Optional param As String = "")
Dim dirnameString As String, fileNameSuffix As String, parentDirnameString As String, datafileDirname As String, sortValue As String
Dim targetSheetName As String, targetFileName As String, dataFile As String, targetDirName As String, inputUser As String, boardName As String
Dim inputUserName As String, inputDate As String, inputGDrive As String, inputFolder As String, folderSheetName As String, outputFolder As String
Dim fileExtension As String, inputFileExtension As String, outputFolderSheet As String, configSheetName As String, tmpPath As String
Dim mondayPrefix As String, mondaySuffix As String, updatesSheetName As String, outputBookname As String, subitemParentFlag As String, statusFilterFlag As String
Dim tmpWorkbook As Workbook, summaryWorkbook As Workbook, codeWorkbook As Workbook
Dim boardWorksheet As Worksheet, tmpWorksheet As Worksheet, sourceFolderSheet As Worksheet, targetFolderSheet As Worksheet, folderSheet As Worksheet
Dim tmpSheet As Worksheet, sourceUpdatesSheet As Worksheet, targetUpdatesSheet As Worksheet
Dim offsetFactor As Integer, i As Integer, numRows As Integer, sectionFolderStart As Integer
Dim columnLink As Range, startItemsCell As Range
Dim fs As Object
Dim boardIdArray As Variant
Dim RV As RibbonVariables

    'On Error GoTo err
    Set RV = New RibbonVariables

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

    inputUserName = GetConfig(RV, "User")
    inputDate = GetConfig(RV, "Config__Input_Date")
    inputGDrive = GetConfig(RV, "Config__Working_Dir")
    
    If DirExist(inputGDrive) = False Then
        MsgBox "Cannot find GDrive " & inputGDrive
        GoTo err
    End If

    outputFolder = GetConfig(RV, "Config__Working_Dir")
    inputFolder = GetConfig(RV, "Config__Working_Dir")
    
    If DirExist(outputFolder) = False Then
        MsgBox "Cannot find " & outputFolder
        GoTo err
    End If
    
    inputOpenFlag = GetConfig(RV, "OpenReport")
    refreshUpdatesFlag = GetConfig(RV, "RefreshUpdates")
    refreshFolderFlag = GetConfig(RV, "RefreshFolders")
    outputFolderSheet = GetConfig(RV, "Config__Output_Folder_Sheet")
    mondayPrefix = GetConfig(RV, "Config__Monday_Email_Prefix")
    mondaySuffix = GetConfig(RV, "Config__Monday_Email_Suffix")
    subitemParentFlag = GetConfig(RV, "SubItemParent")
    
    inputUser = GetConfig(RV, "User")
    latestFlag = GetConfig(RV, "Latest")
    
    statusFilterFlag = GetConfig(RV, "Config__Status_Filter")
    templateFile = GetConfig(RV, "Config__Template_File")
    
    ' build input absolute path and filename
    parentDirnameString = inputGDrive
    datafileDirname = parentDirnameString
    targetDirName = outputFolder
    targetFileName = fs.BuildPath(targetDirName, viewerFileNameSuffix & "_" & inputDate & "_" & Replace(inputUserName, " ", "") & fileExtension)
    
    ' existing reference sheets
    Set sourceFolderSheet = codeWorkbook.Sheets(folderSheetName)
    Set sourceUpdatesSheet = codeWorkbook.Sheets(updatesSheetName)

    ' create the output book with all its named ranges and sheets
    Debug.Print "Creating Viewer Book : " & Now()
    
    If templateFile = "False" Then
        Set summaryWorkbook = CreateViewerBook(codeWorkbook, outputBookname, inputGDrive, targetDirName)
        Set targetUpdatesSheet = summaryWorkbook.Sheets.Add ' where the underlying updates/posts data goes for reference fro final report sheet
        targetUpdatesSheet.Name = "updates"
        targetUpdatesSheet.Visible = xlSheetHidden
    Else:
        Set summaryWorkbook = Workbooks.Open(templateFile)
        Set targetUpdatesSheet = summaryWorkbook.Sheets("updates")
    End If
    
    Set boardWorksheet = summaryWorkbook.Sheets(targetSheetName) 'this is where the final report goes
    
    Set targetFolderSheet = summaryWorkbook.Sheets(outputFolderSheet) ' where the underlying folder data goes
    Set clipboardSheet = summaryWorkbook.Sheets.Add
    clipboardSheet.Name = "___tmp" ' temp store to collate all the input board data


    Debug.Print Now() & " GetSandboxFolder "
    If refreshFolderFlag = "Yes" Then
        GetSandboxFolder inputFolder, sourceFolderSheet
    End If
    
    Debug.Print Now() & " GetUpdates " & boardName & ".txt"
    If refreshUpdatesFlag = "Yes" Then
        GetUpdates datafileDirname, sourceUpdatesSheet
    End If
    
    offsetFactor = 4
    Set startItemsCell = boardWorksheet.Cells(offsetFactor + 1, 1)
    
    boardIdArray = GetBoardIdArray(CBool(latest_flag))
    
    ' for each monday board input file
    For i = 0 To UBound(boardIdArray)
    
        boardName = boardIdArray(i)
        
        dataurl = Workbooks("MV.xlsm").Sheets("Reference").Range("dataurl").value
        
        Application.Run "VBAUtils.xlsm!HTTPDownloadFile", _
                    dataurl + "/Monday/" & boardName & ".txt", _
                    codeWorkbook, _
                    "A", "REFERENCE", 0, "start-of-day", boardName, False
                    
        Set tmpWorksheet = codeWorkbook.Sheets(boardIdArray(i))
        tmpWorksheet.Activate
        
        ' copy the items from the board
        ActiveSheet.UsedRange.Select
        Selection.Copy
        
        ' paste to temp store
        clipboardSheet.Activate
        clipboardSheet.Range("A1").offset(offsetFactor).Select
        clipboardSheet.Paste

        offsetFactor = offsetFactor + Selection.Rows.Count

        Application.CutCopyMode = False
    Next i

    ' copy from temp and paste into main output report, then delete the temp sheet

    Debug.Print Now() & " Loading " & CStr(Selection.Rows.Count)
    clipboardSheet.Activate
    clipboardSheet.UsedRange.Select
    Selection.Copy
    
    boardWorksheet.Activate
    startItemsCell.Select
    startItemsCell.PasteSpecial Paste:=xlPasteValues
    clipboardSheet.Delete

    Debug.Print Now() & " Adding; formulas; """
    AddFormulas summaryWorkbook, boardWorksheet, offsetFactor, startItemsCell.Row
    startItemsCell.Select
    ActiveWindow.ScrollColumn = 1

    ' copy the folder data into the target  book
    sourceFolderSheet.UsedRange.Copy
    targetFolderSheet.Activate
    targetFolderSheet.Cells(1, 1).PasteSpecial Paste:=xlPasteValues
    
    summaryWorkbook.Names.Add "FOLDER_COLUMNS", RefersTo:=targetFolderSheet.UsedRange

    Debug.Print Now() & "Adding Folder formulas"
    AddFoldersFormulas boardWorksheet, offsetFactor + 1, mondayPrefix, mondaySuffix
    
    ' copy the item updates  into the target  book
    Debug.Print Now() & " Copy AllItems to Target"
    Set sourceUpdatesSheet = codeWorkbook.Sheets(updatesSheetName)
    sourceUpdatesSheet.UsedRange.Copy
    targetUpdatesSheet.Activate
    targetUpdatesSheet.Cells(1, 1).PasteSpecial Paste:=xlPasteValues

    If templateFile = "False" Then
    
        summaryWorkbook.Names.Add "UPDATES_UPDATETIME", RefersTo:=targetUpdatesSheet.UsedRange.Columns(1)
        summaryWorkbook.Names.Add "UPDATES_PARENTID", RefersTo:=targetUpdatesSheet.UsedRange.Columns(2)
        summaryWorkbook.Names.Add "UPDATES_ITEMID", RefersTo:=targetUpdatesSheet.UsedRange.Columns(3)
        summaryWorkbook.Names.Add "UPDATES_CREATOR", RefersTo:=targetUpdatesSheet.UsedRange.Columns(4)
        summaryWorkbook.Names.Add "UPDATES_FIRSTLINE", RefersTo:=targetUpdatesSheet.UsedRange.Columns(5)

        Debug.Print Now() & "Adding search text column"
        AddSearchTextColumn 2000, 4, boardWorksheet, summaryWorkbook, codeWorkbook

    End If

    Debug.Print Now() & "Adding Updates formulas"
    AddUpdatesFormulas boardWorksheet, offsetFactor + 1, mondayPrefix, mondaySuffix

        
    ' add the board name column
    AddBoardNameColumn boardWorksheet, offsetFactor + 1, mondayPrefix, mondaySuffix
    
    Debug.Print Now() & "Applying sorts"
    'ApplySort boardWorksheet, "COLUMN_UPDATED_ON"
    sortValue = "COLUMN_" & Split(RV.Sort, "__")(1)
    ApplySort boardWorksheet, sortValue
    
    Debug.Print Now() & "Applying filters"
    statusFilterFlag = Replace(statusFilterFlag, "_", " ")
    ApplyFilter boardWorksheet, "COLUMN_STATUS", Split(statusFilterFlag, ",")
    
    If subitemParentFlag = "No" Then
        ApplyFilter boardWorksheet, "COLUMN_TYPE", Array("item", "subitem")
    End If

    summaryWorkbook.Activate
    summaryWorkbook.SaveAs FileFormat:=xlOpenXMLWorkbookMacroEnabled, filename:=targetFileName
    
    GoTo exitsub
    
err:
    SetEventsOn
    
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
    Set fs = Nothing
    Set codeWorkbook = Nothing
    Set sourceFolderSheet = Nothing
    Set sourceUpdatesSheet = Nothing
    Set summaryWorkbook = Nothing
    Set boardWorksheet = Nothing
    Set targetFolderSheet = Nothing
    Set clipboardSheet = Nothing
    Set targetUpdatesSheet = Nothing
    Set startItemsCell = Nothing
    Set tmpWorkbook = Nothing
    Set tmpWorksheet = Nothing
    Set RV = Nothing
    
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
    
    Set targetColumnRange = targetColumnRange.Resize(maxRow).offset(startRow)
    targetColumnRange.value = outputArray
    
    Set viewerColumnFilterRange = targetSheet.Range("E1")
    Set viewerColumnFilterName = targetWorkbook.Names.Add("COLUMN_FILTER_SEARCHSTR", RefersTo:=viewerColumnFilterRange)
    viewerColumnFilterRange.Style = "input"
    
    ' this column is added separately and the AddFilterCode function assumes that column needs to be +1 so -1 is to compensate
    Set targetColumnRange = targetColumnRange.offset(, -1)
    AddFilterCode targetWorkbook, targetSheet.Name, targetColumnRange, viewerColumnFilterName
    
exitsub:
    Set myWorkbooks = Nothing
    Set searchTermRange = Nothing
    Set searchTermColIndexRange = Nothing
    Set targetColumnRange = Nothing
    Set viewerColumnFilterRange = Nothing
    Set viewerColumnFilterName = Nothing
    Set targetColumnRange = Nothing
    Erase outputArray
    
End Sub

Sub CopyRefData(targetRangeName As String, sourceRangeName As String, refSheet As Worksheet, genSheet As Worksheet, viewerWorkbook As Workbook)
Dim viewerRange As Range

    
    Set viewerRange = genSheet.Range(sourceRangeName)
    viewerWorkbook.Names.Add targetRangeName, RefersTo:=refSheet.Range(viewerRange.Address)
    viewerRange.Copy
    refSheet.Activate
    refSheet.Range(viewerRange.Address).Select
    refSheet.Paste
    Debug.Print Now() & " copy ref data " & sourceRangeName & " " & Str(viewerRange.Rows.Count)
    
exitsub:
    Set viewerRange = Nothing
    
End Sub
Public Function CreateViewerBook(genBook As Workbook, viewerBookname As String, viewerBookpath As String, outputFolder As String) As Workbook
Dim viewerSheet As Worksheet, viewerGenSheet As Worksheet, viewerRefSheet As Worksheet, viewerFoldersSheet As Worksheet, tmpSheet As Worksheet, viewerLogsSheet As Worksheet
Dim viewerColumnNamesRange As Range, viewerColumnNumRange As Range, viewerColumnRange As Range
Dim viewerGenColumnCell As Range, viewerGenColumnWidthRange As Range, formatCell As Range, viewerGenColumnFormat As Range
Dim viewerColumnFilterRange As Range, viewerGenColumnFilterInputRange As Range, typeCell As Range, subitemParentFlag As Range
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
    
    Set viewerLogsSheet = ActiveWorkbook.Sheets.Add
    viewerLogsSheet.Name = "Logs"
    viewerLogsSheet.Visible = xlSheetHidden
    viewerWorkbook.Names.Add "ACTIVITY_LOG", RefersTo:=viewerLogsSheet.Range("A:J")

    'On Error Resume Next ' hide the default sheet if it exists
    Set tmpSheet = viewerWorkbook.Sheets("Sheet1")
    tmpSheet.Visible = xlSheetHidden
    'On Error GoTo 0
    
    Set viewerGenColumnNameRange = viewerGenSheet.Range("VIEWER_COLUMN_NAMES")
    Set viewerGenColumnWidthRange = viewerGenSheet.Range("VIEWER_COLUMN_WIDTH")
    Set viewerGenColumnFormat = viewerGenSheet.Range("VIEWER_COLUMN_FORMAT")
    Set viewerGenColumnVisibleRange = viewerGenSheet.Range("VIEWER_COLUMN_VISIBLE")
    Set viewerGenColumnRange = viewerGenSheet.Range("VIEWER_COLUMN")
    'Set viewerGenDataBoardnameRange = viewerGenSheet.Range("VIEWER_DATA_BOARDNAMES")
    'Set viewerGenDataBoadrdidRange = viewerGenSheet.Range("VIEWER_DATA_BOARDID")
    Set viewerGenColumnFilterInputRange = viewerGenSheet.Range("VIEWER_COLUMN_FILTER_INPUT")
    Set viewerGenColumnTypeRange = viewerGenSheet.Range("VIEWER_COLUMN_TYPE")
    Set viewerGenColumnHeaderFormatRange = viewerGenSheet.Range("VIEWER_COLUMN_HEADER_FORMAT")
    
    viewerSheet.Activate
        
    AddFilterCalbackSub viewerWorkbook, viewerSheet.Name
        
    For Each viewerGenColumnCell In viewerGenColumnNameRange
        Debug.Print "Creating column " & viewerGenColumnCell.value
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

            Set viewerColumnFilterRange = viewerColumnRange.Resize(1).offset(-2)
            Set viewerColumnFilterName = viewerWorkbook.Names.Add("COLUMN_FILTER_" & viewerGenColumnCell.value, RefersTo:=viewerColumnFilterRange)
            viewerColumnFilterRange.Style = "input"
            'viewerGenColumnFilterInputRange.AddComment = "> for multiple values enter a comma separated list i.e. word1,word2" & vbNewLine & "> for wildcards add a asterisk as suffix/prefix i.e. *word* " & vbNewLine & "> to reset to no filters enter an empty string " & vbNewLine & "> for nonblanks enter <>"
        
            AddFilterCode viewerWorkbook, viewerSheet.Name, viewerColumnFilterRange, viewerColumnFilterName, allOffFlag:=True
            
            ' add a tooltip
            viewerColumnFilterRange.Select
            viewerColumnFilterRange.AddCommentThreaded ( _
                "for multiple values enter a comma separated list i.e. word1,word2, " & Chr(10) & "" & Chr(10) & "for wildcards add a asterisk as suffix/prefix i.e. *word*" & Chr(10) & "" & Chr(10) & "to reset to no filters enter an empty string") & Chr(10) & "" & Chr(10) & "> for nonblanks enter <>"
        
        End If
    Next viewerGenColumnCell
    

    Debug.Print "Adding SendMondayUpdateCode"
    AddSendMondayUpdateCode genBook, viewerWorkbook
    'AddSendMondayUpdateCode genBook, viewerWorkbook, outputFolder & "\tmp.txt"
    
    ' copy across boardid/boardname reference data
    
    CopyRefData "DATA_BOARDNAMES", "VIEWER_DATA_BOARDNAMES", viewerRefSheet, viewerGenSheet, viewerWorkbook
    CopyRefData "DATA_BOARDID", "VIEWER_DATA_BOARDID", viewerRefSheet, viewerGenSheet, viewerWorkbook
    
    CopyRefData "DATA_USERID", "VIEWER_DATA_USERID", viewerRefSheet, viewerGenSheet, viewerWorkbook
    CopyRefData "DATA_USERNAME", "VIEWER_DATA_USERNAME", viewerRefSheet, viewerGenSheet, viewerWorkbook
    CopyRefData "DATA_USEREMAIL", "VIEWER_DATA_USEREMAIL", viewerRefSheet, viewerGenSheet, viewerWorkbook
    CopyRefData "DATA_USERTAG", "VIEWER_DATA_USERTAG", viewerRefSheet, viewerGenSheet, viewerWorkbook
    
    CopyRefData "DATA_TAGNAME", "VIEWER_TAG_NAME", viewerRefSheet, viewerGenSheet, viewerWorkbook
    CopyRefData "DATA_TAGID", "VIEWER_TAG_ID", viewerRefSheet, viewerGenSheet, viewerWorkbook
    
    CopyRefData "DATA_GROUPAME", "VIEWER_GROUP_NAME", viewerRefSheet, viewerGenSheet, viewerWorkbook
    CopyRefData "DATA_GROUPID", "VIEWER_GROUP_ID", viewerRefSheet, viewerGenSheet, viewerWorkbook
    
    CopyRefData "DATA_DATATOPICFILTER", "VIEWER_DATA_TOPIC_FILTER", viewerRefSheet, viewerGenSheet, viewerWorkbook
    
    
    viewerSheet.Activate
    Rows("4:4").Select
    Selection.AutoFilter
    
    Range("H5:H5").Select
    ActiveWindow.FreezePanes = True
    
    ActiveWindow.Zoom = 70

exitfunction:

    Set CreateViewerBook = viewerWorkbook
    Set summaryWorkbook = Nothing
    Set viewerWorkbook = Nothing
    Set viewerGenSheet = Nothing
    Set viewerFoldersSheet = Nothing
    Set viewerSheet = Nothing
    Set viewerLogsSheet = Nothing
    Set viewerRefSheet = Nothing
    Set tmpSheet = Nothing
    Set viewerGenColumnNameRange = Nothing
    Set viewerGenColumnWidthRange = Nothing
    Set viewerGenColumnFormat = Nothing
    Set viewerGenColumnVisibleRange = Nothing
    Set viewerGenColumnRange = Nothing
    Set viewerGenDataBoardnameRange = Nothing
    Set viewerGenDataBoadrdidRange = Nothing
    Set viewerGenColumnFilterInputRange = Nothing
    Set viewerGenColumnTypeRange = Nothing
    Set viewerGenColumnHeaderFormatRange = Nothing
    Set viewerColumnRange = Nothing
    Set formatCell = Nothing
    Set typeCell = Nothing
    Set headerFormatCell = Nothing
    Set viewerColumnFilterRange = Nothing
    Set viewerColumnFilterName = Nothing
    
End Function

Public Function GetFolderSelection(initialFolder As String, selectedFolder As String) As String
Dim sFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select File to Import"
        .InitialFileName = initialFolder
    
        If .Show = -1 Then ' if OK is pressed
            If .SelectedItems(1) = "" Then
                sFolder = selectedFolder
            Else
                sFolder = .SelectedItems(1)
            End If
        End If
    End With
    GetFolderSelection = sFolder
End Function


Public Sub DisplayUsers()

End Sub
Public Sub DisplayGroups(tmpSheet As Worksheet, topLeftCell As Range)
Dim boardsColl As Collection, groupsColl As Collection
Dim rs As String, rt As String
Dim board As Variant, boardGroup As Variant, groupItem As Variant
Dim rowIndex As Integer
Dim nextCell As Range, namedRange As Range, lastCell As Range

    Set nextCell = topLeftCell
    rowIndex = 1
    Set boardsColl = GetBoards(rs, rt)
    
    For Each board In boardsColl

        Set groupsColl = GetGroupsForBoard(CStr(board("id")), rs, rt)
        If groupsColl.Count > 1 Then
            For i = 1 To groupsColl.Count
                nextCell.value = board("id")
                nextCell.offset(, 1).value = board("name")
                nextCell.offset(, 2).value = groupsColl(i)("title")
                nextCell.offset(, 3).value = groupsColl(i)("id")

                Set nextCell = nextCell.offset(1)
            Next i
        End If
    Next board
    
    Set lastCell = nextCell.offset(-1)
    
    'Set lastCell = nextCell.Resize(-1)
    Set namedRange = tmpSheet.Range(topLeftCell, lastCell)
    ThisWorkbook.Names.Add Name:="BOARD_IDS", RefersTo:=namedRange

    Set namedRange = namedRange.offset(, 1)
    ThisWorkbook.Names.Add Name:="BOARD_NAMES", RefersTo:=namedRange
    
    Set namedRange = namedRange.offset(, 1)
    ThisWorkbook.Names.Add Name:="GROUP_NAMES", RefersTo:=namedRange
    
    Set namedRange = namedRange.offset(, 1)
    ThisWorkbook.Names.Add Name:="GROUP_IDS", RefersTo:=namedRange
    
exitsub:
    Set nextCell = Nothing
    Set boardsColl = Nothing
    Set lastCell = Nothing
    Set namedRange = Nothing
End Sub

Public Sub DisplayTags(tmpSheet As Worksheet, topLeftCell As Range)
Dim tagColl As Collection, tagsColl As Collection
Dim rs As String, rt As String
Dim tag As Variant
Dim rowIndex As Integer
Dim nextCell As Range

    Set nextCell = topLeftCell
    rowIndex = 1
    Set tagsColl = GetTags(rs, rt)
    
    For i = 1 To tagsColl.Count
        nextCell.value = tagsColl(i)("id")
        nextCell.offset(, 1).value = tagsColl(i)("name")
        Set nextCell = nextCell.offset(1)
    Next i

    
End Sub

Public Sub DisplayAllBoardColumns()
Dim rs As String, rt As String, boardid As String
Dim columnColl As Collection
Dim column As Variant
Dim userAcct As Dictionary
Dim testItemIDRange As Range, testItemBoardIdRange As Range
Dim configSheet As Worksheet
Dim i As Long

    'dropdown6, dropdown8, dropdown81, dropdown
    Set configSheet = ActiveWorkbook.Sheets("Reference")
    Set testItemIDRange = configSheet.Range("VIEWER_TESTITEM_ITEMID")
    Set testItemBoardIdRange = testItemIDRange.offset(, -1)
    
    For i = 1 To testItemIDRange.Rows.Count
        boardid = testItemBoardIdRange.Rows(i)
        
        Set columnColl = GetBoardColumns(boardid, rs, rt)
        
        For Each column In columnColl
            Debug.Print boardid,
            Debug.Print column.item("title"),
            Debug.Print column.item("type"),
            Debug.Print column.item("id")
        Next column
    Next i
    
exitsub:
    Set configSheet = Nothing
    Set testItemIDRange = Nothing
    Set testItemBoardIdRange = Nothing
    Set columnColl = Nothing
    
End Sub

Public Sub DisplaySubItemColumns()
Dim rs As String, rt As String, boardid As String
Dim columnColl As Collection
Dim column As Variant

    Set columnColl = GetSubitemColumns("1140656959", "3969481240", rs, rt)
    For Each column In columnColl
        Debug.Print boardid,
        Debug.Print column.item("text"),
        Debug.Print column.item("value"),
        Debug.Print column.item("title"),
        Debug.Print column.item("id")
    Next column

exitsub:
    Set columnColl = Nothing
End Sub

Public Sub AddFrequencyColumn()
Dim columnid As String, rs As String, rt As String, itemId As String, boardid As String
Dim testItemIDRange As Range, testItemBoardIdRange As Range
Dim configSheet As Worksheet
Dim i As Long

    Set configSheet = ActiveWorkbook.Sheets("Reference")
    Set testItemIDRange = configSheet.Range("VIEWER_TESTITEM_ITEMID")
    Set testItemBoardIdRange = testItemIDRange.offset(, -1)
    
    For i = 1 To testItemIDRange.Rows.Count
        itemId = testItemIDRange.Rows(i)
        boardid = testItemBoardIdRange.Rows(i)
        
        columnid = CreateDropdownColumn(boardid, "Frequency", rs, rt)
        Debug.Print rs, rt
        SetDropdownColumnValues boardid, itemId, columnid, "One-off, Daily,Weekly,Monthly,Quarterly", rs, rt
        Debug.Print rs, rt
    Next i
exitsub:
    Set configSheet = Nothing
    Set testItemIDRange = Nothing
    Set testItemBoardIdRange = Nothing
End Sub
