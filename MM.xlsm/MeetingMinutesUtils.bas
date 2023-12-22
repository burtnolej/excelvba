Attribute VB_Name = "MeetingMinutesUtils"
'GetCellFill
'Color
'CreateRefNamedRanges
'WriteFolderHyperlink
'WriteFileHyperlink
'ChangeInputSheetFocus
'ProcessSelection
'PopupListBox
'PopupTextBox
'SetEventsOff
'SetEventsOn
'SheetExists
'ClearInputContents
'RangeToArray
'RangeFormatsToArray
'ArrayToRange
'IsEmailAddressInternal
'GetEmailDomainFromAddress
'SetConditionalFillFormat
'IsInStr
'ClearNamedRangeContents
'NamedRangeExists
'generateRanges
'insertHeader
'insertValue

    
Sub RefreshMondayData()
Dim outputRange As Range
Dim formulaStr As String

    url = ActiveWorkbook.Sheets("Reference").Range("dataurl").Value + "/Monday/"
    
    'url = "http://172.22.237.138/datafiles/Monday/"


    Application.Run "vbautils.xlsm!SetEventsOff"
    
    On Error Resume Next
    Application.StatusBar = "loading " + url + "5555786972.txt"
    Set outputRange = Application.Run("vbautils.xlsm!HTTPDownloadFile", url + "5555786972.txt", _
                ActiveWorkbook, _
                "", "", 1, "start-of-day", "MONDAY_META", False, 0)
                
    colArray = Array(3, 4, 5)
    formulaArray = Array("""https://veloxfintech.monday.com/boards/""", "B2", """/pulses/""", "G2")
    
    Application.Run "vbautils.xlsm!CreateCalcNamedRange", outputRange.Worksheet, outputRange, "fullItemName", colArray
   
    'https://veloxfintech.monday.com/boards/4973959122/pulses/5409868786
     
    Set outputRange = outputRange.Resize(, outputRange.Columns.count + 1)
    
    Application.Run "vbautils.xlsm!AddNamedRange", outputRange.Worksheet, outputRange, 13, "MONDAY_FULLNAME"
    
    Set outputRange = outputRange.Resize(, outputRange.Columns.count + 1)
    
    Application.Run "vbautils.xlsm!SortRange", outputRange.Worksheet, outputRange, 9
    
    formulaStr = "="
    formulaStr = formulaStr & """https://veloxfintech.monday.com/boards/"""
    formulaStr = formulaStr & " & B1 & "
    formulaStr = formulaStr & """/pulses/"""
    formulaStr = formulaStr & " & G1"
    outputRange.Columns(14).Value = formulaStr
    outputRange.Rows(1).Value = "itemLink"
    
    Application.Run "vbautils.xlsm!AddNamedRange", outputRange.Worksheet, outputRange, 14, "MONDAY_ITEMLINK"
    
    
End Sub
Function RefreshCapsuleData(Optional param As String) As String
Dim outputRange As Range

    On Error GoTo errhandler
    RefreshCapsuleData = "OK"
    'url = "http://172.23.208.38/datafiles/"
    'url = "http://172.22.237.138/datafiles/"

    url = ActiveWorkbook.Sheets("Reference").Range("dataurl").Value
    
    'Application.Run "VBAUtils.xlsm!HTTPDownloadFile", _
    '        dataurl + "/Monday/updates.txt", _
    '        updatesSheet.Parent, _
    '        "", "REFERENCE", 1, "start-of-day", updatesSheet.Name, True
            
    Application.Run "vbautils.xlsm!SetEventsOff"
    
    On Error Resume Next
    Application.StatusBar = "loading " + url + "/entries_meetings.csv"
    Set outputRange = Application.Run("vbautils.xlsm!HTTPDownloadFile", url + "/entries_meetings.csv", _
                ActiveWorkbook, _
                "", "", 0, "start-of-day", "ENTRIES_MEETINGS", False, 0)
                
    Application.Run "vbautils.xlsm!SortRange", outputRange.Worksheet, outputRange, 7
    
    
    Application.StatusBar = "loading " + url + "/datafiles/person.csv"
    Set outputRange = Application.Run("vbautils.xlsm!HTTPDownloadFile", url + "/person.csv", _
                ActiveWorkbook, _
                "", "", 0, "start-of-day", "PERSON", False, 0)
                
    Application.Run "vbautils.xlsm!SortRange", outputRange.Worksheet, outputRange, 13
    
    
    Application.StatusBar = "loading " + url + "/datafiles/opportunities.csv"
    Set outputRange = Application.Run("vbautils.xlsm!HTTPDownloadFile", url + "/opportunities.csv", _
                ActiveWorkbook, _
                "", "", 0, "start-of-day", "OPPORTUNITY", False, 0)
    
    Application.Run "vbautils.xlsm!SortRange", outputRange.Worksheet, outputRange, 15
    
    Application.StatusBar = "loading " + url + "/datafiles/organisation.csv"
    Set outputRange = Application.Run("vbautils.xlsm!HTTPDownloadFile", url + "/organisation.csv", _
                ActiveWorkbook, _
                "", "", 0, "start-of-day", "CLIENT", False, 0)
            
    
    Application.Run "vbautils.xlsm!SortRange", outputRange.Worksheet, outputRange, 7
    'Application.Run "vbautils.xlsm!SortRange", outputRange.Worksheet, outputRange, 8
    On Error GoTo 0
    
    GoTo exitfunc

errhandler:
    RefreshCapsuleData = "FAILED"
    
exitfunc:
    Application.Run "vbautils.xlsm!SetEventsOn"
    Set outputRange = Nothing
    
End Function
Sub insertHeader(tableCursor As Integer, keyName As Variant, wordTable As Word.Table)
    With wordTable.Cell(tableCursor, 1).Range
        .End = .End - 1
        .Collapse wdCollapseEnd
        .Text = keyName
        .Font.Bold = True
        .Shading.BackgroundPatternColor = -553582797
    End With
End Sub

Sub UnpackRGB(rgbString As String, ByRef red As Long, ByRef green As Long, ByRef blue As Long)
Dim colors As Variant

    colors = Split(rgbString, ",")
    red = colors(0)
    green = colors(1)
    blue = colors(2)
    

End Sub
Function GetOneDriveLink(filesystempath As String) As String
Dim pathElements As Variant
Dim urlprefix As String, url As String
Dim i As Integer
Dim pastHeader As Boolean

    pastHeader = False
    pathElements = Split(filesystempath, "\")
    urlprefix = "https://veloxfintechcom.sharepoint.com/sites/VeloxSharedDrive/Shared%20Documents"
    
    url = urlprefix
    For i = 0 To UBound(pathElements)
        If pastHeader = True Then
            url = url & "/" & pathElements(i)
        End If
        If pathElements(i) = "Velox Shared Drive - Documents" Then
            pastHeader = True
        End If

    Next i
    
    GetOneDriveLink = url & "?web=1"


End Function
Sub insertValue(tableCursor As Integer, cellValue As Variant, columNum As Integer, wordTable As Word.Table, wordDoc As Word.Document, Color As String, Optional dollarFormat As Boolean = False)
Dim FSO As Scripting.FileSystemObject
Dim filesystemObjectExists As Boolean
Dim red As Long, green As Long, blue As Long
Dim url As String

    If IsNumeric(cellValue) Then
        If CDec(cellValue) <= 1 Then
            cellValue = Format(cellValue, "#,##0.00%")
        ElseIf dollarFormat = True Then
            cellValue = Format(cellValue, "$#,##0")
        End If
    End If
    
    wordTable.Cell(tableCursor, columNum).Range.InsertAfter cellValue

    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    filesystemObjectExists = False
    On Error Resume Next
    filesystemObjectExists = FSO.folderExists(cellValue)
    If filesystemObjectExists = False Then
        filesystemObjectExists = FSO.fileExists(cellValue)
    End If
    If filesystemObjectExists = False Then
        If Left(cellValue, 4) = "http" Then
            filesystemObjectExists = True
        End If
    End If
    On Error GoTo 0
           
    If filesystemObjectExists = True Then
    
        wordTable.Cell(tableCursor, columNum).Range.Delete
        wordTable.Cell(tableCursor, columNum).Range.InsertAfter FSO.GetBaseName(cellValue)

        'https://veloxfintech.monday.com/boards/4973959122/pulses/5409868786
        'https://www.temenos.com
        'https://www.linkedin.com
        If Left(CStr(cellValue), 5) <> "https" Then
            url = GetOneDriveLink(CStr(cellValue))
        Else
            url = CStr(cellValue)
        End If

        wordDoc.Hyperlinks.Add _
            Anchor:=wordTable.Cell(tableCursor, columNum).Range, _
            Address:=url

    Else
        Debug.Print
    End If
    
    UnpackRGB Color, red, green, blue
    wordTable.Cell(tableCursor, columNum).Shading.BackgroundPatternColor = RGB(red, green, blue)
    
exitsub:
    Set FSO = Nothing
    
End Sub
                    



Sub generateRanges()

    'Application.Run "vbaUtils.xlsm!DeleteNamedRanges", ActiveWorkbook, "CLIENT"
    'Application.Run "vbaUtils.xlsm!DeleteNamedRanges", ActiveWorkbook, "OPPORTUNITY"
    'Application.Run "vbaUtils.xlsm!DeleteNamedRanges", ActiveWorkbook, "PERSONS"
    'Application.Run "vbaUtils.xlsm!DeleteNamedRanges", ActiveWorkbook, "ENTRIES_MEETINGS"
    Application.Run "vbaUtils.xlsm!DeleteNamedRanges", ActiveWorkbook, "INPUT"
    'CreateRefNamedRanges
    
End Sub

Public Sub GetCellFill(Optional param As Variant, Optional param2 As Variant)
Dim length As Long, count As Long
Dim headerRange As Range, resultRange As Range, myCell As Range
Dim resultArray() As Variant
Dim colors As Variant

    Set headerRange = ActiveSheet.Range("COLUMN_HEADERS")
    length = headerRange.Rows.count
    count = 1
    
    ReDim resultArray(1 To length, 1 To 1)
    For Each myCell In headerRange.Cells
        resultArray(count, 1) = Color(ActiveSheet.Range(myCell.Value), 2)
        count = count + 1
    Next myCell
    
    Set resultRange = headerRange.offset(, 2)
    resultRange = resultArray
    
    For Each myCell In resultRange.Cells
        colors = Split(myCell.Value, ",")
        myCell.Interior.Color = RGB(colors(0), colors(1), colors(2))
    Next myCell
    
    
End Sub


Function Color(rng As Range, Optional formatType As Integer = 0) As Variant
    Dim colorVal As Variant

    colorVal = rng.Cells(1, 1).DisplayFormat.Interior.Color
        
    
    Select Case formatType
        Case 1
            Color = WorksheetFunction.Dec2Hex(colorVal, 6)
        Case 2
            Color = (colorVal Mod 256) & ", " & ((colorVal \ 256) Mod 256) & ", " & (colorVal \ 65536)
        Case 3
            Color = rng.Cells(1, 1).Interior.ColorIndex
        Case Else
            Color = colorVal
    End Select
End Function


Sub CreateRefNamedRanges()
Dim numRows As Integer
Dim sheetName As String
Dim rangeName As String, dataRangeName As String
Dim rangeHeight As Long
Dim inputRange As Range
Dim sourceSheet As Worksheet
Dim sourceHeaderRange As Range, dataRange As Range, dataTopCell As Range
Dim dataColumnNum As Integer

Dim i As Integer
    
    'On Error GoTo err
    Set inputRange = ActiveWorkbook.Sheets("RANGES").Range("RANGES")
    numRows = inputRange.Rows.count
    For i = 1 To numRows
        sheetName = inputRange.Cells(i, 1).Value
        rangeName = inputRange.Cells(i, 2).Value
        rangeHeight = inputRange.Cells(i, 3).Value
        
        Set sourceSheet = ActiveWorkbook.Sheets(sheetName)
        Set sourceHeaderRange = sourceSheet.Range("1:1")
        
        dataColumnNum = Application.WorksheetFunction.Match(rangeName, sourceHeaderRange, 0)
        
        Set dataTopCell = sourceSheet.Cells(1, dataColumnNum)
        Set dataTopCell = dataTopCell.offset(1)
        sourceSheet.Activate
        If rangeHeight = -1 Then
            dataTopCell.Select
            Range(Selection, Selection.End(xlDown)).Select
        Else
            Set dataTopCell = dataTopCell.Resize(rangeHeight)
            dataTopCell.Select
        End If
        
        dataRangeName = Replace(UCase(sheetName) & "_" & UCase(rangeName), " ", "_")
        ActiveWorkbook.Names.Add Name:=dataRangeName, RefersTo:=Selection
        
        Debug.Print sheetName, rangeName, dataColumnNum, dataRangeName, Selection.Address
        
    Next i
    GoTo endsub
    
err:
     Debug.Print "error", sheetName, rangeName
     
endsub:
    On Error GoTo 0
     
End Sub
Sub TestCreateMMEmail()
Dim HTMLContent As String
    HTMLContent = "<html lang=" & lang & "><body><table align=" & align & "><tr bgcolor=" & bgcolor & "><td><div align=" & align & "><table bgcolor=" & bgcolor & "><tr><td colspan=" & colspan & ">October, 2023</td></tr><tr><td>foo</td><td>foo</td></tr></table></div></td></tr><table></body></html>"

    CreateMMEmail "E:\Velox Financial Technology\Velox Shared Drive - Documents\General\Sales Team\Meeting Summaries\TT_BUILDING_NEW_UI.htm", _
    "E:\Velox Financial Technology\Velox Shared Drive - Documents\General\Sales Team\Meeting Summaries\TT_BUILDING_NEW_UI.pdf"

End Sub


Function GetClipBoardText() As String
   Dim DataObj As MsForms.DataObject
   Set DataObj = New MsForms.DataObject '<~~ Amended as per jp's suggestion

   On Error GoTo Whoa

   '~~> Get data from the clipboard.
   DataObj.GetFromClipboard

   '~~> Get clipboard contents
   GetClipBoardText = DataObj.GetText(1)

   MsgBox DataObj.GetText(1)
   Exit Function
Whoa:
   If err <> 0 Then MsgBox "Data on clipboard is not text or is empty"
End Function

Sub CreateMMEmail(Optional htmlContentFile As String = "", Optional myAttachment As String = "")
Dim capsuleIdRange As Range, oppClientNameRange As Range, _
    meetingTypeRange As Range, meetingDateRange As Range
Dim capsuleOppEmailAddress As String, emailSubject As String
Dim lang As String, align As String, bgcolor As String, colspan As String, HTMLContent As String

lang = """en"""
align = """center"""
bgcolor = """#22233D"""
colspan = """2"""

    Debug.Print htmlContentFile
    HTMLContent = Application.Run("vbautils.xlsm!FileToString", htmlContentFile)
    'htmlContent = GetClipBoardText
    Set capsuleIdRange = Workbooks("mm.xlsm").Sheets("INPUT_SHEET").Range("CAPSULE_ID")
    Set meetingDateRange = Workbooks("mm.xlsm").Sheets("INPUT_SHEET").Range("MEETING_DATE")
    Set oppClientNameRange = Workbooks("mm.xlsm").Sheets("INPUT_SHEET").Range("OPP_CLIENT_NAME")
    Set meetingTypeRange = Workbooks("mm.xlsm").Sheets("INPUT_SHEET").Range("MEETING_TYPE")
    
    
    emailSubject = oppClientNameRange & " : " & meetingTypeRange & "   [" & CStr(meetingDateRange) & "]"
    capsuleOppEmailAddress = "opportunity+" & capsuleIdRange.Value & "@21628933.veloxfintech.capsulecrm.com"
    Application.Run "vbautils.xlsm!createEmail", "MeetingSummaries@veloxfintech.com", _
            capsuleOppEmailAddress, emailSubject, HTMLContent, myAttachment

    
End Sub
Sub ChangeInputSheetFocus(rangeName As String)
    
    rangeName = Right(rangeName, Len(rangeName) - 3)
    ActiveWindow.ScrollColumn = ActiveSheet.Range(rangeName).Column
    
End Sub
Public Sub ProcessSelection(oppoPartyName As String)
Dim filterString As String

    'ActiveWorkbook.Sheets("LOOKUPS").Range("LOOKUPS_CLIENT_DISPLAY_NAME") = oppoPartyName

End Sub
Public Sub PopupListBox(controlSourceRangeAddress As String, controlSourceSheetName As String, _
                    rowSourceRangeAddress As String, rowSourceSheetName As String, Optional wideFlag As Boolean = False)
    
    If wideFlag = False Then
        UserForm1.ComboBox1.ControlSource = controlSourceSheetName & "!" & controlSourceRangeAddress
        UserForm1.ComboBox1.RowSource = rowSourceSheetName & "!" & rowSourceRangeAddress
        UserForm1.Show
    Else
        UserForm3.ComboBox1.ControlSource = controlSourceSheetName & "!" & controlSourceRangeAddress
        UserForm3.ComboBox1.RowSource = rowSourceSheetName & "!" & rowSourceRangeAddress
        UserForm3.Show
    End If
    
End Sub


Public Sub PopupTextBox(controlSourceRangeAddress As String, controlSourceSheetName As String)
    
 
    UserForm2.TextBox1.ControlTipText = "Type your text here. Enter SHIFT+ENTER to move to a new line."
    UserForm2.TextBox1.ControlSource = controlSourceSheetName & "!" & controlSourceRangeAddress
    UserForm2.TextBox1.MultiLine = True
    UserForm2.Show
    
End Sub

Public Sub WriteFileHyperlink(folderName As String, Target As Range)
Dim hlinkFormula As String, initFolderPath As String, hlinkName As String, DQ As String
Dim selectionResult As Variant
Dim userHomePath As String
Dim FSO As FileSystemObject

    Set FSO = CreateObject("Scripting.FileSystemObject")
    userHomePath = "E:\"
    DQ = Chr(34)

    'https://veloxfintechcom.sharepoint.com/sites/VeloxSharedDrive/Shared%20Documents/General/Sales%20Cycle/In%20Sales%20Process/SANTANDER/082223
    

    initFolderPath = FSO.BuildPath(userHomePath, VELOX_ONEDRIVE_NAME) & _
    folderName & Application.PathSeparator

    selectionResult = Run("VBAUtils.xlsm!GetFileSelection")
    hlinkName = FSO.GetBaseName(selectionResult)
    'hlinkFormula = "=HYPERLINK(" & DQ & selectionResult & DQ & "," & DQ & FSO.GetBaseName(selectionResult) & DQ & ")"
    hlinkFormula = "=HYPERLINK(" & DQ & selectionResult & DQ & "," & DQ & selectionResult & DQ & ")"
    Target.Formula = hlinkFormula

End Sub
Public Sub WriteFolderHyperlink(folderName As String, Target As Range)
Dim hlinkFormula As String, initFolderPath As String, hlinkName As String, DQ As String
Dim selectionResult As Variant
Dim userHomePath As String
Dim FSO As FileSystemObject

    Set FSO = CreateObject("Scripting.FileSystemObject")
    userHomePath = "E:\"
    DQ = Chr(34)


    initFolderPath = FSO.BuildPath(userHomePath, VELOX_ONEDRIVE_NAME) & _
    folderName & Application.PathSeparator

    selectionResult = Run("VBAUtils.xlsm!GetFolderSelection", initFolderPath)
    hlinkName = FSO.GetBaseName(selectionResult)
    'hlinkFormula = "=HYPERLINK(" & DQ & selectionResult & DQ & "," & DQ & FSO.GetBaseName(selectionResult) & DQ & ")"
    hlinkFormula = "=HYPERLINK(" & DQ & selectionResult & DQ & "," & DQ & selectionResult & DQ & ")"
    Target.Formula = hlinkFormula

    'https://veloxfintechcom.sharepoint.com/sites/VeloxSharedDrive/Shared Documents/General/Sales Cycle/In Sales Process/Adaptive/talking points.docx?web=1
    
End Sub


Public Sub SetEventsOn()
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    EVENTSON = True
End Sub

Public Sub SetEventsOff()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    EVENTSON = False
End Sub
Public Function SheetExists(sheetName As String, book As Workbook) As Boolean

    SheetExists = False
    
    For i = 1 To book.Sheets.count
    If LCase(book.Worksheets(i).Name) = LCase(sheetName) Then
        SheetExists = True
        Exit Function
    End If
Next i
End Function


Sub ClearInputContents()
    ClearNamedRangeContents ActiveWorkbook
End Sub
Sub ClearNamedRangeContents(tmpWorkbook As Workbook, Optional rangeNamePrefix As String)
Dim tmpName As Name
Dim tmpNames As Names
Dim inputRangeSet As Range, inputRange As Range

    Set inputRangeSet = ActiveSheet.Range("INPUT_COLUMN_SET")
    
    Set tmpNames = tmpWorkbook.Names
    
    For Each inputRange In inputRangeSet
        On Error Resume Next
        Debug.Print inputRange.Value
        Debug.Print Right(inputRange.Value, Len(inputRange.Value) - 3)
        'Set tmpName = tmpNames.Item(Right(inputRange.Value, Len(inputRange.Value) - 3))
        Set tmpName = tmpNames.Item(inputRange.Value)
        tmpName.RefersToRange.ClearContents
        On Error GoTo 0
    Next inputRange

    
End Sub


Function NamedRangeExists(tmpWorkbook As Workbook, nameName As String) As Boolean
Dim tmpName As Name

    Set tmpName = Nothing
    NamedRangeExists = False
    On Error Resume Next
    Set tmpName = tmpWorkbook.Names(nameName)
    On Error GoTo 0
    
    If tmpName Is Not Nothing Then
        NamedRangeExists = True
    End If
    
End Function
Function IsInStr(findStr As String, searchStr As String, Optional notFlag As Boolean = False)
        
    IsInStr = False
    If InStr(1, searchStr, findStr) <> 0 Then
    
        IsInStr = True
    End If
    
    If notFlag = True Then
        If IsInStr = False Then
            IsInStr = True
        Else
            IsInStr = False
        End If
    End If

End Function


Sub SetConditionalFillFormat()
Dim colorRefArray() As Variant
Dim length As Long, width As Long
Dim numAreas As Integer, areaCount As Integer, colorCount As Integer
Dim dataRange As Range, currentArea As Range, myCell As Range
Dim myThemeColor As Variant, myColor As Variant, myTintAndShade As Variant, myPatternTintAndShade

    SetEventsOff
    
    RangeToArray ActiveWorkbook, "Reference", "COLORRANGE", colorRefArray, length, width
    
    Set dataRange = ActiveWorkbook.Sheets("TopTargetGrid").Range("DATARANGE")
    ActiveWorkbook.Sheets("TopTargetGrid").Activate
    numAreas = dataRange.Areas.count
    For areaCount = 1 To numAreas
        Set currentArea = dataRange.Areas(areaCount)
        For Each myCell In currentArea.Cells
            For colorCount = 1 To width
                If myCell.Value <= colorRefArray(2, colorCount) And myCell.Value >= colorRefArray(1, colorCount) Then
                    Debug.Print myCell.Value
                    myThemeColor = colorRefArray(3, colorCount)
                    myColor = colorRefArray(4, colorCount)
                    myTintAndShade = colorRefArray(5, colorCount)
                    myPatternTintAndShade = colorRefArray(6, colorCount)
                    myCell.Select
                    With Selection.Interior
                        .Pattern = xlSolid
                        .TintAndShade = myTintAndShade
                        .Color = myColor
                        .PatternColorIndex = xlAutomatic
                        .PatternTintAndShade = 0
                        
                        On Error Resume Next
                        .ThemeColor = myThemeColor
                        On Error GoTo 0
                    End With
                End If
            Next colorCount
        Next myCell
    Next areaCount
    
    SetEventsOn

endsub:
    Set dataRange = Nothing
End Sub

Function GetEmailDomainFromAddress(emailAddress As String) As String
Dim domainName As String, msg As String
Dim extArray As Variant
Dim i As Long

    extArray = Array(".com.sa", ".co.uk", ".com", ".io", ".dev", ".net", ".co", ".se", ".dk", ".de", ".uz", ".org", ".eu", ".tech", ".jp", ".us")

    'On Error GoTo err
    domainName = Split(emailAddress, "@")(1)
    
    For i = 0 To UBound(extArray)
        domainName = Replace(domainName, extArray(i), "")
    Next i
    
    'If InStr(1, domainName, ".") <> 0 Then
        'Debug.Print "incorrect domain :" & domainName
    'End If
    
    domainName = Replace(domainName, "'", "") 'on occasion there is an errant '
    
    GetEmailDomainFromAddress = domainName
    'On Error Resume Next

err:
    If err.Number <> 0 Then
        msg = "Error # " & Str(err.Number) & " was generated by " & err.Source & Chr(13) & "Error Line: " & Erl & Chr(13) & err.Description
        'Debug.Print "Error", err.HelpFile, err.HelpContext
    End If
    
End Function

Function IsEmailAddressInternal(emailAddress As String, folderString As String, nextRow As Long) As Boolean

    IsEmailAddressInternal = False
    If Left(LCase(emailAddress), Len("/O=EXCHANGELABS")) = LCase("/O=EXCHANGELABS") Then
        IsEmailAddressInternal = True
        'Application.StatusBar = folderString & ": row:" & Str(nextRow) & emailAddress & ": INTERNAL - SKIPPING"
    End If
End Function



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
    
    tmpRange.Value = tmpArray

    
endfunc:
    Set tmpRange = Nothing
    Set origSheet = Nothing
    Set tmpSheet = Nothing
    
    End Sub



Public Function RangeFormatsToArray(tmpWorkbook As Workbook, sheetNameStr As String, rangeNameStr As String, _
    ByRef tmpArray() As Variant, Optional ByRef length As Long, Optional ByRef width As Long)
Dim tmpRange As Range, tmpCell As Range
Dim tmpSheet As Worksheet
Dim origSheet As Worksheet

    Set origSheet = ActiveSheet

    Set tmpSheet = tmpWorkbook.Sheets(sheetNameStr)
    
    tmpSheet.Activate
    Set tmpRange = tmpSheet.Range(rangeNameStr)
    origSheet.Activate
    
    ReDim tmpArray(1 To tmpRange.Columns.count, 1 To 1)
    For i = 1 To tmpRange.Columns.count
        Debug.Print tmpRange.Cells(1, i).Interior.Color
        If tmpRange.Cells(1, i).Interior.Color = 255 Then
            tmpArray(i, 1) = "255,0,0"
            
        ElseIf tmpRange.Cells(1, i).Interior.Color = 65280 Then
            tmpArray(i, 1) = "0,255,0"
        ElseIf tmpRange.Cells(1, i).Interior.Color = 16711680 Then
             tmpArray(i, 1) = "0,0,255"
        Else
            tmpArray(i, 1) = "255,255,255"
        End If

    Next i
    
    'tmpArray = tmpRange
    width = UBound(tmpArray, 2)
    length = UBound(tmpArray)
    
endfunc:

    Set tmpRange = Nothing
    Set origSheet = Nothing
    Set tmpSheet = Nothing
    
End Function


Public Function RangeToArray(tmpWorkbook As Workbook, sheetNameStr As String, rangeNameStr As String, _
    ByRef tmpArray() As Variant, Optional ByRef length As Long, Optional ByRef width As Long)
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

    Set tmpRange = Nothing
    Set origSheet = Nothing
    Set tmpSheet = Nothing
    
End Function






