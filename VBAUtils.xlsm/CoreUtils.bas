Attribute VB_Name = "CoreUtils"

'Public Function GetNow(Optional format As String = "yymmdd") As String
'Public Sub MakeNamedRangesLocal()
'Public Sub DuplicateLocalNamedRanges()
'Public Sub sortRange(tmpWorksheet As Worksheet, sortRange As Range, sortColumn As Integer)
'Sub DeleteNamedRanges(tmpWorkbook As Workbook, Optional sheetName As String = "ALL")
'Public Function GetSharedPathPrefix(fullPathName As String) As String
'Public Sub testform(controlSourceRangeAddress As String, controlSourceSheetName As String, _
'Sub CreateRefNamedRanges(refSheetName As String, configAddress As String, _
'        headerAddress As String, sourceSheetName As String, Optional rangeOffset As Long = 0, _
'        Optional deleteCurrent As Boolean = True)
'Public Function GetFolderSelection(initFolderPath As String) As String
'Public Function GetFileSelection() As String
'Public Sub SetEventsOn()
'Public Sub SetEventsOff()
'Public Function SheetExists(sheetName As String, book As Workbook) As Boolean
'Function IsInStr(findStr As String, searchStr As String, Optional notFlag As Boolean = False)
'Sub DumpFormats(Optional tmpRange As Range = Nothing)
'Sub SetConditionalFillFormat()

Public Function GetNow(Optional format As String = "yymmdd") As String

    TestGetNow = format(Now(), format)

End Function

Public Sub MakeNamedRangesLocal()
Dim myName As Name
Dim targetWorksheet As Worksheet
Dim targetRange As Range
Dim targetName As String

    For Each myName In ActiveWorkbook.Names
        Set targetRange = myName.RefersToRange
        targetName = myName.Name
        Set targetWorksheet = targetRange.Parent
        myName.Delete
        targetWorksheet.Names.Add Name:=targetName, RefersTo:=targetRange
        
    Next myName
End Sub

Public Sub DuplicateLocalNamedRanges()
Dim myName As Name
Dim targetWorksheet As Worksheet
Dim targetRange As Range
Dim targetRangeAddress As String
Dim targetName As String

    Set targetWorksheet = ActiveWorkbook.Sheets("Next 12 months (20% discount)")

    For Each myName In ActiveSheet.Names
    
        Debug.Print myName.Name
        Set targetRange = targetWorksheet.Range(myName.RefersToRange.Address)
        targetName = Split(myName.Name, "!")(1)

        targetWorksheet.Names.Add Name:=targetName, RefersTo:=targetRange
        
    Next myName
End Sub


Public Sub sortRange(tmpWorksheet As Worksheet, sortRange As Range, sortColumn As Integer)
Dim tmpRange As Range, tmpColumn As Range

    Set tmpColumn = sortRange.Columns(sortColumn)
    
    sortRange.Rows(1).Offset(-1).Select
    Range(Selection, Selection.End(xlDown)).Select
    tmpWorksheet.Sort.SortFields.Clear
    tmpWorksheet.Sort.SortFields.Add2 Key:=tmpColumn, _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With tmpWorksheet.Sort
        .SetRange sortRange.Offset(-1).Resize(sortRange.Rows.Count + 1)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub
Sub DeleteNamedRanges(tmpWorkbook As Workbook, Optional sheetName As String = "ALL")
Dim myName As Name

    For Each myName In tmpWorkbook.Names
        If sheetName = "ALL" Or Left(myName.Name, Len(sheetName)) = UCase(sheetName) Then
            If myName.MacroType = -4142 Then
                If myName.Name = "CLIENTS_ROW_COUNT" Or myName.Name = "OPPORTUNITY_ROW_COUNT" Or myName.Name = "PERSONS_ROW_COUNT" Then
                Else
                    Debug.Print "deleted named range " & myName.Name & " : " & myName.MacroType
                    myName.Delete
                End If
            End If
        'Debug.Print Left(myName, Len(sheetName))
        End If

    Next myName
End Sub


Public Function GetSharedPathPrefix(fullPathName As String) As String
Dim sHostName As String

    sHostName = Environ$("computername")
    
    If Left(fullPathName, 3) = "E:\" Or Left(fullPathName, 3) = "C:\" Then
        GetSharedPathPrefix = fullPathName
    ElseIf sHostName = "DESKTOP-AIODDE8" Then
        GetSharedPathPrefix = "C:\Users\burtn\" & fullPathName
    Else
        GetSharedPathPrefix = "E:\" & fullPathName
    End If
End Function


Public Sub testform(controlSourceRangeAddress As String, controlSourceSheetName As String, _
                    rowSourceRangeAddress As String, rowSourceSheetName As String)
    
    UserForm1.ComboBox1.ControlSource = ActiveWorkbook.Sheets(controlSourceSheetName).Range(controlSourceRangeAddress)
    UserForm1.ComboBox1.RowSource = ActiveWorkbook.Sheets(rowSourceSheetName).Range(rowSourceRangeAddress)
    UserForm1.Show
    
End Sub
'move to vbautils as routine already there

Sub CreateRefNamedRanges(refSheetName As String, configAddress As String, _
        headerAddress As String, sourceSheetName As String, Optional rangeOffset As Long = 0, _
        Optional deleteCurrent As Boolean = True)
Dim numRows As Integer
Dim sheetName As String
Dim rangeName As String, dataRangeName As String
Dim rangeHeight As Long, expRangeHeight As Long
Dim inputRange As Range
Dim sourceSheet As Worksheet
Dim sourceHeaderRange As Range, dataRange As Range, dataTopCell As Range
Dim dataColumnNum As Integer
Dim inputSheet As Worksheet

Dim i As Integer
    
    On Error GoTo err
    'Set inputRange = ActiveWorkbook.Sheets("RANGES").Range("A2:C61")
    Set inputSheet = ActiveWorkbook.Sheets(refSheetName)
    inputSheet.Activate
    Set inputRange = inputSheet.Range(configAddress)

    numRows = inputRange.Rows.Count
    expRangeHeight = 0
    
    For i = 1 To numRows
        sheetName = inputRange.Cells(i, 1).Value
        rangeName = inputRange.Cells(i, 2).Value
        rangeHeight = inputRange.Cells(i, 3).Value
        
        Set sourceSheet = ActiveWorkbook.Sheets(sourceSheetName)
        'Set sourceHeaderRange = sourceSheet.Range("1:1")
        Set sourceHeaderRange = sourceSheet.Range(headerAddress)
        
        
        dataColumnNum = Application.WorksheetFunction.Match(rangeName, sourceHeaderRange, 0)
        
        Set dataTopCell = sourceSheet.Cells(1, dataColumnNum)
        Set dataTopCell = dataTopCell.Offset(rangeOffset)
        sourceSheet.Activate
        
        If rangeHeight = -999 Then
            dataTopCell.Select
            Range(Selection, Selection.End(xlDown)).Select
            expRangeHeight = Selection.Rows.Count
        ElseIf rangeHeight = -1 Then
            If expRangeHeight <> 0 Then
                dataTopCell.Select
                Set dataTopCell = dataTopCell.Resize(expRangeHeight)
                dataTopCell.Select
            Else
                dataTopCell.Select
                Range(Selection, Selection.End(xlDown)).Select
            End If
        Else
            Set dataTopCell = dataTopCell.Resize(rangeHeight)
            dataTopCell.Select
        End If
        
        dataRangeName = Replace(UCase(sheetName) & "_" & UCase(rangeName), " ", "_")
        
        If deleteCurrent = True Then
            On Error Resume Next
            ActiveWorkbook.Names.Item(dataRangeName).Delete
            On Error GoTo 0
        End If
        ActiveWorkbook.Names.Add Name:=dataRangeName, RefersTo:=Selection
        
        Debug.Print sheetName, rangeName, dataColumnNum, dataRangeName, Selection.Address
        
    Next i
    GoTo endsub
    
err:
     Debug.Print "error", sheetName, rangeName
     
endsub:
    On Error GoTo 0
     
End Sub

Public Function GetFolderSelection(initFolderPath As String) As String
Dim sFolder As String
Dim fDialog As FileDialog

    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    fDialog.InitialFileName = initFolderPath
    
    If fDialog.Show = -1 Then ' if OK is pressed
        sFolder = fDialog.SelectedItems(1)
    End If

    GetFolderSelection = sFolder
    
End Function

Public Function GetFileSelection() As String
Dim sFolder As String
    With Application.FileDialog(msoFileDialogFilePicker)
        If .Show = -1 Then ' if OK is pressed
            sFolder = .SelectedItems(1)
        End If
    End With
    GetFileSelection = sFolder
End Function

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
    
    For i = 1 To book.Sheets.Count
    If LCase(book.Worksheets(i).Name) = LCase(sheetName) Then
        SheetExists = True
        Exit Function
    End If
Next i
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


'Selection.Font.Bold = True

Sub DumpFormats(Optional tmpRange As Range = Nothing)
Dim tmpCell As Range

    If tmpRange Is Nothing Then
        Set tmpRange = Selection
        
        For Each tmpCell In tmpRange.Cells
            Debug.Print tmpCell.Address
            With tmpCell.Interior
                Debug.Print "Pattern", .Pattern,
                Debug.Print "TintAndShade", .TintAndShade,
                Debug.Print "Color", .Color,
                Debug.Print "PatternColorIndex", .PatternColorIndex,
                Debug.Print "PatternTintAndShade", .PatternTintAndShade,
                Debug.Print "ThemeColor", .ThemeColor
            End With
        Next tmpCell
    End If
End Sub



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
    numAreas = dataRange.Areas.Count
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
