Attribute VB_Name = "Main"

'PopulateWordDocFromExcel

Sub PopulateWordDocFromExcel()
' Add a reference to the Word-library via VBE > Tools > References > Microsoft Word xx.x Object Library.
' Create a folder named C:\Temp or edit the filnames in the code.
'
        
    Dim arrayDict As Scripting.Dictionary, recordDict As Scripting.Dictionary
    Dim wordApp As Word.Application
    Dim wordDoc As Word.Document
    Dim wordTable As Word.Table
    Dim bWeStartedWord As Boolean, filesystemObjectExists As Boolean
    Dim wordRange As Word.Range
    Dim offset As Integer, itemRowHeight As Integer, tableCursor As Integer, reportNum As Integer, currentMeetingId As Integer, targetColNumber As Integer, _
        numTableColumns As Integer
    Dim meetingId As String, columnHeader As String, formatColumnHeader As String, valueColor As String
    Dim myName As Name
    Dim mySheet As Worksheet
    Dim meetingIdArray() As Variant, headerArray() As Variant, tmpArray() As Variant, columnArray() As Variant, formatHeaderArray() As Variant, rowdataHeaderArray() As Variant, colorArray() As Variant
    Dim columnValue As Variant, keyName As Variant, keyValueElements() As String
    Dim length As Long, width As Long, columnHeaderCursor As Long, numHeaders As Long, numHeadersFormat As Long, widthFormat As Long, numberHeadersColors As Long, widthHeadersFormat As Long
    Dim columnRange As Range, reportNumRange As Range, firstColumnCell As Range, newColumnRange As Range
    Dim cellValue As Variant
    
    Set arrayDict = New Dictionary

    offset = 1
    itemRowHeight = 5
    numTableColumns = 4

    'On Error Resume Next
    'Set wordApp = GetObject(, "Word.Application")
    'On Error GoTo 0
    If wordApp Is Nothing Then
        Set wordApp = CreateObject("Word.Application")
    End If
    wordApp.Visible = False 'optional!
    wordApp.ScreenUpdating = False
    
    Set wordDoc = wordApp.Documents.Open("C:\Users\burtn\MeetingMinutes_Template.docm")
    
    Set wordRange = wordDoc.Range
    Set wordTable = wordDoc.Tables.Add(wordRange, 55, numTableColumns, wdWord9TableBehavior, wdAutoFitFixed)
    wordTable.Style = "Table Grid"

    wordTable.ApplyStyleHeadingRows = True
    wordTable.ApplyStyleLastRow = False
    wordTable.ApplyStyleFirstColumn = True
    wordTable.ApplyStyleLastColumn = False
    wordTable.ApplyStyleRowBands = True
    wordTable.ApplyStyleColumnBands = False

    wordTable.Range.Font.Name = "Roboto Light"
    wordTable.Range.Font.Size = 10
    
    wordDoc.PageSetup.Orientation = wdOrientLandscape

    wordTable.PreferredWidthType = wdPreferredWidthPoints
    wordTable.PreferredWidth = InchesToPoints(8)
        
    RangeToArray ActiveWorkbook, "INPUT_SHEET", "COLUMN_HEADERS", headerArray, numHeaders, width
    RangeToArray ActiveWorkbook, "INPUT_SHEET", "COLUMN_HEADERS_FORMAT", formatHeaderArray, numHeadersFormat, widthFormat
    RangeToArray ActiveWorkbook, "INPUT_SHEET", "COLUMN_COLORS", colorArray, numberHeadersColors, widthHeadersFormat

    For columnHeaderCursor = 1 To numHeaders
        columnHeader = headerArray(columnHeaderCursor, 1)
        If columnHeader = "-1" Then Exit For
        
        Set myName = Nothing
        On Error Resume Next
        Set myName = ActiveWorkbook.Names.Item(columnHeader)
        On Error GoTo 0
        
        Set mySheet = ActiveWorkbook.Sheets("INPUT_SHEET")
        Set columnRange = mySheet.Range("A:Z")
        Set firstColumnCell = columnRange.Cells(5, columnHeaderCursor)
        Set newColumnRange = firstColumnCell.Resize(5)
        
         
        If Not myName Is Nothing Then
            If myName.RefersToRange.Address <> newColumnRange.Address Then
                Debug.Print "Updating address NamedRange " & columnHeader & " : " & newColumnRange.Address
                myName.RefersTo = newColumnRange
            ElseIf myName.Name <> columnHeader Then
                Debug.Print "Updating name NamedRange " & myName.Name & " : " & columnHeader
                myName.Name = columnHeader
            Else
                Debug.Print "Nothing to change " & columnHeader
            End If
        Else
            Debug.Print "Adding NamedRange " & columnHeader
            ActiveWorkbook.Names.Add columnHeader, newColumnRange
        End If
        
        RangeToArray ActiveWorkbook, "INPUT_SHEET", columnHeader, tmpArray, length, width
        arrayDict.Add columnHeader, tmpArray
    Next

    Set reportNumRange = ActiveWorkbook.Sheets("INPUT_SHEET").Range("ROW_NUM")
    reportNum = reportNumRange.Value
    
    For cursorRowNum = offset To length Step itemRowHeight
        Set recordDict = New Dictionary
        
        ' get an early look at which meeting this is so we can quickly skip
        columnArray = arrayDict.Item("INPUT_MEETING_ID")
        If IsEmpty(columnArray(cursorRowNum, 1)) = True Then Exit For
            
        If columnArray(cursorRowNum, 1) <> reportNum Or reportNum = -1 Then
            Debug.Print "skipping " & columnArray(cursorRowNum, 1)
            GoTo nextcursorrownum
        Else
            Debug.Print "generating " & columnArray(cursorRowNum, 1)
        End If
        
        For i = 1 To numHeaders
            columnHeader = headerArray(i, 1)
            formatColumnHeader = formatHeaderArray(i, 1)
            colorHeader = colorArray(i, 1)
            
            If columnHeader = "-1" Then Exit For
            
            columnArray = arrayDict.Item(columnHeader)
            For j = cursorRowNum To cursorRowNum + itemRowHeight - 1
                columnValue = columnArray(j, 1)

                If columnValue = "" Or columnValue = "NONE" Then
                    GoTo nexti ' no more subrows
                End If

                If columnHeader = "INPUT_MEETING_ID" Then
                    recordDict.Add columnHeader, columnValue
                    recordDict.Add columnHeader & "_FORMAT", formatColumnHeader
                    recordDict.Add columnHeader & "_COLOR", colorHeader
                    GoTo nexti ' no more subrows
                End If
                
                If recordDict.Exists(columnHeader) = True Then
                    recordDict.Item(columnHeader) = recordDict.Item(columnHeader) & "^" & columnValue
                Else
                    recordDict.Add columnHeader, columnValue
                    recordDict.Add columnHeader & "_FORMAT", formatColumnHeader
                    recordDict.Add columnHeader & "_COLOR", colorHeader
                End If
            Next j
nexti:
        Next

            
        tableCursor = 1
        For Each keyName In recordDict.Keys
        
            If Right(keyName, 7) <> "_FORMAT" And Right(keyName, 6) <> "_COLOR" Then
                With wordDoc
                    valueColor = recordDict.Item(keyName & "_COLOR")
                    ' if the field name ends with an underbar and a number the 5 field instances should be rendered as a column not a row
                    ' the number signifies the column to write into
                    If Left(Right(keyName, 2), 1) = "_" Then
                        targetColNumber = Right(keyName, 1)
                        tableCursor = origTableCursor ' go back to original first row
                        keyValueElements = Split(recordDict.Item(keyName), "^")
                        insertValue tableCursor, keyValueElements(0), targetColNumber, wordTable, wordDoc, valueColor
                        
                        If UBound(keyValueElements) > 0 Then
                            For k = 1 To UBound(keyValueElements)
                                tableCursor = tableCursor + 1
    
                                insertValue tableCursor, keyValueElements(k), targetColNumber, wordTable, wordDoc, valueColor
        
    
                            Next k
                        End If
                        
                        tableCursor = tableCursor + 1
                        GoTo nextkeyname
    
                    End If
                    
                    If Left(keyName, 7) = "INPUT_X" Then
                        'keyName = Right(keyName, Len(keyName) - 1)
                        wordTable.Cell(tableCursor, 1).Range.InsertAfter recordDict.Item(keyName & "_FORMAT")
                        
                        For k = 1 To numTableColumns
                            wordTable.Cell(tableCursor, k).Range.Shading.ForegroundPatternColor = wdColorAutomatic
                            wordTable.Cell(tableCursor, k).Range.Shading.BackgroundPatternColor = 4006690
                            wordTable.Cell(tableCursor, k).Range.Font.Color = wdColorWhite
                        Next k
                        wordTable.Cell(tableCursor, 1).Merge MergeTo:=wordTable.Cell(tableCursor, numTableColumns)
                        
                        tableCursor = tableCursor + 1
                        GoTo nextkeyname
                    End If
                    
                    If Left(keyName, 7) = "INPUT_Y" Then
                        tableCursor = tableCursor + 1
                        GoTo nextkeyname
                    End If
                    
                    keyValueElements = Split(recordDict.Item(keyName), "^")
                    If UBound(keyValueElements) > 0 Then
                        origTableCursor = tableCursor
                        insertHeader tableCursor, recordDict.Item(keyName & "_FORMAT"), wordTable
                        insertValue tableCursor, keyValueElements(0), 2, wordTable, wordDoc, valueColor
                        tableCursor = tableCursor + 1
                        For k = 1 To UBound(keyValueElements)
                            insertValue tableCursor, keyValueElements(k), 2, wordTable, wordDoc, valueColor
                            tableCursor = tableCursor + 1
                        Next k
                    Else
                        origTableCursor = tableCursor
                        insertHeader tableCursor, recordDict.Item(keyName & "_FORMAT"), wordTable
                        If Right(keyName, 5) = "VALUE" Then ' this means it should be formatted as a dollar field
                            insertValue tableCursor, recordDict.Item(keyName), 2, wordTable, wordDoc, valueColor, True
                        Else
                            insertValue tableCursor, recordDict.Item(keyName), 2, wordTable, wordDoc, valueColor
                        End If
                    
                        
                    End If
                    
                    If Right(keyName, 1) <> "1" Then ' this is just a 2 column row so merge together the remaining columns
                        wordTable.Cell(tableCursor, 2).Merge MergeTo:=wordTable.Cell(tableCursor, numTableColumns)
                    End If
    
                End With
                tableCursor = tableCursor + 1
            End If
nextkeyname:
        Next keyName

    
    'need a new Document at this point
nextcursorrownum:

    wordApp.ScreenUpdating = True
    wordApp.Visible = True 'optional!
    
    Next cursorRowNum
    
exitsub:

    
    wordApp.Run "CopyTableToClipboard"
    Application.Run "vbautils.xlsm!createEmail", "jon.butler@veloxfintech.com", "jon.butler@veloxfintech.com", "foo", "foo"
    
    clientName = arrayDict.Item("INPUT_OPPORTUNITY_NAME")(1, 1)

    wordDoc.ExportAsFixedFormat OutputFileName:="E:\Velox Financial Technology\Velox Shared Drive - Documents\General\Sales Team\Meeting Summaries\" & clientName & ".pdf", ExportFormat:=wdExportFormatPDF
    
    wordDoc.Name = "E:\Velox Financial Technology\Velox Shared Drive - Documents\General\Sales Team\Meeting Summaries\" & clientName & "_" & GetNow() & ".docx"
    wordApp.Quit
    Set wordDoc = Nothing
    Set wordApp = Nothing
    Set wordRange = Nothing
    Set workTable = Nothing
End Sub

