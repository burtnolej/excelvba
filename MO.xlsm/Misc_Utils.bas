Attribute VB_Name = "Misc_Utils"
'Public Sub WriteToMondayAPI(itemID As String, msg As String)
'Public Sub ShowMODisplayables()
'Public Sub myParseJson(responseText)
'Public Function getResponseItemid(responseText, Optional itemType As String = "create_item") As String
'Public Function DirExist(filename As String) As Boolean
'Public Sub SetEventsOn()
'Public Sub SetEventsOff()
'Public Sub SetRangeFormat(myRange As Range, bgColor As Variant, fontColor As Variant)
'Public Sub GetCellFormat(myCell As Range, ByRef bgColor As Variant, fontColor As Variant)
'Public Sub CopyCellFormat(fromCell As Range, toRange As Range)
'Public Sub AddFilterCalbackSub(targetBook As Workbook, sheetName As String)
'Public Sub AddFilterCode(targetBook As Workbook, sheetName As String, filterRange As Range, filterName As Name, Optional filterRangeDepth As Long = 2000)
'Public Function GetMondayTimestamp() As String
'Public Sub AddMondayCallbackCode(targetBook As Workbook)
'Public Sub AddSendMondayUpdateCode(sourceBook As Workbook, targetBook As Workbook, Optional tmpFileName As String = "C:\Users\burtn")
'Public Sub AddVBReferences(targetBook As Workbook)
'Public Sub CopyCodeModule(sourceBook As Workbook, targetBook As Workbook, fromModule As String, fromProc As String, toModule As String)
'Public Sub CopyModule(sourceBook As Workbook, targetBook As Workbook, fromModule As String, tmpFileName As String)
'Sub AddReference(targetBook As Workbook, refName As String, refFileName As String)
'Public Sub BatchUpdateDropdownRefData()
'Public Sub UpdateDropdownRefData(ownerString As String, sourceWBStr As String, sourceWSStr As String, initRowNum As Long, initColNum As Long, targetRangeName As String, targetWBStr As String, targetWSStr As String, ByRef prevCopyRowCount As Long, ByRef prevCopyPasteCount As Long, Optional parentOnlyFlag As Boolean = True, Optional overrideCopySizeFlag As Boolean = False)
'Public Sub CreateGroupNameDropdown(dropDownTarget As Range, inputRangeAddress As String, inputSheetName As String)

Public EVENTSON As Boolean
Public Sub WriteToMondayAPI(itemid As String, msg As String)
apiKey = "eyJhbGciOiJIUzI1NiJ9.eyJ0aWQiOjExMTgyMDkwNiwidWlkIjoxNTE2MzEwNywiaWFkIjoiMjAyMS0wNS0zMFQxMTowMDo1OS4wMDBaIiwicGVyIjoibWU6d3JpdGUiLCJhY3RpZCI6NjY5MDk4NSwicmduIjoidXNlMSJ9.zIeOeoqeaZ2Q8NuKBPPw2LQFh2JRPvPwIkhhn4e5Q08"
Url = "https://api.monday.com/v2"

Dim objHTTP As Object
Dim postData As String
Dim DDQ As String

DDQ = Chr(34)

'postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "query {updates (limit: 3) { text_body id item_id created_at creator { name id } } }" & DDQ & "}"
postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "mutation { create_update (item_id: " & itemid & ", body: " & "\" & DDQ & msg & "\" & DDQ & ") {id}}" & DDQ & "}"


Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
objHTTP.Open "POST", Url, False
objHTTP.setRequestHeader "Authorization", apiKey
objHTTP.setRequestHeader "Content-Type", "application/json"
objHTTP.send postData
Debug.Print postData
Debug.Print objHTTP.responseText

'myParseJson objHTTP.responseText

End Sub

Public Sub ShowMODisplayables()

    Windows("MO.xlsm").DisplayFormulas = False
    Windows("MO.xlsm").tmpWindow.DisplayWorkbookTabs = True
    Windows("MO.xlsm").tmpWindow.DisplayHorizontalScrollBar = True
    Windows("MO.xlsm").tmpWindow.DisplayVerticalScrollBar = True
    Windows("MO.xlsm").tmpWindow.DisplayHeadings = True
    
    'Application.Run "vbautils.xlsm!ShowSheets", "MO.xlsm"
    'Application.Run "vbautils.xlsm!ShowDisplayables", "MO.xlsm"
    

End Sub
Public Sub myParseJson(responseText)
Dim Json As Object
Dim d As Dictionary
Dim l As Variant
Dim itemid As String

Dim i As Integer
Set Json = JsonConverter.ParseJson(responseText)

Set d = Json("data")
Set c = d("create_subitem")
itemid = c("id")
For i = 1 To l.Count
    Set d = l(i)
    Debug.Print d("text_body")
Next i

End Sub

Public Function getResponseItemid(responseText, Optional itemType As String = "create_item", Optional param As String = "id") As String

Dim Json As Object
Dim d As Dictionary, c As Dictionary
Dim l As Variant
Dim itemid As String
Dim i As Integer

    Set Json = JsonConverter.ParseJson(responseText)
    
    Set d = Json("data")
    Set c = d(itemType)
    If param <> "id" Then
        Set c = c(param)
    End If
    getResponseItemid = c("id")
    
   '{
   '"data": {
   '  "create_subitem": {
   '    "id": "5768407847",
   '    "board": {
   '      "id": "4978854654"
   '    }
   '  }
   '},
   '"account_id": 6690985
   '}
End Function

Public Function DirExist(fileName As String) As Boolean
 
    DirExist = True
    If Dir(fileName, vbDirectory) = "" Then
        DirExist = False
    End If

End Function

    
Public Sub SetEventsOn()

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    EVENTSON = True
End Sub

Public Sub SetEventsOff()

    Application.ScreenUpdating = False
    'Application.EnableEvents = False
    'Application.DisplayAlerts = False
    'EVENTSON = False
End Sub

Public Sub SetRangeFormat(myRange As Range, bgColor As Variant, fontColor As Variant)

    myRange.Interior.Color = bgColor
    myRange.Font.Color = fontColor
End Sub
Public Sub GetCellFormat(myCell As Range, ByRef bgColor As Variant, fontColor As Variant)

    fontColor = myCell.Font.Color
    bgColor = myCell.Interior.Color
    
End Sub
Public Sub CopyCellFormat(fromCell As Range, toRange As Range)

    toRange.Font.Color = fromCell.Font.Color
    toRange.Font.Bold = fromCell.Font.Bold
    toRange.Interior.Color = fromCell.Interior.Color
    toRange.Borders(xlEdgeBottom).LineStyle = fromCell.Borders(xlEdgeBottom).LineStyle
    toRange.Borders(xlEdgeBottom).Color = fromCell.Borders(xlEdgeBottom).Color
    
    'toRange.Borders(xlEdgeUp).LineStyle = fromCell.Borders(xlEdgeUp).LineStyle
    toRange.Borders(xlEdgeLeft).LineStyle = fromCell.Borders(xlEdgeLeft).LineStyle
    toRange.Borders(xlEdgeLeft).Color = fromCell.Borders(xlEdgeLeft).Color
    toRange.Borders(xlEdgeRight).LineStyle = fromCell.Borders(xlEdgeRight).LineStyle
    toRange.Borders(xlEdgeRight).Color = fromCell.Borders(xlEdgeRight).Color
    
End Sub

Public Sub AddFilterCalbackSub(targetBook As Workbook, sheetName As String)
    Dim wb As Workbook
    Dim xPro As VBIDE.VBProject
    Dim xCom As VBIDE.VBComponent
    Dim xMod As VBIDE.CodeModule
    Dim xLine As Long

    With targetBook
        Set xPro = .VBProject
        Set xCom = xPro.VBComponents("Sheet3")
        Set xMod = xCom.CodeModule

        With xMod

            .InsertLines 1, "Private Sub Worksheet_Change(ByVal Target As Range)"
            .InsertLines 2, "Dim sFolder As String"
            .InsertLines 3, "on error goto err"
            .InsertLines 4, ""
            .InsertLines 5, ""
            .InsertLines 6, ""
            .InsertLines 7, ""
            .InsertLines 8, ""
            .InsertLines 9, ""
            .InsertLines 10, ""
            .InsertLines 11, ""
            .InsertLines 12, ""
            .InsertLines 13, ""
            .InsertLines 14, ""
            .InsertLines 15, ""
            .InsertLines 16, ""
            .InsertLines 17, ""
            .InsertLines 18, ""
            .InsertLines 19, ""
            .InsertLines 20, ""
            .InsertLines 21, ""
            .InsertLines 22, ""
            .InsertLines 23, ""
            .InsertLines 24, ""
            .InsertLines 25, ""
            .InsertLines 26, ""
            .InsertLines 27, ""
            .InsertLines 28, ""
            .InsertLines 29, ""
            .InsertLines 230, "err:"
            .InsertLines 231, "End Sub"

        End With
    End With

End Sub

Public Sub AddFilterCode(targetBook As Workbook, sheetName As String, filterRange As Range, filterName As Name, Optional filterRangeDepth As Long = 2000)
    Dim wb As Workbook
    Dim xPro As VBIDE.VBProject
    Dim xCom As VBIDE.VBComponent
    Dim xMod As VBIDE.CodeModule
    Dim xLineOffset As Long
    
    Dim DDQ As String, EMPTY_STRING As String, COMMA As String
    
    DDQ = Chr(34)
    COMMA = Chr(44)
    QCOMMA = Chr(34) & Chr(44) & Chr(34)
    EMPTY_STRING = Chr(34) & Chr(34)

    With targetBook
        Set xPro = .VBProject
        Set xCom = xPro.VBComponents("Sheet3")
        Set xMod = xCom.CodeModule
        
        ' append new code to the line before :err
        xLineOffset = xMod.CountOfLines - 2
        
        With xMod
            .InsertLines xLineOffset, "    If Target.Address = ActiveSheet.Range(" & DDQ & filterName.Name & DDQ & ").Address Then"
            .InsertLines xLineOffset + 1, "        if Target.Value = vbNullString Then "
            .InsertLines xLineOffset + 2, "            ActiveSheet.Range(" & DDQ & "$A$4:$AZ$" & filterRangeDepth & DDQ & ").AutoFilter Field:=" & filterRange.Column + 1
            .InsertLines xLineOffset + 3, "        ElseIf InStr(Target.Value, " & QCOMMA & ") <> 0 Then"
            .InsertLines xLineOffset + 4, "            criteriaList = Split(Target.Value, " & QCOMMA & ")"
            .InsertLines xLineOffset + 5, "            ActiveSheet.Range(" & DDQ & "$A$4:$AZ$" & filterRangeDepth & DDQ & ").AutoFilter Field:=" & filterRange.Column + 1 & ", Criteria1:=criteriaList, operator:=xlFilterValues"
            .InsertLines xLineOffset + 6, "        else"
            .InsertLines xLineOffset + 7, "            ActiveSheet.Range(" & DDQ & "$A$4:$AZ$" & filterRangeDepth & DDQ & ").AutoFilter Field:=" & filterRange.Column + 1 & ", Criteria1:=Target.Value, operator:=xlFilterValues"
            .InsertLines xLineOffset + 8, "        end if"
            .InsertLines xLineOffset + 9, "            ActiveWindow.ScrollRow = 1"
            .InsertLines xLineOffset + 10, "    End If"
            
            '.InsertLines 14, "    If Target.Address = ActiveSheet.Range(" & DDQ & filterName.Name & DDQ & ").Address Then"
            '.InsertLines 15, "        if Target.Value = vbNullString Then "
            '.InsertLines 16, "            ActiveSheet.Range(" & DDQ & "$A$4:$AJ$" & filterRangeDepth & DDQ & ").AutoFilter Field:=" & filterRange.Column + 1
            '.InsertLines 17, "        ElseIf InStr(Target.Value, " & QCOMMA & ") <> 0 Then"
            '.InsertLines 18, "            criteriaList = Split(Target.Value, " & QCOMMA & ")"
            '.InsertLines 19, "            ActiveSheet.Range(" & DDQ & "$A$4:$AJ$" & filterRangeDepth & DDQ & ").AutoFilter Field:=" & filterRange.Column + 1 & ", Criteria1:=criteriaList, operator:=xlFilterValues"
            '.InsertLines 20, "        else"
            '.InsertLines 21, "            ActiveSheet.Range(" & DDQ & "$A$4:$AJ$" & filterRangeDepth & DDQ & ").AutoFilter Field:=" & filterRange.Column + 1 & ", Criteria1:=Target.Value, operator:=xlFilterValues"
            '.InsertLines 22, "        end if"
            '.InsertLines 23, "            ActiveWindow.ScrollRow = 1"
            '.InsertLines 24, "    End If"
        End With
        
        Set xPro = Nothing
        Set xCom = Nothing
        Set xMod = Nothing
        
    End With

End Sub


Public Sub test2()
    Debug.Print GetMondayTimestamp
End Sub


Public Function GetMondayTimestamp() As String
    
    GetMondayTimestamp = Format(Now(), "YYYY-MM-DD" & "T" & "hh:mm:ss" & "Z")
    'GetMondayTimestamp = Format(Now(), "YYYY-MM-DD") & "T" & Format(Now(), "hh:mm:ss") & "Z"

End Function
Public Sub AddMondayCallbackCode(targetBook As Workbook)
    Dim wb As Workbook
    Dim xPro As VBIDE.VBProject
    Dim xCom As VBIDE.VBComponent
    Dim xMod As VBIDE.CodeModule
    Dim xLine As Long
    
    Dim DDQ As String, EMPTY_STRING As String, COMMA As String
    
    DDQ = Chr(34)
    COMMA = Chr(44)
    QCOMMA = Chr(34) & Chr(44) & Chr(34)
    EMPTY_STRING = Chr(34) & Chr(34)

    With targetBook
        Set xPro = .VBProject
        Set xCom = xPro.VBComponents("Sheet3")
        Set xMod = xCom.CodeModule


        With xMod
        
            .InsertLines 4, "    dim tmpVal as string, itemid as string, boardid as string, responseStatus as string, responseText as string, newStatus As String,itemType as string,subitemid as string"
            .InsertLines 5, "    If Target.Column = ActiveSheet.Range(" & DDQ & "COLUMN_UPDATES_MONDAY_WRITE" & DDQ & ").Column Then"
            .InsertLines 6, "        If Not Target.Value = vbNullString Then"
            .InsertLines 7, "            tmpVal = ActiveSheet.Range(" & DDQ & "COLUMN_ITEMID" & DDQ & ").Rows(Target.Row - 3).Value"

            .InsertLines 8, "            PostUpdateMonday tmpVal, Target.Value,responseStatus,responseText"
            .InsertLines 9, "            If responseStatus=" & DDQ & "200" & DDQ & "Then"
            .InsertLines 10, "                Application.StatusBar = " & DDQ & "Posted update to [" & DDQ & " & tmpVal & " & DDQ & "] to " & DDQ & " & Target.Value"
            .InsertLines 11, "                ActiveSheet.Range(" & DDQ & "COLUMN_UPDATES_UPDATETIME" & DDQ & ").Rows(Target.Row - 3).Value = Format(Now(), " & DDQ & "YYYY-MM-DD" & DDQ & " & " & DDQ & "T" & DDQ & " & " & DDQ & "hh:mm:ss" & DDQ & " & " & DDQ & "Z" & DDQ & ")"
            .InsertLines 12, "                ActiveSheet.Range(" & DDQ & "COLUMN_UPDATES_FIRSTLINE" & DDQ & ").Rows(Target.Row - 3).Value = Target.Value"
            
            .InsertLines 13, "                ActiveSheet.Range(" & DDQ & "COLUMN_UPDATES_UPDATETIME" & DDQ & ").Rows(Target.Row - 3).Font.Color = RGB(255, 0, 0)"
            .InsertLines 14, "                ActiveSheet.Range(" & DDQ & "COLUMN_UPDATES_UPDATETIME" & DDQ & ").Rows(Target.Row - 3).Font.Bold = True"
            .InsertLines 15, "                ActiveSheet.Range(" & DDQ & "COLUMN_UPDATES_FIRSTLINE" & DDQ & ").Rows(Target.Row - 3).Font.Color = RGB(255, 0, 0)"
            .InsertLines 16, "                ActiveSheet.Range(" & DDQ & "COLUMN_UPDATES_FIRSTLINE" & DDQ & ").Rows(Target.Row - 3).Font.Bold = True"
                
            .InsertLines 17, "                Target.Value = " & DDQ & DDQ
                
            .InsertLines 18, "           Else"
            .InsertLines 19, "                Application.StatusBar = " & DDQ & "Failed to post update to [" & DDQ & " & tmpVal & " & DDQ & "] & responseText"
            .InsertLines 20, "           End If"
            .InsertLines 21, "       End If"
            .InsertLines 22, "    End If"

            .InsertLines 23, "    If Target.Column = ActiveSheet.Range(" & DDQ & "COLUMN_UPDATES_MONDAY_STATUS" & DDQ & ").Column Then"
            .InsertLines 24, "        If Not Target.Value = vbNullString Then"
            .InsertLines 25, "            itemid = ActiveSheet.Range(" & DDQ & "COLUMN_ITEMID" & DDQ & ").Rows(Target.Row - 3).Value"
            .InsertLines 26, "            boardid = ActiveSheet.Range(" & DDQ & "COLUMN_BOARDID" & DDQ & ").Rows(Target.Row - 3).Value"
            .InsertLines 26, "            itemType = ActiveSheet.Range(" & DDQ & "COLUMN_TYPE" & DDQ & ").Rows(Target.Row - 3).Value"
            .InsertLines 27, "            if itemType = " & DDQ & "subitem" & DDQ & " then boardid = GetBoardId(CStr(itemid), responseStatus, responseText)"
            
            .InsertLines 28, "            If Target.Value = " & DDQ & "Working" & DDQ & " Then newStatus = " & DDQ & "0" & DDQ & " Else If Target.Value = " & DDQ & "Completed" & DDQ & " Then newStatus = " & DDQ & "1" & DDQ & " Else newStatus = " & DDQ & "3" & DDQ
        
            .InsertLines 29, "            UpdateStatusMonday boardid, itemid, newStatus,responseStatus,responseText"
            
            .InsertLines 30, "            If responseStatus=" & DDQ & "200" & DDQ & "Then"
            .InsertLines 31, "                Application.StatusBar = " & DDQ & "Posted update to [ " & DDQ & " & itemid & " & DDQ & " ] to " & DDQ & " & Target.Value"
            .InsertLines 32, "                ActiveSheet.Range(" & DDQ & "COLUMN_UPDATES_UPDATETIME" & DDQ & ").Rows(Target.Row - 3).Value = Format(Now(), " & DDQ & "YYYY-MM-DD" & DDQ & " & " & DDQ & "T" & DDQ & " & " & DDQ & "hh:mm:ss" & DDQ & " & " & DDQ & "Z" & DDQ & ")"
            .InsertLines 33, "                ActiveSheet.Range(" & DDQ & "COLUMN_STATUS" & DDQ & ").Rows(Target.Row - 3).Value = Target.Value"
            
            .InsertLines 34, "                ActiveSheet.Range(" & DDQ & "COLUMN_UPDATES_UPDATETIME" & DDQ & ").Rows(Target.Row - 3).Font.Color = RGB(255, 0, 0)"
            .InsertLines 35, "                ActiveSheet.Range(" & DDQ & "COLUMN_UPDATES_UPDATETIME" & DDQ & ").Rows(Target.Row - 3).Font.Bold = True"
            .InsertLines 36, "                ActiveSheet.Range(" & DDQ & "COLUMN_STATUS" & DDQ & ").Rows(Target.Row - 3).Font.Color = RGB(255, 0, 0)"
            .InsertLines 37, "                ActiveSheet.Range(" & DDQ & "COLUMN_STATUS" & DDQ & ").Rows(Target.Row - 3).Font.Bold = True"
                
            .InsertLines 38, "                Target.Value = " & DDQ & DDQ
            
            .InsertLines 39, "           Else"
            .InsertLines 40, "                Application.StatusBar = " & DDQ & "Failed to post update to [" & DDQ & " & tmpVal & " & DDQ & "] & responseText"
            .InsertLines 41, "           End If"
            .InsertLines 42, "       End If"
            .InsertLines 43, "    End If"
            
        End With
        
        Set xPro = Nothing
        Set xCom = Nothing
        Set xMod = Nothing
    End With
End Sub

Public Sub AddSendMondayUpdateCode(sourceBook As Workbook, targetBook As Workbook, Optional tmpFileName As String = "C:\Users\burtn")
        
    AddMondayCallbackCode targetBook
    
    CopyCodeModule sourceBook, targetBook, "FORCOPY", "WriteToMondayAPI", "Monday_Utils"
    CopyCodeModule sourceBook, targetBook, "FORCOPY", "PostUpdateMonday", "Monday_Utils"
    CopyCodeModule sourceBook, targetBook, "FORCOPY", "UpdateStatusMonday", "Monday_Utils"
    CopyCodeModule sourceBook, targetBook, "FORCOPY", "GetBoardId", "Monday_Utils"
    CopyCodeModule sourceBook, targetBook, "FORCOPY", "UpdateItemAttributeMonday", "Monday_Utils"
    
    CopyModule sourceBook, targetBook, "JsonConverter", tmpFileName
    
    AddVBReferences targetBook
    
End Sub

Public Sub AddVBReferences(targetBook As Workbook)
    AddReference targetBook, "Scripting", "C:\WINDOWS\system32\scrrun.dll"
    AddReference targetBook, "VBScript_RegExp_55", "C:\Windows\System32\vbscript.dll\3"
    AddReference targetBook, "VBScript_RegExp_10", "C:\Windows\System32\vbscript.dll\2"
    AddReference targetBook, "VBIDE", "C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB"
End Sub


Public Sub CopyCodeModule(sourceBook As Workbook, targetBook As Workbook, _
        fromModule As String, fromProc As String, toModule As String)
    Dim wb As Workbook
    Dim xPro As VBIDE.VBProject
    Dim xComSource As VBIDE.VBComponent, xComTarget As VBIDE.VBComponent
    Dim xModTarget As VBIDE.CodeModule, xModSource As VBIDE.CodeModule
    Dim xCodeSource As String
    Dim xLine As Long, xFirstLine As Long, xProcLength As Long, xTargetLine As Long
    
    With sourceBook
        Set xPro = .VBProject
        Set xComSource = xPro.VBComponents(fromModule)
        Set xModSource = xComSource.CodeModule
        xCodeSource = xModSource.ProcOfLine(1, vbext_pk_Proc)
        xFirstLine = xModSource.ProcBodyLine(fromProc, vbext_pk_Proc)
        xProcLength = xModSource.ProcCountLines(fromProc, vbext_pk_Proc)
    End With
    
    With targetBook
        Set xPro = .VBProject
        
        On Error Resume Next
        Set xComTarget = xPro.VBComponents(toModule)
        On Error GoTo 0
    
        If xComTarget Is Nothing Then
            Set xComTarget = xPro.VBComponents.Add(vbext_ct_StdModule)
            xComTarget.Name = "Monday_Utils"
        End If
        Set xModTarget = xComTarget.CodeModule
        
        With xModTarget
            xTargetLine = 1
            For i = xFirstLine To xFirstLine + xProcLength - 1
                .InsertLines xTargetLine, xModSource.Lines(i, 1)
                xTargetLine = xTargetLine + 1
            Next i
        End With
    End With
    
    Set xPro = Nothing
    Set xComSource = Nothing
    Set xModSource = Nothing
    Set xComTarget = Nothing
    Set xModTarget = Nothing

End Sub
Public Sub CopyModule(sourceBook As Workbook, targetBook As Workbook, fromModule As String, tmpFileName As String)
    Dim wb As Workbook
    Dim xPro As VBIDE.VBProject
    Dim xComSource As VBIDE.VBComponent, xComTarget As VBIDE.VBComponent

    With sourceBook
        Set xPro = .VBProject
        Set xComSource = xPro.VBComponents(fromModule)
        xComSource.Export tmpFileName
        'xComSource.Export "C:/Users/burtn/tmp.txt"
    End With
    
    With targetBook
        Set xPro = .VBProject
        xPro.VBComponents.Import tmpFileName
    End With
    
    Set xPro = Nothing
    Set xComSource = Nothing
    Set xComTarget = Nothing

End Sub

Sub AddReference(targetBook As Workbook, refName As String, refFileName As String)
    Dim VBAEditor As VBIDE.VBE
    Dim vbProj As VBIDE.VBProject
    Dim chkRef As VBIDE.Reference
    Dim BoolExists As Boolean

    Set VBAEditor = Application.VBE
    Set vbProj = targetBook.VBProject

    For Each chkRef In vbProj.References
        If chkRef.Name = refName Then
            BoolExists = True
            GoTo exitsub
        End If
    Next

    vbProj.References.AddFromFile refFileName
exitsub:

    Set vbProj = Nothing
    Set VBAEditor = Nothing
End Sub


Public Function RefreshItemsExec(sourceWBStr As String) As Dictionary
Dim prevCopyRowCount As Long: prevCopyRowCount = 0
Dim prevCopyPasteCount As Long: prevCopyPasteCount = 0
'Dim sourceWBStr As String
Dim sourceWB As Workbook
Dim fso As New FileSystemObject
Dim fileName As String
Dim resultsDict As New Dictionary

    SetEventsOff
    
    'sourceWBStr = ActiveSheet.Range("REFSHEET").value

    fileName = fso.GetFileName(sourceWBStr)

    On Error Resume Next
    Set sourceWB = Workbooks(sourceWBStr)
    On Error GoTo 0
    
    If sourceWB Is Nothing Then
        'Set sourceWB = Workbooks.Open(ActiveSheet.Range("REFSHEET_DIR").value & "\" & sourceWBStr & ".xlsm")
        Set sourceWB = Workbooks.Open(sourceWBStr)
    End If
    'Application.Run ("'Master Calc with Macro.xlsm'!SummarizeMaster")
    
    Application.Run ("'" & fileName & "'!ExpandAll")
    
    'groups
    'UpdateDropdownRefData "Jon Butler", sourceWBStr, "Viewer", 5, 4, "ITEM_GROUP_NAMES", "MondayAddItems.xlsm", "Reference", prevCopyRowCount, prevCopyPasteCount
    resultsDict.Add "ITEM_GROUP_NAMES", UpdateDropdownRefData("Jon Butler", sourceWB.Sheets("Viewer"), 5, 4, "ITEM_GROUP_NAMES", "MO.xlsm", "Reference", prevCopyRowCount, prevCopyPasteCount)
    
    'items
    resultsDict.Add "ITEM_NAMES", UpdateDropdownRefData("Jon Butler", sourceWB.Sheets("Viewer"), 5, 5, "ITEM_NAMES", "MO.xlsm", "Reference", prevCopyRowCount, prevCopyPasteCount)
    resultsDict.Add "ITEM_ITEMIDS", UpdateDropdownRefData("Jon Butler", sourceWB.Sheets("Viewer"), 5, 8, "ITEM_ITEMIDS", "MO.xlsm", "Reference", prevCopyRowCount, prevCopyPasteCount)
    'UpdateDropdownRefData "Jon Butler", sourceWBStr, "Viewer", 5, 11, "ITEM_ITEMIDS", "MondayAddItems.xlsm", "Reference", prevCopyRowCount, prevCopyPasteCount
    
    'subitems
    resultsDict.Add "SUBITEM_ITEMNAMES", UpdateDropdownRefData("Jon Butler", sourceWB.Sheets("Viewer"), 5, 5, "SUBITEM_ITEMNAMES", "MO.xlsm", "Reference", prevCopyRowCount, prevCopyPasteCount, False)
    resultsDict.Add "SUBITEM_SUBITEMIDS", UpdateDropdownRefData("Jon Butler", sourceWB.Sheets("Viewer"), 5, 8, "SUBITEM_SUBITEMIDS", "MO.xlsm", "Reference", prevCopyRowCount, prevCopyPasteCount, False)
    'UpdateDropdownRefData "Jon Butler", sourceWBStr, "Viewer", 5, 11, "SUBITEM_SUBITEMIDS", "MondayAddItems.xlsm", "Reference", prevCopyRowCount, prevCopyPasteCount, False
    resultsDict.Add "SUBITEM_SUBITEMNAMES", UpdateDropdownRefData("Jon Butler", sourceWB.Sheets("Viewer"), 5, 6, "SUBITEM_SUBITEMNAMES", "MO.xlsm", "Reference", prevCopyRowCount, prevCopyPasteCount, False, True)
    
    CloseBook sourceWB
    
    SetEventsOn

    Set RefreshItemsExec = resultsDict
    Set resultsDict = Nothing
    Set fso = Nothing
End Function
Public Function UpdateDropdownRefData(ownerString As String, sourceWS As Worksheet, initRowNum As Long, initColNum As Long, targetRangeName As String, targetWBStr As String, targetWSStr As String, _
                   ByRef prevCopyRowCount As Long, ByRef prevCopyPasteCount As Long, Optional parentOnlyFlag As Boolean = True, Optional overrideCopySizeFlag As Boolean = False) As Long
Dim targetWB As Workbook
Dim targetWS As Worksheet
Dim initSourceCell As Range, initTargetCell As Range, targetRange As Range
Dim targetNamedRange As Name
Dim updatedNamedRangeAddress As String
Dim initTargetCol As Long, initTargetRow As Long

    'SetEventsOff
        
    'On Error Resume Next
    'Set sourceWB = Workbooks(sourceWBStr & ".xlsm")
    'On Error GoTo 0
    
    'If sourceWB Is Nothing Then
    '    Set sourceWB = Workbooks.Open(ActiveSheet.Range("REFSHEET_DIR").value & "\" & sourceWBStr & ".xlsm")
    'End If
    
    'Set sourceWS = sourceWB.Sheets(sourceWSStr)
    
    Set initCell = sourceWS.Cells(initRowNum, initColNum)
    
    Set targetWB = Workbooks(targetWBStr)
    Set targetWS = targetWB.Sheets(targetWSStr)
    
    sourceWS.Activate
    
    Rows("4:4").Select
    Selection.AutoFilter
    Selection.AutoFilter
    
    If parentOnlyFlag = True Then
        sourceWS.Rows("4:4").AutoFilter Field:=1, Criteria1:="=item_parent", Operator:=xlOr, Criteria2:="=item"
        'sourceWS.Rows("4:4").AutoFilter Field:=15, Criteria1:="=Working", Operator:=xlOr, Criteria2:="=Not Started", Criteria2:="=Ongoing"
        'sourceWS.Rows("4:4").AutoFilter Field:=14, Criteria1:="=" & ownerString
    Else
        sourceWS.Range("$A$4:$AQ$1655").AutoFilter Field:=1, Criteria1:="=subitem"
        'sourceWS.Rows("4:4").AutoFilter Field:=15, Criteria1:="=Working", Operator:=xlOr, Criteria2:="=Not Started", Criteria2:="=Ongoing"
        'sourceWS.Rows("4:4").AutoFilter Field:=14, Criteria1:="=" & ownerString
    End If
    
    initCell.Select
    If overrideCopySizeFlag = False Then
        ' if not zero then num of rows was pasted in
        Range(Selection, Selection.End(xlDown)).Select
        Range(Selection, Selection.End(xlDown)).Select
        
        Debug.Print Selection.Address
        initTargetRow = Selection.Rows.SpecialCells(xlCellTypeVisible).Count
        initTargetCol = Selection.Columns.Count
        prevCopyRowCount = Selection.Rows.Count
        prevCopyPasteCount = prevCopyRowCount
    Else
        initCell.Resize(prevCopyRowCount).Select
        initTargetCol = Selection.Columns.Count
        initTargetRow = Selection.Rows.SpecialCells(xlCellTypeVisible).Count
    End If

    Selection.Copy
    
    Set targetNamedRange = targetWB.Names.Item(targetRangeName)
    targetNamedRange.RefersToRange.ClearContents
    
    Set initTargetCell = targetWS.Range(targetNamedRange.RefersToRange.Address(0, 0))
    
    Set targetRange = initTargetCell.Resize(initTargetRow, initTargetCol)
    updatedNamedRangeAddress = "=" & targetWSStr & "!" & targetRange.Address

    targetNamedRange.RefersTo = updatedNamedRangeAddress
    
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy ' for some reason it drops the selection so re select
    'Windows("Reference").Activate
    targetWS.Activate
    targetWS.Activate
    targetNamedRange.RefersToRange.Rows(1).Select
    ActiveSheet.Paste


exitsub:
    UpdateDropdownRefData = targetRange.Rows.Count
    Set sourceWB = Nothing
    Set sourceWS = Nothing
    Set initCell = Nothing
    
    'SetEventsOn
    
End Function

Public Sub CreateGroupNameDropdown(dropDownTarget As Range, inputRangeAddress As String, inputSheetName As String)
Dim rangeAddress As String, sheetName As String, dropdownAddress As String
Dim listRange As Range, inputRange As Range
Dim rangeLength As Integer, outputStartRow As Integer
Dim tmpWorksheet As Worksheet

    'sheetName = "AddNewItems"
    Set tmpWorksheet = ActiveWorkbook.Sheets(inputSheetName)

    outputStartRow = 4

    ActiveWorkbook.Sheets("AddNewItems").Activate
    'ActiveWorkbook.Sheets(inputSheetName).Activate
     
    'Set inputRange = tmpWorksheet.Range(inputRangeAddress)
    Set inputRange = ActiveWorkbook.Sheets("AddNewItems").Range(inputRangeAddress)
    
    
    outputStartColumn = inputRange.Column
    'rangeLength = WorksheetFunction.CountA(tmpWorksheet.Range(inputRangeAddress)) - 1
    rangeLength = WorksheetFunction.CountA(ActiveWorkbook.Sheets("AddNewItems").Range(inputRangeAddress))
    
    'Set listRange = tmpWorksheet.Range(Cells(outputStartRow, outputStartColumn), Cells(outputStartRow + rangeLength, outputStartColumn))
    Set listRange = ActiveWorkbook.Sheets("AddNewItems").Range(Cells(outputStartRow, outputStartColumn), Cells(outputStartRow + rangeLength, outputStartColumn))
    
    'ActiveWorkbook.Sheets("Meeting Notes").Activate
    tmpWorksheet.Activate
    
    tmpWorksheet.Range(dropDownTarget.Address).Select
    'dropDownTarget.Select
    
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=AddNewItems!" & listRange.Address
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
    dropDownTarget.value = listRange.Rows(1) 'set the cell to the first value in the drop down
End Sub

