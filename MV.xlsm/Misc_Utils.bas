Attribute VB_Name = "Misc_Utils"
Public EVENTSON As Boolean
Public Sub WriteToMondayAPI(itemId As String, msg As String)
apiKey = "eyJhbGciOiJIUzI1NiJ9.eyJ0aWQiOjExMTgyMDkwNiwidWlkIjoxNTE2MzEwNywiaWFkIjoiMjAyMS0wNS0zMFQxMTowMDo1OS4wMDBaIiwicGVyIjoibWU6d3JpdGUiLCJhY3RpZCI6NjY5MDk4NSwicmduIjoidXNlMSJ9.zIeOeoqeaZ2Q8NuKBPPw2LQFh2JRPvPwIkhhn4e5Q08"
Url = "https://api.monday.com/v2"

Dim objHTTP As Object
Dim postData As String
Dim DDQ As String

DDQ = Chr(34)

'postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "query {updates (limit: 3) { text_body id item_id created_at creator { name id } } }" & DDQ & "}"
postData = "{" & DDQ & "query" & DDQ & ":" & DDQ & "mutation { create_update (item_id: " & itemId & ", body: " & "\" & DDQ & msg & "\" & DDQ & ") {id}}" & DDQ & "}"


Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
objHTTP.Open "POST", Url, False
objHTTP.setRequestHeader "Authorization", apiKey
objHTTP.setRequestHeader "Content-Type", "application/json"
objHTTP.send postData
Debug.Print postData
Debug.Print objHTTP.responseText

'myParseJson objHTTP.responseText

End Sub

Public Sub myParseJson(responseText)
Dim Json As Object
Dim d As Dictionary
Dim l As Variant
Dim itemId As String

Dim i As Integer
Set Json = JsonConverter.ParseJson(responseText)

Set d = Json("data")
Set C = d("create_subitem")
itemId = C("id")
For i = 1 To l.Count
    Set d = l(i)
    Debug.Print d("text_body")
Next i

End Sub

Public Function getResponseItemid(responseText, Optional itemType As String = "create_item") As String

Dim Json As Object
Dim d As Dictionary
Dim l As Variant
Dim itemId As String
Dim i As Integer

    Set Json = JsonConverter.ParseJson(responseText)
    
    Set d = Json("data")
    Set C = d(itemType)
    getResponseItemid = C("id")
End Function

Public Function DirExist(filename As String) As Boolean
 
    DirExist = True
    If Dir(filename, vbDirectory) = "" Then
        DirExist = False
    End If

End Function

    
Public Sub SetEventsOn()

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    EVENTSON = True
End Sub

Public Sub SetEventsOff()

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    EVENTSON = False
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
        Set xCom = xPro.VBComponents("Sheet4")
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
            .InsertLines 24, "err:"
            .InsertLines 25, "exitsub:"
            .InsertLines 26, "End Sub"

        End With
    End With

End Sub

Public Sub AddFilterCode(targetBook As Workbook, sheetName As String, filterRange As Range, filterName As Name, _
            Optional filterRangeDepth As Long = 2000, Optional allOffFlag As Boolean = False)
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
        Set xCom = xPro.VBComponents("Sheet4")
        Set xMod = xCom.CodeModule
        
        ' append new code to the line before :err
        xLineOffset = xMod.CountOfLines - 2
        
        With xMod
            .InsertLines xLineOffset, "    If Target.Address = ActiveSheet.Range(" & DDQ & filterName.Name & DDQ & ").Address Then"
            .InsertLines xLineOffset + 1, "        if Target.Value = vbNullString Then "
            .InsertLines xLineOffset + 2, "            ActiveSheet.Range(" & DDQ & "$A$4:$AZ$" & filterRangeDepth & DDQ & ").AutoFilter Field:=" & filterRange.column + 1
            If allOffFlag = True Then
                    .InsertLines xLineOffset + 3, "        ElseIf InStr(Target.Value, " & DDQ & "!!" & DDQ & ") <> 0 Then"
                    .InsertLines xLineOffset + 4, "            ActiveSheet.Range(" & DDQ & "$A$4:$AZ$" & filterRangeDepth & DDQ & ").AutoFilter.ShowAllData"
                    .InsertLines xLineOffset + 5, "        Target.Value = " & DDQ & DDQ
            End If
            .InsertLines xLineOffset + 6, "        ElseIf InStr(Target.Value, " & QCOMMA & ") <> 0 Then"
            .InsertLines xLineOffset + 7, "            criteriaList = Split(Target.Value, " & QCOMMA & ")"
            .InsertLines xLineOffset + 8, "            ActiveSheet.Range(" & DDQ & "$A$4:$AZ$" & filterRangeDepth & DDQ & ").AutoFilter Field:=" & filterRange.column + 1 & ", Criteria1:=criteriaList, operator:=xlFilterValues"
            .InsertLines xLineOffset + 9, "        else"
            .InsertLines xLineOffset + 10, "            ActiveSheet.Range(" & DDQ & "$A$4:$AZ$" & filterRangeDepth & DDQ & ").AutoFilter Field:=" & filterRange.column + 1 & ", Criteria1:=Target.Value, operator:=xlFilterValues"
            .InsertLines xLineOffset + 11, "        end if"
            .InsertLines xLineOffset + 12, "            ActiveWindow.ScrollRow = 1"
            .InsertLines xLineOffset + 13, "    End If"

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

Sub AddResponseCode(ByRef xMod As CodeModule, offset As Long, rangeStringA As String, rangeStringB As String)
Dim DDQ As String

    DDQ = Chr(34)
    
    With xMod
            .InsertLines offset + 1, "            If responseStatus=" & DDQ & "200" & DDQ & "Then"
            .InsertLines offset + 2, "                Application.StatusBar = " & DDQ & "Successfullly updated [" & DDQ & " & itemid & " & DDQ & "] to " & DDQ & " & Target.Value & " & DDQ & "[" & DDQ & " & responseText & " & DDQ & "]" & DDQ
            .InsertLines offset + 3, "                ActiveSheet.Range(" & DDQ & rangeStringA & DDQ & ").Rows(Target.Row - 3).Value = Format(Now(), " & DDQ & "YYYY-MM-DD" & DDQ & " & " & DDQ & "T" & DDQ & " & " & DDQ & "hh:mm:ss" & DDQ & " & " & DDQ & "Z" & DDQ & ")"
            .InsertLines offset + 4, "                ActiveSheet.Range(" & DDQ & rangeStringB & DDQ & ").Rows(Target.Row - 3).Value = Target.Value"
            
            .InsertLines offset + 5, "                ActiveSheet.Range(" & DDQ & rangeStringA & DDQ & ").Rows(Target.Row - 3).Font.Color = RGB(255, 0, 0)"
            .InsertLines offset + 6, "                ActiveSheet.Range(" & DDQ & rangeStringA & DDQ & ").Rows(Target.Row - 3).Font.Bold = True"
            .InsertLines offset + 7, "                ActiveSheet.Range(" & DDQ & rangeStringB & DDQ & ").Rows(Target.Row - 3).Font.Color = RGB(255, 0, 0)"
            .InsertLines offset + 8, "                ActiveSheet.Range(" & DDQ & rangeStringB & DDQ & ").Rows(Target.Row - 3).Font.Bold = True"
                
            .InsertLines offset + 9, "                Target.Value = " & DDQ & DDQ
                
            .InsertLines offset + 10, "           Else"
            .InsertLines offset + 11, "                Application.StatusBar = " & DDQ & "Failed to update  [" & DDQ & " & itemid & " & DDQ & "] to " & DDQ & " & Target.Value & " & DDQ & "[" & DDQ & " & responseText & " & DDQ & "]" & DDQ
            .InsertLines offset + 12, "           End If"
    End With
End Sub


            
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
        Set xCom = xPro.VBComponents("Sheet4")
        Set xMod = xCom.CodeModule


        With xMod

            .InsertLines 4, "    dim tmpVal as string, itemid as string, boardid as string, responseStatus as string, responseText as string, newStatus As String,itemType as string,subitemid as string"
            .InsertLines 5, "    Dim userId As Variant, tagId As Variant, tag As Variant"
            .InsertLines 6, "    Dim userIdRange As Range, tagNameRange As Range"
           
            ' ----------------------------------------------------------------
            ' ----------------------------------------------------------------
            .InsertLines 7, "    If Target.Column = ActiveSheet.Range(" & DDQ & "COLUMN_UPDATES_MONDAY_WRITE" & DDQ & ").Column Then"
            .InsertLines 8, "        If Not Target.Value = vbNullString Then"
            .InsertLines 9, "            tmpVal = ActiveSheet.Range(" & DDQ & "COLUMN_ITEMID" & DDQ & ").Rows(Target.Row - 3).Value"
            
            .InsertLines 10, "            PostUpdateMonday tmpVal, Target.Value,responseStatus,responseText"
            
            AddResponseCode xMod, 10, "COLUMN_UPDATES_UPDATETIME", "COLUMN_UPDATES_FIRSTLINE"

            .InsertLines 23, "       End If"
            .InsertLines 24, "       goto ExitSub"
            .InsertLines 25, "    End If"

            ' ----------------------------------------------------------------
            ' ----------------------------------------------------------------
            .InsertLines 26, "    If Target.Column = ActiveSheet.Range(" & DDQ & "COLUMN_UPDATES_MONDAY_STATUS" & DDQ & ").Column Then"
            .InsertLines 27, "        If Not Target.Value = vbNullString Then"
            .InsertLines 28, "            itemid = ActiveSheet.Range(" & DDQ & "COLUMN_ITEMID" & DDQ & ").Rows(Target.Row - 3).Value"
            .InsertLines 29, "            boardid = ActiveSheet.Range(" & DDQ & "COLUMN_BOARDID" & DDQ & ").Rows(Target.Row - 3).Value"
            .InsertLines 30, "            itemType = ActiveSheet.Range(" & DDQ & "COLUMN_TYPE" & DDQ & ").Rows(Target.Row - 3).Value"
            .InsertLines 31, "            if itemType = " & DDQ & "subitem" & DDQ & " then boardid = GetBoardId(CStr(itemid), responseStatus, responseText)"
            .InsertLines 32, "            If target.Value = " & DDQ & "Working" & DDQ & " Then newStatus = " & DDQ & "0" & DDQ & "Else If target.Value = " & DDQ & "Completed" & DDQ & " Then newStatus = " & DDQ & "1" & DDQ & "Else If target.Value = " & DDQ & "Duplicate" & DDQ & " Then newStatus = " & DDQ & "6" & DDQ & " Else If target.Value = " & DDQ & "Ongoing" & DDQ & " Then newStatus = " & DDQ & "7" & DDQ & "Else newStatus = " & DDQ & "7" & DDQ
            .InsertLines 33, "            UpdateStatusMonday boardid, itemid, newStatus,responseStatus,responseText"
            AddResponseCode xMod, 33, "COLUMN_UPDATES_UPDATETIME", "COLUMN_STATUS"
            .InsertLines 46, "       End If"
            .InsertLines 47, "       goto ExitSub"
            .InsertLines 48, "    End If"
            
            ' ----------------------------------------------------------------
            ' ----------------------------------------------------------------
            .InsertLines 49, "     If target.Column = ActiveSheet.Range(" & DDQ & "COLUMN_UPDATES_MONDAY_NAME" & DDQ & ").Column Then"
            .InsertLines 50, "        If Not Target.Value = vbNullString Then"
            .InsertLines 51, "            itemid = ActiveSheet.Range(" & DDQ & "COLUMN_ITEMID" & DDQ & ").Rows(Target.Row - 3).Value"
            .InsertLines 52, "            boardid = ActiveSheet.Range(" & DDQ & "COLUMN_BOARDID" & DDQ & ").Rows(Target.Row - 3).Value"
            
            .InsertLines 53, "            UpdateItemAttributeMonday boardid, itemid, " & DDQ & "name" & DDQ & ", " & "target.Value" & " , responseStatus, responseText"

            AddResponseCode xMod, 53, "COLUMN_UPDATES_UPDATETIME", "COLUMN_ITEM_NAME"
            
            .InsertLines 66, "         End If"
            .InsertLines 67, "       goto ExitSub"
            .InsertLines 68, "     End If"
            
            ' ----------------------------------------------------------------
            ' ----------------------------------------------------------------
            .InsertLines 69, "     If target.Column = ActiveSheet.Range(" & DDQ & "COLUMN_UPDATES_MONDAY_OWNER" & DDQ & ").Column Then"
            .InsertLines 70, "         If Not Target.Value = vbNullString Then"
            .InsertLines 71, "            itemid = ActiveSheet.Range(" & DDQ & "COLUMN_ITEMID" & DDQ & ").Rows(target.Row - 3).Value"
            .InsertLines 72, "            boardID = ActiveSheet.Range(" & DDQ & "COLUMN_BOARDID" & DDQ & ").Rows(target.Row - 3).Value"
            
            .InsertLines 73, "            itemType = ActiveSheet.Range(" & DDQ & "COLUMN_TYPE" & DDQ & ").Rows(target.Row - 3).Value"
            .InsertLines 74, "            If itemType = " & DDQ & "subitem" & DDQ & " Then boardID = GetBoardId(CStr(itemid), responseStatus, responseText)"
            
            
            .InsertLines 75, "            Set userIdRange = Worksheets(" & DDQ & "Reference" & DDQ & ").Range(" & DDQ & "DATA_USERNAME" & DDQ & ")"
            
            .InsertLines 76, "            If IsError(Application.Match(target.Value, userIdRange, 0)) Then"
            .InsertLines 77, "                MsgBox (target.Value & " & DDQ & "not found" & DDQ & ")"
            .InsertLines 78, "                GoTo exitsub"
            .InsertLines 79, "            Else"
            .InsertLines 80, "                userRow = Application.Match(target.Value, userIdRange, 0)"
            .InsertLines 81, "                userId = userIdRange.Rows(userRow).offset(, 3).Value"
            .InsertLines 82, "            End If"
            .InsertLines 83, "            UpdateOwnerMonday boardID, itemid, UserID, responseStatus, responseText"
            
            AddResponseCode xMod, 83, "COLUMN_UPDATES_UPDATETIME", "COLUMN_OWNER"

            .InsertLines 96, "         End If"
            .InsertLines 97, "       goto ExitSub"
            .InsertLines 98, "     End If"
            
            
            
            
            ' ----------------------------------------------------------------
            ' ----------------------------------------------------------------
            .InsertLines 99, "    If Target.Column = ActiveSheet.Range(" & DDQ & "COLUMN_UPDATES_MONDAY_TAGS" & DDQ & ").Column Then"
            .InsertLines 100, "        If Not Target.Value = vbNullString Then"
            .InsertLines 101, "            itemid = ActiveSheet.Range(" & DDQ & "COLUMN_ITEMID" & DDQ & ").Rows(Target.Row - 3).Value"
            .InsertLines 102, "            boardid = ActiveSheet.Range(" & DDQ & "COLUMN_BOARDID" & DDQ & ").Rows(Target.Row - 3).Value"
            .InsertLines 103, "            itemType = ActiveSheet.Range(" & DDQ & "COLUMN_TYPE" & DDQ & ").Rows(Target.Row - 3).Value"
            .InsertLines 104, "            If itemType = " & DDQ & "subitem" & DDQ & " Then boardid = GetBoardId(CStr(itemid), responseStatus, responseText)"
            
            
            .InsertLines 105, "            Set tagNameRange = Worksheets(" & DDQ & "Reference" & DDQ & ").Range(" & DDQ & "DATA_TAGNAME" & DDQ & ")"
            .InsertLines 106, "            tagRow = Application.Match(target.Value, tagNameRange, 0)"
            .InsertLines 107, "            tag = tagNameRange.Rows(tagRow).offset(, -1).Value"

            .InsertLines 108, "            UpdateTagsMonday boardid, itemid, tag,responseStatus,responseText"

            AddResponseCode xMod, 108, "COLUMN_UPDATES_UPDATETIME", "COLUMN_STATUS"
            
            .InsertLines 121, "       End If"
            .InsertLines 122, "       goto ExitSub"
            .InsertLines 123, "    End If"

            
            
        End With
        
        Set xPro = Nothing
        Set xCom = Nothing
        Set xMod = Nothing
    End With
End Sub

Public Sub AddSendMondayUpdateCode(sourceBook As Workbook, targetBook As Workbook, Optional tmpFileName As String = "C:\Users\burtn\tmp.txt")
Dim xProSource As VBIDE.VBProject, xProTarget As VBIDE.VBProject
Dim xComSource As VBIDE.VBComponent, xComTarget As VBIDE.VBComponent


    AddMondayCallbackCode targetBook
    
    Set xProSource = sourceBook.VBProject
    Set xProTarget = targetBook.VBProject
    
    On Error Resume Next
    Set xComTarget = xProTarget.VBComponents("Monday_Utils")
    On Error GoTo 0

    If xComTarget Is Nothing Then
        With xProTarget
            Set xComTarget = .VBComponents.Add(vbext_ct_StdModule)
            xComTarget.Name = "Monday_Utils"
        End With
    End If
        
    Set xComTarget = xProTarget.VBComponents("Monday_Utils")
    Set xComSource = xProSource.VBComponents("FORCOPY")
    
    CopyCodeModule xComSource, xComTarget, sourceBook, targetBook, "FORCOPY", "WriteToMondayAPI", "Monday_Utils"

    CopyCodeModule xComSource, xComTarget, sourceBook, targetBook, "FORCOPY", "PostUpdateMonday", "Monday_Utils"
    CopyCodeModule xComSource, xComTarget, sourceBook, targetBook, "FORCOPY", "UpdateStatusMonday", "Monday_Utils"
    CopyCodeModule xComSource, xComTarget, sourceBook, targetBook, "FORCOPY", "GetBoardId", "Monday_Utils"
    CopyCodeModule xComSource, xComTarget, sourceBook, targetBook, "FORCOPY", "UpdateItemAttributeMonday", "Monday_Utils"
    CopyCodeModule xComSource, xComTarget, sourceBook, targetBook, "FORCOPY", "LogIt", "Monday_Utils"

    CopyCodeModule xComSource, xComTarget, sourceBook, targetBook, "FORCOPY", "UpdateOwnerMonday", "Monday_Utils"
    CopyCodeModule xComSource, xComTarget, sourceBook, targetBook, "FORCOPY", "UpdateTagsMonday", "Monday_Utils"
    CopyCodeModule xComSource, xComTarget, sourceBook, targetBook, "FORCOPY", "BatchUpdateMondayStatus", "Monday_Utils"
    CopyCodeModule xComSource, xComTarget, sourceBook, targetBook, "FORCOPY", "BatchUpdateMondayOwner", "Monday_Utils"
    
    CopyModule sourceBook, targetBook, "JsonConverter", tmpFileName
    
    AddVBReferences targetBook
    
    Set xProS = Nothing
    Set xProT = Nothing
    Set xComTarget = Nothing
    Set xComSource = Nothing
    
End Sub

Public Sub AddVBReferences(targetBook As Workbook)
    AddReference targetBook, "Scripting", "C:\WINDOWS\system32\scrrun.dll"
    AddReference targetBook, "VBScript_RegExp_55", "C:\Windows\System32\vbscript.dll\3"
    AddReference targetBook, "VBScript_RegExp_10", "C:\Windows\System32\vbscript.dll\2"
    AddReference targetBook, "VBIDE", "C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB"
End Sub


Public Sub CopyCodeModule(xComSource As VBIDE.VBComponent, xComTarget As VBIDE.VBComponent, sourceBook As Workbook, targetBook As Workbook, _
        fromModule As String, fromProc As String, toModule As String)
    Dim wb As Workbook
    
    'Dim xComSource As VBIDE.VBComponent, xComTarget As VBIDE.VBComponent
    Dim xModTarget As VBIDE.CodeModule, xModSource As VBIDE.CodeModule
    Dim xCodeSource As String
    Dim xLine As Long, xFirstLine As Long, xProcLength As Long, xTargetLine As Long
    
    Debug.Print "Copying Code " & fromModule & "-" & fromProc
    
    With sourceBook
        'Set xPro = .VBProject
        'Set xComSource = xProSource.VBComponents(fromModule)
        Set xModSource = xComSource.CodeModule
        xCodeSource = xModSource.ProcOfLine(1, vbext_pk_Proc)
        xFirstLine = xModSource.ProcBodyLine(fromProc, vbext_pk_Proc)
        xProcLength = xModSource.ProcCountLines(fromProc, vbext_pk_Proc)
    End With
    
    With targetBook
        'Set xPro = .VBProject
        
        'On Error Resume Next
        'Set xComTarget = xProTarget.VBComponents(toModule)
        'On Error GoTo 0
    
        'If xComTarget Is Nothing Then
        '    Set xComTarget = xProTarget.VBComponents.Add(vbext_ct_StdModule)
        '    xComTarget.Name = "Monday_Utils"
        'End If
        Set xModTarget = xComTarget.CodeModule
        
        With xModTarget
            xTargetLine = 1
            For i = xFirstLine To xFirstLine + xProcLength - 1
                .InsertLines xTargetLine, xModSource.Lines(i, 1)
                xTargetLine = xTargetLine + 1
            Next i
        End With
    End With
    
    'Set xPro = Nothing
    'Set xComSource = Nothing
    Set xModSource = Nothing
    'Set xComTarget = Nothing
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
        Debug.Print Now() & " adding reference " & refName
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


Public Sub BatchUpdateDropdownRefData()
Dim prevCopyRowCount As Long: prevCopyRowCount = 0
Dim prevCopyPasteCount As Long: prevCopyPasteCount = 0
Dim sourceWBStr As String

    sourceWBStr = ActiveSheet.Range("REFSHEET").Value
    
    'groups
    UpdateDropdownRefData "Jon Butler", sourceWBStr, "Viewer", 5, 4, "ITEM_GROUP_NAMES", "monday_report_gen_DEV_v1.10.xlsm", "Reference", prevCopyRowCount, prevCopyPasteCount
    
    'items
    UpdateDropdownRefData "Jon Butler", sourceWBStr, "Viewer", 5, 5, "ITEM_NAMES", "monday_report_gen_DEV_v1.10.xlsm", "Reference", prevCopyRowCount, prevCopyPasteCount
    UpdateDropdownRefData "Jon Butler", sourceWBStr, "Viewer", 5, 10, "ITEM_ITEMIDS", "monday_report_gen_DEV_v1.10.xlsm", "Reference", prevCopyRowCount, prevCopyPasteCount
    
    'subitems
    UpdateDropdownRefData "Jon Butler", sourceWBStr, "Viewer", 5, 5, "SUBITEM_ITEMNAMES", "monday_report_gen_DEV_v1.10.xlsm", "Reference", prevCopyRowCount, prevCopyPasteCount, False
    UpdateDropdownRefData "Jon Butler", sourceWBStr, "Viewer", 5, 10, "SUBITEM_SUBITEMIDS", "monday_report_gen_DEV_v1.10.xlsm", "Reference", prevCopyRowCount, prevCopyPasteCount, False
    UpdateDropdownRefData "Jon Butler", sourceWBStr, "Viewer", 5, 6, "SUBITEM_SUBITEMNAMES", "monday_report_gen_DEV_v1.10.xlsm", "Reference", prevCopyRowCount, prevCopyPasteCount, False, True
    
    
End Sub
Public Sub UpdateDropdownRefData(ownerString As String, sourceWBStr As String, sourceWSStr As String, initRowNum As Long, initColNum As Long, targetRangeName As String, targetWBStr As String, targetWSStr As String, _
                   ByRef prevCopyRowCount As Long, ByRef prevCopyPasteCount As Long, Optional parentOnlyFlag As Boolean = True, Optional overrideCopySizeFlag As Boolean = False)
Dim sourceWB As Workbook, targetWB As Workbook
Dim sourceWS As Worksheet, targetWS As Worksheet
Dim initSourceCell As Range, initTargetCell As Range, targetRange As Range
Dim targetNamedRange As Name
Dim updatedNamedRangeAddress As String
Dim initTargetCol As Long

    'SetEventsOff
        
    Set sourceWB = Workbooks(sourceWBStr)
    Set sourceWS = sourceWB.Sheets(sourceWSStr)
    Set initCell = sourceWS.Cells(initRowNum, initColNum)
    
    Set targetWB = Workbooks(targetWBStr)
    Set targetWS = targetWB.Sheets(targetWSStr)
    
    sourceWS.Activate
    
    Rows("4:4").Select
    Selection.AutoFilter
    Selection.AutoFilter
    
    If parentOnlyFlag = True Then
        sourceWS.Rows("4:4").AutoFilter Field:=1, Criteria1:="=item_parent", Operator:=xlOr, Criteria2:="=item"
        sourceWS.Rows("4:4").AutoFilter Field:=14, Criteria1:="=Working", Operator:=xlOr, Criteria2:="=Not Started", Criteria2:="=Ongoing"
        sourceWS.Rows("4:4").AutoFilter Field:=13, Criteria1:="=" & ownerString
    Else
        sourceWS.Range("$A$4:$AQ$1655").AutoFilter Field:=1, Criteria1:="=subitem"
        sourceWS.Rows("4:4").AutoFilter Field:=14, Criteria1:="=Working", Operator:=xlOr, Criteria2:="=Not Started", Criteria2:="=Ongoing"
        sourceWS.Rows("4:4").AutoFilter Field:=13, Criteria1:="=" & ownerString
    End If
    
    initCell.Select
    If overrideCopySizeFlag = False Then
        ' if not zero then num of rows was pasted in
        Range(Selection, Selection.End(xlDown)).Select
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
    
    Set targetNamedRange = targetWB.Names.item(targetRangeName)
    targetNamedRange.RefersToRange.ClearContents
    
    Set initTargetCell = targetWS.Range(targetNamedRange.RefersToRange.Address(0, 0))
    
    Set targetRange = initTargetCell.Resize(initTargetRow, initTargetCol)
    updatedNamedRangeAddress = "=" & targetWSStr & "!" & targetRange.Address

    targetNamedRange.RefersTo = updatedNamedRangeAddress
    
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy ' for some reason it drops the selection so re select
    Windows("monday_report_gen_DEV_v1.10.xlsm").Activate
    targetNamedRange.RefersToRange.Rows(1).Select
    ActiveSheet.Paste


exitsub:
    Set sourceWB = Nothing
    Set sourceWS = Nothing
    Set initCell = Nothing
    
    'SetEventsOn
    
End Sub

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
    
    
    outputStartColumn = inputRange.column
    'rangeLength = WorksheetFunction.CountA(tmpWorksheet.Range(inputRangeAddress)) - 1
    rangeLength = WorksheetFunction.CountA(ActiveWorkbook.Sheets("AddNewItems").Range(inputRangeAddress)) - 1
    
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
    
    dropDownTarget.Value = listRange.Rows(1) 'set the cell to the first value in the drop down
End Sub

Sub DumpReferences()
    Dim VBAEditor As VBIDE.VBE
    Dim vbProj As VBIDE.VBProject
    Dim chkRef As VBIDE.Reference
    Dim BoolExists As Boolean

    Set VBAEditor = Application.VBE
    Set vbProj = ActiveWorkbook.VBProject

    For Each chkRef In vbProj.References
        Debug.Print chkRef.Name, chkRef.FullPath
    Next
End Sub

Public Sub AddReferences()
Dim nameArray As Variant, pathArray As Variant
Dim i As Integer
Dim VBAEditor As VBIDE.VBE
Dim vbProj As VBIDE.VBProject
Dim chkRef As VBIDE.Reference
Dim BoolExists As Boolean
Dim resultStr As String

Set VBAEditor = Application.VBE
ActiveWorkbook.Activate
Set vbProj = ActiveWorkbook.VBProject


nameArray = Array("VBA", "Excel", "stdole", "Office", "MSForms", "Outlook", "VBIDE", "Scripting", "PowerPoint", "VBScript_RegExp_55", "VBScript_RegExp_10", "Word")
pathArray = Array("C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\VBE7.DLL", "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE", "C:\Windows\System32\stdole2.tlb", "C:\Program Files\Common Files\Microsoft Shared\OFFICE16\MSO.DLL", "C:\WINDOWS\system32\FM20.DLL", "C:\Program Files\Microsoft Office\root\Office16\MSOUTL.OLB", "C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB", "C:\Windows\System32\scrrun.dll", "C:\Program Files\Microsoft Office\root\Office16\MSPPT.OLB", "C:\Windows\System32\vbscript.dll\2", "C:\Program Files\Microsoft Office\root\Office16\MSWORD.OLB", "C:\Windows\System32\vbscript.dll\3")

    resultStr = "Adding required VBA references " & vbCrLf
    resultStr = resultStr & "==========================" & vbCrLf
    
    For i = 0 To UBound(nameArray)
        Debug.Print Now() & " Adding Reference " & nameArray(i)
        If Not CheckReference(nameArray(i)) Then
            resultStr = resultStr & nameArray(i) & " adding vba reference " & Now() & vbCrLf
            vbProj.References.AddFromFile pathArray(i)
        Else
            resultStr = resultStr & nameArray(i) & " vba reference already exists " & Now() & vbCrLf
        End If
    Next i
    
    MsgBox resultStr
    
End Sub

Function CheckReference(refName As Variant) As Boolean
    Dim VBAEditor As VBIDE.VBE
    Dim vbProj As VBIDE.VBProject
    Dim chkRef As VBIDE.Reference
    Dim BoolExists As Boolean

    Set VBAEditor = Application.VBE
    Set vbProj = ActiveWorkbook.VBProject

    CheckReference = False
    For Each chkRef In vbProj.References
        If chkRef.Name = refName Then
            CheckReference = True
            GoTo exitsub
        End If
    Next
exitsub:
    Set VBAEditor = Nothing
    Set vbProj = Nothing
End Function

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
  IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function

