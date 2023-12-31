Attribute VB_Name = "RibbonUtils"
'Declare PtrSafe Function GetSystemMetrics32 Lib "USER32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
'Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal length As Long)
'Public Sub RefreshDownloadFiles(Optional param As Variant
'Sub rbx_onLoad(ribbon As IRibbonUI)
'Sub dropDown_onAction(control As IRibbonControl, id As String, index As Integer)
'Sub dropDown_getText(control As IRibbonControl, ByRef returnedVal)
'Sub editBox_onChange(control As IRibbonControl, Text As String)
'Sub editBox_getText(control As IRibbonControl, ByRef returnedVal)
'Sub chkBox_onAction(control As IRibbonControl, isPressed As Boolean)
'Public Sub fncGetPressed(control As IRibbonControl, ByRef bolReturn)
'Sub btns_onAction(control As IRibbonControl)
'Sub ToolActionOpenExec(appname As String)
'Sub ToolActionCloseExec(appname As String)
'Sub btns_onActionOLD(control As IRibbonControl)



Option Explicit

Dim rbxUI As IRibbonUI
Dim RV As RibbonVariables
Dim SpinValue As Long
Dim manifestFiles() As Variant
Dim folderListDict As Dictionary, appListDict As Dictionary, urlListDict As Dictionary, checkboxVals As Dictionary
Dim x As Long
Dim y As Long
Dim height As Long
Dim width As Long
Dim rootpath As String
Dim dataurl As String

Declare PtrSafe Function GetSystemMetrics32 Lib "USER32" _
    Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    destination As Any, _
    source As Any, _
    ByVal length As Long)
    
Sub GetScreenRes(ByRef w As Long, ByRef h As Long)
    w = GetSystemMetrics32(0) ' width in points
    h = GetSystemMetrics32(1) ' height in points
End Sub

Public Sub RefreshDownloadFiles(Optional param As Variant)
Dim outputRange As Range
Dim url As String
Dim colArray() As Variant

    'url = "http://172.23.208.38/datafiles/"
    url = "http://172.22.237.138/datafiles/"
    
    Application.Run "vbautils.xlsm!SetEventsOff"
    
    On Error Resume Next
    Application.StatusBar = "loading http://172.22.237.138/datafiles/manifest.csv"
    Set outputRange = HTTPDownloadFile(url + "manifest.csv", _
                ActiveWorkbook, _
                "", "", 1, "start-of-day", "MANIFEST", False, 0)
                
    Application.Run "vbautils.xlsm!SortRange", outputRange.Worksheet, outputRange, 2
    
    
    colArray = Array(1, 2)
    
    'Application.Run "vbautils.xlsm!CreateCalcNamedRange", outputRange.Worksheet, outputRange, "fullFileName", colArray
    
    'Set outputRange = outputRange.Resize(, outputRange.Columns.Count + 1)
    
    Application.Run "vbautils.xlsm!AddNamedRange", outputRange.Worksheet, outputRange, 1, "FILENAME"
    
    manifestFiles = outputRange.Offset(1).Resize(outputRange.Rows.count - 1)
    
    'manifestFiles = Application.Run("vbautils.xlsm!RangeToArray", ActiveWorkbook, "FILE_ALLDATA", "FILE_FULLNAME", manifestFiles)

End Sub
 
Sub rbx_onLoad(ribbon As IRibbonUI)
Dim persistedVars() As Variant
Dim varValues As Dictionary
    
    Set rbxUI = ribbon
    
    On Error Resume Next
    CommandBars("Document Recovery").Visible = False
    On Error GoTo 0
    
    rbxUI.ActivateTab "tab3"
    
    Set folderListDict = New Dictionary
    Set appListDict = New Dictionary
    Set urlListDict = New Dictionary
    Set checkboxVals = New Dictionary
    
    RangeToDict ActiveWorkbook, "Persist", "FOLDERS", folderListDict
    RangeToDict ActiveWorkbook, "Persist", "APPS", appListDict
    RangeToDict ActiveWorkbook, "Persist", "URLS", urlListDict
    
    Set RV = New RibbonVariables
    
    CallByName RV, "RibbonPointer", VbLet, rbxUI

End Sub

'Callback for choosetool onAction
Sub dropDown_onAction(control As IRibbonControl, id As String, index As Integer)
    Debug.Print "dropDown_onAction", id, index
    Set RV = New RibbonVariables
    CallByName RV, control.id, VbLet, id
    Set RV = Nothing
End Sub

'Callback for choosetool getSelectedItemID
Sub dropDown_getText(control As IRibbonControl, ByRef returnedVal)
    Set RV = New RibbonVariables
    returnedVal = CallByName(RV, control.id, VbGet)
    Set RV = Nothing
End Sub


' Set default value of editBox to 0

Sub editBox_onChange(control As IRibbonControl, Text As String)
    Debug.Print "editBox_onChange", Text
    Set RV = New RibbonVariables
    CallByName RV, control.id, VbLet, Text
    Set RV = Nothing
End Sub

' Return value of editBox

Sub editBox_getText(control As IRibbonControl, ByRef returnedVal)
    Set RV = New RibbonVariables
    returnedVal = CallByName(RV, control.id, VbGet)
    Set RV = Nothing
   
End Sub



Sub chkBox_onAction(control As IRibbonControl, isPressed As Boolean)
    Set RV = New RibbonVariables
    CallByName RV, control.id, VbLet, isPressed
    Set RV = Nothing
End Sub

Public Sub fncGetPressed(control As IRibbonControl, ByRef bolReturn)

    Set RV = New RibbonVariables
    bolReturn = CallByName(RV, control.id, VbGet)
    Set RV = Nothing

End Sub

Sub btns_onAction(control As IRibbonControl)
Dim tag As String, action As String, param As String
Dim tagSplit As Variant, functionSplit As Variant
    Debug.Print "CAbtns_onAction", control.id, control.tag
    tag = control.tag
    
    Set RV = New RibbonVariables
    If tag = "" Then
        CallByName RV, control.id, VbLet, control.id
    Else
        tagSplit = Split(tag, "_")
        action = tagSplit(0)
        If UBound(tagSplit) > 0 Then
            param = tagSplit(1)
        Else
            param = ""
        End If
    
        functionSplit = Split(param, "^")
        CallByName RV, functionSplit(1), VbLet, control.id
    End If
    Set RV = Nothing
End Sub
    
Sub btns_onActionOLD(control As IRibbonControl)
Dim tag As String, action As String, param As String, foldername As String, bookname As String
Dim tagSplit As Variant, functionSplit As Variant
Dim w As Long, h As Long
Dim persistedVars() As Variant
Dim varValues As Dictionary
Dim args() As Variant


    foldername = RootPathValue
    
    'foldername = "E:\Velox Financial Technology\Velox Shared Drive - Documents\General\Tools"

    tag = control.tag
    
    tagSplit = Split(tag, "_")
    action = tagSplit(0)
    If UBound(tagSplit) > 0 Then
        param = tagSplit(1)
    Else
        param = ""
    End If

    CustomRibbon.Invalidate

    Select Case action

        Case "runfunction"
            functionSplit = Split(param, "^")
            functionSplit = Application.Run(functionSplit(0) & ".xlsm!" & functionSplit(1), functionSplit(2))
            
        Case "pickfolder"
            rootpath = Application.Run("VBAUtils.xlsm!GetFolderSelection", Environ("OneDrive"))
            CustomRibbon.InvalidateControl "editBox5"
            'rbxUI.InvalidateControl "editBox5"
            
        Case "runapp"
            functionSplit = Split(param, "^")
            Set appListDict = New Dictionary
            
            RangeToDict ActiveWorkbook, "Reference", "APPS", appListDict

            args = appListDict(functionSplit(0))
            LaunchApp CStr(args(1)), CStr(args(2))
            
         Case "runurl"
            functionSplit = Split(param, "^")
            Set urlListDict = New Dictionary
            
            RangeToDict ActiveWorkbook, "Reference", "URLS", urlListDict

            args = urlListDict(functionSplit(0))
            LaunchBrowser CStr(args(1)), x, y, width, height
            

        Case "killapp"
            functionSplit = Split(param, "^")
            Set appListDict = New Dictionary
            
            RangeToDict ActiveWorkbook, "Reference", "APPS", appListDict

            args = appListDict(functionSplit(0))
            KillApp CStr(args(3))
            

        
        
    End Select
    
    CustomRibbon.Invalidate
End Sub
