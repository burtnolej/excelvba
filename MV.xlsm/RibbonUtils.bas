Attribute VB_Name = "RibbonUtils"

'Dim rbxUI As IRibbonUI

'Declare PtrSafe Function GetSystemMetrics32 Lib "USER32" _
'    Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
'Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
'    destination As Any, _
'    source As Any, _
'    ByVal length As Long)
    
Sub MVdropDown_getText(control As IRibbonControl, ByRef returnedVal)
Dim RV As MVRibbonVariables
    Debug.Print "dropDown_getText"
    Set RV = New MVRibbonVariables
    returnedVal = CallByName(RV, control.id, VbGet)
    Set RV = Nothing
    
End Sub

'Callback for customUI.onLoad
Sub MVrbx_onLoad(ribbon As IRibbonUI)
Dim RV As MVRibbonVariables
    Debug.Print "rbx_onLoad"
    Set RV = New MVRibbonVariables
    
    CallByName RV, "RibbonPointer", VbLet, ribbon
    
    ribbon.ActivateTab "MVtab4"
    Set RV = Nothing
End Sub

'Callback for btns_btn2 onAction
Sub MVbtns_onAction(control As IRibbonControl)
Dim tag As String, action As String, param As String
Dim tagSplit As Variant, functionSplit As Variant
    tag = control.tag
    
    tagSplit = Split(tag, "_")
    action = tagSplit(0)
    If UBound(tagSplit) > 0 Then
        param = tagSplit(1)
    Else
        param = ""
    End If

    'CustomRibbon.Invalidate

    Select Case action
        Case "runfunction"
            functionSplit = Split(param, "^")
            functionSplit = Application.Run(functionSplit(0) & ".xlsm!" & functionSplit(1), functionSplit(2))
    End Select
End Sub

Sub RehydrateConfig(Optional param As String = "")
Dim RV As MVRibbonVariables
    Set RV = New MVRibbonVariables
    RV.Rehydrate
    Set RV = Nothing
    
End Sub

Sub PersistConfig(Optional param As String = "")
Dim RV As MVRibbonVariables
    Set RV = New MVRibbonVariables
    RV.Persist
    Set RV = Nothing
End Sub

'Callback for dropDown3 onAction
Sub MVdropDown_onAction(control As IRibbonControl, id As String, index As Integer)
Dim RV As MVRibbonVariables
    Debug.Print "dropDown_onAction", id, index
    Set RV = New MVRibbonVariables
    CallByName RV, control.id, VbLet, id
    Set RV = Nothing
End Sub

'Callback for working onAction
Sub MVchkBox_onAction(control As IRibbonControl, pressed As Boolean)
Dim RV As MVRibbonVariables
    Debug.Print "chkBox_onAction"
    Set RV = New MVRibbonVariables
    CallByName RV, control.id, VbLet, CStr(pressed)
    RV.RibbonPointer.InvalidateControl "config__Status_Filter"
    Set RV = Nothing
    
End Sub

'Callback for workingdir getText
Sub MVeditBox_getText(control As IRibbonControl, ByRef returnedVal)
Dim RV As MVRibbonVariables
    Debug.Print "editBox_getText"
    Set RV = New MVRibbonVariables
    returnedVal = CallByName(RV, control.id, VbGet)
    
    If control.id = "config__Working_Dir" And returnedVal = "" Then
        Debug.Print "defaulting config__Working_Dir"
        CallByName RV, control.id, VbLet, Environ("USERPROFILE") & "\Deploy"
        returnedVal = Environ("USERPROFILE") & "\Deploy"
    ElseIf control.id = "config__Template_File" And returnedVal = "" Then
        Debug.Print "defaulting config__Template_File"
        CallByName RV, control.id, VbLet, Environ("USERPROFILE") & "\Deploy\MondayViewUpdate_Template.xlsm"
        returnedVal = Environ("USERPROFILE") & "\Deploy\MondayViewUpdate_Template.xlsm"
    Else
        Debug.Print "returnedval=" & returnedVal
    End If
    Set RV = Nothing
End Sub

'Callback for workingdir onChange
Sub MVeditBox_onChange(control As IRibbonControl, text As String)
Dim RV As MVRibbonVariables
    Debug.Print "editBox_onChange"
    Set RV = New MVRibbonVariables
    CallByName RV, control.id, VbLet, text
    Set RV = Nothing
End Sub


'Callback for working getPressed
Sub MVfncGetPressed(control As IRibbonControl, ByRef returnedVal)
Dim RV As MVRibbonVariables
    Debug.Print "fncGetPressed"
    Set RV = New MVRibbonVariables
    returnedVal = CallByName(RV, control.id, VbGet)
    If returnedVal = "" Then
        returnedVal = False
    Else
        returnedVal = True
    End If
    Set RV = Nothing
End Sub
