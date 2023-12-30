Attribute VB_Name = "RibbonUtils"
Dim RV As RibbonVariables
Dim rbxUI As IRibbonUI


'Callback for customUI.onLoad
Sub CArbx_onLoad(ribbon As IRibbonUI)
    Debug.Print "rbx_onLoad"
    'Set RV = New RibbonVariables
    '
    'CallByName RV, "RibbonPointer", VbLet, ribbon
    
End Sub


'Callback for searchby__type onAction
Sub CAdropDown_onAction(control As IRibbonControl, id As String, index As Integer)
    Debug.Print "CAdropDown_onAction", id, index
    Set RV = New RibbonVariables
    CallByName RV, control.id, VbLet, id
    Set RV = Nothing
    'CallByName RV, functionSplit(1), VbLet, functionSplit(2)
End Sub

'Callback for searchby__type getSelectedItemID
Sub CAdropDown_getText(control As IRibbonControl, ByRef returnedVal)
    Set RV = New RibbonVariables
    returnedVal = CallByName(RV, control.id, VbGet)
    Set RV = Nothing
End Sub


'Callback for btns_btn1 onAction
Public Sub CAbtns_onAction(control As IRibbonControl)
Dim tag As String, action As String, param As String
Dim tagSplit As Variant, functionSplit As Variant
    Debug.Print "CAbtns_onAction", control.id, control.tag
    tag = control.tag
    
    tagSplit = Split(tag, "_")
    action = tagSplit(0)
    If UBound(tagSplit) > 0 Then
        param = tagSplit(1)
    Else
        param = ""
    End If

    Set RV = New RibbonVariables
    functionSplit = Split(param, "^")
    'CallByName RV, functionSplit(1), VbLet, functionSplit(2)
    CallByName RV, functionSplit(1), VbLet, control.id
    Set RV = Nothing
    
End Sub


'Sub RehydrateConfig(Optional param As String = "")
'    Set RV = New RibbonVariables
'    RV.Rehydrate
    
'End Sub

'Sub PersistConfig(Optional param As String = "")
'    Set RV = New RibbonVariables
'    RV.Persist
'End Sub


'Callback for searchby__id getText
Sub CAeditBox_getText(control As IRibbonControl, ByRef returnedVal)
    Debug.Print "CAeditBox_getText", id, index
    Set RV = New RibbonVariables
    returnedVal = CallByName(RV, control.id, VbGet)
    Set RV = Nothing
End Sub

'Callback for searchby__id onChange
Sub CAeditBox_onChange(control As IRibbonControl, text As String)
    Debug.Print "CAeditBox_onChange", id, text
    Set RV = New RibbonVariables
    CallByName RV, control.id, VbLet, text
    Set RV = Nothing
End Sub

