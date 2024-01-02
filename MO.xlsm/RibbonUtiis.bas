Attribute VB_Name = "RibbonUtiis"
Option Explicit

'Callback for customUI.onLoad
Sub MOrbx_onLoad(ribbon As IRibbonUI)
Dim RV As MORibbonVariables
    Debug.Print "rbx_onLoad"
    Set RV = New MORibbonVariables
    
    CallByName RV, "RibbonPointer", VbLet, ribbon
    Set RV = Nothing
End Sub

'Callback for actions__additem onAction
Sub MObtns_onAction(control As IRibbonControl)
Dim tag As String, action As String, param As String
Dim tagSplit As Variant, functionSplit As Variant
Dim RV As MORibbonVariables
    Debug.Print "MObtns_onAction", control.ID, control.tag
    tag = control.tag
    
    tagSplit = Split(tag, "_")
    action = tagSplit(0)
    If UBound(tagSplit) > 0 Then
        param = tagSplit(1)
    Else
        param = ""
    End If

    Set RV = New MORibbonVariables
    functionSplit = Split(param, "^")
    Set RV = New MORibbonVariables
    CallByName RV, control.ID, VbLet, control.ID
    Set RV = Nothing
End Sub

'Callback for refreshdata__mvreport getText
Sub MOeditBox_getText(control As IRibbonControl, ByRef returnedVal)
Dim RV As MORibbonVariables
    Debug.Print "MOeditBox_getText"
    Set RV = New MORibbonVariables
    returnedVal = CallByName(RV, control.ID, VbGet)
    Set RV = Nothing
End Sub

'Callback for refreshdata__mvreport onChange
Sub MOeditBox_onChange(control As IRibbonControl, text As String)
Dim RV As MORibbonVariables
    Debug.Print "MOeditBox_onChange", text
    Set RV = New MORibbonVariables
    CallByName RV, control.ID, VbLet, text
    Set RV = Nothing
End Sub


