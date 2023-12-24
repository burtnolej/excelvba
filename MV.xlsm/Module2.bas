Attribute VB_Name = "Module2"
Dim RV As RibbonVariables
Dim rbxUI As IRibbonUI

Declare PtrSafe Function GetSystemMetrics32 Lib "USER32" _
    Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    destination As Any, _
    source As Any, _
    ByVal length As Long)
    
Sub dropDown_getText(control As IRibbonControl, ByRef returnedVal)

    Debug.Print "dropDown_getText"
    Set RV = New RibbonVariables
    returnedVal = CallByName(RV, control.id, VbGet)

End Sub

'Callback for customUI.onLoad
Sub rbx_onLoad(ribbon As IRibbonUI)
    Debug.Print "rbx_onLoad"
    Set RV = New RibbonVariables
    
    CallByName RV, "RibbonPointer", VbLet, ribbon
    
    ribbon.InvalidateControl "debugflagval"
    ribbon.InvalidateControl "userval"
    ribbon.InvalidateControl "agefilterval"
    ribbon.InvalidateControl "sortval"
    
End Sub

'Callback for btns_btn2 onAction
Sub btns_onAction(control As IRibbonControl)
End Sub

'Callback for dropDown3 onAction
Sub dropDown_onAction(control As IRibbonControl, id As String, index As Integer)
    Debug.Print "dropDown_onAction", id, index
    CallByName RV, control.id, VbLet, id
End Sub

'Callback for working onAction
Sub chkBox_onAction(control As IRibbonControl, pressed As Boolean)
    Debug.Print "chkBox_onAction"
    CallByName RV, control.id, VbLet, CStr(pressed)
End Sub


'Callback for working getPressed
Sub fncGetPressed(control As IRibbonControl, ByRef returnedVal)
    Debug.Print "fncGetPressed"
    Set RV = New RibbonVariables
    returnedVal = CallByName(RV, control.id, VbGet)
    If returnedVal = "" Then
        returnedVal = False
    Else
        returnedVal = True
    End If
End Sub

Sub InvalidateRibbon()
    'ribbon.InvalidateControl "debugflagval"
    'ribbon.InvalidateControl "userval"
    'ribbon.InvalidateControl "agefilterval"
    'ribbon.InvalidateControl "sortval"
End Sub
