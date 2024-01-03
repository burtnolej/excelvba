Attribute VB_Name = "RibbonUtils"
Option Explicit

Declare PtrSafe Function GetSystemMetrics32 Lib "USER32" _
    Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    destination As Any, _
    source As Any, _
    ByVal length As Long)
    
'Callback for customUI.onLoad
Sub TArbx_onLoad(ribbon As IRibbonUI)
Dim RV As TARibbonVariables
    Debug.Print "rbx_onLoad"
    Set RV = New TARibbonVariables
    
    CallByName RV, "RibbonPointer", VbLet, ribbon
    
    ribbon.ActivateTab "tab23"
    Set RV = Nothing
    
    ActiveWorkbook.Sheets("RawData").Activate
End Sub

'Callback for action_updateitems onAction
Sub TAbtns_onAction(control As IRibbonControl)
Dim tag As String, action As String, param As String
Dim tagSplit As Variant, functionSplit As Variant
Dim RV As TARibbonVariables
    Debug.Print "TAbtns_onAction", control.id, control.tag
    tag = control.tag
    
    tagSplit = Split(tag, "_")
    action = tagSplit(0)
    If UBound(tagSplit) > 0 Then
        param = tagSplit(1)
    Else
        param = ""
    End If

    Set RV = New TARibbonVariables
    functionSplit = Split(param, "^")
    Set RV = New TARibbonVariables
    CallByName RV, control.id, VbLet, control.id
    Set RV = Nothing
End Sub

'Callback for settings__startdate getText
Sub TAeditBox_getText(control As IRibbonControl, ByRef returnedVal)
Dim RV As TARibbonVariables
    Debug.Print "TAeditBox_getText"
    Set RV = New TARibbonVariables
    returnedVal = CallByName(RV, control.id, VbGet)
    Set RV = Nothing
End Sub

'Callback for settings__startdate onChange
Sub TAeditBox_onChange(control As IRibbonControl, text As String)
Dim RV As TARibbonVariables

    Debug.Print "TAeditBox_onChange", text
    Set RV = New TARibbonVariables
    CallByName RV, control.id, VbLet, text
    
    RV.RibbonPointer.InvalidateControl "settings__enddate"
    Set RV = Nothing
End Sub

'Callback for useremail onAction
Sub TAdropDown_onAction(control As IRibbonControl, id As String, index As Integer)
Dim RV As TARibbonVariables
    Debug.Print "TAdropDown_onAction", id, index
    Set RV = New TARibbonVariables
    CallByName RV, control.id, VbLet, id
    Set RV = Nothing
End Sub

'Callback for useremail getSelectedItemID
Sub TAdropDown_getText(control As IRibbonControl, ByRef returnedVal)
Dim RV As TARibbonVariables
    Set RV = New TARibbonVariables
    returnedVal = CallByName(RV, control.id, VbGet)
    Set RV = Nothing
End Sub


'Callback for additem__mondayitem getItemCount
Sub TACombo_getItemCount(control As IRibbonControl, ByRef returnedVal)
Dim tmpRange As Range
    If control.id = "additem__mondayitem" Then
        Set tmpRange = ActiveWorkbook.Sheets("Monday Data").Range("MONDAY_ITEMS_DISPLAY")
        returnedVal = 300
    ElseIf control.id = "additem__mondaysubitem" Then
        Set tmpRange = ActiveWorkbook.Sheets("Monday Data").Range("MONDAY_SUBITEMS_DISPLAY")
        returnedVal = 300
    ElseIf control.id = "additem__category" Then
        Set tmpRange = ActiveWorkbook.Sheets("Reference").Range("CATEGORY_LIST")
        returnedVal = tmpRange.Rows.Count
        
    End If
    
    Set tmpRange = Nothing
    
End Sub

'Callback for additem__mondayitem getItemID
Sub TACombo_getItemID(control As IRibbonControl, index As Integer, ByRef returnedVal)
    returnedVal = CStr(index) & "val"
End Sub

'Callback for additem__mondayitem getItemLabel
Sub TACombo_getItemLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)
Dim tmpRange As Range, mondayItemName As Range, mondaySubItemName As Range

    If control.id = "additem__mondayitem" Then
        Set mondayItemName = ActiveWorkbook.Sheets("Monday Data").Range("MONDAY_ITEMS_DISPLAY")
        returnedVal = mondayItemName.Rows(index).value
    ElseIf control.id = "additem__mondaysubitem" Then
        Set mondaySubItemName = ActiveWorkbook.Sheets("Monday Data").Range("MONDAY_SUBITEMS_DISPLAY")
        returnedVal = mondaySubItemName.Rows(index).value
    ElseIf control.id = "additem__category" Then
        Set tmpRange = ActiveWorkbook.Sheets("Reference").Range("CATEGORY_LIST")
        returnedVal = tmpRange.Rows(index).value
    End If

    Set tmpRange = Nothing
    Set mondayItemName = Nothing
    
End Sub

'Callback for additem__mondayitem getText
Sub TACombo_getText(control As IRibbonControl, ByRef returnedVal)
    returnedVal = ""
End Sub

'Callback for additem__mondayitem onChange
Sub TACombo_onChange(control As IRibbonControl, text As String)
Dim RV As TARibbonVariables
    Debug.Print "TACombo_onChange", text
    Set RV = New TARibbonVariables
    CallByName RV, control.id, VbLet, text
    Set RV = Nothing
End Sub


