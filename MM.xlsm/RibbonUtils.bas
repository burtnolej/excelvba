Attribute VB_Name = "RibbonUtils"
Option Explicit

Dim rbxUI As IRibbonUI

'Callback for customUI.onLoad
Sub rbx_onLoad(ribbon As IRibbonUI)

    Set rbxUI = ribbon
    rbxUI.ActivateTab "tab3"
End Sub

'Callback for splitBtn_btn16 onAction
Sub splitBtn_onAction(control As IRibbonControl)
End Sub

'Callback for togBtn_btn2 onAction
Sub togBtn_onAction(control As IRibbonControl, pressed As Boolean)

    btns_onAction control
End Sub

'Callback for btns_btn1 onAction
Sub btns_onAction(control As IRibbonControl)
Dim tag As String, action As String, param As String, result As String
Dim tagSplit() As String, functionSplit() As String
Dim Folders() As Variant
Dim folderListDict As Dictionary

    tag = control.tag
    
    tagSplit = Split(tag, "_")
    action = tagSplit(0)
    If UBound(tagSplit) > 0 Then
        param = tagSplit(1)
    Else
        param = ""
    End If
    
    Select Case action
    
        Case "runfunction"
            functionSplit = Split(param, "^")
            result = Application.Run(functionSplit(0) & ".xlsm!" & functionSplit(1), functionSplit(2))
            MsgBox result
        
        Case "openfolder"
            functionSplit = Split(param, "^")
            Set folderListDict = New Dictionary
            
            Application.Run "vbautils.xlsm!RangeToDict", ActiveWorkbook, "Reference", "FOLDERS", folderListDict

            Folders = folderListDict(functionSplit(0))
            Application.Run "vbautils.xlsm!LaunchExplorer", CStr(Folders(1)), CStr(Folders(2))
            
        Case "togglesheet"
            param = Replace(param, "^", "_")
            If ActiveWorkbook.Sheets(UCase(param)).Visible = True Then
                ActiveWorkbook.Sheets(UCase(param)).Visible = False
            Else
                ActiveWorkbook.Sheets(UCase(param)).Visible = True
            End If
            
        Case "gotoattr"
            param = Replace(param, "^", "_")
            ChangeInputSheetFocus param
            
            
    End Select
    
End Sub


