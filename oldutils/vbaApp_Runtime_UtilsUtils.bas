Attribute VB_Name = "App_Runtime_Utils"
Option Explicit
Const C_MODULE_NAME = "App_Runtime_Utils"
Private clsAppRuntimeGlobal As App_Runtime
Public Sub ResetAppRuntimeGlobal()
    Set clsAppRuntimeGlobal = Nothing
End Sub
Public Function InitAppRuntimeGlobal(Optional dAppRuntimeValues As Dictionary) As App_Runtime
Dim clsAppRuntime As New App_Runtime
Dim vKey As Variant

    clsAppRuntime.InitProperties
    
    If IsSet(dAppRuntimeValues) Then
        For Each vKey In dAppRuntimeValues
            CallByName clsAppRuntime, vKey, VbLet, dAppRuntimeValues.Item(vKey)
        Next vKey
    End If
    
    Set InitAppRuntimeGlobal = clsAppRuntime
End Function
Public Sub LetAppRuntimeGlobal(clsAppRuntime As App_Runtime)
Dim sFuncName As String
    sFuncName = C_MODULE_NAME & "." & "LetAppRuntimeGlobal"
    If IsInstance(clsAppRuntime, vbAppRuntime) = False Then
        err.Raise ErrorMsgType.BAD_ARGUMENT, Description:="arg is not of type App_Runtime"
    End If
    
    Set clsAppRuntimeGlobal = clsAppRuntime
    FuncLogIt sFuncName, "Setting GLOBAL Quad_Utils.clsAppRuntimeGlobal", C_MODULE_NAME, LogMsgType.INFO
End Sub
Public Function GetAppRuntimeGlobal(Optional bInitFlag As Boolean = False, _
                                     Optional dAppRuntimeValues As Dictionary) As App_Runtime
Dim sFuncName As String
    sFuncName = C_MODULE_NAME & "." & "GetAppRuntimeGlobal"
    
    If IsSet(clsAppRuntimeGlobal) Then
        Set GetAppRuntimeGlobal = clsAppRuntimeGlobal
        FuncLogIt sFuncName, "GETTING GLOBAL Quad_Utils.clsAppRuntimeGlobal", C_MODULE_NAME, LogMsgType.INFO
    Else
        If bInitFlag = True Then
            Set GetAppRuntimeGlobal = InitAppRuntimeGlobal(dAppRuntimeValues:=dAppRuntimeValues)
            FuncLogIt sFuncName, "Initializating GLOBAL Quad_Utils.clsAppRuntimeGlobal", C_MODULE_NAME, LogMsgType.INFO
        Else
            Set GetAppRuntimeGlobal = New App_Runtime
            GetAppRuntimeGlobal.InitProperties bInitializeCache:=False
            FuncLogIt sFuncName, "Recovering from cache GLOBAL Quad_Utils.clsAppRuntimeGlobal", C_MODULE_NAME, LogMsgType.INFO
        End If
    End If
End Function

Public Function New_AppRuntime() As App_Runtime
    Set New_AppRuntime = New App_Runtime
End Function

