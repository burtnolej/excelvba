Attribute VB_Name = "App_Misc"
Option Explicit

Public Function get_misc_timeperiod(clsAppRuntime As App_Runtime, clsExecProc As Exec_Proc, _
                          Optional bInTable As Boolean = True) As Worksheet
Dim dArgs As New Dictionary
                          
    AddArgs dArgs, False, "clsAppRuntime", clsAppRuntime, _
                          "bInTable", bInTable, _
                          "clsExecProc", clsExecProc, _
                          "eQuadDataType", QuadDataType.Misc, _
                          "eQuadSubDataType", QuadSubDataType.TimePeriod, _
                          "eQuadScope", QuadScope.all
                                           
    clsExecProc.ExecProc "GetPersonData", dArgs
    Set get_misc_timeperiod = dArgs.Item("result")
    
End Function
Public Function get_misc_prep(clsAppRuntime As App_Runtime, clsExecProc As Exec_Proc, _
                          Optional bInTable As Boolean = True) As Worksheet
Dim dArgs As New Dictionary
                          
    AddArgs dArgs, False, "clsAppRuntime", clsAppRuntime, _
                          "bInTable", bInTable, _
                          "clsExecProc", clsExecProc, _
                          "eQuadDataType", QuadDataType.Misc, _
                          "eQuadSubDataType", QuadSubDataType.Prep, _
                          "eQuadScope", QuadScope.all
                                           
    clsExecProc.ExecProc "GetPersonData", dArgs
    Set get_misc_prep = dArgs.Item("result")
    
End Function
Public Function get_misc_day(clsAppRuntime As App_Runtime, clsExecProc As Exec_Proc, _
                          Optional bInTable As Boolean = True) As Worksheet
Dim dArgs As New Dictionary
                          
    AddArgs dArgs, False, "clsAppRuntime", clsAppRuntime, _
                          "bInTable", bInTable, _
                          "clsExecProc", clsExecProc, _
                          "eQuadDataType", QuadDataType.Misc, _
                          "eQuadSubDataType", QuadSubDataType.Day, _
                          "eQuadScope", QuadScope.all
                          
    clsExecProc.ExecProc "GetPersonData", dArgs
    Set get_misc_day = dArgs.Item("result")
    
End Function

Public Function get_misc_location(clsAppRuntime As App_Runtime, clsExecProc As Exec_Proc, _
                          Optional bInTable As Boolean = True) As Worksheet
Dim dArgs As New Dictionary
                          
    AddArgs dArgs, False, "clsAppRuntime", clsAppRuntime, _
                          "bInTable", bInTable, _
                          "clsExecProc", clsExecProc, _
                          "eQuadDataType", QuadDataType.Misc, _
                          "eQuadSubDataType", QuadSubDataType.Location, _
                          "eQuadScope", QuadScope.all
                                           
    clsExecProc.ExecProc "GetPersonData", dArgs
    Set get_misc_location = dArgs.Item("result")
    
End Function




