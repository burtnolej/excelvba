Attribute VB_Name = "Enum_TestResult"


Function Test_OK() As Integer
    Test_OK = 9
End Function
Function Test_Failure() As Integer
    Test_Failure = 8
End Function
Function Test_Error() As Integer
    Test_Error = 7
End Function


Sub ListEnum()
Dim dValues As Dictionary

    Set dValues = GetProcsInModules(Workbooks("vba_source_new.xlsm"), "Enum_TestResult")

    Debug.Print Join(dValues.Keys, ",")

End Sub
