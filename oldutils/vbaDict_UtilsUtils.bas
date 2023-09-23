Attribute VB_Name = "Dict_Utils"


Public Sub AddDict(ByRef dTmp As Dictionary, sKey As String, vVal As Variant, Optional bUpdate As Boolean = True)
    
    If bUpdate = True Then
        If dTmp.Exists(sKey) Then
            dTmp.Remove (sKey)
        End If
    End If
    
    dTmp.Add sKey, vVal
End Sub
Function Dict2Array(dTmp As Dictionary, aFields() As String) As String()
' For 2 level nested dictionary. aKeys is a string array containing the sub fields we want to pull out from the next level dict
Dim aKeys() As String
Dim sKey As Variant, sField As Variant
Dim iLength As Integer, iWidth As Integer, iFieldCount As Integer, iRowCount As Integer
Dim aResult() As String

    aKeys = GetDictKeys(dTmp)
    
    iLength = UBound(aKeys)
    iWidth = UBound(aFields)
    ReDim aResult(0 To iLength, 0 To iWidth + 1)
    
    For Each sKey In aKeys
        aResult(iRowCount, 0) = sKey
        iFieldCount = 1
        For Each sField In aFields
            aResult(iRowCount, iFieldCount) = dTmp.Item(sKey).Item(sField)
            iFieldCount = iFieldCount + 1
        Next sField
        iRowCount = iRowCount + 1
    Next sKey

    Dict2Array = aResult
End Function
Function GetDictKeys(dDict As Dictionary) As String()
Dim aResult() As String
Dim iCount As Integer
Dim vKey As Variant

    ReDim aResult(0 To 1000)
    For Each vKey In dDict.Keys
        aResult(iCount) = vKey
        iCount = iCount + 1
    Next
    
    ReDim Preserve aResult(0 To iCount - 1)
    
    GetDictKeys = aResult

End Function

Function GetDictVals(dDict As Dictionary) As Variant()
Dim aResult() As Variant
Dim iCount As Integer
Dim vKey As Variant

    ReDim aResult(0 To 1000)
    For Each vKey In dDict.Keys
        aResult(iCount) = dDict.Item(vKey)
        iCount = iCount + 1
    Next
    
    ReDim Preserve aResult(0 To iCount - 1)
    GetDictVals = aResult

End Function

Sub DumpDict(dDict As Dictionary)
Dim vKey As Variant

    ReDim aResult(0 To 1000)
    For Each vKey In dDict.Keys
        Debug.Print CStr(vKey) & "=" & dDict.Item(vKey)
    Next

End Sub

Function InitDict(aKeys() As String, aValues As Variant) As Dictionary
Dim dResult As New Dictionary
    
    For i = 0 To UBound(aKeys)
        dResult.Add aKeys(i), aValues(i)
    Next i
    
    Set InitDict = dResult
End Function
Function CompareDict(dDict1 As Dictionary, dDict2 As Dictionary) As Boolean
Dim vKey As Variant

    CompareDict = True

    For Each vKey In dDict1.Keys
        If dDict2.Exists(vKey) <> True Then
            CompareDict = False
            Exit Function
        ElseIf dDict2.Item(vKey) <> dDict1.Item(vKey) Then
            CompareDict = False
            Exit Function
        End If
    Next vKey

End Function
