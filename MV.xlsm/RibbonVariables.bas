VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RibbonVariables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private debugflagval As String
Private userval As String
Private agefilterval As String
Private sortval As String

Private statusfiltercompletedval As String
Private statusfilterdoneval As String
Private statusfilterworkingval As String
Private statusfilternotstartedval As String

Private bookname As String
Private sheetname As String

Private ribbonpointerval As IRibbonUI

' Ribbon Pointer  ''''''''''''''''''''''''''''''
Property Get RibbonPointer() As IRibbonUI
Dim aPtr As LongPtr
    Debug.Print "RibbonPointer"

    If Not ribbonpointerval Is Nothing Then
        Set RibbonPointer = ribbonpointerval
    Else
        aPtr = GetVariableSheetValue("ribbonpointerval")
        CopyMemory RibbonPointer, aPtr, LenB(aPtr)
    End If
    
End Property
Property Let RibbonPointer(value As IRibbonUI)
    Debug.Print "LetRibbonPointer", ObjPtr(value)
    
    LetVariableSheetValue "ribbonpointerval", ObjPtr(value)
End Property


' Status Filter Completed  ''''''''''''''''''''''''''''''
Property Get StatusFilterCompleted() As String
    Debug.Print "StatusFilterCompleted"
    If statusfiltercompletedval = "" Then
        StatusFilterCompleted = GetVariableSheetValue("statusfiltercompletedval")
    Else
        StatusFilterCompleted = statusfiltercompletedval
    End If
End Property
Property Let StatusFilterCompleted(value As String)
    Debug.Print "LetStatusFilterCompleted", value
    LetVariableSheetValue "statusfiltercompletedval", value
End Property


' Status Filter Done ''''''''''''''''''''''''''''''
Property Get StatusFilterDone() As String
    Debug.Print "StatusFilterDone"
    If statusfiltercompletedval = "" Then
        StatusFilterDone = GetVariableSheetValue("statusfilterdoneval")
    Else
        StatusFilterDone = statusfilterdoneval
    End If
End Property
Property Let StatusFilterDone(value As String)
    Debug.Print "LetStatusFilterDone", value
    LetVariableSheetValue "statusfilterdoneval", value
End Property


' Status Filter Working ''''''''''''''''''''''''''''''
Property Get StatusFilterWorking() As String
    Debug.Print "StatusFilterWorking"
    If statusfilterworkingval = "" Then
        StatusFilterWorking = GetVariableSheetValue("statusfilterworkingval")
    Else
        StatusFilterWorking = statusfilterworkingval
    End If
End Property
Property Let StatusFilterWorking(value As String)
    Debug.Print "LetStatusFilterWorking", value
    LetVariableSheetValue "statusfilterworkingval", value
End Property


' Status Filter Not Started ''''''''''''''''''''''''''''''
Property Get StatusFilterNotStarted() As String
    Debug.Print "StatusFilterWorking"
    If statusfilternotstartedval = "" Then
        StatusFilterNotStarted = GetVariableSheetValue("statusfilternotstartedval")
    Else
        StatusFilterNotStarted = statusfilternotstartedval
    End If
End Property
Property Let StatusFilterNotStarted(value As String)
    Debug.Print "LetStatusFilterNotStarted", value
    LetVariableSheetValue "statusfilternotstartedval", value
End Property

' Debug Flag ''''''''''''''''''''''''''''''
Property Get DebugFlag() As String
    Debug.Print "GetDebugFlag"
    If debugflagval = "" Then
        DebugFlag = GetVariableSheetValue("debugflagval")
    Else
        DebugFlag = debugflagval
    End If
    
End Property
Property Let DebugFlag(value As String)
    Debug.Print "LetDebugFlag", value
    LetVariableSheetValue "debugflagval", value
    
End Property


' User ''''''''''''''''''''''''''''''
Property Get User() As String
    Debug.Print "GetUser"
    If userval = "" Then
        User = GetVariableSheetValue("userval")
    Else
        User = userval
    End If
    
End Property
Property Let User(value As String)
    Debug.Print "LetUser", value
    LetVariableSheetValue "userval", value
    
End Property


' Age Filter  ''''''''''''''''''''''''''''''
Property Get AgeFilter() As String
    Debug.Print "GetAgeFilter"
    If agefilterval = "" Then
        AgeFilter = GetVariableSheetValue("agefilterval")
    Else
        AgeFilter = agefilterval
    End If
    
End Property
Property Let AgeFilter(value As String)
    Debug.Print "LetAgeFilter", value
    LetVariableSheetValue "agefilterval", value
    
End Property

' Age Filter  ''''''''''''''''''''''''''''''
Property Get Sort() As String
    If sortval = "" Then
        Sort = GetVariableSheetValue("sortval")
    Else
        Sort = sortval
    End If
    Debug.Print "GetSort"
End Property
Property Let Sort(value As String)
    Debug.Print "LetSort", value
    LetVariableSheetValue "sortval", value
    
End Property


Sub LetVariableSheetValue(varname As String, value As String)
    Workbooks(bookname).Sheets(sheetname).Range(varname).value = value
    Debug.Print "LetVariableSheetValue", varname, value
End Sub

Function GetVariableSheetValue(varname As String) As String
     GetVariableSheetValue = Workbooks(bookname).Sheets(sheetname).Range(varname).value
     Debug.Print "GetVariableSheetValue", varname
End Function


Private Sub Class_Initialize()
    Debug.Print "Class_Initialize"
    sheetname = "Persist"
    bookname = "MV.xlsm"
    
    debugflagval = ""
    userval = ""
    agefilterval = ""
    sortval = ""
    
    statusfiltercompletedval = ""
    
End Sub


