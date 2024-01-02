VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MORibbonVariables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private persistsheetname As String
Private persistfilename As String
Private persistrangename As String
Private persistrangelen As String
Private bookname As String

Private actions__sddItemval As String
Private refreshdata__usersval As String
Private refreshdata__tagsval As String
Private refreshdata__groupsval As String
Private refreshdata__Itemsval As String
Private refreshdata__mvreportval As String
Private config__checkinchangesval As String
Private config__dataurlval As String
Private ribbonpointervalval As IRibbonUI


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


' Actions__AddItem  ''''''''''''''''''''''''''''''
Property Get Actions__AddItem() As String
    Debug.Print "Actions__AddItem"
    If Actions__AddItemval = "" Then
        Actions__AddItem = GetVariableSheetValue("actions__additemval")
    Else
        Actions__AddItem = Actions__AddItemval
    End If
End Property
Property Let Actions__AddItem(value As String)
Dim rs As String, rt As String, sirs As String, sirt As String
Dim resultsDict As New Dictionary

    Debug.Print "LetActions__AddItem", value
    AddItemExec rs, rt, sirs, sirt
    resultsDict.Add "rs", rs
    resultsDict.Add "rt", rt
    resultsDict.Add "sirs", sirs
    resultsDict.Add "sirt", sirt
    
    LetVariableSheetValue "actions__additemval", value, resultsDict
End Property




' RefreshData__Users  ''''''''''''''''''''''''''''''
Property Get RefreshData__Users() As String
    Debug.Print "RefreshData__Users"
    If refreshdata__usersval = "" Then
        RefreshData__Users = GetVariableSheetValue("refreshdata__usersval")
    Else
        RefreshData__Users = refreshdata__usersval
    End If
End Property
Property Let RefreshData__Users(value As String)
Dim results As Long

    Debug.Print "LetRefreshData__Users", value
    results = RefreshUsersExec
    
    LetVariableSheetValue "refreshdata__usersval", CStr(results)
End Property




' RefreshData__Tags  ''''''''''''''''''''''''''''''
Property Get RefreshData__Tags() As String
    Debug.Print "RefreshData__Tags"
    If refreshdata__tagsval = "" Then
        RefreshData__Tags = GetVariableSheetValue("refreshdata__Tagsval")
    Else
        RefreshData__Tags = refreshdata__tagsval
    End If
End Property
Property Let RefreshData__Tags(value As String)
Dim results As Long
    Debug.Print "LetRefreshData__Tags", value
    results = RefreshTagsExec
    LetVariableSheetValue "refreshdata__Tagsval", CStr(results)
End Property




' RefreshData__Groups  ''''''''''''''''''''''''''''''
Property Get RefreshData__Groups() As String
    Debug.Print "RefreshData__Groups"
    If refreshdata__groupsval = "" Then
        RefreshData__Groups = GetVariableSheetValue("refreshdata__Groupsval")
    Else
        RefreshData__Groups = refreshdata__groupsval
    End If
End Property
Property Let RefreshData__Groups(value As String)
Dim results As Long
    Debug.Print "LetRefreshData__Groups", value
    results = RefreshGroupsExec
    LetVariableSheetValue "refreshdata__Groupsval", CStr(results)
End Property




' RefreshData__Items  ''''''''''''''''''''''''''''''
Property Get RefreshData__Items() As String
    Debug.Print "RefreshData__Items"
    If refreshdata__Itemsval = "" Then
        RefreshData__Items = GetVariableSheetValue("refreshdata__Itemsval")
    Else
        RefreshData__Items = refreshdata__Itemsval
    End If
End Property
Property Let RefreshData__Items(value As String)
Dim resultsDict As Dictionary
    Debug.Print "LetRefreshData__Items", value
    Set resultsDict = RefreshItemsExec(Me.RefreshData__MVReport)
    LetVariableSheetValue "refreshdata__Itemsval", value, resultsDict
    Set resultsDict = Nothing
End Property



' RefreshData__MVReport  ''''''''''''''''''''''''''''''
Property Get RefreshData__MVReport() As String
    Debug.Print "RefreshData__MVReport"
    If refreshdata__mvreportval = "" Then
        RefreshData__MVReport = GetVariableSheetValue("refreshdata__MVReportval")
    Else
        RefreshData__MVReport = refreshdata__mvreportval
    End If
End Property
Property Let RefreshData__MVReport(value As String)
    Debug.Print "LetRefreshData__MVReport", value
    LetVariableSheetValue "refreshdata__MVReportval", value
End Property



' Config__CheckInChanges  ''''''''''''''''''''''''''''''
Property Get Config__CheckInChanges() As String
    Debug.Print "Config__CheckInChanges"
    If config__checkinchangesval = "" Then
        Config__CheckInChanges = GetVariableSheetValue("config__checkinchangesval")
    Else
        Config__CheckInChanges = config__checkinchangesval
    End If
End Property
Property Let Config__CheckInChanges(value As String)
    Debug.Print "LetConfig__CheckInChanges", value
    Application.Run "VBAUtils.xlsm!CheckInChangesExec"
    LetVariableSheetValue "config__checkinchangesval", value
End Property



' Config__DataUrl  ''''''''''''''''''''''''''''''
Property Get Config__DataUrl() As String
    Debug.Print "Config__DataUrl"
    If config__dataurlval = "" Then
        Config__DataUrl = GetVariableSheetValue("config__DataUrlval")
    Else
        Config__DataUrl = config__dataurlval
    End If
End Property
Property Let Config__DataUrl(value As String)
    Debug.Print "LetConfig__DataUrl", value
    LetVariableSheetValue "config__DataUrlval", value
End Property
Sub LetVariableSheetValue(varname As String, value As String, Optional resultsDict As Dictionary = Nothing)
Dim resultsRange As Range
Dim key As Variant

    Debug.Print "LetVariableSheetValue", varname, value
    Set resultsRange = Workbooks(bookname).Sheets(persistsheetname).Range(varname)
    
    If Not resultsDict Is Nothing Then
        For Each key In resultsDict.Keys()
            resultsRange.value = key & ":" & resultsDict(key)
            Set resultsRange = resultsRange.Offset(, 1)
        Next key
    Else
        resultsRange.value = value
    End If
    
End Sub

Function GetVariableSheetValue(varname As String) As String
     GetVariableSheetValue = Workbooks(bookname).Sheets(persistsheetname).Range(varname).value
     Debug.Print "GetVariableSheetValue", varname
End Function

Sub Rehydrate()
    Application.Run "vbautils.xlsm!RehydrateRangeFromFile", bookname, sheetName, persistrangename, persistfilename, persistrangelen
End Sub

Sub Persist()
    Application.Run "vbautils.xlsm!PersistRangeToFile", bookname, sheetName, persistrangename, persistfilename
End Sub

Private Sub Class_Initialize()
    Debug.Print "Class_Initialize"
    bookname = "MO.xlsm"
    persistsheetname = "Persist"
    persistfilename = Environ("USERPROFILE") & "\Deploy\.MO_persist.csv"
    persistrangename = "persistdata"
    persistrangelen = Workbooks(bookname).Sheets(persistsheetname).Range(persistrangename).Rows.Count

End Sub



