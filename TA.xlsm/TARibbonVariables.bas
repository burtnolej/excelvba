VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TARibbonVariables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private persistsheetname As String
Private persistfilename As String
Private persistrangename As String
Private persistrangelen As String
Private bookname As String

Private action__updateitemsval As String
Private action__loadappointmentsval As String
Private action__refreshmondayval As String
Private action__addcolorcodingval As String

Private settings__startdateval As String
Private settings__offsetval As String
Private settings__enddateval As String

Private useremail__jonbutlerval As String
Private useremail__alisonhoodval As String
Private useremailval As String

Private additem__startdateval As String
Private additem__starttimeval As String
Private additem__subjectval As String
Private additem__mondayitemval As String
Private additem__mondaysubitemval As String
Private additem__categoryval As String
Private additem__durationval As String
Private additem__mondayitemidval As String
Private additem__mondaysubitemidval As String

Private admin__checkinchangesval As String

Private additem__addcalendaritemval As String
Private ribbonpointerval As IRibbonControl

Private Sub Class_Initialize()
    Debug.Print "Class_Initialize"
    bookname = "TA.xlsm"
    persistsheetname = "Persist"
    persistfilename = Environ("USERPROFILE") & "\Deploy\.TA_persist.csv"
    persistrangename = "persistdata"
    persistrangelen = Workbooks(bookname).Sheets(persistsheetname).Range(persistrangename).Rows.Count

End Sub



' Admin__CheckInChanges  ''''''''''''''''''''''''''''''
Property Get Admin__CheckInChanges() As String
    Debug.Print "Admin__CheckInChanges"
    If admin__checkinchangesval = "" Then
        Admin__CheckInChanges = GetVariableSheetValue("admin__checkinchangesval")
    Else
        Admin__CheckInChanges = admin__checkinchangesval
    End If
End Property
Property Let Admin__CheckInChanges(value As String)
Dim rs As String, rt As String, sirs As String, sirt As String, mondayItemId As String
Dim resultsDict As New Dictionary

    Debug.Print "Admin__CheckInChanges", value
    Application.Run "vbautils.xlsm!CheckInChangesExec"
    LetVariableSheetValue "admin__checkinchangesval", value
End Property


' Action__AddCalendarItem  ''''''''''''''''''''''''''''''
Property Get AddItem__AddCalendarItem() As String
    Debug.Print "AddItem__AddCalendarItem"
    If additem__addcalendaritemval = "" Then
        AddItem__AddCalendarItem = GetVariableSheetValue("additem__addcalendaritemval")
    Else
        AddItem__AddCalendarItem = additem__addcalendaritemval
    End If
End Property
Property Let AddItem__AddCalendarItem(value As String)
Dim rs As String, rt As String, sirs As String, sirt As String, mondayItemId As String
Dim resultsDict As New Dictionary

    Debug.Print "AddItem__AddCalendarItem", value
    
    If Me.AddItem__MondayItemId = "" Then
        mondayItemId = Me.AddItem__MondaySubItemId
    Else
        mondayItemId = Me.AddItem__MondayItemId
    End If
    
    AddMeetingExec Me.AddItem__StartDate, Me.AddItem__StartTime, Me.AddItem__Duration, mondayItemId, Me.AddItem__Category, Me.AddItem__Subject
    LetVariableSheetValue "additem__addcalendaritemval", value
End Property

' AddItem__Duration  ''''''''''''''''''''''''''''''
Property Get AddItem__Duration() As String
    Debug.Print "AddItem__Duration"
    If useremailval = "" Then
        AddItem__Duration = GetVariableSheetValue("additem__Durationval")
    Else
        AddItem__Duration = additem__durationval
    End If
End Property
Property Let AddItem__Duration(value As String)
Dim rs As String, rt As String, sirs As String, sirt As String
Dim resultsDict As New Dictionary

    Debug.Print "AddItem__Duration", value
    LetVariableSheetValue "additem__Durationval", value
End Property


' AddItem__StartDate  ''''''''''''''''''''''''''''''
Property Get AddItem__StartDate() As String
    Debug.Print "AddItem__StartDate"
    If additem__startdateval = "" Then
        AddItem__StartDate = GetVariableSheetValue("additem__startdateval")
    Else
        AddItem__StartDate = additem__startdateval
    End If
End Property
Property Let AddItem__StartDate(value As String)
Dim rs As String, rt As String, sirs As String, sirt As String
Dim resultsDict As New Dictionary

    Debug.Print "AddItem__StartDate", value
    LetVariableSheetValue "additem__startdateval", value
End Property


' AddItem__StartTime  ''''''''''''''''''''''''''''''
Property Get AddItem__StartTime() As String
    Debug.Print "AddItem__StartTime"
    If additem__starttimeval = "" Then
        AddItem__StartTime = Format(GetVariableSheetValue("additem__StartTimeval"), "hh:mm:ss")
    Else
        AddItem__StartTime = additem__starttimeval
    End If
End Property
Property Let AddItem__StartTime(value As String)
Dim rs As String, rt As String, sirs As String, sirt As String
Dim resultsDict As New Dictionary

    Debug.Print "AddItem__StartTime", value
    LetVariableSheetValue "additem__StartTimeval", value
End Property


' AddItem__Subject  ''''''''''''''''''''''''''''''
Property Get AddItem__Subject() As String
    Debug.Print "AddItem__Subject"
    If additem__subjectval = "" Then
        AddItem__Subject = GetVariableSheetValue("additem__Subjectval")
    Else
        AddItem__Subject = additem__subjectval
    End If
End Property
Property Let AddItem__Subject(value As String)
Dim rs As String, rt As String, sirs As String, sirt As String
Dim resultsDict As New Dictionary

    Debug.Print "AddItem__Subject", value
    LetVariableSheetValue "additem__Subjectval", value
End Property


' AddItem__MondayItem  ''''''''''''''''''''''''''''''
Property Get AddItem__MondayItem() As String
    Debug.Print "AddItem__MondayItem"
    If additem__mondayitemval = "" Then
        AddItem__MondayItem = GetVariableSheetValue("additem__MondayItemval")
    Else
        AddItem__MondayItem = additem__mondayitemval
    End If
End Property
Property Let AddItem__MondayItem(value As String)
Dim rs As String, rt As String, sirs As String, sirt As String
Dim resultsDict As New Dictionary

    Debug.Print "AddItem__MondayItem", value
    LetVariableSheetValue "additem__MondayItemval", value
End Property

' AddItem__MondayItemId  ''''''''''''''''''''''''''''''
Property Get AddItem__MondayItemId() As String
    Debug.Print "AddItem__MondayItemId"
    If additem__mondayitemidval = "" Then
        AddItem__MondayItemId = GetVariableSheetValue("additem__MondayItemIdval")
    Else
        AddItem__MondayItemId = additem__mondayitemidval
    End If
End Property
Property Let AddItem__MondayItemId(value As String)
Dim rs As String, rt As String, sirs As String, sirt As String
Dim resultsDict As New Dictionary

    Debug.Print "AddItem__MondayItemId", value
    LetVariableSheetValue "additem__MondayItemIdval", value
End Property

' AddItem__MondaySubItem  ''''''''''''''''''''''''''''''
Property Get AddItem__MondaySubItem() As String
    Debug.Print "AddItem__MondaySubItem"
    If additem__mondaysubitemval = "" Then
        AddItem__MondaySubItem = GetVariableSheetValue("additem__MondaysubItemval")
    Else
        AddItem__MondaySubItem = additem__mondaysubitemval
    End If
End Property
Property Let AddItem__MondaySubItem(value As String)
Dim rs As String, rt As String, sirs As String, sirt As String
Dim resultsDict As New Dictionary

    Debug.Print "AddItem__MondaySubItem", value
    LetVariableSheetValue "additem__MondaysubItemval", value
End Property


' AddItem__MondaySubItemId  ''''''''''''''''''''''''''''''
Property Get AddItem__MondaySubItemId() As String
    Debug.Print "AddItem__MondaySubItemId"
    If additem__mondaysubitemidval = "" Then
        AddItem__MondaySubItemId = GetVariableSheetValue("additem__MondaySubItemIdval")
    Else
        AddItem__MondaySubItemId = additem__mondaysubitemidval
    End If
End Property
Property Let AddItem__MondaySubItemId(value As String)
Dim rs As String, rt As String, sirs As String, sirt As String
Dim resultsDict As New Dictionary

    Debug.Print "AddItem__MondaySubItemId", value
    LetVariableSheetValue "additem__MondaySubItemIdval", value
End Property


' AddItem__Category  ''''''''''''''''''''''''''''''
Property Get AddItem__Category() As String
    Debug.Print "AddItem__Category"
    If additem__categoryval = "" Then
        AddItem__Category = GetVariableSheetValue("additem__Categoryval")
    Else
        AddItem__Category = additem__categoryval
    End If
End Property
Property Let AddItem__Category(value As String)
Dim rs As String, rt As String, sirs As String, sirt As String
Dim resultsDict As New Dictionary

    Debug.Print "AddItem__Category", value
    LetVariableSheetValue "additem__Categoryval", value
End Property

' UserEmail  ''''''''''''''''''''''''''''''
Property Get UserEmail() As String
    Debug.Print "UserEmail"
    If useremailval = "" Then
        UserEmail = GetVariableSheetValue("useremailval")
    Else
        UserEmail = useremailval
    End If
End Property
Property Let UserEmail(value As String)
Dim rs As String, rt As String, sirs As String, sirt As String
Dim resultsDict As New Dictionary

    Debug.Print "UserEmail", value
    LetVariableSheetValue "useremailval", value
End Property



' Settings__StartDate  ''''''''''''''''''''''''''''''
Property Get Settings__StartDate() As String
    Debug.Print "Settings__StartDate"
    If settings__startdateval = "" Then
        Settings__StartDate = GetVariableSheetValue("settings__startdateval")
    Else
        Settings__StartDate = settings__startdateval
    End If
End Property
Property Let Settings__StartDate(value As String)
Dim rs As String, rt As String, sirs As String, sirt As String
Dim resultsDict As New Dictionary

    Debug.Print "Settings__StartDate", value
    LetVariableSheetValue "settings__startdateval", value
End Property



' Settings__Offset  ''''''''''''''''''''''''''''''
Property Get Settings__Offset() As String
    Debug.Print "Settings__Offset"
    If settings__offsetval = "" Then
        Settings__Offset = GetVariableSheetValue("settings__offsetval")
    Else
        Settings__Offset = settings__offsetval
    End If
End Property
Property Let Settings__Offset(value As String)
Dim rs As String, rt As String, sirs As String, sirt As String
Dim resultsDict As New Dictionary

    Debug.Print "LetSettings__Offset", value
    LetVariableSheetValue "settings__offsetval", value
End Property


' Settings__EndDate  ''''''''''''''''''''''''''''''
Property Get Settings__EndDate() As String
    Debug.Print "Settings__EndDate"
    If settings__enddateval = "" Then
        Settings__EndDate = GetVariableSheetValue("settings__enddateval")
    Else
        Settings__EndDate = settings__enddateval
    End If
End Property
Property Let Settings__EndDate(value As String)
Dim rs As String, rt As String, sirs As String, sirt As String
Dim resultsDict As New Dictionary

    Debug.Print "LetSettings__EndDate", value
    LetVariableSheetValue "settings__enddateval", value
End Property



' Actions__UpdateItem  ''''''''''''''''''''''''''''''
Property Get Action__UpdateItems() As String
    Debug.Print "Action__UpdateItems"
    If action__updateitemsval = "" Then
        Action__UpdateItems = GetVariableSheetValue("action__updateitemsval")
    Else
        Action__UpdateItems = action__updateitemsval
    End If
End Property
Property Let Action__UpdateItems(value As String)
Dim rs As String, rt As String, sirs As String, sirt As String
Dim resultsDict As New Dictionary

    Debug.Print "LetAction__UpdateItems", value
    UpdateItemsExec
    LetVariableSheetValue "action__updateitemsval", value
End Property



' Action__LoadAppointments  ''''''''''''''''''''''''''''''
Property Get Action__LoadAppointments() As String
    Debug.Print "Action__LoadAppointments"
    If action__loadappointmentsval = "" Then
        Action__LoadAppointments = GetVariableSheetValue("action__LoadAppointmentsval")
    Else
        Action__LoadAppointments = action__loadappointmentsval
    End If
End Property
Property Let Action__LoadAppointments(value As String)
Dim rs As String, rt As String, sirs As String, sirt As String, useremailval As String
Dim resultsDict As New Dictionary

    Debug.Print "LetAction__LoadAppointments", value
    useremailval = Right(Me.UserEmail, Len(Me.UserEmail) - 11) & "@veloxfintech.com"
    
    LoadAppointmentsExec Me.Settings__EndDate, Me.Settings__StartDate, useremailval
    LetVariableSheetValue "action__LoadAppointmentsval", value
End Property




' Action__RefreshMonday  ''''''''''''''''''''''''''''''
Property Get Action__RefreshMonday() As String
    Debug.Print "Action__RefreshMonday"
    If action__refreshmondayval = "" Then
        Action__RefreshMonday = GetVariableSheetValue("action__RefreshMondayval")
    Else
        Action__RefreshMonday = action__refreshmondayval
    End If
End Property
Property Let Action__RefreshMonday(value As String)
Dim rs As String, rt As String, sirs As String, sirt As String
Dim resultsDict As New Dictionary

    Debug.Print "LetAction__RefreshMonday", value
    RefreshMondayExec
    LetVariableSheetValue "action__RefreshMondayval", value
End Property


' Action__AddColorCoding  ''''''''''''''''''''''''''''''
Property Get Action__AddColorCoding() As String
    Debug.Print "Action__AddColorCoding"
    If action__addcolorcodingval = "" Then
        Action__AddColorCoding = GetVariableSheetValue("action__AddColorCodingval")
    Else
        Action__AddColorCoding = action__addcolorcodingval
    End If
End Property
Property Let Action__AddColorCoding(value As String)
Dim rs As String, rt As String, sirs As String, sirt As String
Dim resultsDict As New Dictionary

    Debug.Print "LetAction__AddColorCoding", value
    AddColorCodingExec
    LetVariableSheetValue "action__AddColorCodingval", value
End Property






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

Sub Rehydrate()
    Application.Run "vbautils.xlsm!RehydrateRangeFromFile", bookname, sheetName, persistrangename, persistfilename, persistrangelen
End Sub

Sub Persist()
    Application.Run "vbautils.xlsm!PersistRangeToFile", bookname, sheetName, persistrangename, persistfilename
End Sub

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
