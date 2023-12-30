VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RibbonVariables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private bookname As String
Private persistsheetname As String
Private persistfilename As String
Private persistrangename As String
Private persistrangelen As String

Private action__refresh_capsuledataval As String
Private action__add_recordval As String
Private action__delete_recordval As String
Private action__get_recordval As String
Private action__update_recordval As String
Private action__checkin_changesval As String

Private pick__entityopportunityval As String
Private pick__entitypersonval As String
Private pick__entityorganisationval As String

Private searchby__idval As String
Private searchby__nameval As String
Private searchby__typeval As String

Private fastadd__Organisationval As String
Private fastadd__Opportunityval As String
Private fastadd__Personval As String

Private config__dataurlval As String

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



' Fast Add Organisation  ''''''''''''''''''''''''''''''
Property Get Fastadd__Organisation() As String
    Debug.Print "Fastadd__Organisation"
    If fastadd__Organisationval = "" Then
        Fastadd__Organisation = GetVariableSheetValue("fastadd__organisationval")
    Else
        Fastadd__Organisation = fastadd__Organisationval
    End If
End Property
Property Let Fastadd__Organisation(value As String)
    Debug.Print "LetFastadd__Organisation", value
    LetVariableSheetValue "fastadd__organisationval", value
End Property


' Fast Add Person  ''''''''''''''''''''''''''''''
Property Get fastadd__Person() As String
    Debug.Print "Fastadd__Person"
    If fastadd__Personval = "" Then
        fastadd__Person = GetVariableSheetValue("fastadd__personval")
    Else
        fastadd__Person = fastadd__Personval
    End If
End Property
Property Let fastadd__Person(value As String)
    Debug.Print "LetFastadd__Person", value
    LetVariableSheetValue "fastadd__personval", value
End Property


' Fast Add Opportunity  ''''''''''''''''''''''''''''''
Property Get fastadd__Opportunity() As String
    Debug.Print "Fastadd__Opportunity"
    If fastadd__Opportunityval = "" Then
        fastadd__Opportunity = GetVariableSheetValue("fastadd__opportunityval")
    Else
        fastadd__Opportunity = fastadd__Opportunityval
    End If
End Property
Property Let fastadd__Opportunity(value As String)
    Debug.Print "LetFastadd__Opportunity", value
    LetVariableSheetValue "fastadd__opportunityval", value
End Property


' Config Dataurl  ''''''''''''''''''''''''''''''
Property Get Config__Dataurl() As String
    Debug.Print "Config__Dataurl"
    If config__dataurlval = "" Then
        Config__Dataurl = GetVariableSheetValue("config__dataurlval")
    Else
        Config__Dataurl = config__dataurlval
    End If
End Property
Property Let Config__Dataurl(value As String)
    Debug.Print "LetConfig__Dataurl", value
    LetVariableSheetValue "config__dataurlval", value
End Property


' Search By ID  ''''''''''''''''''''''''''''''
Property Get SearchBy__ID() As String
    Debug.Print "SearchBy__ID"
    If searchby__idval = "" Then
        SearchBy__ID = GetVariableSheetValue("searchby__idval")
    Else
        SearchBy__ID = searchby__idval
    End If
End Property
Property Let SearchBy__ID(value As String)
    Debug.Print "LetSearchBy__ID", value
    LetVariableSheetValue "searchby__idval", value
End Property

' Search By Type  ''''''''''''''''''''''''''''''
Property Get SearchByType() As String
    Debug.Print "SearchByType"
    If searchby__typeval = "" Then
        SearchByType = GetVariableSheetValue("searchby__typeval")
    Else
        SearchByType = searchby__typeval
    End If
End Property
Property Let SearchByType(value As String)
    Debug.Print "LetSearchByType", value
    LetVariableSheetValue "searchby__typeval", value
End Property

' Search By Name  ''''''''''''''''''''''''''''''
Property Get SearchBy__Name() As String
    Debug.Print "SearchBy__Name"
    If searchby__nameval = "" Then
        SearchBy__Name = GetVariableSheetValue("searchby__nameval")
    Else
        SearchBy__Name = searchby__nameval
    End If
End Property
Property Let SearchBy__Name(value As String)
    Debug.Print "LetSearchBy__Name", value
    LetVariableSheetValue "searchby__nameval", value
End Property

' RefreshCapsule  ''''''''''''''''''''''''''''''
Property Get RefreshCapsuleData() As String
    Debug.Print "RefreshCapsuleData"
    If action__refresh_capsuledataval = "" Then
        RefreshCapsuleData = GetVariableSheetValue("action__refresh_capsuledataval")
    Else
        RefreshCapsuleData = action__refresh_capsuledataval
    End If
End Property
Property Let RefreshCapsuleData(value As String)
Dim resultsDict As New Dictionary
    Debug.Print "LetRefreshCapsuleData", value
    RefreshCapsuleDataExec resultsDict
    LetVariableSheetValue "action__refresh_capsuledataval", value, resultsDict
End Property



' AddRecord  ''''''''''''''''''''''''''''''
Property Get AddRecord() As String
    Debug.Print "AddRecord"
    If action__add_recordval = "" Then
        AddRecord = GetVariableSheetValue("action__add_recordval")
    Else
        AddRecord = action__add_recordval
    End If
End Property
Property Let AddRecord(value As String)
Dim responseStatus As String, responseText As String, outputStr As String, resultStr As String, datastr As String
    Debug.Print "LetAddRecord", value
    AddRecordExec responseStatus, responseText, datastr
    LetVariableSheetValue "action__add_recordval", responseText
End Property




' DeleteRecord  ''''''''''''''''''''''''''''''
Property Get DeleteRecord() As String
    Debug.Print "DeleteRecord"
    If action__delete_recordval = "" Then
        DeleteRecord = GetVariableSheetValue("action__delete_recordval")
    Else
        DeleteRecord = action__delete_recordval
    End If
End Property
Property Let DeleteRecord(value As String)
Dim rs As String, rt As String, lookupId As String, recordType As String
    Set RV = New RibbonVariables
    Debug.Print "LetDeleteRecord", value
    lookupId = CallByName(RV, "SearchBy__Id", VbGet)
    recordType = CallByName(RV, "SearchByType", VbGet)
    recordType = Split(Split(recordType, "__")(1), "_")(1)
    
    DeleteRecordExec lookupId, recordType, rs, rt
    LetVariableSheetValue "action__delete_recordval", rs
    
End Property



' GetRecord  ''''''''''''''''''''''''''''''
Property Get GetRecord() As String
    Debug.Print "GetRecord"
    If action__get_recordval = "" Then
        GetRecord = GetVariableSheetValue("action__get_recordval")
    Else
        GetRecord = action__get_recordval
    End If
End Property
Property Let GetRecord(value As String)
Dim rs As String, rt As String, lookupId As String, recordType As String
    Set RV = New RibbonVariables
    Debug.Print "LetGetRecord", value
    lookupId = CallByName(RV, "SearchBy__Id", VbGet)
    recordType = CallByName(RV, "SearchByType", VbGet)
    recordType = Split(Split(recordType, "__")(1), "_")(1)
    GetRecordExec lookupId, recordType, rs, rt
    LetVariableSheetValue "action__get_recordval", rt
    Set RV = Nothing
End Property



' UpdateRecord  ''''''''''''''''''''''''''''''
Property Get UpdateRecord() As String
    Debug.Print "UpdateRecord"
    If action__update_recordval = "" Then
        UpdateRecord = GetVariableSheetValue("action__update_recordval")
    Else
        UpdateRecord = action__update_recordval
    End If
End Property
Property Let UpdateRecord(value As String)
Dim rs As String, rt As String, lookupId As String, recordType As String
    Set RV = New RibbonVariables
    Debug.Print "LetUpdateRecord", value

    lookupId = CallByName(RV, "SearchBy__Id", VbGet)
    recordType = CallByName(RV, "SearchByType", VbGet)
    recordType = Split(Split(recordType, "__")(1), "_")(1)
    UpdateRecordExec lookupId, recordType, rs, rt
    LetVariableSheetValue "action__update_recordval", value
    
    
    'LetVariableSheetValue "action__get_recordval", rt
    Set RV = Nothing
    
End Property



' CheckInChanges  ''''''''''''''''''''''''''''''
Property Get CheckInChanges() As String
    Debug.Print "CheckInChanges"
    If action__checkin_changesval = "" Then
        CheckInChanges = GetVariableSheetValue("action__checkin_changesval")
    Else
        CheckInChanges = action__checkin_changesval
    End If
End Property
Property Let CheckInChanges(value As String)
    Debug.Print "LetCheckInChanges", value
    CheckInChangesExec
    LetVariableSheetValue "action__checkin_changesval", value
End Property


' EntityOpportunity  ''''''''''''''''''''''''''''''
Property Get EntityOpportunity() As String
    Debug.Print "EntityOpportunity"
    If pick__entityopportunityval = "" Then
        EntityOpportunity = GetVariableSheetValue("pick__entityopportunityval")
    Else
        EntityOpportunity = pick__entityopportunityval
    End If
End Property
Property Let EntityOpportunity(value As String)
    Debug.Print "LetEntityOpportunity", value
    SetupRecordDefaults "opportunity"
    LetVariableSheetValue "pick__entityopportunityval", value
End Property


' EntityPerson  ''''''''''''''''''''''''''''''
Property Get EntityPerson() As String
    Debug.Print "EntityPerson"
    If pick__entitypersonval = "" Then
        EntityPerson = GetVariableSheetValue("pick__entitypersonval")
    Else
        EntityPerson = pick__entitypersonval
    End If
End Property
Property Let EntityPerson(value As String)
    Debug.Print "EntityPerson", value
    
    SetupRecordDefaults "person"
    
    LetVariableSheetValue "pick__entitypersonval", value
End Property



' EntityOrganisation  ''''''''''''''''''''''''''''''
Property Get EntityOrganisation() As String
    Debug.Print "EntityOrganisation"
    If pick__entityorganisationval = "" Then
        EntityPerson = GetVariableSheetValue("pick__entityorganisationval")
    Else
        EntityPerson = pick__entityorganisationval
    End If
End Property
Property Let EntityOrganisation(value As String)
    Debug.Print "EntityOrganisation", value
    SetupRecordDefaults "organisation"
    LetVariableSheetValue "pick__entityorganisationval", value
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
    Application.Run "vbautils.xlsm!RehydrateRangeFromFile", bookname, sheetname, persistrangename, persistfilename, persistrangelen
End Sub

Sub Persist()
    Application.Run "vbautils.xlsm!PersistRangeToFile", bookname, sheetname, persistrangename, persistfilename
End Sub


Private Sub Class_Initialize()
    Debug.Print "Class_Initialize"
    bookname = "CA.xlsm"
    persistsheetname = "Persist"
    persistfilename = Environ("USERPROFILE") & "\Deploy\.CA_persist.csv"
    persistrangename = "persistdata"
    persistrangelen = Workbooks(bookname).Sheets(persistsheetname).Range(persistrangename).Rows.Count

End Sub



