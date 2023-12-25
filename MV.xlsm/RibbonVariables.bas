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

Private config__working_dirval As String
Private config__max_monday_itemsval As String
Private config__input_dateval As String
Private Config__monday_email_suffixval As String
Private config__monday_email_prefixval As String
Private config__output_folder_sheetval As String
Private config__template_fileval As String
Private config__status_filterval As String


Private batchupdateval As String

Private refreshfoldersval As String
Private refreshupdatesval As String
Private subitemparentval As String
Private latestval As String
Private openreportval As String
Private savereportval As String

Private statusfilter__completedval As String
Private statusfilter__doneval As String
Private statusfilter__workingval As String
Private statusfilter__not_startedval As String


Private search__subitem_namesval As String
Private search__item_namesval As String
Private search__allval As String

Private bookname As String
Private sheetname As String
Private persistfilename As String


Private persistrangename As String
Private persistrangelen As Long

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

' Batch Update  ''''''''''''''''''''''''''''''
Property Get BatchUpdate() As String
    Debug.Print "BatchUpdate"
    If batchupdateval = "" Then
        BatchUpdate = GetVariableSheetValue("batchupdateval")
    Else
        BatchUpdate = batchupdateval
    End If
End Property
Property Let BatchUpdate(value As String)
    Debug.Print "LetBatchUpdate", value
    LetVariableSheetValue "batchupdateval", value
End Property

' Search All  ''''''''''''''''''''''''''''''
Property Get Search__All() As String
    Debug.Print "Search__All"
    If search__allval = "" Then
        Search__All = GetVariableSheetValue("search__allval")
    Else
        Search__All = search__allval
    End If
End Property
Property Let Search__All(value As String)
    Debug.Print "LetSearch__All", value
    LetVariableSheetValue "search__allval", value
End Property

' Search Item Names  ''''''''''''''''''''''''''''''
Property Get Search__Item_Names() As String
    Debug.Print "Search__Item_Names"
    If search__item_namesval = "" Then
        Search__Item_Names = GetVariableSheetValue("search__item_namesval")
    Else
        Search__Item_Names = search__item_namesval
    End If
End Property
Property Let Search__Item_Names(value As String)
    Debug.Print "LetSearch__Item_Names", value
    LetVariableSheetValue "search__item_namesval", value
End Property

' Search Sub Item Names  ''''''''''''''''''''''''''''''
Property Get Search__Sub_Item_Names() As String
    Debug.Print "Search__Sub_Item_Names"
    If search__sub_item_namesval = "" Then
        Search__Sub_Item_Names = GetVariableSheetValue("search__sub_item_namesval")
    Else
        Search__Sub_Item_Names = search__sub_item_namesval
    End If
End Property
Property Let Search__Sub_Item_Names(value As String)
    Debug.Print "LetSearch__Sub_Item_Names", value
    LetVariableSheetValue "search__sub_item_namesval", value
End Property

' Open Report  ''''''''''''''''''''''''''''''
Property Get OpenReport() As String
    Debug.Print "OpenReport"
    If openreportval = "" Then
        OpenReport = GetVariableSheetValue("openreportval")
    Else
        OpenReport = openreportval
    End If
End Property
Property Let OpenReport(value As String)
    Debug.Print "LetOpenReport", value
    LetVariableSheetValue "openreportval", value
End Property

' Save Report  ''''''''''''''''''''''''''''''
Property Get SaveReport() As String
    Debug.Print "SaveReport"
    If savereportval = "" Then
        SaveReport = GetVariableSheetValue("savereportval")
    Else
        SaveReport = savereportval
    End If
End Property
Property Let SaveReport(value As String)
    Debug.Print "LetSaveReport", value
    LetVariableSheetValue "savereportval", value
End Property


' Latest  ''''''''''''''''''''''''''''''
Property Get Latest() As String
    Debug.Print "Latest"
    If latestval = "" Then
        Latest = GetVariableSheetValue("latestval")
    Else
        Latest = latestval
    End If
End Property
Property Let Latest(value As String)
    Debug.Print "LetLatest", value
    LetVariableSheetValue "latestval", value
End Property

' Refresh Folders  ''''''''''''''''''''''''''''''
Property Get RefreshFolders() As String
    Debug.Print "RefreshFolders"
    If refreshfoldersval = "" Then
        RefreshFolders = GetVariableSheetValue("refreshfoldersval")
    Else
        RefreshFolders = refreshfoldersval
    End If
End Property
Property Let RefreshFolders(value As String)
    Debug.Print "LetRefreshFolders", value
    LetVariableSheetValue "refreshfoldersval", value
End Property

' Refresh Updates  ''''''''''''''''''''''''''''''
Property Get RefreshUpdates() As String
    Debug.Print "RefreshUpdates"
    If refreshupdatesval = "" Then
        RefreshUpdates = GetVariableSheetValue("refreshupdatesval")
    Else
        RefreshUpdates = refreshupdatesval
    End If
End Property
Property Let RefreshUpdates(value As String)
    Debug.Print "LetRefreshUpdates", value
    LetVariableSheetValue "refreshupdatesval", value
End Property

' Sub Item Parent   ''''''''''''''''''''''''''''''
Property Get SubItemParent() As String
    Debug.Print "SubItemParent"
    If subitemparentval = "" Then
        SubItemParent = GetVariableSheetValue("subitemparentval")
    Else
        SubItemParent = subitemparentval
    End If
End Property
Property Let SubItemParent(value As String)
    Debug.Print "LetSubItemParent", value
    LetVariableSheetValue "subitemparentval", value
End Property

' Input Date  ''''''''''''''''''''''''''''''
Property Get Config__Input_Date() As String
    Debug.Print "config__Working_Dirval"
    todaysdate = Application.Run("vbautils.xlsm!GetNow", "yyyymmdd")
    
    If config__input_dateval = "" Then
        Config__Input_Date = todaysdate
        'GetVariableSheetValue("config__Input_Dateval")
    Else
        Config__Input_Date = config__input_dateval
    End If
End Property
Property Let Config__Input_Date(value As String)
    Debug.Print "LetConfig__Input_Date", value
    LetVariableSheetValue "config__Input_Dateval", value
End Property


' Working Directory  ''''''''''''''''''''''''''''''
Property Get Config__Working_Dir() As String
    Debug.Print "config__Working_Dirval"
    If config__working_dirval = "" Then
        Config__Working_Dir = GetVariableSheetValue("config__Working_Dirval")
    Else
        Config__Working_Dir = config__working_dirval
    End If
End Property
Property Let Config__Working_Dir(value As String)
    Debug.Print "LetConfig__Working_Dir", value
    LetVariableSheetValue "config__Working_Dirval", value
End Property

' Max Monday Items  ''''''''''''''''''''''''''''''
Property Get Config__Max_Monday_Items() As String
    Debug.Print "Config__Max_Monday_Items"
    If config__max_monday_itemsval = "" Then
        Config__Max_Monday_Items = GetVariableSheetValue("config__Max_Monday_Itemsval")
    Else
        Config__Max_Monday_Items = config__max_monday_itemsval
    End If
End Property
Property Let Config__Max_Monday_Items(value As String)
    Debug.Print "LetConfig__Max_Monday_Items", value
    LetVariableSheetValue "config__Max_Monday_Itemsval", value
End Property


' Status Filter String  ''''''''''''''''''''''''''''''
Property Get Config__Status_Filter() As String
Dim tmpstring As String
Dim filterStatusArray As Variant
Dim filterValue As String

    filterStatusArray = Array("Completed", "Done", "Not_Started", "Working")
    For i = 0 To UBound(filterStatusArray)
        filterValue = CallByName(Me, "StatusFilter__" + filterStatusArray(i), VbGet)
        If filterValue = "True" Then
            If tmpstring = "" Then
                tmpstring = filterStatusArray(i)
            Else
                tmpstring = tmpstring + "," + filterStatusArray(i)
            End If
        End If
    Next i
    Debug.Print "Config__Status_Filter"

    Config__Status_Filter = tmpstring

End Property
Property Let Config__Status_Filter(value As String)
    Debug.Print "LetConfig__Status_Filter", value
    LetVariableSheetValue "config__status_filterval", value
End Property


' Template File  ''''''''''''''''''''''''''''''
Property Get Config__Template_File() As String
    Debug.Print "Config__Template_File"
    If config__template_fileval = "" Then
        Config__Template_File = GetVariableSheetValue("config__template_fileval")
    Else
        Config__Template_File = config__template_fileval
    End If
End Property
Property Let Config__Template_File(value As String)
    Debug.Print "LetConfig__Template_File", value
    LetVariableSheetValue "config__template_fileval", value
End Property

' Monday Email Suffix  ''''''''''''''''''''''''''''''
Property Get Config__Monday_Email_Suffix() As String
    Debug.Print "Config__Monday_Email_Suffix"
    If Config__monday_email_suffixval = "" Then
        Config__Monday_Email_Suffix = GetVariableSheetValue("config__monday_email_suffixval")
    Else
        Config__Monday_Email_Suffix = Config__monday_email_suffixval
    End If
End Property
Property Let Config__Monday_Email_Suffix(value As String)
    Debug.Print "LetConfig__Monday_Email_Suffix", value
    LetVariableSheetValue "config__monday_email_suffixval", value
End Property

' Monday Email Prefix  ''''''''''''''''''''''''''''''
Property Get Config__Monday_Email_Prefix() As String
    Debug.Print "Config__Monday_Email_Prefix"
    If config__monday_email_prefixval = "" Then
        Config__Monday_Email_Prefix = GetVariableSheetValue("config__monday_email_prefixval")
    Else
        Config__Monday_Email_Prefix = config__monday_email_prefixval
    End If
End Property
Property Let Config__Monday_Email_Prefix(value As String)
    Debug.Print "LetConfig__Monday_Email_Prefix", value
    LetVariableSheetValue "config__monday_email_prefixval", value
End Property

' Output Folder Sheet  ''''''''''''''''''''''''''''''
Property Get Config__Output_Folder_Sheet() As String
    Debug.Print "Config__Output_Folder_Sheet"
    If config__output_folder_sheetval = "" Then
        Config__Output_Folder_Sheet = GetVariableSheetValue("config__output_folder_sheetval")
    Else
        Config__Output_Folder_Sheet = config__output_folder_sheetval
    End If
End Property
Property Let Config__Output_Folder_Sheet(value As String)
    Debug.Print "LetConfig__Output_Folder_Sheet", value
    LetVariableSheetValue "config__output_folder_sheetval", value
End Property

' Status Filter Completed  ''''''''''''''''''''''''''''''
Property Get StatusFilter__Completed() As String
    Debug.Print "StatusFilter__Completed"
    If statusfilter__completedval = "" Then
        StatusFilter__Completed = GetVariableSheetValue("statusfilter__Completedval")
    Else
        StatusFilter__Completed = statusfilter__completedval
    End If
End Property
Property Let StatusFilter__Completed(value As String)
    Debug.Print "LetStatusFilter__Completed", value
    LetVariableSheetValue "statusfilter__Completedval", value
End Property


' Status Filter Done ''''''''''''''''''''''''''''''
Property Get StatusFilter__Done() As String
    Debug.Print "StatusFilter__Done"
    If statusfilter__doneval = "" Then
        StatusFilter__Done = GetVariableSheetValue("statusfilter__Doneval")
    Else
        StatusFilter__Done = statusfilter__doneval
    End If
End Property
Property Let StatusFilter__Done(value As String)
    Debug.Print "LetStatusFilter__Done", value
    LetVariableSheetValue "statusfilter__Doneval", value
End Property


' Status Filter Working ''''''''''''''''''''''''''''''
Property Get StatusFilter__Working() As String
    Debug.Print "StatusFilter__Working"
    If statusfilter__workingval = "" Then
        StatusFilter__Working = GetVariableSheetValue("statusfilter__Workingval")
    Else
        StatusFilter__Working = statusfilter__workingval
    End If
End Property
Property Let StatusFilter__Working(value As String)
    Debug.Print "LetStatusFilter__Working", value
    LetVariableSheetValue "statusfilter__Workingval", value
End Property


' Status Filter Not Started ''''''''''''''''''''''''''''''
Property Get StatusFilter__Not_Started() As String
    Debug.Print "StatusFilter__Not_Started"
    If statusfilter__not_startedval = "" Then
        StatusFilter__Not_Started = GetVariableSheetValue("statusfilter__Not_Startedval")
    Else
        StatusFilter__Not_Started = statusfilter__not_startedval
    End If
End Property
Property Let StatusFilter__Not_Started(value As String)
    Debug.Print "LetStatusFilter__Not_Started", value
    LetVariableSheetValue "statusfilter__Not_Startedval", value
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

Sub Rehydrate()
    Application.Run "vbautils.xlsm!RehydrateRangeFromFile", bookname, sheetname, persistrangename, persistfilename, persistrangelen
End Sub

Sub Persist()
    Application.Run "vbautils.xlsm!PersistRangeToFile", bookname, sheetname, persistrangename, persistfilename
End Sub


Private Sub Class_Initialize()
    Debug.Print "Class_Initialize"
    sheetname = "Persist"
    bookname = "MV.xlsm"
    persistfilename = Environ("USERPROFILE") & "\Deploy\.MV_persist.csv"
    persistrangename = "persistdata"
    persistrangelen = Workbooks(bookname).Sheets(sheetname).Range(persistrangename).Rows.Count
    
    debugflagval = ""
    userval = ""
    agefilterval = ""
    sortval = ""
    
    statusfiltercompletedval = ""
    
End Sub


