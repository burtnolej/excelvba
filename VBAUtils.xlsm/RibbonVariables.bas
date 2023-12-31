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

Private docs__doc1val As String
Private windowsize__xval As String
Private windowsize__yval As String
Private windowsize__heightval As String
Private windowsize__widthval As String
Private windowSizeRight__xval As String
Private windowSizeRight__yval As String
Private windowSizeRight__heightval As String
Private windowSizeRight__widthval As String
Private windowSizeLeft__xval As String
Private windowSizeLeft__yval As String
Private windowSizeLeft__heightval As String
Private windowSizeLeft__widthval As String

Private settings__rootpathval As String
Private settings__dataurlval As String

Private toolaction__openval As String
Private toolaction__closeval As String
Private toolaction__minval As String
Private toolaction__maxval As String
Private toolaction__hideval As String
Private toolaction__showval As String

Private runningapps__CAval As String
Private runningapps__MOval As String
Private runningapps__MMval As String
Private runningapps__MVval As String
Private runningapps__ESval As String
Private runningapps__TAval As String
Private runningapps__DVval As String

Private utils__setxlontopval As String
Private utils__setxlnormalval As String
Private utils__showtoolsval As String
Private utils__hidetoolsval As String
Private utils__runribboneditorval As String
Private utils__displayvbeval As String
Private utils__closeribboneditorval As String
Private utils__resizethiswindowval As String
Private utils__checkinchangesval As String
Private utils__packuptoolsval As String
Private utils__editnewsletterval As String

Private choosetoolval As String
Private windowlocationval As String
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


' Utils__RunRibbonEditor  ''''''''''''''''''''''''''''''
Property Get Utils__RunRibbonEditor() As String
    Debug.Print "Utils__RunRibbonEditor"
    If utils__runribboneditorval = "" Then
        Utils__RunRibbonEditor = GetVariableSheetValue("utils__runribboneditorval")
    Else
        Utils__RunRibbonEditor = utils__runribboneditorval
    End If
End Property

Property Let Utils__RunRibbonEditor(value As String)
Dim bookname As String
    Debug.Print "LetUtils__RunRibbonEditor", value
    RunRibbonEditorExec "vbautils.xlsm"
    LetVariableSheetValue "utils__runribboneditorval", value
End Property


' Utils__ShowTools  ''''''''''''''''''''''''''''''
Property Get Utils__ShowTools() As String
    Debug.Print "Utils__ShowTools"
    If windowlocationval = "" Then
        Utils__ShowTools = GetVariableSheetValue("utils__showtoolsval")
    Else
        Utils__ShowTools = utils__showtoolsval
    End If
End Property

Property Let Utils__ShowTools(value As String)
Dim bookname As String
    Debug.Print "LetUtils__ShowTools", value
    ShowToolsExec "vbautils.xlsm"
    LetVariableSheetValue "utils__showtoolsval", value
End Property

' Utils__HideTools  ''''''''''''''''''''''''''''''
Property Get Utils__HideTools() As String
    Debug.Print "Utils__HideTools"
    If windowlocationval = "" Then
        Utils__HideTools = GetVariableSheetValue("utils__HideToolsval")
    Else
        Utils__HideTools = utils__hidetoolsval
    End If
End Property

Property Let Utils__HideTools(value As String)
Dim bookname As String
    Debug.Print "LetUtils__HideTools", value
    HideToolsExec "vbautils.xlsm"
    LetVariableSheetValue "utils__HideToolsval", value
End Property

' Utils__DisplayVBE  ''''''''''''''''''''''''''''''
Property Get Utils__DisplayVBE() As String
    Debug.Print "Utils__DisplayVBE"
    If windowlocationval = "" Then
        Utils__DisplayVBE = GetVariableSheetValue("utils__DisplayVBEval")
    Else
        Utils__DisplayVBE = utils__displayvbeval
    End If
End Property

Property Let Utils__DisplayVBE(value As String)
Dim bookname As String
    Debug.Print "LetUtils__DisplayVBE", value
    DisplayVBEExec
    LetVariableSheetValue "utils__DisplayVBEval", value
End Property


' Utils__EditNewsletter  ''''''''''''''''''''''''''''''
Property Get Utils__EditNewsletter() As String
    Debug.Print "Utils__EditNewsletter"
    If windowlocationval = "" Then
        Utils__EditNewsletter = GetVariableSheetValue("utils__EditNewsletterval")
    Else
        Utils__EditNewsletter = utils__editnewsletterval
    End If
End Property

Property Let Utils__EditNewsletter(value As String)
Dim bookname As String
    Debug.Print "LetUtils__EditNewsletter", value
    EditNewsletterExec
    LetVariableSheetValue "utils__EditNewsletterval", value
End Property


' Utils__CloseRibboneditor  ''''''''''''''''''''''''''''''
Property Get Utils__CloseRibboneditor() As String
    Debug.Print "Utils__CloseRibboneditor"
    If windowlocationval = "" Then
        Utils__CloseRibboneditor = GetVariableSheetValue("utils__CloseRibboneditorval")
    Else
        Utils__CloseRibboneditor = utils__closeribboneditorval
    End If
End Property

Property Let Utils__CloseRibboneditor(value As String)
Dim bookname As String
    Debug.Print "LetUtils__CloseRibboneditor", value
    CloseRibbonEditorExec
    LetVariableSheetValue "utils__CloseRibboneditorval", value
End Property


' Utils__SetXLOnTop  ''''''''''''''''''''''''''''''
Property Get Utils__SetXLOnTop() As String
    Debug.Print "Utils__SetXLOnTop"
    If windowlocationval = "" Then
        Utils__SetXLOnTop = GetVariableSheetValue("utils__SetXLOnTopval")
    Else
        Utils__SetXLOnTop = utils__setxlontopval
    End If
End Property

Property Let Utils__SetXLOnTop(value As String)
Dim bookname As String
    Debug.Print "LetUtils__SetXLOnTop", value
    SetXLOnTopExec
    LetVariableSheetValue "utils__SetXLOnTopval", value
End Property







' Utils__SetXLNormal  ''''''''''''''''''''''''''''''
Property Get Utils__SetXLNormal() As String
    Debug.Print "Utils__SetXLNormal"
    If windowlocationval = "" Then
        Utils__SetXLNormal = GetVariableSheetValue("utils__SetXLNormalval")
    Else
        Utils__SetXLNormal = utils__setxlnormalval
    End If
End Property

Property Let Utils__SetXLNormal(value As String)
Dim bookname As String
    Debug.Print "LetUtils__SetXLNormal", value
    SetXLNormalExec
    LetVariableSheetValue "utils__SetXLNormalval", value
End Property





' Utils__ResizeThiswindow  ''''''''''''''''''''''''''''''
Property Get Utils__ResizeThiswindow() As String
    Debug.Print "Utils__ResizeThiswindow"
    If windowlocationval = "" Then
        Utils__ResizeThiswindow = GetVariableSheetValue("utils__ResizeThiswindowval")
    Else
        Utils__ResizeThiswindow = utils__resizethiswindowval
    End If
End Property

Property Let Utils__ResizeThiswindow(value As String)
Dim bookname As String
    'bookname = Right(Me.ChooseTool, Len(Me.ChooseTool) - 12)
    Debug.Print "LetUtils__ResizeThiswindow", value
    ResizeWindowExec "vbautils.xlsm", Me.WindowSize__Width, Me.WindowSize__Height
    LetVariableSheetValue "utils__ResizeThiswindowval", value
End Property







' Utils__CheckInChanges  ''''''''''''''''''''''''''''''
Property Get Utils__CheckInChanges() As String
    Debug.Print "Utils__CheckInChanges"
    If windowlocationval = "" Then
        Utils__CheckInChanges = GetVariableSheetValue("utils__CheckInChangesval")
    Else
        Utils__CheckInChanges = utils__checkinchangesval
    End If
End Property

Property Let Utils__CheckInChanges(value As String)
Dim bookname As String
    Debug.Print "LetUtils__CheckInChanges", value
    CheckInChangesExec "vbautils.xlsm"
    LetVariableSheetValue "utils__CheckInChangesval", value
End Property




' Utils__PackupTools  ''''''''''''''''''''''''''''''
Property Get Utils__PackupTools() As String
    Debug.Print "Utils__PackupTools"
    If windowlocationval = "" Then
        Utils__PackupTools = GetVariableSheetValue("utils__PackupToolsval")
    Else
        Utils__PackupTools = utils__packuptoolsval
    End If
End Property

Property Let Utils__PackupTools(value As String)
Dim bookname As String
    Debug.Print "LetUtils__PackupTools", value
    LaunchPackupToolsExec
    LetVariableSheetValue "utils__PackupToolsval", value
End Property



' WindowLocation  ''''''''''''''''''''''''''''''
Property Get WindowLocation() As String
    Debug.Print "WindowLocation"
    If windowlocationval = "" Then
        WindowLocation = GetVariableSheetValue("windowlocationval")
    Else
        WindowLocation = windowlocationval
    End If
End Property
Property Let WindowLocation(value As String)
Dim bookname As String
    Debug.Print "LetWindowLocation", value
    LetVariableSheetValue "windowlocationval", value
End Property


' RunningApps__CA  ''''''''''''''''''''''''''''''
Property Get RunningApps__CA() As Boolean
    Debug.Print "RunningApps__CA"
    If runningapps__CAval = "" Then
        RunningApps__CA = GetVariableSheetValue("runningapps__CAval")
    Else
        RunningApps__CA = runningapps__CAval
    End If
End Property
Property Let RunningApps__CA(value As Boolean)
Dim bookname As String
    Debug.Print "LetRunningApps__CA", value
    LetVariableSheetValue "runningapps__CAval", value
End Property


' RunningApps__MO  ''''''''''''''''''''''''''''''
Property Get RunningApps__MO() As Boolean
    Debug.Print "RunningApps__MO"
    If runningapps__MOval = "" Then
        RunningApps__MO = GetVariableSheetValue("runningapps__MOval")
    Else
        RunningApps__MO = runningapps__MOval
    End If
End Property
Property Let RunningApps__MO(value As Boolean)
Dim bookname As String
    Debug.Print "LetRunningApps__MO", value
    LetVariableSheetValue "runningapps__MOval", value
End Property


' RunningApps__MM  ''''''''''''''''''''''''''''''
Property Get RunningApps__MM() As Boolean
    Debug.Print "RunningApps__MM"
    If runningapps__MMval = "" Then
        RunningApps__MM = GetVariableSheetValue("runningapps__MMval")
    Else
        RunningApps__MM = runningapps__MMval
    End If
End Property
Property Let RunningApps__MM(value As Boolean)
Dim bookname As String
    Debug.Print "LetRunningApps__MM", value
    LetVariableSheetValue "runningapps__MMval", value
End Property


' RunningApps__MV  ''''''''''''''''''''''''''''''
Property Get RunningApps__MV() As Boolean
    Debug.Print "RunningApps__MV"
    If runningapps__MVval = "" Then
        RunningApps__MV = GetVariableSheetValue("runningapps__MVval")
    Else
        RunningApps__MV = runningapps__MVval
    End If
End Property
Property Let RunningApps__MV(value As Boolean)
Dim bookname As String
    Debug.Print "LetRunningApps__MV", value
    LetVariableSheetValue "runningapps__MVval", value
End Property

' RunningApps__ES  ''''''''''''''''''''''''''''''
Property Get RunningApps__ES() As Boolean
    Debug.Print "RunningApps__ES"
    If runningapps__ESval = "" Then
        RunningApps__ES = GetVariableSheetValue("runningapps__ESval")
    Else
        RunningApps__ES = runningapps__ESval
    End If
End Property
Property Let RunningApps__ES(value As Boolean)
Dim bookname As String
    Debug.Print "LetRunningApps__ES", value
    LetVariableSheetValue "runningapps__ESval", value
End Property


' RunningApps__TA  ''''''''''''''''''''''''''''''
Property Get RunningApps__TA() As Boolean
    Debug.Print "RunningApps__TA"
    If runningapps__TAval = "" Then
        RunningApps__TA = GetVariableSheetValue("runningapps__TAval")
    Else
        RunningApps__TA = runningapps__TAval
    End If
End Property
Property Let RunningApps__TA(value As Boolean)
Dim bookname As String
    Debug.Print "LetRunningApps__TA", value
    LetVariableSheetValue "runningapps__TAval", value
End Property


' RunningApps__DV  ''''''''''''''''''''''''''''''
Property Get RunningApps__DV() As Boolean
    Debug.Print "RunningApps__DV"
    If runningapps__DVval = "" Then
        RunningApps__DV = GetVariableSheetValue("runningapps__DVval")
    Else
        RunningApps__DV = runningapps__DVval
    End If
End Property
Property Let RunningApps__DV(value As Boolean)
Dim bookname As String
    Debug.Print "LetRunningApps__DV", value
    LetVariableSheetValue "runningapps__DVval", value
End Property


' ToolAction__Open  ''''''''''''''''''''''''''''''
Property Get ToolAction__Open() As String
    Debug.Print "ToolAction__Open"
    If toolaction__showval = "" Then
        ToolAction__Open = GetVariableSheetValue("toolaction__showval")
    Else
        ToolAction__Open = toolaction__showval
    End If
End Property
Property Let ToolAction__Open(value As String)
Dim bookname As String
    Debug.Print "LetToolAction__Open", value
    bookname = Right(Me.ChooseTool, Len(Me.ChooseTool) - 12)
    ToolActionOpenExec bookname
    LetVariableSheetValue "toolaction__showval", value
End Property


' ToolAction__Close  ''''''''''''''''''''''''''''''
Property Get ToolAction__Close() As String
    Debug.Print "ToolAction__Close"
    If toolaction__closeval = "" Then
        ToolAction__Close = GetVariableSheetValue("toolaction__closeval")
    Else
        ToolAction__Close = toolaction__closeval
    End If
End Property
Property Let ToolAction__Close(value As String)
Dim bookname As String
    Debug.Print "LetToolAction__Close", value
    bookname = Right(Me.ChooseTool, Len(Me.ChooseTool) - 12)
    ToolActionCloseExec bookname
    LetVariableSheetValue "toolaction__closeval", value
End Property



' ToolAction__Min  ''''''''''''''''''''''''''''''
Property Get ToolAction__Min() As String
    Debug.Print "ToolAction__Min"
    If toolaction__minval = "" Then
        ToolAction__Min = GetVariableSheetValue("toolaction__minval")
    Else
        ToolAction__Min = toolaction__minval
    End If
End Property
Property Let ToolAction__Min(value As String)
Dim bookname As String
    Debug.Print "LetToolAction__Min", value
    bookname = Right(Me.ChooseTool, Len(Me.ChooseTool) - 12)
    MinBookExec bookname, Me.WindowSize__Width, Me.WindowSize__Height, Me.WindowSize__X, Me.WindowSize__Y
    LetVariableSheetValue "toolaction__minval", value
End Property



' ToolAction__Max  ''''''''''''''''''''''''''''''
Property Get ToolAction__Max() As String
    Debug.Print "ToolAction__Max"
    If toolaction__maxval = "" Then
        ToolAction__Max = GetVariableSheetValue("toolaction__maxval")
    Else
        ToolAction__Max = toolaction__maxval
    End If
End Property
Property Let ToolAction__Max(value As String)
Dim bookname As String
    Debug.Print "LetToolAction__Max", value
    bookname = Right(Me.ChooseTool, Len(Me.ChooseTool) - 12)
    MaxBookExec bookname, Me.WindowSize__Width, Me.WindowSize__Height
    LetVariableSheetValue "toolaction__maxval", value
End Property







' ToolAction__Show  ''''''''''''''''''''''''''''''
Property Get ToolAction__Show() As String
    Debug.Print "ToolAction__Show"
    If toolaction__showval = "" Then
        ToolAction__Show = GetVariableSheetValue("toolaction__showval")
    Else
        ToolAction__Show = toolaction__showval
    End If
End Property
Property Let ToolAction__Show(value As String)
Dim bookname As String
    Debug.Print "LetToolAction__Show", value
    bookname = Right(Me.ChooseTool, Len(Me.ChooseTool) - 12)
    ShowBookExec bookname, Me.WindowSize__Width, Me.WindowSize__Height, Me.WindowSize__X, Me.WindowSize__Y
    LetVariableSheetValue "toolaction__showval", value
End Property

' ToolAction__Hide  ''''''''''''''''''''''''''''''
Property Get ToolAction__Hide() As String
    Debug.Print "ToolAction__Hide"
    If toolaction__hideval = "" Then
        ToolAction__Hide = GetVariableSheetValue("toolaction__hideval")
    Else
        ToolAction__Hide = toolaction__hideval
    End If
End Property
Property Let ToolAction__Hide(value As String)
Dim bookname As String
    Debug.Print "LetToolAction__Hide", value
    bookname = Right(Me.ChooseTool, Len(Me.ChooseTool) - 12)
    HideBookExec bookname
    LetVariableSheetValue "toolaction__hideval", value
End Property



' ChooseTool  ''''''''''''''''''''''''''''''
Property Get ChooseTool() As String
    Debug.Print "Settings__rootpath"
    If choosetoolval = "" Then
        ChooseTool = GetVariableSheetValue("choosetoolval")
    Else
        ChooseTool = choosetoolval
    End If
End Property
Property Let ChooseTool(value As String)
    Debug.Print "LetChooseTool", value
    LetVariableSheetValue "choosetoolval", value
End Property

' Settings__rootpath  ''''''''''''''''''''''''''''''
Property Get Settings__rootpath() As String
    Debug.Print "Settings__rootpath"
    If settings__rootpathval = "" Then
        Settings__rootpath = GetVariableSheetValue("settings__rootpathval")
    Else
        Settings__rootpath = settings__rootpathval
    End If
End Property
Property Let Settings__rootpath(value As String)
    Debug.Print "LetSettings__rootpath", value
    LetVariableSheetValue "settings__rootpathval", value
End Property


' Settings__dataurl  ''''''''''''''''''''''''''''''
Property Get Settings__dataurl() As String
    Debug.Print "Settings__dataurl"
    If settings__dataurlval = "" Then
        Settings__dataurl = GetVariableSheetValue("settings__dataurlval")
    Else
        Settings__dataurl = settings__dataurlval
    End If
End Property
Property Let Settings__dataurl(value As String)
    Debug.Print "LetSettings__dataurl", value
    LetVariableSheetValue "settings__dataurlval", value
End Property



' WindowSize__X  ''''''''''''''''''''''''''''''
Property Get WindowSize__X() As String
    Debug.Print "WindowSize__X"
    If windowsize__xval = "" Then
        WindowSize__X = GetVariableSheetValue("windowsize__xval")
    Else
        WindowSize__X = windowsize__xval
    End If
End Property
Property Let WindowSize__X(value As String)
    Debug.Print "LetWindowSize__X", value
    LetVariableSheetValue "windowsize__xval", value
End Property


' WindowSize__Y  ''''''''''''''''''''''''''''''
Property Get WindowSize__Y() As String
    Debug.Print "Urls__url1"
    If windowsize__yval = "" Then
        WindowSize__Y = GetVariableSheetValue("windowsize__yval")
    Else
        WindowSize__Y = windowsize__yval
    End If
End Property
Property Let WindowSize__Y(value As String)
    Debug.Print "LetWindowSize__Y", value
    LetVariableSheetValue "windowsize__yval", value
End Property


' WindowSize__width  ''''''''''''''''''''''''''''''
Property Get WindowSize__Width() As String
    Debug.Print "WindowSize__Width"
    If windowsize__widthval = "" Then
        WindowSize__Width = GetVariableSheetValue("windowsize__widthval")
    Else
        WindowSize__Width = windowsize__widthval
    End If
End Property
Property Let WindowSize__Width(value As String)
    Debug.Print "LetWindowSize__Width", value
    LetVariableSheetValue "windowsize__widthval", value
End Property

' height  ''''''''''''''''''''''''''''''
Property Get WindowSize__Height() As String
    Debug.Print "WindowSize__Height"
    If windowsize__heightval = "" Then
        WindowSize__Height = GetVariableSheetValue("windowsize__heightval")
    Else
        WindowSize__Height = windowsize__heightval
    End If
End Property

Property Let WindowSize__Height(value As String)
    Debug.Print "LetWindowSize__Height", value
    LetVariableSheetValue "windowsize__heightval", value
End Property





' WindowSizeRight__X  ''''''''''''''''''''''''''''''
Property Get WindowSizeRight__X() As String
    Debug.Print "WindowSizeRight__X"
    If windowSizeRight__xval = "" Then
        WindowSizeRight__X = GetVariableSheetValue("windowSizeRight__xval")
    Else
        WindowSizeRight__X = windowSizeRight__xval
    End If
End Property
Property Let WindowSizeRight__X(value As String)
    Debug.Print "LetWindowSizeRight__X", value
    LetVariableSheetValue "windowSizeRight__xval", value
End Property


' WindowSizeRight__Y  ''''''''''''''''''''''''''''''
Property Get WindowSizeRight__Y() As String
    Debug.Print "Urls__url1"
    If windowSizeRight__yval = "" Then
        WindowSizeRight__Y = GetVariableSheetValue("windowSizeRight__yval")
    Else
        WindowSizeRight__Y = windowSizeRight__yval
    End If
End Property
Property Let WindowSizeRight__Y(value As String)
    Debug.Print "LetWindowSizeRight__Y", value
    LetVariableSheetValue "windowSizeRight__yval", value
End Property


' WindowSizeRight__width  ''''''''''''''''''''''''''''''
Property Get WindowSizeRight__Width() As String
    Debug.Print "WindowSizeRight__Width"
    If windowSizeRight__widthval = "" Then
        WindowSizeRight__Width = GetVariableSheetValue("windowSizeRight__widthval")
    Else
        WindowSizeRight__Width = windowSizeRight__widthval
    End If
End Property
Property Let WindowSizeRight__Width(value As String)
    Debug.Print "LetWindowSizeRight__Width", value
    LetVariableSheetValue "windowSizeRight__widthval", value
End Property

' height  ''''''''''''''''''''''''''''''
Property Get WindowSizeRight__Height() As String
    Debug.Print "WindowSizeRight__Height"
    If windowSizeRight__heightval = "" Then
        WindowSizeRight__Height = GetVariableSheetValue("windowSizeRight__heightval")
    Else
        WindowSizeRight__Height = windowSizeRight__heightval
    End If
End Property

Property Let WindowSizeRight__Height(value As String)
    Debug.Print "LetWindowSizeRight__Height", value
    LetVariableSheetValue "windowSizeRight__heightval", value
End Property



' WindowSizeLeft__X  ''''''''''''''''''''''''''''''
Property Get WindowSizeLeft__X() As String
    Debug.Print "WindowSizeLeft__X"
    If windowSizeLeft__xval = "" Then
        WindowSizeLeft__X = GetVariableSheetValue("windowSizeLeft__xval")
    Else
        WindowSizeLeft__X = windowSizeLeft__xval
    End If
End Property
Property Let WindowSizeLeft__X(value As String)
    Debug.Print "LetWindowSizeLeft__X", value
    LetVariableSheetValue "windowSizeLeft__xval", value
End Property


' WindowSizeLeft__Y  ''''''''''''''''''''''''''''''
Property Get WindowSizeLeft__Y() As String
    Debug.Print "Urls__url1"
    If windowSizeLeft__yval = "" Then
        WindowSizeLeft__Y = GetVariableSheetValue("windowSizeLeft__yval")
    Else
        WindowSizeLeft__Y = windowSizeLeft__yval
    End If
End Property
Property Let WindowSizeLeft__Y(value As String)
    Debug.Print "LetWindowSizeLeft__Y", value
    LetVariableSheetValue "windowSizeLeft__yval", value
End Property


' WindowSizeLeft__width  ''''''''''''''''''''''''''''''
Property Get WindowSizeLeft__Width() As String
    Debug.Print "WindowSizeLeft__Width"
    If windowSizeLeft__widthval = "" Then
        WindowSizeLeft__Width = GetVariableSheetValue("windowSizeLeft__widthval")
    Else
        WindowSizeLeft__Width = windowSizeLeft__widthval
    End If
End Property
Property Let WindowSizeLeft__Width(value As String)
    Debug.Print "LetWindowSizeLeft__Width", value
    LetVariableSheetValue "windowSizeLeft__widthval", value
End Property

' height  ''''''''''''''''''''''''''''''
Property Get WindowSizeLeft__Height() As String
    Debug.Print "WindowSizeLeft__Height"
    If windowSizeLeft__heightval = "" Then
        WindowSizeLeft__Height = GetVariableSheetValue("windowSizeLeft__heightval")
    Else
        WindowSizeLeft__Height = windowSizeLeft__heightval
    End If
End Property

Property Let WindowSizeLeft__Height(value As String)
    Debug.Print "LetWindowSizeLeft__Height", value
    LetVariableSheetValue "windowSizeLeft__heightval", value
End Property


' ur1  ''''''''''''''''''''''''''''''
Property Get Urls__url1() As String
    Debug.Print "Urls__url1"
    If urls__url1val = "" Then
        Urls__url1 = GetVariableSheetValue("urls__url1val")
    Else
        Urls__url1 = urls__url1val
    End If
End Property
Property Let Urls__url1(value As String)
    Debug.Print "LetUrls__url1", value
    LetVariableSheetValue "urls__url1val", value
End Property


' docs1  ''''''''''''''''''''''''''''''
Property Get LaunchDocs() As String
    Debug.Print "Docs__docs1"
    If docs__doc1val = "" Then
        LaunchDocs = GetVariableSheetValue("docs__doc1val")
    Else
        LaunchDocs = docs__doc1val
    End If
End Property
Property Let LaunchDocs(value As String)
    Debug.Print "LetLaunchDocs", value
    LaunchBrowser value, Me.WindowSize__X, Me.WindowSize__Y, Me.WindowSize__Width, Me.WindowSize__Height
    LetVariableSheetValue "docs__doc1val", value
End Property


Sub LetVariableSheetValue(varname As String, value As Variant, Optional resultsDict As Dictionary = Nothing)
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
    bookname = "VBAUtils.xlsm"
    persistsheetname = "Persist"
    persistfilename = Environ("USERPROFILE") & "\Deploy\.VBAUtils_persist.csv"
    persistrangename = "persistdata"
    persistrangelen = Workbooks(bookname).Sheets(persistsheetname).Range(persistrangename).Rows.count

End Sub
