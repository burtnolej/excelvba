Attribute VB_Name = "RibbonUtils"
Option Explicit

Dim rbxUI As IRibbonUI
Dim SpinValue As Long
Dim manifestFiles() As Variant
Dim folderListDict As Dictionary, appListDict As Dictionary, urlListDict As Dictionary, checkboxVals As Dictionary
Dim x As Long
Dim y As Long
Dim height As Long
Dim width As Long
Dim rootpath As String
Dim dataurl As String

Declare PtrSafe Function GetSystemMetrics32 Lib "USER32" _
    Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    destination As Any, _
    source As Any, _
    ByVal length As Long)
    
Sub GetScreenRes(ByRef w As Long, ByRef h As Long)
    w = GetSystemMetrics32(0) ' width in points
    h = GetSystemMetrics32(1) ' height in points
End Sub


'https://stackoverflow.com/questions/73586725/dynamically-change-ribbon-image


Sub InitCheckBoxVals()
Dim runningAppsRange As Range
Dim i As Integer

    Set runningAppsRange = ActiveWorkbook.Sheets("Reference").Range("RUNNINGAPPS")
    For i = 1 To runningAppsRange.Rows.count()
        'runningAppsRange.Cells(i, 1).value = False
        runningAppsRange.Cells(i, 2).value = False
    Next i
exitsub:
    Set runningAppsRange = Nothing
End Sub
Sub PersistCheckBoxVals(checkboxVals As Dictionary)
Dim runningAppsRange As Range
Dim key As Variant
Dim count As Integer

    Set runningAppsRange = Workbooks("vbautils.xlsm").Sheets("Reference").Range("RUNNINGAPPS")

    count = 1
    For Each key In checkboxVals.Keys()
        runningAppsRange.Cells(count, 1).value = key
        runningAppsRange.Cells(count, 2).value = checkboxVals.Item(key)
        count = count + 1
    Next key
        
exitsub:
    Set runningAppsRange = Nothing
    
End Sub

Function RetrieveCheckBoxVals() As Dictionary
Dim runningAppsRange As Range
Dim tmpDict As Dictionary
Dim i As Integer

    Set runningAppsRange = Workbooks("vbautils.xlsm").Sheets("Reference").Range("RUNNINGAPPS")

    Set tmpDict = New Dictionary
    
    For i = 1 To runningAppsRange.Rows.count()
        If tmpDict.Exists(runningAppsRange.Cells(i, 1).value) = False Then
            tmpDict.Add runningAppsRange.Cells(i, 1).value, runningAppsRange.Cells(i, 2).value
        End If
    Next i
    
    Set RetrieveCheckBoxVals = tmpDict
exitsub:

    Set tmpDict = Nothing
    Set runningAppsRange = Nothing
End Function
Private Property Get CheckboxValues() As Dictionary

    If checkboxVals Is Nothing Then
        'Set CheckboxValues = New Dictionary
        Set CheckboxValues = RetrieveCheckBoxVals()
    Else
        Set CheckboxValues = checkboxVals
    End If
End Property

Function RetrieveRootPathVals() As String
Dim persistedVars() As Variant
Dim varValues As Dictionary

    persistedVars = Array("rootpath")
    Set varValues = RetreiveVars(persistedVars)
    RetrieveRootPathVals = varValues("rootpath")
    
End Function

Function RetrieveDataurlVals() As String
Dim persistedVars() As Variant
Dim varValues As Dictionary

    persistedVars = Array("dataurl")
    Set varValues = RetreiveVars(persistedVars)
    RetrieveDataurlVals = varValues("dataurl")
    
End Function

Private Property Get RootPathValue() As String

    If rootpath = "" Then
        'Set CheckboxValues = New Dictionary
        RootPathValue = RetrieveRootPathVals()
    Else
        RootPathValue = rootpath
    End If
End Property

Private Property Get DataURLValue() As String
    If dataurl = "" Then
        DataURLValue = RetrieveDataurlVals()
    Else
        DataURLValue = dataurl
    End If
End Property

Private Property Get CustomRibbon() As IRibbonUI
Dim persistedVars() As Variant
Dim varValues As Dictionary
Dim aPtr As LongPtr
Dim ribUI As Object

    'On Error GoTo EH
    
    If Not rbxUI Is Nothing Then
        Set CustomRibbon = rbxUI
        Exit Function
    End If
    
    persistedVars = Array("ribbonui")
    Set varValues = RetreiveVars(persistedVars)
    MsgBox varValues("ribbonui")
    aPtr = varValues("ribbonui")
    
    'aPtr = wsSettings.Range("A1").Value2
    CopyMemory ribUI, aPtr, LenB(aPtr)
    Set rbxUI = ribUI
    Set ribUI = Nothing
    
    Set CustomRibbon = rbxUI
    Exit Function
EH:
End Property

Public Sub RefreshDownloadFiles(Optional param As Variant)
Dim outputRange As Range
Dim url As String
Dim colArray() As Variant

    'url = "http://172.23.208.38/datafiles/"
    url = "http://172.22.237.138/datafiles/"
    
    Application.Run "vbautils.xlsm!SetEventsOff"
    
    On Error Resume Next
    Application.StatusBar = "loading http://172.22.237.138/datafiles/manifest.csv"
    Set outputRange = HTTPDownloadFile(url + "manifest.csv", _
                ActiveWorkbook, _
                "", "", 1, "start-of-day", "MANIFEST", False, 0)
                
    Application.Run "vbautils.xlsm!SortRange", outputRange.Worksheet, outputRange, 2
    
    
    colArray = Array(1, 2)
    
    'Application.Run "vbautils.xlsm!CreateCalcNamedRange", outputRange.Worksheet, outputRange, "fullFileName", colArray
    
    'Set outputRange = outputRange.Resize(, outputRange.Columns.Count + 1)
    
    Application.Run "vbautils.xlsm!AddNamedRange", outputRange.Worksheet, outputRange, 1, "FILENAME"
    
    manifestFiles = outputRange.Offset(1).Resize(outputRange.Rows.count - 1)
    
    'manifestFiles = Application.Run("vbautils.xlsm!RangeToArray", ActiveWorkbook, "FILE_ALLDATA", "FILE_FULLNAME", manifestFiles)

End Sub
 
 
Sub PersistVar(key As String, value As Variant)
Dim tmpWorksheet As Worksheet

    Set tmpWorksheet = ActiveWorkbook.Sheets("Reference")
    tmpWorksheet.Range(key) = value
End Sub
Function RetreiveVars(persistedVars() As Variant) As Dictionary
Dim tmpWorksheet As Worksheet
Dim tmpDict As New Dictionary
Dim i As Variant

    Set tmpWorksheet = Workbooks("vbautils.xlsm").Sheets("Reference")
    For i = 0 To UBound(persistedVars)
        tmpDict.Add persistedVars(i), tmpWorksheet.Range(persistedVars(i)).value
    Next i
    Set RetreiveVars = tmpDict
    
exitsub:
    Set tmpDict = Nothing
    Set tmpWorksheet = Nothing
End Function

Sub rbx_onLoad(ribbon As IRibbonUI)
Dim persistedVars() As Variant
Dim varValues As Dictionary
    
    Set rbxUI = ribbon

    PersistVar "ribbonui", ObjPtr(ribbon)
    
    On Error Resume Next
    CommandBars("Document Recovery").Visible = False
    On Error GoTo 0
    'CommandBars("Document Recovery").Visible = False
    
    persistedVars = Array("x", "y", "width", "height", "onedrive", "ribbonui")
    Set varValues = RetreiveVars(persistedVars)
    

    x = varValues("x")
    y = varValues("y")
    width = varValues("width")
    height = varValues("height")
    PersistVar "onedrive", Environ("OneDrive")
    rootpath = Environ("rootpath")
    
    rbxUI.ActivateTab "tab3"
    
    Set folderListDict = New Dictionary
    Set appListDict = New Dictionary
    Set urlListDict = New Dictionary
    Set checkboxVals = New Dictionary
    
    RangeToDict ActiveWorkbook, "Reference", "FOLDERS", folderListDict
    RangeToDict ActiveWorkbook, "Reference", "APPS", appListDict
    RangeToDict ActiveWorkbook, "Reference", "URLS", urlListDict
    
    'SetXLOnTop
    
    InitCheckBoxVals
    
End Sub


' Set default value of editBox to 0

Sub editBox_onChange(control As IRibbonControl, Text As Variant)

    Select Case control.tag
    
        Case "x"
            x = Text
        Case "y"
            y = Text
        Case "height"
            height = Text
        Case "width"
            width = Text
        Case "rootpath"
            rootpath = Text
    End Select

    PersistVar control.tag, Text
    
End Sub

' Return value of editBox

Sub editBox_getText(control As IRibbonControl, ByRef returnedVal)
Dim persistedVars() As Variant
Dim varValues As Dictionary

    If height = 0 Then
        persistedVars = Array("x", "y", "width", "height")
        Set varValues = RetreiveVars(persistedVars)
        x = varValues("x")
        y = varValues("y")
        width = varValues("width")
        height = varValues("height")
    End If
    
    Select Case control.tag
    
        Case "x"
             returnedVal = x
        Case "y"
            returnedVal = y
        Case "height"
            returnedVal = height
        Case "width"
            returnedVal = width
        Case "rootpath"
            'returnedVal = rootpath
            returnedVal = RootPathValue
    End Select
    
   
End Sub



Sub chkBox_onAction(control As IRibbonControl, isPressed As Boolean)

    'Display status of checkbox in cell: TRUE or FALSE
    
    Sheet1.Range("H9").value = isPressed
    
End Sub

Public Sub fncGetPressed(control As IRibbonControl, ByRef bolReturn)
'Callback Checkbox State
'Select Case control.id
'Case "MA"
'Here do you change the condition of bolReturn conforms to your form
If CheckboxValues.Exists(control.id) Then
    If CheckboxValues.Item(control.id) = True Then
        bolReturn = True
    Else
        bolReturn = False
    End If
End If
'End Select
End Sub

Sub btns_onAction(control As IRibbonControl)
Dim tag As String, action As String, param As String, foldername As String, bookname As String
Dim tagSplit As Variant, functionSplit As Variant
Dim w As Long, h As Long
Dim persistedVars() As Variant
Dim varValues As Dictionary
Dim args() As Variant


    foldername = RootPathValue
    
    'foldername = "E:\Velox Financial Technology\Velox Shared Drive - Documents\General\Tools"

    tag = control.tag
    
    tagSplit = Split(tag, "_")
    action = tagSplit(0)
    If UBound(tagSplit) > 0 Then
        param = tagSplit(1)
    Else
        param = ""
    End If

    CustomRibbon.Invalidate

    Select Case action
    
        Case "openbook"
            SetEventsOff
            bookname = OpenWorkbook(foldername & "\" & param & ".xlsm")
            ResizeWindow bookname, width, height
            MoveWindow bookname, x, y
            SetEventsOn
            If CheckboxValues.Exists(param) = False Then
                CheckboxValues.Add param, True
            Else
                CheckboxValues.Item(param) = True
            End If
            PersistCheckBoxVals CheckboxValues
            
            'Workbooks(bookname).Sheets("REFERENCE").Activate
            'Set tmpWB = Workbooks(bookname).Sheets("REFERENCE")
            'tmpWB.Activate
            Workbooks(bookname).Sheets("REFERENCE").Range("dataurl").value = DataURLValue
            
        Case "closebook"
            CloseWorkbook param
            If CheckboxValues.Exists(param) Then
                CheckboxValues.Remove param
                PersistCheckBoxVals CheckboxValues
            End If
            
        
        Case "maxbook"
            GetScreenRes w, h
            ResizeWindow param, w, h
            MoveWindow param, 0, 0
        
        Case "minbook"
            ResizeWindow param, width, height
            MoveWindow param, x, y
            
            
        Case "runfunction"
            functionSplit = Split(param, "^")
            functionSplit = Application.Run(functionSplit(0) & ".xlsm!" & functionSplit(1), functionSplit(2))
            
        Case "pickfolder"
            rootpath = Application.Run("VBAUtils.xlsm!GetFolderSelection", Environ("OneDrive"))
            CustomRibbon.InvalidateControl "editBox5"
            'rbxUI.InvalidateControl "editBox5"
            
        Case "showtools"
            ShowDisplayables param
            ShowMenuBar param
            ShowSheets param
            ShowFormulaBar param
        
        Case "hidetools"
            HideDisplayables param
            HideMenuBar param
            HideSheets param
            HideFormulaBar param

        Case "hidebook"
            HideBook param
            

        Case "runapp"
            functionSplit = Split(param, "^")
            Set appListDict = New Dictionary
            
            RangeToDict ActiveWorkbook, "Reference", "APPS", appListDict

            args = appListDict(functionSplit(0))
            LaunchApp CStr(args(1)), CStr(args(2))
            
         Case "runurl"
            functionSplit = Split(param, "^")
            Set urlListDict = New Dictionary
            
            RangeToDict ActiveWorkbook, "Reference", "URLS", urlListDict

            args = urlListDict(functionSplit(0))
            LaunchBrowser CStr(args(1)), x, y, width, height
            

        Case "killapp"
            functionSplit = Split(param, "^")
            Set appListDict = New Dictionary
            
            RangeToDict ActiveWorkbook, "Reference", "APPS", appListDict

            args = appListDict(functionSplit(0))
            KillApp CStr(args(3))
            
        Case "showbook"
            ShowBook param
            ResizeWindow param, width, height
            MoveWindow param, x, y
        
        
        
    End Select
    
    CustomRibbon.Invalidate
End Sub

Sub btnGrp_onAction(control As IRibbonControl)
    Select Case control.id

        'Buttons 1-9
        
        Case "btnGrp_btn1", "btnGrp4_btn1"
            Sheet1.Activate

    End Select
End Sub

Sub togBtn_onAction(control As IRibbonControl, isPressed As Boolean)
    
    Select Case control.id
        
        Case "togBtn_btn1"
            
    End Select
    
End Sub

Sub dropDown_onAction(control As IRibbonControl, id As String, index As Integer)

    Select Case id
        
        Case "itm1"
            Sheet1.Activate
        Case "itm2"
            Sheet2.Activate
        Case "itm3"
            Sheet3.Activate
    End Select
    
End Sub
Sub splitBtn_onAction(control As IRibbonControl)
    Select Case control.id
        
        Case "splitBtn_btn1"
            MsgBox "This is a button!"

            
    End Select
End Sub


'https://medium.com/codex/how-to-build-a-custom-ribbon-in-excel-a3bc531551e1
'https://github.com/fernandreu/office-ribbonx-editor/releases/tag/v1.9.0

Sub LaunchCA(bookname As String)
    'Environ ("VELOXTOOLS")
    'OpenWorkbook
    
    Debug.Print bookname
End Sub



'Callback for customUI.onLoad
Sub Initialize(ribbon As IRibbonUI)
    Set rbxUI = ribbon
End Sub

'Callback for Combo3 getItemCount (called once when the combobox is invalidated)
Sub Combo3_getItemCount(control As IRibbonControl, ByRef returnedVal)
    'returnedVal = 10 'the number of items for combobox
    If (Not Not manifestFiles) = 0 Then
        RefreshDownloadFiles
    End If
    
    RefreshDownloadFiles
    returnedVal = UBound(manifestFiles)
End Sub

'Callback for Combo3 getItemID (called 10 times when combobox is invalidated)
Public Sub Combo3_getItemID(control As IRibbonControl, index As Integer, ByRef id)
    'id = "ComboboxItem" & index + 1
    id = manifestFiles(index + 1, 1)
End Sub

'Callback for Combo3 getItemLabel (called 10 times when combobox is invalidated)
Sub Combo3_getItemLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)
    'returnedVal = "Item" & index + 1
    returnedVal = manifestFiles(index + 1, 1)
End Sub

'Callback for Combo3 getText
Sub Combo3_getText(control As IRibbonControl, ByRef returnedVal)
    returnedVal = "" 'clears the text from the combobox
End Sub

'Callback for Combo3 onChange
Sub Combo3_onChange(control As IRibbonControl, filename As String)
Dim textSplit As Variant
Dim sheetname As String, url As String
Dim outputRange As Range

    sheetname = UCase(Split(filename, ".")(0))
    
    url = "http://172.22.237.138/datafiles/"

    Application.Run "vbautils.xlsm!SetEventsOff"
    
    Set outputRange = Application.Run("vbautils.xlsm!HTTPDownloadFile", url + filename, _
                ActiveWorkbook, _
                "", "", 1, "start-of-day", sheetname, False, 0)
    Application.Run "vbautils.xlsm!SetEventsOn"
End Sub

Sub UpdateCombo3()
    myRibbon.InvalidateControl "Combo3" 'invalidates the cache for the combobox
End Sub



Sub LoadCustRibbon()

Dim hFile As Long
Dim path As String, filename As String, ribbonXML As String, user As String

hFile = FreeFile
user = Environ("Username")
path = "C:\Users\" & user & "\AppData\Local\Microsoft\Office\"
filename = "Excel.officeUI"

ribbonXML = "<mso:customUI      xmlns:mso='http://schemas.microsoft.com/office/2009/07/customui'>" & vbNewLine
ribbonXML = ribbonXML + "  <mso:ribbon>" & vbNewLine
ribbonXML = ribbonXML + "    <mso:qat/>" & vbNewLine
ribbonXML = ribbonXML + "    <mso:tabs>" & vbNewLine
ribbonXML = ribbonXML + "      <mso:tab id='reportTab' label='Reports' insertBeforeQ='mso:TabFormat'>" & vbNewLine
ribbonXML = ribbonXML + "        <mso:group id='reportGroup' label='Reports' autoScale='true'>" & vbNewLine
ribbonXML = ribbonXML + "          <mso:button id='runReport' label='Desktop' " & vbNewLine
ribbonXML = ribbonXML + "imageMso='AppointmentColor3'      onAction='E:\Velox Financial Technology\Velox Shared Drive - Documents\General\Tools\VBAUtils!OpenDesktop'/>" & vbNewLine
ribbonXML = ribbonXML + "        </mso:group>" & vbNewLine
ribbonXML = ribbonXML + "      </mso:tab>" & vbNewLine
ribbonXML = ribbonXML + "    </mso:tabs>" & vbNewLine
ribbonXML = ribbonXML + "  </mso:ribbon>" & vbNewLine
ribbonXML = ribbonXML + "</mso:customUI>"

ribbonXML = Replace(ribbonXML, """", "")

Open path & filename For Output Access Write As hFile
Print #hFile, ribbonXML
Close hFile

End Sub

Sub ClearCustRibbon()

Dim hFile As Long
Dim path As String, filename As String, ribbonXML As String, user As String

hFile = FreeFile
user = Environ("Username")
path = "C:\Users\" & user & "\AppData\Local\Microsoft\Office\"
filename = "Excel.officeUI"

ribbonXML = "<mso:customUI           xmlns:mso=""http://schemas.microsoft.com/office/2009/07/customui"">" & _
"<mso:ribbon></mso:ribbon></mso:customUI>"

Open path & filename For Output Access Write As hFile
Print #hFile, ribbonXML
Close hFile

End Sub

