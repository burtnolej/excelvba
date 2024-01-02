Attribute VB_Name = "DesktopUtils"
'Sub ShowXLOnTop(ByVal OnTop As Boolean)
'Sub SetXLOnTopExec(Optional param As Variant)
'Sub SetXLNormalExec(Optional param As Variant)
'Public Sub LaunchExplorer(foldername As String, Optional param As String)
'Public Sub LaunchGitBash(Optional startdir As String = "C:\Users\burtn\Development")
'Public Sub LaunchPackupTools(Optional startdir As String = "C:\Users\burtn\Development")
'Function IsCapitalized(char As String) As Boolean
'Sub CreateCustomNamedRanges(headerRange As Range, targetRange As Range, targetSheet As Worksheet)
'Public Sub LaunchBrowser(urlname As String, x As Long, y As Long, width As Long, height As Long)
'Public Sub GetDataFile(filename)
'Sub RunPython()
'Sub RunPowershell(arg1 As String)
'Function GetMondayFolders(workingdir As String) As Variant
'Public Sub LaunchApp(appname As String, param As String)
'Public Sub ShowHideables()
'Public Sub HideHideables()
'Public Sub HideFormulaBar(bookname As String)
'Public Sub ShowFormulaBar(bookname As String)
'Public Sub HideDisplayables(bookname As String)
'Sub HideSheets(bookname As String, Optional visibleSheet As String = "BLANK")
'Public Sub ShowSheets(Optional bookname As String = "")
'Public Sub ShowDisplayables(bookname As String)
'Public Sub HideMenuBar(bookname As String)
'Public Sub ShowMenuBar(bookname As String)
'Public Sub CloseWorkbook(bookname As String)
'Public Function OpenWorkbook(bookFullPath As String) As String
'Public Sub DisplayCommandBars()
'Public Sub DisplayWindow(bookname As String)
'Public Sub ZoomWindow(bookname As String, zoom As Double)
'Public Sub HideWindow(bookname As String)
'Public Sub HideBook(bookname As String)
'Public Sub ShowBook(bookname As String)
'Public Sub ResizeWindow(bookname As String, Optional width As Long = 1000, Optional height As Long = 1000)
'Public Sub MoveWindow(bookname As String, Optional top As Long = 0, Optional left As Long = 0)
'Public Function GetWorkbooks() As Variant

'Sub KillApp(sTaskName As String)
'Sub ToolActionOpenExec(appname As String)
'Sub ToolActionCloseExec(appname As String)
'Public Sub MaxBookExec(param As String, width As Long, height As Long)
'Public Sub MinBookExec(param As String, width As Long, height As Long, x As Long, y As Long)
'Public Sub HideBookExec(param As String)
'Public Sub ShowBookExec(param As String, width As Long, height As Long, x As Long, y As Long)
'Public Sub ShowToolsExec(param As String)
'Public Sub HideToolsExec(param As String)
'Sub DisplayVBEExec(Optional param As Variant)
'Sub RunRibbonEditorExec(bookname As String)
'Sub CloseRibbonEditorExec()
Option Explicit

Public Declare PtrSafe Function SetWindowPos _
    Lib "USER32" ( _
        ByVal hWnd As LongPtr, _
        ByVal hwndInsertAfter As LongPtr, _
        ByVal x As Long, ByVal y As Long, _
        ByVal cx As Long, ByVal cy As Long, _
        ByVal wFlags As Long) _
As Long

Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2


Function KillApp(sTaskName As String) As Variant
Dim result As Variant
    result = CreateObject("WScript.Shell").Run("taskkill /f /im " & sTaskName, 0, True)
End Function

Sub ShowXLOnTop(ByVal OnTop As Boolean)
    Dim xStype As Long
    #If Win64 Then
        Dim xHwnd As LongPtr
    #Else
        Dim xHwnd As Long
    #End If
    If OnTop Then
        xStype = HWND_TOPMOST
    Else
        xStype = HWND_NOTOPMOST
    End If
    Call SetWindowPos(Application.hWnd, xStype, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
End Sub
Sub SetXLOnTopExec(Optional param As Variant)
    ShowXLOnTop True
End Sub
Sub SetXLNormalExec(Optional param As Variant)
    ShowXLOnTop False
End Sub

Sub ToolActionOpenExec(appname As String)
Dim bookname As String, controlid As String
Dim RV As RibbonVariables
Dim rootpath As String, dataurl As String
Dim width As Long, height As Long, x As Long, y As Long

    Set RV = New RibbonVariables
    
    SetEventsOff
    controlid = "runningapps__" & UCase(appname)
    
    CallByName RV, controlid, VbLet, True
    'RV.RibbonPointer.InvalidateControl controlid
    rootpath = RV.Settings__rootpath
    width = RV.WindowSize__Width
    height = RV.WindowSize__Height
    x = RV.WindowSize__X
    y = RV.WindowSize__Y
    dataurl = RV.Settings__dataurl
    Set RV = Nothing
    
    Application.Wait Now + #12:00:01 AM#

    bookname = OpenWorkbook(rootpath & "\" & appname & ".xlsm")
    ResizeWindowExec bookname, width, height
    MoveWindow bookname, x, y
    SetEventsOn
    
    On Error Resume Next ' moving over the location of these values to the Persist sheet
    Workbooks(bookname).Sheets("REFERENCE").Range("dataurl").value = dataurl
    Workbooks(bookname).Sheets("Persist").Range("dataurl").value = dataurl
    On Error GoTo 0
    
    'Workbooks("VBAUtils.xlsm").Activate
    'Set RV = New RibbonVariables
    'CallByName RV, controlid, VbLet, True
    'RV.RibbonPointer.InvalidateControl controlid
    'Set RV = Nothing
End Sub

Sub ToolActionCloseExec(appname As String)
Dim bookname As String, controlid As String
Dim RV As RibbonVariables
    SetEventsOff
    controlid = "runningapps__" & UCase(appname)
    
    

    
    CloseWorkbook appname & ".xlsm"
    
    'Workbooks("VBAUtils.xlsm").Activate
    'Set RV = New RibbonVariables
    'CallByName RV, controlid, VbLet, False
    'RV.RibbonPointer.InvalidateControl controlid
    'Set RV = Nothing
    

    SetEventsOn
End Sub
Public Sub LaunchExplorer(foldername As String, Optional param As String)
    Shell "C:\WINDOWS\explorer.exe """ & foldername & "", vbNormalFocus
End Sub

Public Sub LaunchGitBash(Optional startdir As String = "C:\Users\burtn\Development")

Dim execStr As String
Dim objShell As Object
Dim psexepath As String, execPath As String

    psexepath = "POWERSHELL.exe -noexit"
    execPath = """C:\Users\burtn\Development\ps\Launch-GitBash.ps1"""
    Set objShell = VBA.CreateObject("Wscript.Shell")
    execStr = psexepath & " " & execPath & " " & startdir
    objShell.Run execStr, vbHide
    
End Sub

Public Sub LaunchPackupToolsExec(Optional startdir As String = "C:\Users\burtn\Development")
Dim execStr As String, psexepath As String, execPath As String
Dim objShell As Object
    
    psexepath = "POWERSHELL.exe"
    execPath = """C:\Users\burtn\Development\ps\Packup-Tools.ps1"""
    Set objShell = VBA.CreateObject("Wscript.Shell")

    ChDir "C:\Users\burtn\Development"
    execStr = psexepath & " " & execPath
    objShell.Run execStr
    
End Sub


Function IsCapitalized(char As String) As Boolean
Dim theChar As String
    theChar = left(char, 1)
    If theChar Like "*[A-Z]*" Then IsCapitalized = True Else IsCapitalized = False
End Function

Sub CreateCustomNamedRanges(headerRange As Range, targetRange As Range, targetSheet As Worksheet)
Dim tmpCell As Range, tmpCol As Range
Dim tmpColHeight As Long
Dim tmpColName As String, tmpSheetName As String

    tmpSheetName = targetSheet.Name
    For Each tmpCell In headerRange.Cells
        tmpColName = tmpSheetName & "_" & UCase(tmpCell.value)
        tmpColName = Replace(tmpColName, " ", "_")
        If IsCapitalized(tmpCell.value) = True Then
            Set tmpCol = targetRange.Columns(tmpCell.Column)
            tmpColHeight = tmpCol.Rows.count
            Set tmpCol = tmpCol.Offset(1).Resize(tmpColHeight - 1)
            targetSheet.Names.Add tmpColName, tmpCol
        End If
    Next tmpCell
End Sub
Public Sub GetDataFile(filename)
Dim tmpSheet As Worksheet
Dim outputRange As Range
    sheetname = UCase(Split(filename, ".")(0))
    
    If envUrl = "" Then
        url = "http://172.22.237.138/datafiles/"
    Else
        url = envUrl
    End If
    

    Application.Run "DV.xlsm!SetEventsOff"
    
    Set outputRange = Application.Run("DV.xlsm!HTTPDownloadFile", url + filename, _
                ActiveWorkbook, _
                "", "", 0, "start-of-day", sheetname, False, 0)
    Application.Run "DV.xlsm!SetEventsOn"
    
    Set tmpSheet = ActiveWorkbook.Sheets(sheetname)
    tmpSheet.Names.Add UCase(sheetname) & "_DATA", outputRange
    tmpSheet.Names.Add UCase(sheetname) & "_DATA_HEADER", outputRange.Rows(1)

exitsub:
    Set tmpSheet = Nothing
    
End Sub


Public Sub LaunchBrowser(urlname As String, x As Long, y As Long, width As Long, height As Long)
Dim execStr As String
Dim objShell As Object

    pypath = """C:\Users\burtn\AppData\Local\Microsoft\WindowsApps\python3.11.exe"""
    proctitle = """Mozilla Firefox"""
    pyLauncher = """E:\new_onedrive\Velox Financial Technology\Velox Shared Drive - Documents\General\Tools\py\resize_window.py"""
    execproc = """C:\\Program Files\\Mozilla Firefox\\firefox.exe"""
    args = """-foreground -tab """ & urlname

    Set objShell = VBA.CreateObject("Wscript.Shell")
    execStr = execStr & pypath
    execStr = execStr & " "
    execStr = execStr & pyLauncher
    execStr = execStr & " "
    execStr = execStr & "x=" & CStr(x)
    execStr = execStr & " "
    execStr = execStr & "y=" & CStr(y)
    execStr = execStr & " "
    execStr = execStr & "width=" & CStr(width)
    execStr = execStr & " "
    execStr = execStr & "height=" & CStr(height)
    execStr = execStr & " "
    execStr = execStr & "proctitle=" & proctitle
    execStr = execStr & " "
    execStr = execStr & "launch=True"
    execStr = execStr & " "
    execStr = execStr & "execproc=" & execproc
    execStr = execStr & " "
    execStr = execStr & "args=" & args
    
    objShell.Run execStr, vbHide

End Sub


Sub RunPython()

Dim objShell As Object
Dim PythonExe, PythonScript As String
    
    Set objShell = VBA.CreateObject("Wscript.Shell")

    PythonExe = """C:\Users\burtn\AppData\Local\Microsoft\WindowsApps\python3.11.exe"""
    PythonScript = """E:\new_onedrive\Velox Financial Technology\Velox Shared Drive - Documents\General\Tools\py\resize_window.py"""
    
    objShell.Run PythonExe & " " & PythonScript
    
End Sub

Sub RunPowershell(arg1 As String)
Dim objShell As Object
Dim PSExe, PSScript As String
    
    Set objShell = VBA.CreateObject("Wscript.Shell")

    PythonExe = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
    PythonScript = "" & Environ("USERPROFILE") & "\Deploy\Upload-MMReport.ps1"
    
    Debug.Print PythonExe & " " & PythonScript & " " & arg1

    objShell.Run PythonExe & " " & PythonScript & " " & arg1

End Sub

Function GetMondayFolders(workingdir As String) As Variant
Dim objShell As Object
Dim PSExe, PSScript As String
Dim folderString As String
Dim outputfilepath As String
Dim outputFileArray() As String
Dim linecount As Long
Dim lineSplit() As String

    Set objShell = VBA.CreateObject("Wscript.Shell")

    outputfilepath = workingdir & "\.folders.csv"
    
    PSExe = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
    PSScript = "" & workingdir & "\GetFolder-Monday-Nodep.ps1 " & outputfilepath
    
    Debug.Print PSExe & " " & PSScript
    objShell.Run PSExe & " " & PSScript, 1, True
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set File = FSO.OpenTextFile(outputfilepath, 1, True)
    
    ReDim outputFileArray(1 To 1000, 1 To 9)
    linecount = 1
    Do While File.AtEndOfStream = False
        lineSplit = Split(File.ReadLine, ",")
        If UBound(lineSplit) = 6 Then
            outputFileArray(linecount, 1) = lineSplit(0)
            outputFileArray(linecount, 2) = lineSplit(1)
            outputFileArray(linecount, 3) = lineSplit(2)
            outputFileArray(linecount, 4) = lineSplit(3)
            outputFileArray(linecount, 5) = lineSplit(4)
            outputFileArray(linecount, 6) = lineSplit(5)
            outputFileArray(linecount, 7) = lineSplit(6)
            outputFileArray(linecount, 8) = "https://veloxfintechcom.sharepoint.com/" & lineSplit(3)
            outputFileArray(linecount, 9) = "a" & left(lineSplit(0), 10)

            
            linecount = linecount + 1
        End If
    Loop

    
    GetMondayFolders = outputFileArray
    File.Close

endsub:
    Set FSO = Nothing
    Set File = Nothing
    Set objShell = Nothing
    Erase outputFileArray
    
End Function

Public Sub EditNewsletterExec()

    LaunchApp "C:\Program Files\Notepad++\notepad++.exe", _
        "\\wsl.localhost\Ubuntu\home\burtnolej\sambashare\veloxmon\websitepy\output_articles\generated_docs.html"

End Sub
Public Sub LaunchApp(appname As String, param As String)
Dim execStr As String

    execStr = appname & " " & """" & param & """"
    Shell execStr, vbNormalFocus

    
End Sub


Public Sub ShowHideables()
Dim controls As CommandBarControls
Dim thiscontrol As Variant

    ShowFormulaBar
    ShowDisplayables "vbautils.xlsm"
    ShowMenuBar "vbautils.xlsm"
    Set controls = CommandBars.FindControls
    
    For Each thiscontrol In controls
         Debug.Print thiscontrol.Caption
    Next thiscontrol

End Sub

Public Sub HideHideables()
Dim controls As CommandBarControls
Dim thiscontrol As Variant

    HideFormulaBar
    HideDisplayables "vbautils.xlsm"
    HideMenuBar "vbautils.xlsm"
    Set controls = CommandBars.FindControls
    
    For Each thiscontrol In controls
         Debug.Print thiscontrol.Caption
    Next thiscontrol

    
End Sub


Public Sub HideFormulaBar(bookname As String)
Dim tmpWindow As Window
    Set tmpWindow = Windows(bookname)
    
    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = False
End Sub
Public Sub ShowFormulaBar(bookname As String)
Dim tmpWindow As Window

    If Right(bookname, 4) <> "xlsm" Then
        Set tmpWindow = Windows(bookname & ".xlsm")
    Else
        Set tmpWindow = Windows(bookname)
    End If
    
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True
End Sub

Public Sub HideDisplayables(bookname As String)
Dim tmpWindow As Window
    Set tmpWindow = Windows(bookname)
    
    tmpWindow.DisplayWorkbookTabs = False
    tmpWindow.DisplayFormulas = False
    tmpWindow.DisplayHorizontalScrollBar = False
    tmpWindow.DisplayVerticalScrollBar = False
    
    tmpWindow.DisplayHeadings = False
    
exitsub:
    Set tmpWindow = Nothing
    
End Sub

Sub DisplayVBEExec(Optional param As Variant)
    Dim isEnabled As Boolean
    ' where 21 is the "Visual Basic" CommandBar and 4 is the "Visual Basic Editor" CommandBarButton
    With Application.CommandBars(21).controls(4)
        isEnabled = .Enabled
        .Enabled = True
        .Execute
        .Enabled = isEnabled
    End With
End Sub

Sub RunRibbonEditorExec(bookname As String)
    execStr = "C:\Program Files\Office RibbonX Editor\OfficeRibbonXEditor.exe" & " " & """" & bookname & """"
    Shell execStr, vbNormalFocus
End Sub
Sub CloseRibbonEditorExec()
    If KillApp("OfficeRibbonXEditor.exe") = 0 Then MsgBox "Terminated" Else MsgBox "Failed"
End Sub


Sub HideSheets(bookname As String, Optional visibleSheet As String = "BLANK")
Dim i As Integer
Dim tmpBook As Workbook
    Set tmpBook = Workbooks(bookname)
    
    For i = 1 To tmpBook.Sheets.count
        If tmpBook.Sheets(i).Name <> visibleSheet Then
            On Error Resume Next
            tmpBook.Sheets(i).Visible = False
            On Error GoTo 0
        End If
    Next i
    
End Sub

Public Sub ShowSheets(Optional bookname As String = "")
Dim tmpBook As Workbook
Dim i As Integer
    If bookname <> "" Then
        If Right(bookname, 4) <> "xlsm" Then
            Set tmpBook = Workbooks(bookname & ".xlsm")
        Else
            Set tmpBook = Workbooks(bookname)
        End If
    Else
        Set tmpBook = ActiveWorkbook
    End If
    
    For i = 1 To tmpBook.Sheets.count
        Debug.Print tmpBook.Sheets(i).Name
        tmpBook.Sheets(i).Visible = True
    Next i
    
End Sub

Public Sub MaxBookExec(param As String, width As Long, height As Long)
    GetScreenRes width, height
    ResizeWindowExec param, width, height
    MoveWindow param, 0, 0
End Sub
Public Sub MinBookExec(param As String, width As Long, height As Long, x As Long, y As Long)
    ResizeWindowExec param, width, height
    MoveWindow param, x, y
End Sub
Public Sub HideBookExec(param As String)
    HideBook param
End Sub
Public Sub ShowBookExec(param As String, width As Long, height As Long, x As Long, y As Long)
    ShowBook param
    ResizeWindowExec param, width, height
    MoveWindow param, x, y
End Sub
            
Public Sub ShowToolsExec(param As String)
    ShowDisplayables param
    ShowMenuBar param
    ShowSheets param
    ShowFormulaBar param
End Sub

Public Sub HideToolsExec(param As String)
    HideDisplayables param
    HideMenuBar param
    HideSheets param
    HideFormulaBar param
End Sub

            
Public Sub ShowDisplayables(bookname As String)
Dim tmpWindow As Window

    If Right(bookname, 4) <> "xlsm" Then
        Set tmpWindow = Windows(bookname & ".xlsm")
    Else
        Set tmpWindow = Windows(bookname)
    End If
    
    tmpWindow.DisplayWorkbookTabs = True
    tmpWindow.DisplayFormulas = True
    tmpWindow.DisplayHorizontalScrollBar = True
    tmpWindow.DisplayVerticalScrollBar = True
    tmpWindow.DisplayHeadings = True


exitsub:
    Set tmpWindow = Nothing
    
End Sub
Public Sub HideMenuBar(bookname As String)
Dim tmpWorkbook As Workbook
Dim tmpCommandBar As CommandBar
    Set tmpWorkbook = Application.Workbooks(bookname)

    Set tmpCommandBar = CommandBars("Worksheet Menu Bar")
    
exitsub:
    
    tmpCommandBar.Enabled = False
    Set tmpWorkbook = Nothing
    Set tmpCommandBar = Nothing
End Sub
Public Sub ShowMenuBar(bookname As String)
Dim tmpWorkbook As Workbook
Dim tmpCommandBar As CommandBar

    If Right(bookname, 4) <> "xlsm" Then
        Set tmpWorkbook = Workbooks(bookname & ".xlsm")
    Else
        Set tmpWorkbook = Workbooks(bookname)
    End If
    
    Set tmpCommandBar = CommandBars("Worksheet Menu Bar")
    
exitsub:
    
    tmpCommandBar.Enabled = True
    Set tmpWorkbook = Nothing
    Set tmpCommandBar = Nothing
End Sub

Public Sub CloseWorkbook(bookname As String)
Dim tmpWorkbook As Workbook
    
    If Right(bookname, 4) <> "xlsm" Then
        Set tmpWorkbook = Workbooks(bookname & ".xlsm")
    Else
        Set tmpWorkbook = Workbooks(bookname)
    End If
    
    tmpWorkbook.Close
    
exitsub:
    
    Set tmpWorkbook = Nothing

End Sub


Public Function OpenWorkbook(bookFullPath As String) As String
Dim tmpWorkbook As Workbook
Dim currentWorkbook As Workbook
Dim splitpath() As String

    Set currentWorkbook = ActiveWorkbook
    splitpath = Split(bookFullPath, "\")
    If splitpath(0) = "$HOME" Then
        splitpath(0) = Environ("USERPROFILE")
        bookFullPath = Join(splitpath, "\")
    End If

    Set tmpWorkbook = Application.Workbooks.Open(bookFullPath)
    currentWorkbook.Activate
    
exitsub:
    OpenWorkbook = tmpWorkbook.Name
    Set tmpWorkbook = Nothing

End Function


Public Sub DisplayCommandBars()
Dim tmpCommandBar As CommandBar

    For i = 1 To CommandBars.count
        Set tmpCommandBar = CommandBars(i)
        Debug.Print tmpCommandBar.Name
    Next i

exitsub:
    Set tmpCommandBar = Nothing
    
End Sub
Public Sub DisplayWindow(bookname As String)
Dim tmpWindow As Window
    Set tmpWindow = Windows(bookname)
    tmpWindow.Activate

exitsub:
    Set tmpWindow = Nothing

End Sub
Public Sub ZoomWindow(bookname As String, zoom As Double)
Dim tmpWindow As Window
    Set tmpWindow = Windows(bookname)
    tmpWindow.zoom = zoom * 100

exitsub:
    Set tmpWindow = Nothing

End Sub
Public Sub HideWindow(bookname As String)
Dim tmpWindow As Window
    Set tmpWindow = Windows(bookname)
    tmpWindow.Visible = False

exitsub:
    Set tmpWindow = Nothing

End Sub
Public Sub HideBook(bookname As String)
Dim tmpWorkbook As Workbook
Dim filename As String
    Set tmpWorkbook = ActiveWorkbook
    
    filename = bookname & ".xlsm"
    Workbooks(filename).Activate
    ActiveWindow.WindowState = xlMinimized
    
exitsub:
    Set tmpWorkbook = Nothing

End Sub
Public Sub ShowBook(bookname As String)
Dim tmpWorkbook As Workbook
Dim filename As String

    Set tmpWorkbook = ActiveWorkbook
    filename = bookname & ".xlsm"
    Workbooks(filename).Activate
    ActiveWindow.WindowState = xlMaximized
    
exitsub:
    Set tmpWorkbook = Nothing

End Sub
Public Sub ResizeWindowExec(bookname As String, Optional width As Long = 1000, Optional height As Long = 1000)
Dim tmpWindow As Window
Dim tmpWorkbook As Workbook

    If Right(bookname, 4) <> "xlsm" Then
        Set tmpWindow = Windows(bookname & ".xlsm")
    Else
        Set tmpWindow = Windows(bookname)
    End If
    tmpWindow.WindowState = xlNormal
    tmpWindow.width = width
    tmpWindow.height = height
    
exitsub:
    Set tmpWindow = Nothing

End Sub
Public Sub MoveWindow(bookname As String, Optional top As Long = 0, Optional left As Long = 0)
Dim tmpWindow As Window
Dim tmpWorkbook As Workbook

    If Right(bookname, 4) <> "xlsm" Then
        Set tmpWindow = Windows(bookname & ".xlsm")
    Else
        Set tmpWindow = Windows(bookname)
    End If
    tmpWindow.WindowState = xlNormal
    tmpWindow.top = top
    tmpWindow.left = left
    
exitsub:
    Set tmpWindow = Nothing

End Sub
Public Function GetWorkbooks() As Variant
Dim books() As Variant
Dim book As Workbook
Dim bookCount As Integer

    ReDim books(0 To Application.Workbooks.count - 1)
    
    For Each book In Application.Workbooks
        books(bookCount) = book.Name
        bookCount = bookCount + 1
    Next book
    
    GetWorkbooks = books
End Function
