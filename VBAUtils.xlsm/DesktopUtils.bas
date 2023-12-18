Attribute VB_Name = "DesktopUtils"


'python3.11.exe  .\resize_window.py x=1, y=1,width=200,height=200,proctitle='Mozilla Firefox',launch=True, execproc="C:\\Program Files\\Mozilla Firefox\\firefox.exe"
'python3.11.exe  .\resize_window.py x=1, y=1,width=200,height=200,proctitle='VBAUtils - Excel'
'python3.11.exe  .\resize_window.py x=1, y=1,width=200,height=200,proctitle='Mozilla Firefox'





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
Sub SetXLOnTop(Optional param As Variant)
    ShowXLOnTop True
End Sub
Sub SetXLNormal(Optional param As Variant)
    ShowXLOnTop False
End Sub

Sub temp()

CommandBars.ExecuteMso "MaximizeRibbon"

End Sub

Public Sub LaunchExplorer(foldername As String, Optional param As String)


Shell "C:\WINDOWS\explorer.exe """ & foldername & "", vbNormalFocus
End Sub

'Sub TestLaunchBrowser()
'    LaunchBrowser "www.bbc.com", 1, 500, 3000, 1000, """Mozilla Firefox""", """C:\\Program Files\\Mozilla Firefox\\firefox.exe""", _
'        """E:\new_onedrive\Velox Financial Technology\Velox Shared Drive - Documents\General\Tools\py\resize_window.py""", _
'        """C:\Users\burtn\AppData\Local\Microsoft\WindowsApps\python3.11.exe"""
'End Sub


Sub TestLaunchGitBash()
    LaunchGitBash "C:\Users\burtn\Development\py"
End Sub
Public Sub LaunchGitBash(Optional startdir As String = "C:\Users\burtn\Development")

Dim execStr As String
Dim objShell As Object
    
    psexepath = "POWERSHELL.exe -noexit"
    execpath = """C:\Users\burtn\Development\ps\Launch-GitBash.ps1"""
    'startdir = """C:\Users\burtn\Development"""
    Set objShell = VBA.CreateObject("Wscript.Shell")

    execStr = psexepath & " " & execpath & " " & startdir
    objShell.Run execStr, vbHide
    
    
End Sub

Sub TestLaunchBrowser()
    Debug.Print IsCapitalized("ddsdsd")
    'LaunchBrowser "www.bbc.com", 1, 312, 2566, 1389
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
    
    'CreateCustomNamedRanges outputRange.Rows(1), outputRange, tmpSheet
exitsub:
    Set tmpSheet = Nothing
    
End Sub


Public Sub LaunchBrowser(urlname As String, x As Long, y As Long, width As Long, height As Long)
'Public Sub LaunchBrowser(urlname As String, x As Long, y As Long, width As Long, height As Long, proctitle As String, execproc As String, pyLauncher As String, pypath As String)
Dim execStr As String
Dim objShell As Object
    'Shell "C:\Program Files\Google\Chrome\Application\chrome" & " " & urlname
    
    
    pypath = """C:\Users\burtn\AppData\Local\Microsoft\WindowsApps\python3.11.exe"""
    proctitle = """Mozilla Firefox"""
    pyLauncher = """E:\new_onedrive\Velox Financial Technology\Velox Shared Drive - Documents\General\Tools\py\resize_window.py"""
    execproc = """C:\\Program Files\\Mozilla Firefox\\firefox.exe"""
    args = """-foreground -tab """ & urlname

    Set objShell = VBA.CreateObject("Wscript.Shell")
    
    'execStr = execStr & """C:\Users\burtn\AppData\Local\Microsoft\WindowsApps\python3.11.exe"""
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
    
    'Debug.Print execStr
    'objShell.Run execStr
    objShell.Run execStr, vbHide

End Sub


Sub RunPython()

Dim objShell As Object
Dim PythonExe, PythonScript As String
    
    Set objShell = VBA.CreateObject("Wscript.Shell")

    PythonExe = """C:\Users\burtn\AppData\Local\Microsoft\WindowsApps\python3.11.exe"""
    PythonScript = """E:\new_onedrive\Velox Financial Technology\Velox Shared Drive - Documents\General\Tools\py\resize_window.py"""
    
    objShell.Run PythonExe & " " & PythonScript
    'PythonExe & PythonScript
    
End Sub
Public Sub LaunchApp(appname As String, param As String)
Dim execStr As String

    execStr = appname & " " & """" & param & """"
    Shell execStr, vbNormalFocus

End Sub

Sub Test()
    If TaskKill("OfficeRibbonXEditor.exe") = 0 Then MsgBox "Terminated" Else MsgBox "Failed"
End Sub

Sub KillApp(sTaskName As String)
Dim result As Variant
    result = CreateObject("WScript.Shell").Run("taskkill /f /im " & sTaskName, 0, True)
End Sub

Sub testcommandbars()

    CommandBars("Document Recovery").Visible = False
    Debug.Print CommandBars("Document Recovery").Name
    
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
    
    'CommandBars.ExecuteMso "MaximizeRibbon"
    
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
    
    'CommandBars.ExecuteMso "MaximizeRibbon"
    
End Sub

Public Sub Workspace()
Dim books As Variant
Dim tmpWindow As Window
Dim tmpWorkbook As Workbook

    books = GetWorkbooks

    For i = 0 To UBound(books)
        ResizeWindow CStr(books(i))
        ToggleDisplayables CStr(books(i))
        HideMenuBar CStr(books(i))
        DisplayWindow CStr(books(i))
        WaitSecs (4)
    Next i

exitsub:
    Erase books

End Sub

Public Sub HideFormulaBar(bookname As String)
Dim tmpWindow As Window
    Set tmpWindow = Windows(bookname)
    
    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = False
End Sub
Public Sub ShowFormulaBar(bookname As String)
Dim tmpWindow As Window
    'Set tmpWindow = Windows(bookname)

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
    Set tmpWorkbook = Nothing
    
End Sub

Sub DisplayVbe(Optional param As Variant)
    Dim isEnabled As Boolean
    ' where 21 is the "Visual Basic" CommandBar and 4 is the "Visual Basic Editor" CommandBarButton
    With Application.CommandBars(21).controls(4)
        isEnabled = .Enabled
        .Enabled = True
        .Execute
        .Enabled = isEnabled
    End With
End Sub

Sub HideSheets(bookname As String, Optional visibleSheet As String = "BLANK")

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

    
        'Set tmpBook = Workbooks(bookname)
        
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
    Set tmpWorkbook = Nothing
    
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


    

    
    'Set tmpWorkbook = Application.Workbooks(bookname)

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
    'Set tmpWorkbook = Application.Workbooks(bookname)
    
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
    
    'CommandBars("Worksheet Menu Bar").Enabled = False
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
    'ActiveWorkbook.Windows(1).Visible = False
    tmpWindow.Visible = False

exitsub:
    Set tmpWindow = Nothing

End Sub
Public Sub HideBook(bookname As String)
Dim tmpWorkbook As Workbook
Dim filename As String
    Set tmpWorkbook = ActiveWorkbook
    'ActiveWorkbook.Windows(1).Visible = False
    
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
    'ActiveWorkbook.Windows(1).Visible = False
    filename = bookname & ".xlsm"
    Workbooks(filename).Activate
    ActiveWindow.WindowState = xlMaximized
    
exitsub:
    Set tmpWorkbook = Nothing

End Sub
Public Sub ResizeWindow(bookname As String, Optional width As Long = 1000, Optional height As Long = 1000)
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
