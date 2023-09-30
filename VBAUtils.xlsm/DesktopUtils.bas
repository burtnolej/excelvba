Attribute VB_Name = "DesktopUtils"

Sub temp()

CommandBars.ExecuteMso "MaximizeRibbon"

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

Public Sub HideFormulaBar()
    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = False
End Sub
Public Sub ShowFormulaBar()
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True
End Sub

Public Sub HideDisplayables(bookName As String)
Dim tmpWindow As Window
    Set tmpWindow = Windows(bookName)
    
    tmpWindow.DisplayWorkbookTabs = False
    tmpWindow.DisplayFormulas = False
    tmpWindow.DisplayHorizontalScrollBar = False
    tmpWindow.DisplayVerticalScrollBar = False
exitsub:
    Set tmpWorkbook = Nothing
    
End Sub
Public Sub ShowDisplayables(bookName As String)
Dim tmpWindow As Window
    Set tmpWindow = Windows(bookName)
    
    tmpWindow.DisplayWorkbookTabs = True
    tmpWindow.DisplayFormulas = True
    tmpWindow.DisplayHorizontalScrollBar = True
    tmpWindow.DisplayVerticalScrollBar = True
exitsub:
    Set tmpWorkbook = Nothing
    
End Sub
Public Sub HideMenuBar(bookName As String)
Dim tmpWorkbook As Workbook
Dim tmpCommandBar As CommandBar
    Set tmpWorkbook = Application.Workbooks(bookName)

    Set tmpCommandBar = CommandBars("Worksheet Menu Bar")
    
exitsub:
    
    tmpCommandBar.Enabled = False
    Set tmpWorkbook = Nothing
    Set tmpCommandBar = Nothing
End Sub
Public Sub ShowMenuBar(bookName As String)
Dim tmpWorkbook As Workbook
Dim tmpCommandBar As CommandBar
    Set tmpWorkbook = Application.Workbooks(bookName)

    Set tmpCommandBar = CommandBars("Worksheet Menu Bar")
    
exitsub:
    
    tmpCommandBar.Enabled = True
    Set tmpWorkbook = Nothing
    Set tmpCommandBar = Nothing
End Sub

Public Sub CloseWorkbook(bookName As String)
Dim tmpWorkbook As Workbook
    Set tmpWorkbook = Application.Workbooks(bookName)
    tmpWorkbook.Close
    
exitsub:
    
    Set tmpWorkbook = Nothing

End Sub
Public Sub OpenWorkbook(bookFullPath As String)
Dim tmpWorkbook As Workbook
Dim currentWorkbook As Workbook

    Set currentWorkbook = ActiveWorkbook
    Set tmpWorkbook = Application.Workbooks.Open(bookFullPath)
    currentWorkbook.Activate
    
exitsub:
    
    Set tmpWorkbook = Nothing

End Sub

Public Sub DisplayCommandBars()
Dim tmpCommandBar As CommandBar

    For i = 1 To CommandBars.Count
        Set tmpCommandBar = CommandBars(i)
        Debug.Print tmpCommandBar.Name
    Next i
    
    'CommandBars("Worksheet Menu Bar").Enabled = False
exitsub:
    Set tmpCommandBar = Nothing
    
End Sub
Public Sub DisplayWindow(bookName As String)
Dim tmpWindow As Window
    Set tmpWindow = Windows(bookName)
    tmpWindow.Activate

exitsub:
    Set tmpWindow = Nothing

End Sub
Public Sub ZoomWindow(bookName As String, zoom As Double)
Dim tmpWindow As Window
    Set tmpWindow = Windows(bookName)
    tmpWindow.zoom = zoom * 100

exitsub:
    Set tmpWindow = Nothing

End Sub
Public Sub HideWindow(bookName As String)
Dim tmpWindow As Window
    Set tmpWindow = Windows(bookName)
    'ActiveWorkbook.Windows(1).Visible = False
    tmpWindow.Visible = False

exitsub:
    Set tmpWindow = Nothing

End Sub
Public Sub ResizeWindow(bookName As String, Optional width As Long = 1000, Optional height As Long = 1000)
Dim tmpWindow As Window
Dim tmpWorkbook As Workbook

    Set tmpWindow = Windows(bookName)
    tmpWindow.WindowState = xlNormal
    tmpWindow.width = width
    tmpWindow.height = height
    
exitsub:
    Set tmpWindow = Nothing

End Sub
Public Sub MoveWindow(bookName As String, Optional top As Long = 0, Optional left As Long = 0)
Dim tmpWindow As Window
Dim tmpWorkbook As Workbook

    Set tmpWindow = Windows(bookName)
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

    ReDim books(0 To Application.Workbooks.Count - 1)
    
    For Each book In Application.Workbooks
        books(bookCount) = book.Name
        bookCount = bookCount + 1
    Next book
    
    GetWorkbooks = books
End Function
