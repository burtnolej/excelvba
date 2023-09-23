Attribute VB_Name = "LogUtils"
'Function EnumLogMsgType(i As Long) As String
'Public Sub FilterLogs(Optional sFilter As String = "0,1,2,3")
'Public Sub PurgeLogs()
'Public Function FuncLogIt(ByVal sFuncName As String, sLogMsg As String, sModuleName, eLogMsgType As LogMsgType, _
            Optional sKey As String, Optional lLastTick As Long = -1) As Long
'Function GetLogFile(Optional sLogFileName As String = "excel_log.txt") As Object
'Public Sub CloseLogFile()
'Sub SetLogFile(Optional sLogFileName As String = "excel_log.txt")
'Sub WriteToLog(aVals() As String, Optional aWidths As Variant, Optional sCaller As String = "", Optional sDelim As String = "|")
'Public Sub ViewLogs()


Public fLogFile As Object

Const C_LOG_MSG_TYPE = "FATAL,Error,Failure,INFO,OK,,,DEBUGGING,FAIL_TEST,PASS_TEST,,INFUNC,OUTFUNC,DEBUGGING2"

Public LogFilter As String
Public StartTick As Long

Function EnumLogMsgType(i As Long) As String
    EnumLogMsgType = Split(C_LOG_MSG_TYPE, COMMA)(i)
End Function

Public Sub FilterLogs(Optional sFilter As String = "0,1,2,3")
    Log_Utils.LogFilter = sFilter
End Sub
Public Sub PurgeLogs()
'<<<
'purpose: remove any log files from the runtime directory. a log is a file that ends in _log
'>>>
Dim sLogPath As String, sFuncName As String, sSheetName As String
Dim vFileNames() As String

setup:
    sFuncName = C_MODULE_NAME & "." & "DoPurgeLogs"
    sSheetName = "Logs"
    sLogPath = Environ("MYHOME") & "\runtime\"
    On Error GoTo err
        
    If SheetExists(ActiveWorkbook, sSheetName) = True Then
        deleteSheet ActiveWorkbook, sSheetName
    End If

main:
    
    GetLogFile
    CloseLogFile
    vFileNames = GetFolderFiles(sLogPath)

    For Each sFileName In vFileNames
        If InStr(sFileName, "_log") <> 0 Then
            DeleteFile sLogPath & CStr(sFileName)
        End If
    Next sFileName
    Exit Sub
    
err:
    FuncLogIt sFuncName, "Error [ " & err.Description & "]", C_MODULE_NAME, LogMsgType.Error
End Sub
Public Sub ViewLogs()
Dim vFileNames() As String
Dim sLogPath As String, sFuncName As String, sSheetName As String
Dim sFileName As Variant, iColWidth As Variant
Dim iCount As Integer, iRowNum As Integer
Dim vFile As Variant, aColWidths() As Integer
Dim wsTmp As Worksheet
Dim rSource As Range

    aColWidths = InitIntArray(Array(8, 6, 5.29, 11, 1, 15, 26, 100, 6))
    
    sFuncName = C_MODULE_NAME & "." & "DoViewLogs"
    sSheetName = "Logs"
    
    sLogPath = Environ("MYHOME") & "\runtime\"
    
    vFileNames = GetFolderFiles(sLogPath)
    
    iRowNum = 1
    
    If SheetExists(ActiveWorkbook, sSheetName) = True Then
        deleteSheet ActiveWorkbook, sSheetName
    End If
    
    Set wsTmp = CreateSheet(Application.ActiveWorkbook, sSheetName)
    
    For Each sFileName In vFileNames
        If InStr(sFileName, "_log") <> 0 Then
            FuncLogIt sFuncName, "Found log [" & sFileName & "] loading", C_MODULE_NAME, LogMsgType.OK
            
            vFile = ReadFile2Array(sLogPath & sFileName, sFieldDelim:="|", bVariant:=True)
            
            Set rSource = RangeFromStrArray(vFile, wsTmp, iRowNum, 0)
            'Set rSource = rSource.Resize(, 1).Offset(, 3)

            iRowNum = iRowNum + UBound(vFile) + 1
        End If
    Next sFileName
    
    iCount = 1
    For Each iColWidth In aColWidths
        SetColumnWidth iCount, CInt(iColWidth), sSheetName
        iCount = iCount + 1
    Next iColWidth
    
    With wsTmp
        Set rSource = .Range(.Cells(1, 1), .Cells(iRowNum, 10))
        RangeSort sSheetName, rSource, aSortColumns:=InitIntArray(Array(1, 2))
    End With
    
    CreateFilter Application.ActiveWorkbook, sSheetName, rSource.Rows(1), iRowNum + 1
    
End Sub

Public Function FuncLogIt(ByVal sFuncName As String, sLogMsg As String, sModuleName, _
            Optional sKey As String, Optional lLastTick As Long = -1) As Long
Dim aLogVals() As String, aLogFilter() As String
Dim aLogWidths() As Integer
Dim sNowTime As String, sNowDate As String
Dim lTick As Long, lDuration As Long

    lDuration = 0
    If UBound(Split(sFuncName, Period)) = 1 Then
        sFuncName = Split(sFuncName, Period)(1)
    End If
    
    lTick = GetTicks() - StartTick
    
    If lLastTick <> -1 Then
        lDuration = lTick - lLastTick
    End If
    
    If LogFilter = "" Then
        aLogFilter = InitStringArray(Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 12, 13))
    Else
        aLogFilter = Split(LogFilter, ",")
    End If
    
    If InArray(aLogFilter, eLogMsgType) = True Then
        sNowTime = GetDateString(Now(), "HH:MM:SS")
        sNowDate = GetDateString(Now(), "ddmmyy")

        ' take the calling function, a message and some id and write a formatted record to the log
        aLogVals = InitStringArray(Array(sNowTime, lTick, lDuration, EnumLogMsgType(eLogMsgType), BLANK, sModuleName, sFuncName, sLogMsg, sNowDate))
        aLogWidths = InitIntArray(Array(10, 5, 5, 8, 1, 1, 30, 60, 50))
        'Call WriteToLog(aLogVals, aLogWidths)
        Call WriteToLog(aLogVals)
    End If
    FuncLogIt = lTick
End Function

Sub SetLogFile(Optional sLogFileName As String = "excel_log.txt")
'Set the module member var to be a pointer to a file
    Set fLogFile = CreateFile(sLogFileName)
End Sub
Function GetLogFile(Optional sLogFileName As String = "excel_log.txt") As Object
' If the module member var has already been set, return it, otherwise set it
    If fLogFile Is Nothing Then
        If FileExists(sLogFileName) = True Then
            Set fLogFile = OpenFile(sLogFileName, 8)
        Else
            Call SetLogFile(sLogFileName)
        End If
    End If
    
    StartTick = GetTicks
        
    Set GetLogFile = fLogFile
End Function
Public Sub CloseLogFile()
    If Not fLogFile Is Nothing Then
        fLogFile.Close
        Set fLogFile = Nothing
    End If
End Sub
Sub WriteToLog(aVals() As String, Optional aWidths As Variant, Optional sCaller As String = "", Optional sDelim As String = "|")
' Add default elements to the passed log message attributes and either write to the log
' or display in the Immediate window if a log file has not been setup
Dim sLogStr As String
Dim sTimestamp As String
Dim sBookName As String
Dim aHeaderVals() As String
Dim aHeaderWidths() As Integer

    sLogStr = ""
    
    If sCaller <> "" Then
        sTimestamp = GetDateString(Now(), "ddmmyyHHSS")
        sBookName = ActiveWorkbook.Name
        aHeaderVals = InitStringArray(Array(sTimestamp, sBookName, sCaller))
        aHeaderWidths = InitIntArray(Array(14, 10, 20))
        sLogStr = Array2String(aHeaderVals, aWidths:=aHeaderWidths, sDelim:=sDelim) & sDelim
    End If
    
    sLogStr = sLogStr & Array2String(aVals, aWidths:=aWidths, sDelim:=sDelim)
    
    If fLogFile Is Nothing Then
        Debug.Print sLogStr
    Else
        fLogFile.Write sLogStr & vbCrLf
    End If
End Sub







