Attribute VB_Name = "FileUtils"
'Public Function GetFolderFiles(folderDir As String, folderSheet As Worksheet)
'Public Function DirExists(sPath As String) As Boolean
'Public Function GetFileFromPath(sPath As String) As String
'Public Sub FileMove(sFileName As String, sSourcePath As String, sTargetPath As String)


'Public Sub FileCopy(sFileName As String, sSourcePath As String, sTargetPath As String)
'Public Function GetFolderFiles(sPath As String, Optional bDateSorted As Boolean = False, _
'                Optional vExtensions As Variant) As String()

'Public Function CreateDir(sPath As String) As Object
'Public Sub RemoveDir(sPath As String)
'Public Function ReadFile(sPath As String) As String
'Public Function ReadFile2Array(sPath As String, _
'                                Optional sFieldDelim As String = "^", _
'                                Optional bSingleCol As Boolean = False, _
'                                Optional bVariant As Boolean = False) As Variant
'Public Function InitFileArray(sFilePath As String, _
'                             iNumLines As Integer, _
'                    Optional sInitVal As String = " ", _
'                    Optional bCreateFile As Boolean = True, _
'                    Optional bCloseFile As Boolean = True) As Object
'Public Sub WriteArray2File(vSource() As String, sFilePath As String)
'Public Function FileExists(sPath As String) As Boolean
'Public Function OpenFile(sPath As String, iRWFlag As Integer) As Object
'Public Sub AppendFile(sPath As String, sText As String)
'Public Function CreateFile(sPath As String) As Object
'Public Sub TouchFile(sPath As String)
'Public Function DeleteFile(sFileName As String, Optional sPath As String)
'Public Function WriteFile(sPath As String, sText As String)
'Public Function WriteFileObject(oFile As Object, sText As String)
'Public Function FilesAreSame(ByVal fFirst As String, ByVal fSecond As String) As Boolean

Public Function GetFolderFiles(folderDir As String, folderSheet As Worksheet)
 
Dim oFSO As Object, oFolders As Object, oFile As Object
Dim oFolder As Folder
Dim resultArray() As String, fileList As String, DDQ As String
Dim outputRange As Range, columnLink As Range, fillRange As Range, itemLink As Range
Dim i As Integer

'On Error GoTo err
    i = 0
    DDQ = """"

    folderSheet.UsedRange.Offset(1).ClearContents

    ReDim resultArray(0 To 600, 0 To 8)
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(folderDir)

    
    For Each oFile In oFolder.Files
        resultArray(i, 0) = oFile.Name
        resultArray(i, 1) = format(CDate(oFile.DateCreated), "YYYY/MM/DD")
        resultArray(i, 2) = format(CDate(oFile.DateLastModified), "YYYY/MM/DD")
        resultArray(i, 3) = "=hyperlink(" & DDQ & oFile.path & DDQ & ")"
        resultArray(i, 4) = left(oFile.Name, Len(oFile.Name) - 12)
        
        If InStr(1, oFile.Name, "_") <> 0 Then
            resultArray(i, 5) = left(oFile.Name, InStr(1, oFile.Name, "_") - 1)
        End If
        
        resultArray(i, 6) = resultArray(i, 4) & " [" & resultArray(i, 1) & "]"

        i = i + 1
    Next oFile
    
    folderSheet.Activate
    With folderSheet
        Set outputRange = .Range(Cells(2, 1), Cells(i + 1, 9))
        outputRange = resultArray
    End With
    
    sortRange ActiveSheet, outputRange, 2
    GoTo exitsub
    
err:
    MsgBox err.Number & ": " & err.Description, , ThisWorkbook.Name & ": GetSandboxFolder"
    #If IsDebug Then
        Stop            ' Used for troubleshooting - Then press F8 to step thru code
        Resume          ' Resume will take you to the line that errored out
    #Else
        Resume exitsub ' Exit procedure during normal running
    #End If
    
    
exitsub:
    Set oFSO = Nothing
    Set outputRange = Nothing
    Set folderSheet = Nothing
    Set fillRange = Nothing
    Set oFolder = Nothing
    Set oFile = Nothing
    Erase resultArray
    
End Function


Public Function DirExists(sPath As String) As Boolean
Dim objFSO As Object
Dim oFile As Object

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FolderExists(sPath) Then
        DirExists = True
        Exit Function
    End If
    
    DirExists = False
End Function
Public Function GetFileFromPath(sPath As String) As String
Dim FSO As New FileSystemObject
Dim filename As String
    GetFileFromPath = FSO.GetFileName(sPath)
End Function

Public Sub FileMove(sFileName As String, sSourcePath As String, sTargetPath As String)
Dim objFSO As Object
Dim sFuncName As String
    sFuncName = "FileMove"
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    On Error GoTo err
    objFSO.MoveFile sSourcePath & sFileName, sTargetPath & sFileName
    On Error GoTo 0
    FuncLogIt sFuncName, "Moved [" & sFileName & "] from  [" & sSourcePath & "] to [" & sTargetPath & "]", C_MODULE_NAME, LogMsgType.Failure
    Exit Sub
err:
    FuncLogIt sFuncName, "Failed to move [" & sFileName & "] from  [" & sSourcePath & "] to [" & sTargetPath & "] with err [" & err.Description & "]", C_MODULE_NAME, LogMsgType.Failure

End Sub

Public Sub FileCopy(sFileName As String, sSourcePath As String, sTargetPath As String)
Dim objFSO As Object
Dim sFuncName As String
Dim sSourceFilePath As String, sTargetFilePath As String

    sFuncName = "FileCopy"
    
    sSourceFilePath = sSourcePath & sFileName
    sTargetFilePath = sTargetPath & sFileName
    
    If FileExists(sSourceFilePath) = False Then
        GoTo err
    End If
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    On Error GoTo err
    objFSO.CopyFile sSourceFilePath, sTargetFilePath
    On Error GoTo 0
    'FuncLogIt sFuncName, "Copied from  [" & sSourceFilePath & "] to [" & sTargetFilePath & "]", C_MODULE_NAME, LogMsgType.Failure
    Exit Sub
err:
    'FuncLogIt sFuncName, "Failed to copy [" & sSourceFilePath & "] to  [" & sTargetFilePath & "]", C_MODULE_NAME, LogMsgType.Failure

End Sub


Public Function CreateDir(sPath As String) As Object
Dim objFSO As Object
Dim oDir As Object
Dim sFuncName As String

    sFuncName = "CreateDir"
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    On Error GoTo err
    Set oDir = objFSO.CreateFolder(sPath)
    On Error GoTo 0
    'FuncLogIt sFuncName, "Created Dir [" & sPath & "]", C_MODULE_NAME, LogMsgType.OK
    Set CreateDir = oDir
    Exit Function
err:
    'FuncLogIt sFuncName, "Failed to create Dir [" & sPath & "] with err [" & err.Description & "]", C_MODULE_NAME, LogMsgType.Failure
    

End Function
Public Sub RemoveDir(sPath As String)
Dim objFSO As Object
Dim oDir As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    objFSO.DeleteFolder (sPath)
End Sub
Public Function ReadFile(sPath As String) As String
Dim iLineNum As Integer

    Set oFile = OpenFile(sPath, 1)
    iLineNum = 1
    Do While oFile.AtEndOfStream = False
        If iLineNum = 1 Then
            ReadFile = oFile.ReadLine
        Else
            ReadFile = ReadFile & vbLf & oFile.ReadLine
        End If
        iLineNum = iLineNum + 1
    Loop
End Function
'Public Function ReadFile2Array(sPath As String, _
'                                Optional sFieldDelim As String = "^", _
'                                Optional bSingleCol As Boolean = False) As String()

Public Function ReadFile2Array(sPath As String, _
                                Optional sFieldDelim As String = "^", _
                                Optional bSingleCol As Boolean = False, _
                                Optional bVariant As Boolean = False) As Variant
'<<<
' purpose: take a flat file and represent in an array; default is a 2d array with
'        : full line in the first col (_,0)
' param  : sPath, string; file path to parse
' param  : sFieldDelim (optional), split the line by delim and store in n columns (_,n)
' param  : bSingleCol (optional), force into a 1d array
' returns: array of strings;
'>>>
Dim iCol As Integer, iRow As Integer
'Dim aTmpRow() As String, aTmp() As String
Dim aTmpRow As Variant, aTmp As Variant

    If bVariant = False Then
        If bSingleCol = True Then
            ReDim aTmp(0 To 30000) As String
        Else
            ReDim aTmp(0 To 30000, 0 To 100) As String
        End If
    Else
        If bSingleCol = True Then
            ReDim aTmp(0 To 30000) As Variant
        Else
            ReDim aTmp(0 To 30000, 0 To 100) As Variant
        End If
    End If
    
    Set oFile = OpenFile(sPath, 1)
    iRow = 0
    Do While oFile.AtEndOfStream = False
        If bSingleCol = True Then
            aTmp(iRow) = oFile.ReadLine
        Else
            aTmpRow = Split(oFile.ReadLine, sFieldDelim)
            For iCol = 0 To UBound(aTmpRow)
                If IsInt(aTmpRow(iCol)) = False Or bVariant = False Then
                    aTmp(iRow, iCol) = aTmpRow(iCol)
                Else
                    aTmp(iRow, iCol) = CLng(aTmpRow(iCol))
                End If
            Next iCol
        End If

        iRow = iRow + 1
    Loop

    If bVariant = False Then
        If bSingleCol = True Then
            ReDim Preserve aTmp(0 To iRow - 1)
        Else
            aTmp = ReDim2DArray(aTmp, iRow, iCol)
        End If
    Else
        If bSingleCol = True Then
            ReDim Preserve aTmp(0 To iRow - 1) As Variant
        Else
            aTmp = ReDim2DArray(aTmp, iRow, iCol, bVariant:=True)
        End If
    End If
    oFile.Close
    
    ReadFile2Array = aTmp
End Function

Public Function InitFileArray(sFilePath As String, _
                             iNumLines As Integer, _
                    Optional sInitVal As String = " ", _
                    Optional bCreateFile As Boolean = True, _
                    Optional bCloseFile As Boolean = True) As Object
'<<<
' purpose: create a file that is indexed so its easy to read/write to a specific line
' param  : sFilePath, string; file path to create
' param  : iNumLines, integer; the length of the file (in lines)
' param  : sInitVal (optional), string; default value in each line (cant have nothing)
' param  : bCreateFile (optional), whether or not to create the file before writing
' param  : bCloseFile (optional), whether or not to leave the file open
' returns: array of strings;
'>>>
Dim oFile As Object
Dim vArray() As String
Dim sFuncName As String

setup:
    sFuncName = C_MODULE_NAME & "." & "InitFileArray"
    ReDim vArray(0 To iNumLines - 1)
    ' ASSERTIONS ----------------------------------------
    If sInitVal = BLANK Then
        err.Raise ErrorMsgType.BAD_ARGUMENT, Description:="init val cannot be BLANK"
    Else
        FuncLogIt sFuncName, "init val cannot be BLANK", C_MODULE_NAME, LogMsgType.Info
    End If
    ' END ASSERTIONS -------------------------------------
    
main:
    If bCreateFile = True Then
        Set oFile = CreateFile(sFilePath)
        oFile.Close
    Else
        If FileExists(sFilePath) = True Then
            Set oFile = OpenFile(sFilePath, 2)
        Else
            err.Raise ErrorMsgType.BAD_ARGUMENT, Description:="file [" & sFilePath & "] does not exist"
        End If
    End If
    
    For i = 0 To iNumLines - 1
        vArray(i) = sInitVal
    Next i
    
    WriteArray2File vArray, sFilePath
    
    Set InitFileArray = oFile
End Function
Public Sub WriteArray2File(vSource() As String, sFilePath As String)
'<<<
' purpose: take a 1d array of strings and write directly to a file; 1 array item to 1 line
' param  : vSource, array of strings;
' param  : sFilePath, string; path to file
'>>>
Dim oFile As Object
Dim sArray As String, sFuncName As String

setup:
    sFuncName = C_MODULE_NAME & "." & "WriteArray2File"
    ' ASSERTIONS ----------------------------------------
    If FileExists(sFilePath) = False Then
        err.Raise ErrorMsgType.BAD_ARGUMENT, Description:="file does not exist"
    Else
        FuncLogIt sFuncName, "file [" & sFilePath & "] does not exist", C_MODULE_NAME, LogMsgType.Info
    End If
    ' END ASSERTIONS -------------------------------------

    sArray = Array2String(vSource, sDelim:=vbNewLine)
    Set oFile = OpenFile(sFilePath, 2)
    oFile.Write sArray
    oFile.Close
    
End Sub
Public Function FileExists(sPath As String) As Boolean
    If Dir(sPath) <> "" Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function
Public Function OpenFile(sPath As String, iRWFlag As Integer) As Object
Dim objFSO As Object
Dim oFile As Object

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set oFile = objFSO.OpenTextFile(sPath, iRWFlag)
    
    Set OpenFile = oFile
End Function
Public Sub AppendFile(sPath As String, sText As String)
Dim oFile As Object
    Set oFile = OpenFile(sPath, 8)
    oFile.Write (sText)
    Set oFile = Nothing
End Sub
Public Function CreateFile(sPath As String) As Object
Dim objFSO As Object
Dim oFile As Object

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set oFile = objFSO.CreateTextFile(sPath)
    
    Set CreateFile = oFile
    
    Set oFile = Nothing
    Set objFSO = Nothing
End Function

Public Sub TouchFile(sPath As String)
' iRWFlag = 1 for reading and 2 for writing
Dim objFSO As Object
Dim oFile As Object

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set oFile = objFSO.CreateTextFile(sPath)
    oFile.Close
    
    Set oFile = Nothing
    Set objFSO = Nothing
End Sub

Public Function DeleteFile(sFileName As String, Optional sPath As String)
Dim objFSO As Object
Dim oFile As Object
Dim sFuncName As String

    If sPath <> "" Then
        If Right(sPath, 1) <> "\" Then
        sFileName = sPath & "\\" & sFileName
        Else
            sFileName = sPath & sFileName
        End If
    End If
        
    sFuncName = "DeleteFile"
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    On Error GoTo err
    objFSO.DeleteFile sFileName
    On Error GoTo 0
    Exit Function
    
err:
    FuncLogIt sFuncName, "Failed to delete [" & sFileName & "] with err [" & err.Description & "]", C_MODULE_NAME, LogMsgType.Failure
    Debug.Print err.Description
End Function

Public Function WriteFile(sPath As String, sText As String)
Dim oFile As Object
Dim sFuncName As String

    sFuncName = "WriteFile"
    Set oFile = OpenFile(sPath, 2)
    oFile.Write (sText)
    Set oFile = Nothing
    
    FuncLogIt sFuncName, "Written [sText=" & sText & "] to  err [sPath=" & sPath & "]", C_MODULE_NAME, LogMsgType.DEBUGGING
    
End Function
Public Function WriteFileObject(oFile As Object, sText As String)
    oFile.Write (sText)
End Function
Public Function FilesAreSame(ByVal fFirst As String, ByVal fSecond As String) As Boolean
Dim lLen1 As Long, lLen2 As Long
Dim iFileNum1 As Integer
Dim iFileNum2 As Integer
Dim bytArr1() As Byte, bytArr2() As Byte
Dim lCtr As Long, lStart As Long
Dim bAns As Boolean
Dim sFuncName As String

    If Dir(fFirst) = "" Or Dir(fSecond) = "" Then
        FuncLogIt sFuncName, "Cannot find files find file [" & CStr(File1) & "]  [" & CStr(File2) & "]", C_MODULE_NAME, LogMsgType.OK
        Exit Function
    End If
        
    lLen1 = FileLen(fFirst)
    lLen2 = FileLen(fSecond)

    If lLen1 <> lLen2 Then
        FilesAreSame = False
        FuncLogIt sFuncName, "Files are not same length len1 [" & CStr(lLen1) & "] != [" & CStr(lLen2) & "]", C_MODULE_NAME, LogMsgType.OK
        Exit Function
    Else
        iFileNum1 = FreeFile
        Open fFirst For Binary Access Read As #iFileNum1
        iFileNum2 = FreeFile
        Open fSecond For Binary Access Read As #iFileNum2

        'put contents of both into byte Array
        bytArr1() = InputB(LOF(iFileNum1), #iFileNum1)
        bytArr2() = InputB(LOF(iFileNum2), #iFileNum2)
        lLen1 = UBound(bytArr1)
        lStart = LBound(bytArr1)
    
        bAns = True
        For lCtr = lStart To lLen1
            If bytArr1(lCtr) <> bytArr2(lCtr) Then
                bAns = False
                FuncLogIt sFuncName, "Bytes are not the same at char [" & CStr(lCtr) & "] [" & CStr(bytArr1(lCtr)) & "] != [" & CStr(bytArr2(lCtr)) & "]", C_MODULE_NAME, LogMsgType.OK
                Exit For
            End If
            
        Next
        FilesAreSame = bAns
       
    End If
 
    If iFileNum1 > 0 Then Close #iFileNum1
    If iFileNum2 > 0 Then Close #iFileNum2
    
End Function



