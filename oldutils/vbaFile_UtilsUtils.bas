Attribute VB_Name = "File_Utils"
Const C_MODULE_NAME = "File_Utils"
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
Dim fso As New FileSystemObject
Dim FileName As String
    GetFileFromPath = fso.GetFileName(sPath)
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
    FuncLogIt sFuncName, "Failed to copy [" & sSourceFilePath & "] to  [" & sTargetFilePath & "]", C_MODULE_NAME, LogMsgType.Failure

End Sub

Public Function GetFolderFiles(sPath As String, Optional bDateSorted As Boolean = False, _
                Optional vExtensions As Variant) As String()
Dim vFileNames() As String
Dim i As Integer
Dim objFSO As Object
Dim sExtension As String

    ReDim vFileNames(0 To 10000)
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    File = Dir(sPath)
    While (File <> "")
        
        If IsSet(vExtensions) = True Then
            sExtension = Split(File, ".")(UBound(Split(File, ".")))
            If InArray(vExtensions, sExtension) = True Then
                vFileNames(i) = File
                i = i + 1
                GoTo nextiter
            End If
        Else
            vFileNames(i) = File
            i = i + 1
            GoTo nextiter
        End If

        
nextiter:
    File = Dir
    Wend
    ReDim Preserve vFileNames(0 To i - 1)
    GetFolderFiles = vFileNames
End Function
Public Function CreateDir(sPath As String) As Object
Dim objFSO As Object
Dim oDir As Object
Dim sFuncName As String

    sFuncName = "CreateDir"
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    On Error GoTo err
    Set oDir = objFSO.CreateFolder(sPath)
    On Error GoTo 0
    FuncLogIt sFuncName, "Created Dir [" & sPath & "]", C_MODULE_NAME, LogMsgType.OK
    Set CreateDir = oDir
    Exit Function
err:
    FuncLogIt sFuncName, "Failed to create Dir [" & sPath & "] with err [" & err.Description & "]", C_MODULE_NAME, LogMsgType.Failure
    

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
        FuncLogIt sFuncName, "init val cannot be BLANK", C_MODULE_NAME, LogMsgType.INFO
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
        FuncLogIt sFuncName, "file [" & sFilePath & "] does not exist", C_MODULE_NAME, LogMsgType.INFO
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

