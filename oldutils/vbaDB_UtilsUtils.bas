Attribute VB_Name = "DB_Utils"
Public Function GetQueryFromRange(Optional rSource As Range) As String
Dim iType As Integer
Dim sFuncName As String

    If rSource Is Nothing Then
        Set rSource = Selection
    End If
        
    sFuncName = "GetQueryFromRange"
    
    iType = vbYesNoCancel + vbQuestion + vbDefaultButton2
    iReturnCode = PopUpWindow("Query:" & rSource.value, "Verify Query", iType)
    
    If iReturnCode = 6 Then
        GetQueryFromRange = rSource.value
    Else
        GetQueryFromRange = "-1"
    End If
End Function

Public Function Encode(sValue As String, sEncoding As String) As String

    If sEncoding = "uu" Then
        Encode = UUEncode(sValue)
        Exit Function
    End If
    
    Encode = sValue
    
End Function
Public Function Decode(sValue As String, sEncoding As String) As String

    If sEncoding = "uu" Then
        Decode = UUDecode(sValue)
        Exit Function
    End If
    
    Decode = sValue
    
End Function
Public Sub CreatePySqliteArgsFile( _
    sDatabaseName As String, _
    sTableName As String, _
    Optional bDeleteFlag As Boolean = False, _
    Optional bDecodeFlag As Boolean = False, _
    Optional aColumns As Variant, _
    Optional aColumnDefns As Variant, _
    Optional sEncoding As String = "uni", _
    Optional aRows As Variant, _
    Optional sQryStr As String = "", _
    Optional sFileName As String, _
    Optional sRuntimeDir As String, _
    Optional sResultFileName As String)
    
Dim PYTHONPATH As String, MYHOME As String
Dim sTmp As String

    PYTHONPATH = LCase(Environ("PYTHONPATH"))
    MYHOME = Environ("MYHOME")
    
    ' these are here because you cannot have a const as a default value form optional arguments
    If sFileName = "" Then
        sFileName = MYHOME & "\Development\pyshell.args.txt"
    End If
    
    If sRuntimeDir = "" Then
        sRuntimeDir = MYHOME & "\runtime"
    End If
    
    'create a file of the following format
    'values with square brackets need to be uuencoded
    
    'database_name  :[dbname]\n
    'table_name     :[tablename]\n
    'delete_flag    :[False|True]\n # delete database after access; used in testing
    'decode_flag    :[False|True]\n # store as
    'columns        :[colA]^[colB]^[colC]\n
    'column_defns   :[colA^colAtype]$$[colB^colBtype]$$[colC^colCtype]\n
    'qry_str        :[select ?? from ?? where ?? = ?? etc
    'rows           :[valA]^[valB]^[valC]$$[valA1]^[valB1]^[valC1]\n

    On Error Resume Next 'in case running for first time and nothing to delete
    Call DeleteFile(sFileName)
    On Error GoTo 0
    Call TouchFile(sFileName)

    If bDecodeFlag = True Then
        Call AppendFile(sFileName, "database_name:" & EncodeBase64(sDatabaseName) & vbCrLf)
        Call AppendFile(sFileName, "table_name:" & EncodeBase64(sTableName) & vbCrLf)
        Call AppendFile(sFileName, "delete_flag:" & EncodeBase64(CStr(bDeleteFlag)) & vbCrLf)
        Call AppendFile(sFileName, "decode_flag:" & EncodeBase64(CStr(bDecodeFlag)) & vbCrLf)
        Call AppendFile(sFileName, "runtime_dir:" & EncodeBase64(sRuntimeDir) & vbCrLf)

        If sQryStr <> "" Then
            Call AppendFile(sFileName, "qry_str:" & EncodeBase64(sQryStr) & vbCrLf)
        End If

        If sResultFileName <> "" Then
            Call AppendFile(sFileName, "result_file:" & EncodeBase64(sResultFileName) & vbCrLf)
        End If
        
    Else
        Call AppendFile(sFileName, "database_name:" & Encode(sDatabaseName, sEncoding) & vbCrLf)
        Call AppendFile(sFileName, "table_name:" & Encode(sTableName, sEncoding) & vbCrLf)
        Call AppendFile(sFileName, "delete_flag:" & Encode(CStr(bDeleteFlag), sEncoding) & vbCrLf)
        Call AppendFile(sFileName, "decode_flag:" & Encode(CStr(bDecodeFlag), sEncoding) & vbCrLf)
        Call AppendFile(sFileName, "runtime_dir:" & Encode(sRuntimeDir, sEncoding) & vbCrLf)

        If sQryStr <> "" Then
            Call AppendFile(sFileName, "qry_str:" & Encode(sQryStr, sEncoding) & vbCrLf)
        End If
        If sResultFileName <> "" Then
            Call AppendFile(sFileName, "result_file:" & Encode(sResultFileName, sEncoding) & vbCrLf)
        End If
        
    End If
    
    If Not IsMissing(aColumnDefns) Then
        Call AppendFile(sFileName, "column_defns:" & ArrayNDtoString(aColumnDefns, bUUEncode:=bDecodeFlag) & vbCrLf)
    End If
    
    If Not IsMissing(aColumns) Then
        Call AppendFile(sFileName, "columns:" & ArrayNDtoString(aColumns, bUUEncode:=bDecodeFlag) & vbCrLf)
    End If
    
    If Not IsMissing(aRows) Then
        
        sTmp = ArrayNDtoString(aRows, bUUEncode:=bDecodeFlag)
        
        If bDecodeFlag = True Then
            sTmp = AsciiReplace(sTmp, 10, 43, iToCount:=3)
        Else
            sTmp = Replace(sTmp, "'", "")
        End If

        Call AppendFile(sFileName, "rows:" & sTmp & vbCrLf)
    
    End If


End Sub

Public Function ParsePySqliteArgsFile(sFileName As String, Optional sEncoding As String = "uni") As Dictionary
Dim dResult As New Dictionary
Dim sFileStr As String
Dim aResultRows() As String, aRows() As String, aCols() As String
Dim aResult As Variant
Dim iRow As Integer, iCol As Integer

    sFileAsStr = ReadFile(sFileName)
    
    aResultRows = Split(sFileAsStr, vbLf)
    
    For Each sResultRow In aResultRows
        sKey = Split(sResultRow, ":")(0)

        If InArray(Array("rows", "columns", "column_defns"), sKey) Then
            ReDim aResult(0 To 10000, 0 To 100)
            sValue = Split(sResultRow, ":")(1)
            aRows = Split(sValue, "$$")
            For iRow = 0 To UBound(aRows)
                aCols = Split(aRows(iRow), "^")
                For iCol = 0 To UBound(aCols)
                    aResult(iRow, iCol) = Encode(aCols(iCol), sEncoding)
                Next iCol
            Next iRow
            
            
            aResult = ReDim2DArray(aResult, iRow, iCol)
            'ReDim Preserve aResult(0 To iRow - 1, 0 To iCol - 1)
            dResult.Add sKey, aResult
        Else:
            sValue = Encode(CStr(Split(sResultRow, ":")(1)), sEncoding)
            dResult.Add sKey, sValue
        End If
    Next sResultRow
    
    Set ParsePySqliteArgsFile = dResult

End Function
Public Function DBInsert(sDatabaseName As String, _
                    sTableName As String, _
                    bDecodeFlag As Boolean, _
                    aColumns() As String, _
                    aColumnsDefns() As String, _
                    aRows As Variant, _
                    Optional sFileName As String)
       
Dim sExecPath As String
Dim iCurrentNumRows As Integer, iNumRows As Integer

    ' these are here because you cannot have a const as a default value form optional arguments
    If sFileName = "" Then
        sFileName = MYHOME & "\args.txt"
    End If
    
    sExecPath = Environ("MYHOME") & "\GitHub\quadviewer\utils\"
    sRuntimePath = Environ("MYHOME") & "\runtime\"
    CreatePySqliteArgsFile sDatabaseName, _
                            sTableName, _
                            bDecodeFlag:=bDecodeFlag, _
                            aColumns:=aColumns, _
                            aColumnDefns:=aColumnsDefns, _
                            aRows:=aRows, _
                            sQryStr:="select * from " & sTableName, _
                            sFileName:=sFileName
                            
    aArgs = InitStringArray(Array("python", _
            sExecPath & "excel_database_util.py", _
            "--access_type query", _
            "--input_filename " & sFileName, _
            "--runtime_path " & sRuntimePath))
                   
    iCurrentNumRows = UBound(Split(ShellRun_Foreground(aArgs), DOUBLEDOLLAR)) + 1
    
    
    ' create the database and table
    aArgs = InitStringArray(Array("python", _
            sExecPath & "excel_database_util.py", _
            "--access_type create", _
            "--input_filename " & sFileName, _
            "--runtime_path " & sRuntimePath))

    sResults = ShellRun_Foreground(aArgs)
    
    aArgs = InitStringArray(Array("python", _
            sExecPath & "excel_database_util.py", _
            "--access_type insert", _
            "--input_filename " & sFileName, _
            "--runtime_path " & sRuntimePath))


    sResults = ShellRun_Foreground(aArgs)
    
    aArgs = InitStringArray(Array("python", _
            sExecPath & "excel_database_util.py", _
            "--access_type query", _
            "--input_filename " & sFileName, _
            "--runtime_path " & sRuntimePath))
    

     iNumRows = UBound(Split(ShellRun_Foreground(aArgs), DOUBLEDOLLAR)) + 1
     
     DBInsert = iNumRows
End Function
Public Function DBQuery(sDatabaseName As String, _
                        sTableName As String, _
                        bDeleteFlag As Boolean, _
                        sQryStr As String, _
                        Optional bDecodeFlag As Boolean = False, _
                        Optional bResultFile As Boolean = False, _
                        Optional sFileName As String = "unipyshellargs.txt", _
                        Optional sResultFileName As String = "pyshellresults.txt") As String
Dim aArgs() As String
Dim sExecPath As String

    sRuntimePath = Environ("MYHOME") & "\runtime\"
    sExecPath = Environ("MYHOME") & "\GitHub\quadviewer\utils\excel\"
                        
    CreatePySqliteArgsFile sDatabaseName, _
                            sTableName, _
                            bDeleteFlag:=bDeleteFlag, _
                            sQryStr:=sQryStr, _
                            sFileName:=sFileName
                            
    aArgs = InitStringArray(Array("python", _
            sExecPath & "excel_database_util.py", _
            "--access_type query", _
            "--input_filename " & sFileName, _
            "--runtime_path " & sRuntimePath))
            
    If bResultFile = True Then
        ReDim Preserve aArgs(0 To UBound(aArgs) + 1)
        aArgs(UBound(aArgs)) = "--result_file " & sRuntimePath & sResultFileName

        ShellRun_Foreground aArgs
        ' return location of result file to caller
        DBQuery = sRuntimePath & sResultFileName
    Else
        ' return results directly to caller
        DBQuery = CleanString(ShellRun_Foreground(aArgs))
    End If
    
End Function

Public Function AutoParseInputRange(Optional ws As Worksheet) As Range
Dim sAddress As String
Dim rSource As Range
Dim iDataLength As Integer

    If IsSet(ws) = False Then
        Set ws = ActiveWorkbook.ActiveSheet
    End If
    With ws
        .Activate
        Set rSource = .UsedRange
        iDataLength = rSource.Rows.Count - 2
        sAddress = rSource.Rows(1).Address
        sAddress = sAddress & COMMA & rSource.Rows(2).Address
        sAddress = sAddress & COMMA & rSource.Offset(2).Resize(iDataLength).Address
        Set AutoParseInputRange = .Range(sAddress)
    End With
End Function
