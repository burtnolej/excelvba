Attribute VB_Name = "Exec_Proc"

Option Explicit

Const C_MODULE_NAME = "Exec_Proc"

' base case must always exist in the first passed book

Private pProcStore() As String
Private pProcIndex As Variant
Private pSourceBook As Workbook
Private pSourceBookName As String
Private pSourceBookPath As String
Private pSourceBook2 As Workbook
Private pSourceBook2Name As String
Private pSourceBook2Path As String

Enum ExecProcErrorMsgType
    BASE_VERSION_DOES_NOT_EXIST = 13001
    CALLED_FUNC_FAILED = 13002
End Enum

Public Property Get SourceBook() As Workbook
    Set SourceBook = pSourceBook
End Property
Public Property Let SourceBook(value As Workbook)
    Set pSourceBook = value
End Property
Public Property Get SourceBookName() As String
    SourceBookName = pSourceBookName
End Property
Public Property Let SourceBookName(value As String)
    pSourceBookName = value
End Property
Public Property Get SourceBookPath() As String
    SourceBookPath = pSourceBookPath
End Property
Public Property Let SourceBookPath(value As String)
    pSourceBookPath = value
End Property


Public Property Get SourceBook2Path() As String
    SourceBook2Path = pSourceBook2Path
End Property
Public Property Let SourceBook2Path(value As String)
    pSourceBook2Path = value
End Property

Public Property Get SourceBook2() As Workbook
    Set SourceBook2 = pSourceBook2
End Property
Public Property Let SourceBook2(value As Workbook)
    Set pSourceBook2 = value
End Property
Public Property Get SourceBook2Name() As String
    SourceBook2Name = pSourceBook2Name
End Property
Public Property Let SourceBook2Name(value As String)
    pSourceBook2Name = value
End Property


Public Property Get ProcStore() As String()
    ProcStore = pProcStore
End Property
Public Property Let ProcStore(value() As String)
    pProcStore = value
End Property

Public Property Get ProcIndex() As Variant
    ProcIndex = pProcIndex
End Property
Public Property Let ProcIndex(value As Variant)
    pProcIndex = value
End Property

Function GetProcInstances(sProcName As String) As Dictionary
'<<<
'purpose: retreive all the versions of a proc; version syntax is [version tag] & __ & [base case name]
'param  : sProcName, String; suffix of the proc family to be retreived (i.e BookCreate when versions would
'       : Quad__BookCreate.
'rtype  : Dictionary; keys syntax [bookname] & ^ & [modulename]; value is comma delim string of proc names
'       : i.e. key=test_procs.xlsm^tmp1; value="CreateBook,Quad__CreateBook"
'>>>
Dim vInstances As Variant, iIndex As Variant, iCount As Integer
Dim vProcStore() As String, vProcModules() As String
Dim dTmp As New Dictionary
Dim sFullProcName As String, sModuleName As String, sInstances As String, sKeyName As String, sBookName As String, sFuncName As String
Dim bBaseExists As Boolean
Dim lStartTick As Long

setup:
    sFuncName = C_MODULE_NAME & "." & "GetProcInstances"
    lStartTick = FuncLogIt(sFuncName, "", C_MODULE_NAME, LogMsgType.INFUNC)
    On Error GoTo err

    bBaseExists = False
    
main:
    vProcStore = Me.ProcStore
    vInstances = IndexArrayMulti(Me.ProcIndex, sProcName, bEndsWith:=True)
    
    If UBound(vInstances) = 0 And vInstances(0) = "" Then
        ' if 0 length then no matching procs found so base must not exist
        GoTo cleanup
    End If
    
    ReDim vProcModules(0 To UBound(vInstances))

    For Each iIndex In vInstances
        sFullProcName = vProcStore(CInt(iIndex), 0)
        sModuleName = vProcStore(CInt(iIndex), 1)
        sBookName = vProcStore(CInt(iIndex), 2)
        sKeyName = sBookName & HAT & sModuleName
        If dTmp.Exists(sKeyName) = False Then
            sInstances = sFullProcName
        Else
            sInstances = dTmp.Item(sKeyName) & COMMA & sFullProcName
            dTmp.Remove sKeyName
        End If
        
        ' check that base version (no suffix) actually exists
        If LCase(sFullProcName) = LCase(sProcName) Then
            bBaseExists = True
        End If
        
        dTmp.Add sKeyName, sInstances

        iCount = iCount + 1
    Next iIndex
    
cleanup:
    If bBaseExists = False Then
        FuncLogIt sFuncName, "no base case found [sProcName" & sProcName & "]", C_MODULE_NAME, LogMsgType.Failure
        err.Raise ExecProcErrorMsgType.BASE_VERSION_DOES_NOT_EXIST, Description:="cannot find a proc named [" & sProcName & "]"
    End If
    
    FuncLogIt sFuncName, "procs found in books [" & Join(GetDictKeys(dTmp), COMMA) & "] procs found [" & Join(GetDictVals(dTmp), COMMA) & "]", C_MODULE_NAME, LogMsgType.DEBUGGING2
    Set GetProcInstances = dTmp
    FuncLogIt sFuncName, "", C_MODULE_NAME, LogMsgType.OUTFUNC, lLastTick:=lStartTick
    Exit Function
        
err:
    FuncLogIt sFuncName, "[" & err.Description & "]  raised", C_MODULE_NAME, LogMsgType.Error
    err.Raise err.Number, err.Source, err.Description ' cannot recover from this
    
    
End Function

Sub ExecProc(sProcName As String, ByRef dArgs As Dictionary)
'<<<
'purpose: exec the best version of a procedure. if version is not found call the base case
'param  : sProcName, String; suffix of the proc family to be retreived (i.e BookCreate when versions would
'       : Quad__BookCreate.
'param  : dArgs, Dictionary; contains key="ver_series" value=??? if a version to be attempted to call
'       : key/value of named params required by the called function
'rtype  : Dictionary; adds "result" key to dArgs containing object returned from function called
'       : i.e. key=result; value=workbook object
'>>>
' base function must be in 1st passed workbook
Dim dProcs As Dictionary
Dim sFullProcPath As String, sVerSeries As String, sExecName As String, sProcNames As String, sVerSeriesTmp As String, sFuncName As String, sModuleName As String, sBookName As String, sErrorDesc As String
Dim vProcNames() As String
Dim iVerNum As Integer, iVerNumTmp As Integer, i As Integer
Dim vKey As Variant
Dim lStartTick As Long, lEndtTick As Long

setup:
    sFuncName = C_MODULE_NAME & "." & "sProcName"
    'On Error GoTo err
    
main:
    sVerSeries = dArgs.Item("ver_series")
    sFullProcPath = sVerSeries & "__" & sProcName
    
    On Error GoTo err1
    Set dProcs = Me.GetProcInstances(sProcName)
    On Error GoTo 0
    
    If sVerSeries = "" Then
        ' we know base exists if have got this far
        sBookName = Me.SourceBookName
        sFullProcPath = sProcName
        GoTo exec
    End If
        
    For Each vKey In dProcs.Keys
        sModuleName = Split(vKey, HAT)(1)
        sBookName = Split(vKey, HAT)(0)

        vProcNames = Split(dProcs.Item(vKey), COMMA)
        For i = LBound(vProcNames) To UBound(vProcNames)
            If InStr(vProcNames(i), "__") <> 0 Then
                sVerSeriesTmp = Split(vProcNames(i), "__")(0)
                If sVerSeriesTmp = sVerSeries Then
                    GoTo exec
                End If
            End If
        Next i
    Next vKey
    
    ' dropped so call base case (not addinf version & __ as suffix
    sVerSeries = ""
    sBookName = Me.SourceBookName
    sFullProcPath = sProcName
    
exec:
    sExecName = "" & QUOTE & sBookName & QUOTE & BANG & sFullProcPath & ""

    lStartTick = FuncLogIt(sExecName, "", sModuleName, LogMsgType.INFUNC)
    'dArgs.Add "result", Application.Run(sExecName, dArgs)
    On Error GoTo err2
    Application.Run sExecName, dArgs
    On Error GoTo 0
    
    AddDict dArgs, "exec_version", sVerSeries
    AddDict dArgs, "exec_book", sBookName
    AddDict dArgs, "exec_module", sModuleName

    ' this is needed because errors cannot be thrown through Application.Run
    
    If dArgs.Exists("error_code") Then
        FuncLogIt sFuncName, "error thrown while calling [" & sProcName & "]", C_MODULE_NAME, LogMsgType.Error
        err.Raise dArgs.Item("error_code"), Description:=dArgs.Item("error_desc") & "[" & CStr(sProcName) & "]" ' cannot recover from this
    End If
    FuncLogIt sFuncName, "", C_MODULE_NAME, LogMsgType.OUTFUNC, lLastTick:=lStartTick
    Exit Sub

cleanup:
    FuncLogIt sFuncName, "Executed [sProcName=" & sProcName & "] [Result type=" & CStr(MyVarType(dArgs.Item("result"))) & "]", C_MODULE_NAME, LogMsgType.DEBUGGING2
    Exit Sub
        
    
err2:
    If err.Number = 449 Then
        sErrorDesc = "[sFullProcPath=" & sFullProcPath & "] needs to support dictionary args / ExecProc"
        FuncLogIt sFuncName, "[sErrorDesc=" & sErrorDesc & "]  raised", C_MODULE_NAME, LogMsgType.Error
        err.Raise ErrorMsgType.FUNC_NEEDS_TO_SUPPORT_EXEC_PROC_ARGS, Description:=sErrorDesc
    End If
    
err1:
    dArgs.Add "result", -1
    dArgs.Add "exec_version", "None"
    FuncLogIt sFuncName, "[" & err.Description & "]  raised", C_MODULE_NAME, LogMsgType.Error
    
    If err.Number = ExecProcErrorMsgType.BASE_VERSION_DOES_NOT_EXIST Then
        err.Raise err.Number, err.Source, err.Description ' cannot recover from this
    End If
    
End Sub

Sub InitProcStore()
'<<<
'purpose: search code to find all procedures available; store in a table with n rows and 3 columns
'         columns are ProcName, ModuleName, Bookname
'rtype  : store the 3 col array in a member attribute and then also store the 1st col(procName) in a memattr
'>>>
Dim dProc As Dictionary
Dim aResult() As String, aResult2() As String, aAllResult() As String
Dim sFuncName As String
Dim lStartTick As Long

setup:
    sFuncName = C_MODULE_NAME & "." & "InitProcStore"
    lStartTick = FuncLogIt(sFuncName, "", C_MODULE_NAME, LogMsgType.INFUNC)
    On Error GoTo err

main:
    Set dProc = GetProcsInModules(Me.SourceBook, bAddBookName:=True)
    ReDim aResult2(0 To 0, 0 To 2)
    aResult = Dict2Array(dProc, InitStringArray(Array("ModuleName", "BookName")))
        
    If IsSet(Me.SourceBook2) = True Then
        FuncLogIt sFuncName, "Procs from additional book loaded [bookname=" & Me.SourceBook2.Name & "]", C_MODULE_NAME, LogMsgType.DEBUGGING
        Set dProc = New Dictionary
        Set dProc = GetProcsInModules(Me.SourceBook2, bAddBookName:=True)
        aResult2 = Dict2Array(dProc, InitStringArray(Array("ModuleName", "BookName")))
    End If
    
    aAllResult = AddArrays(3, aResult, aResult2)
    Me.ProcStore = aAllResult
    Me.ProcIndex = GetColumnFrom2DArray(aAllResult, 0)

cleanup:
    FuncLogIt sFuncName, "loaded [num_procs=" & CStr(UBound(aAllResult) - 1) & "]", C_MODULE_NAME, LogMsgType.DEBUGGING2
    FuncLogIt sFuncName, "", C_MODULE_NAME, LogMsgType.OUTFUNC, lLastTick:=lStartTick
    Exit Sub
        
err:
    FuncLogIt sFuncName, "[" & err.Description & "]  raised", C_MODULE_NAME, LogMsgType.Error
    err.Raise err.Number, err.Source, err.Description ' cannot recover from this
    
End Sub
Public Function IsAExecProc() As Boolean
    IsAExecProc = True
End Function
Public Sub InitProperties(Optional wbTmp As Workbook, Optional wbTmp2 As Workbook)
'<<<
'purpose: instantiate class with known procs and details;
'param  : wbTmp, Workbook; book contains all base classes and some or all versions
'param  : wbTmp2, Workbook; book contains additional versions
'rtype  : store the 3 col array in a member attribute and then also store the 1st col(procName) in a memattr
'>>>
Dim sFuncName As String
Dim lStartTick As Long
setup:
    sFuncName = C_MODULE_NAME & "." & "InitProperties"
    On Error GoTo err
    lStartTick = FuncLogIt(sFuncName, "", C_MODULE_NAME, LogMsgType.INFUNC)
    'On Error GoTo err
    
main:

    If IsSet(wbTmp) = False Then
        Set wbTmp = ActiveWorkbook
    End If
    
    Me.SourceBookName = wbTmp.Name
    Me.SourceBookPath = wbTmp.Path
    
    If BookOpen(Me.SourceBookName) = True Then
        Me.SourceBook = Workbooks(Me.SourceBookName)
    Else
        Me.SourceBook = OpenBook(Me.SourceBookName, sPath:=Me.SourceBookPath)
    End If
    
    If IsSet(wbTmp2) = True Then
        FuncLogIt sFuncName, "Additional book registered [bookname=" & wbTmp2.Name & "]", C_MODULE_NAME, LogMsgType.DEBUGGING

        Me.SourceBook2Name = wbTmp2.Name
        Me.SourceBook2Path = wbTmp2.Path
        
        If BookOpen(Me.SourceBook2Name) = True Then
            Me.SourceBook2 = Workbooks(Me.SourceBook2Name)
        Else
            Me.SourceBook2 = OpenBook(Me.SourceBook2Name, sPath:=Me.SourceBook2Path)
        End If
    
    End If
    
    Me.InitProcStore
    
    LetExecProcGlobal Me
    
    'Workbooks(Me.SourceBookName).Sheets("config").Cells(2, 10).value = ObjPtr(Me)
    
cleanup:
    FuncLogIt sFuncName, "Object Created", C_MODULE_NAME, LogMsgType.DEBUGGING2
    FuncLogIt sFuncName, "", C_MODULE_NAME, LogMsgType.OUTFUNC, lLastTick:=lStartTick
    Exit Sub
        
err:
    FuncLogIt sFuncName, "[" & err.Description & "]  raised", C_MODULE_NAME, LogMsgType.Error
    err.Raise err.Number, err.Source, err.Description ' cannot recover from this
    
End Sub
    

