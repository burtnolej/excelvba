Attribute VB_Name = "Module_Utils"
Const C_MODULE_NAME = "Module_Utils"

Sub DeleteModules(xlwb As Workbook)
Dim VBProj As VBIDE.VBProject
Dim VBComp As VBIDE.VBComponent
Dim vModulesNames() As String
Dim iCount As Integer
ReDim vModulesNames(0 To 100)
    
    If sModuleName <> "" Then
        vModulesNames(iCount) = sModuleName
        iCount = iCount + 1
    Else
        Set VBProj = xlwb.VBProject
        For Each VBComp In VBProj.VBComponents
            If VBComp.Type = vbext_ct_StdModule Or VBComp.Type = vbext_ct_ClassModule Then
                DeleteModule xlwb, VBComp.Name
            End If
        Next VBComp
    End If
End Sub
Public Function CreateModule(xlwb As Workbook, sModuleName As String, sCode As String) As VBComponent
Dim module As VBComponent
    Set module = xlwb.VBProject.VBComponents.Add(vbext_ct_StdModule)
    module.Name = sModuleName
    module.CodeModule.AddFromString sCode
    Set CreateModule = module
End Function
Public Sub AddCode2Module(xlwb As Workbook, sModuleName As String, sCode As String)
Dim VBProj As VBIDE.VBProject
Dim VBComp As VBIDE.VBComponent
    
    Set VBProj = xlwb.VBProject
    Set VBComp = VBProj.VBComponents(sModuleName)
    VBComp.CodeModule.AddFromString sCode
    
    Set VBProj = Nothing
    Set VBComp = Nothing

End Sub

Sub DeleteModule(xlwb As Workbook, sModuleName As String)
Dim VBProj As VBIDE.VBProject
Dim VBComp As VBIDE.VBComponent
    
    Set VBProj = xlwb.VBProject
    Set VBComp = VBProj.VBComponents(sModuleName)
    On Error Resume Next
    VBProj.VBComponents.Remove VBComp
    On Error GoTo 0
End Sub
Function GetModule(xlwb As Workbook, sModuleName As String) As Variant
Dim VBProj As VBIDE.VBProject
Dim VBComp As VBIDE.VBComponent
Dim VBRef As VBIDE.Reference
    
    Set VBProj = xlwb.VBProject
    
    On Error GoTo err
    Set VBComp = VBProj.VBComponents(sModuleName)
    Set GetModule = VBComp
    Exit Function
    
err:
    On Error GoTo err2
    Set VBRef = VBProj.References(sModuleName)
    Set GetModule = VBRef
    Exit Function
    
err2:
    Set GetModule = Nothing
    
End Function

Function ModuleExists(xlwb As Workbook, sModuleName As String) As Boolean
Dim VBProj As VBIDE.VBProject
Dim VBComp As VBIDE.VBComponent
    
    Set VBProj = xlwb.VBProject
    
    On Error GoTo err
    Set VBComp = VBProj.VBComponents(sModuleName)
    ModuleExists = True
    Exit Function
    
err:
    ModuleExists = False
    
End Function
Function ExportModules(xlwb As Workbook, sDirectory As String, sSuffix As String, Optional sModuleName As String) As String()
Dim VBProj As VBIDE.VBProject
Dim VBComp As VBIDE.VBComponent
Dim vModulesNames() As String
Dim iCount As Integer
ReDim vModulesNames(0 To 100)
    
    If sModuleName <> "" Then
        vModulesNames(iCount) = sModuleName
        iCount = iCount + 1
    Else
        Set VBProj = xlwb.VBProject
        For Each VBComp In VBProj.VBComponents
            vModulesNames(iCount) = VBComp.Name
            iCount = iCount + 1
        Next VBComp
    End If
    ReDim Preserve vModulesNames(0 To iCount - 1)
    
    ExportModules = vModulesNames

    For i = 0 To UBound(vModulesNames)
        Set VBComp = GetModule(xlwb, vModulesNames(i))
        Path = sDirectory & vModulesNames(i) & sSuffix & ".bas"
        
        Call VBComp.Export(Path)
    Next i

End Function

Function ImportModules(xlwb As Workbook, sDirectory As String, _
                    Optional sModuleName As String, _
                    Optional bOverwrite As Boolean = True, _
                    Optional sIgnoreModules As String, _
                    Optional bDryRun As Boolean = False) As Integer
Dim VBProj As VBIDE.VBProject
Dim VBComps As VBIDE.VBComponents
Dim VBComp As VBIDE.VBComponent
Dim vFileNames() As String, vIgnoreModules() As String
Dim iCount As Integer
Dim sFuncName As String

    sFuncName = C_MODULE_NAME & "." & "ImportModules"
    If sModuleName <> "" Then
        ReDim vModulesNames(0 To 0)
        vFileNames(0) = sDirectory & "/" & sModuleName
    Else
        vFileNames = GetFolderFiles(sDirectory & "/")
    End If

    Set VBComps = xlwb.VBProject.VBComponents
    
    For Each sFile In vFileNames
        sModuleName = Split(sFile, ".")(0)
        vIgnoreModules = Split(sIgnoreModules, ",")
        If InArray(vIgnoreModules, sModuleName) = False Then
            If ModuleExists(xlwb, sModuleName) = True And bOverwrite = False Then
                FuncLogIt sFuncName, "skipping " & sModuleName & " as exists and bOverwrite = False", C_MODULE_NAME, LogMsgType.INFO
            ElseIf ModuleExists(xlwb, sModuleName) = True And bOverwrite = True Then
                FuncLogIt sFuncName, "deleting [" & sModuleName & "] as exists but overwrite=True", C_MODULE_NAME, LogMsgType.INFO
                If bDryRun = False Then
                    DeleteModule xlwb, sModuleName
                    VBComps.Import sDirectory & "/" & sFile
                End If
            Else
                On Error Resume Next
                If bDryRun = False Then
                    VBComps.Import sDirectory & "/" & sFile
                End If
                iCount = iCount + 1
                FuncLogIt sFuncName, "importing [" & sModuleName & "]", C_MODULE_NAME, LogMsgType.INFO
                On Error GoTo 0
            End If
        Else
            FuncLogIt sFuncName, "skipping [" & sFile & "] as in ignore list", C_MODULE_NAME, LogMsgType.INFO
        End If
    Next sFile
    ImportModules = iCount
    
End Function
Function GetProcsInModules(wb As Workbook, Optional sModuleName As String, _
            Optional bTestsOnly As Boolean = False, _
            Optional bAddBookName As Boolean = False) As Dictionary
Dim VBProj As VBIDE.VBProject
'Dim VBComp As VBIDE.VBComponent
Dim VBComp As Variant
Dim vModuleNames() As String
Dim iCount As Integer
ReDim vModuleNames(0 To 100)
Dim sFuncName As String
Dim dProc As New Dictionary
Dim dDetails As Dictionary
Dim iNumProcs As Integer
Dim sProcName As String
Dim sModName As Variant
Dim sComments As String

setup:
    sFuncName = "GetProcsInModules"

main:

    If sModuleName <> "" Then
        vModuleNames = InitStringArray(Array(sModuleName))
    Else
        vModuleNames = GetModules(wb)
    End If
    
    For Each sModName In vModuleNames
        Set VBComp = GetModule(wb, CStr(sModName))
        
        For i = 1 To VBComp.CodeModule.CountOfLines
            sProcName = VBComp.CodeModule.ProcOfLine(i, vbext_pk_Proc)
                       
            If bTestsOnly = True And Left(sProcName, 4) <> "Test" Then
                GoTo nextproc
            End If
                
            If sProcName = BLANK Then
                ' pass
            ElseIf VBComp.CodeModule.Lines(i, 1) <> BLANK Then ' official start of proc can be blank line above the proc
                If dProc.Exists(sProcName) = False Then
                    Set dDetails = New Dictionary
                    dDetails.Add "ModuleName", sModName
                    dDetails.Add "FirstLine", i
                    dDetails.Add "Args", VBComp.CodeModule.Lines(i, 1)
                    'dDetails.Add "BodyLine", VBComp.CodeModule.ProcBodyLine(sProcName, vbext_pk_Proc)
                    dDetails.Add "VBComp", VBComp
                    dDetails.Add "CodeModule", VBComp.CodeModule
                    
                    If bAddBookName = True Then
                        dDetails.Add "BookName", wb.Name
                    End If
                    dProc.Add sProcName, dDetails
                End If
            End If
nextproc:
        Next i
    Next sModName
    
    Set GetProcsInModules = dProc
    
End Function
Public Function GetProcCode(wb As Workbook, sCompName As String, sProcName As String, _
        Optional eProcType As vbext_ProcKind = vbext_ProcKind.vbext_pk_Proc) As String
Dim VBProj As VBIDE.VBProject
Dim VBComp As VBIDE.VBComponent
Dim iProcLength As Integer, iProcStartLine As Integer
Dim sCode As String, sFuncName As String

setup:
    On Error GoTo err
    sFuncName = "GetProcCode"

main:
    Set VBProj = wb.VBProject
    Set VBComp = VBProj.VBComponents(sCompName)
    iProcLength = VBComp.CodeModule.ProcCountLines(sProcName, vbext_pk_Proc)
    iProcStartLine = VBComp.CodeModule.ProcStartLine(sProcName, vbext_pk_Proc)
    GetProcCode = VBComp.CodeModule.Lines(iProcStartLine, iProcLength)
    Exit Function
    
err:
    FuncLogIt sFuncName, "[" & err.Description & "] not retreieve code for [sCompName" & sCompName & "] [sProcName=" & sProcName & "]", C_MODULE_NAME, LogMsgType.Error
    GetProcCode = "-1"
End Function

Public Function InsertProcCode(wb As Workbook, sCompName As String, sProcName As String, vCode() As String, _
        Optional eProcType As vbext_ProcKind = vbext_ProcKind.vbext_pk_Proc) As String
Dim VBProj As VBIDE.VBProject
Dim VBComp As VBIDE.VBComponent
Dim iProcLength As Integer, iProcStartLine As Integer
Dim sCode As String, sFuncName As String
Dim iLine


setup:
    On Error GoTo err
    sFuncName = "GetProcCode"

main:
    Set VBProj = wb.VBProject
    Set VBComp = VBProj.VBComponents(sCompName)
    iProcLength = VBComp.CodeModule.ProcCountLines(sProcName, vbext_pk_Proc)
    iProcStartLine = VBComp.CodeModule.ProcStartLine(sProcName, vbext_pk_Proc)
    
    iLine = iProcStartLine + 1
    
    Do While Left(VBComp.CodeModule.Lines(iLine, 1), 3) = "Dim"
        iLine = iLine + 1
    Loop
    
    For i = 0 To UBound(vCode)
        VBComp.CodeModule.InsertLines iLine, vCode(i)
        iLine = iLine + 1
    Next i

    Exit Function
    
err:
    FuncLogIt sFuncName, "[" & err.Description & "] not retreieve code for [sCompName" & sCompName & "] [sProcName=" & sProcName & "]", C_MODULE_NAME, LogMsgType.Error
    InsertProcCode = "-1"
End Function

Public Function GetProcAnalysis(wb As Workbook, dProc As Dictionary) As Dictionary
Dim sProcName As Variant
Dim sModuleName As String
Dim VBComp As VBIDE.VBComponent
Dim VBCodeModule As VBIDE.CodeModule
Dim iLineNum As Integer
Dim sComments As String
Dim dDetail As Dictionary
Dim sInComments As String 'In,Out or None

    For Each sProcName In dProc.Keys
        sComments = ""
        Set dDetail = dProc.Item(sProcName)
        sProcName = Replace(sProcName, Space, BLANK)
        sModuleName = Replace(dDetail.Item("ModuleName"), Space, BLANK)

        Set VBCodeModule = dDetail.Item("CodeModule")

        sInComments = "None"
        For iLineNum = dDetail.Item("FirstLine") To dDetail.Item("FirstLine") + 10
            If Left(VBCodeModule.Lines(iLineNum, 1), 4) = "'<<<" Then
                sInComments = "In"
                GoTo nextilinenum
            End If
            
            If Left(VBCodeModule.Lines(iLineNum, 1), 4) = "'>>>" Then
                sInComments = "Out"
                GoTo nextproc
            End If
            
            If sInComments = "In" Then
                If sComments = BLANK Then
                    sComments = VBCodeModule.Lines(iLineNum, 1)
                Else
                    sComments = sComments & vbCrLf & VBCodeModule.Lines(iLineNum, 1)
                End If
            End If

nextilinenum:
        Next iLineNum

nextproc:
    dDetail.Add "Comments", sComments
    dProc.Remove sProcName
    dProc.Add sProcName, dDetail

    Next sProcName
    

    Set GetProcAnalysis = dProc
End Function
Function GetModules(wb As Workbook) As String()
Dim VBProj As VBIDE.VBProject
Dim VBComp As VBIDE.VBComponent
Dim VBRef As VBIDE.Reference

Dim vModuleNames() As String
Dim iCount As Integer

setup:
    sFuncName = "GetModules"
    ReDim vModuleNames(0 To 150)

main:

    Set VBProj = wb.VBProject

    For Each VBComp In VBProj.VBComponents
        vModuleNames(iCount) = VBComp.Name
        iCount = iCount + 1
    Next VBComp
    
    'For Each VBRef In VBProj.References
    '    If VBRef.Type = vbext_rk_Project Then
    '        vModuleNames(iCount) = VBRef.Name
    '        iCount = iCount + 1
    '    End If
    'Next VBRef
    
    ReDim Preserve vModuleNames(0 To iCount - 1)
    
    GetModules = vModuleNames
End Function
