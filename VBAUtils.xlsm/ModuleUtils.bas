Attribute VB_Name = "ModuleUtils"
'Function ExportModules(xlwb As Workbook, sDirectory As String, sSuffix As String, Optional sModuleName As String) As String()
'Function GetModule(xlwb As Workbook, sModuleName As String) As Variant
'Function ImportModules(xlwb As Workbook, sDirectory As String, _
                    Optional sModuleName As String, _
                    Optional bOverwrite As Boolean = True, _
                    Optional sIgnoreModules As String, _
                    Optional bDryRun As Boolean = False) As Integer
                    


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
