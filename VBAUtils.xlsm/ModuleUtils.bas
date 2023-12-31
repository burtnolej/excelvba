Attribute VB_Name = "ModuleUtils"
'Public Sub ExportAllModules()
'Function ExportModules(xlwb As Workbook, sDirectory As String, sSuffix As String, Optional sModuleName As String) As String()
'Function GetModule(xlwb As Workbook, sModuleName As String) As Variant
'Function ImportModules(xlwb As Workbook, sDirectory As String, _
                    Optional sModuleName As String, _
                    Optional bOverwrite As Boolean = True, _
                    Optional sIgnoreModules As String, _
                    Optional bDryRun As Boolean = False) As Integer
                    
Public Function GetFilePath(fullpath As String) As String
Dim oFSO As FileSystemObject, oFile As File
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    GetFilePath = left(fullpath, Len(fullpath) - Len(oFSO.GetFileName(fullpath)))
End Function

Sub TestCheckInChanges()
    CheckInChangesExec ActiveWorkbook.Name
End Sub

Sub CheckInChangesExec(Optional bookname As String)
Dim sourcepath As String

    ActiveWorkbook.Save
    
    If bookname = "" Then
        bookname = ActiveWorkbook.Name
    End If
    
    sourcepath = ExportAllModules(ActiveWorkbook.Name)
    LaunchGitBash sourcepath
End Sub

    
Public Function ExportAllModules(Optional bookname As String, Optional param As Variant) As String
Dim ubuntubookpath As String, ubuntuhome As String, sourcepath As String, siteaddress As String, NewSourcePath As String
Dim tmpWorkbook As Workbook
    
    Set tmpWorkbook = ActiveWorkbook
    siteaddress = "https://veloxfintechcom.sharepoint.com/sites/VeloxSharedDrive/Shared Documents/"
    
    If bookname = "" Then
        bookname = ActiveWorkbook.Name
    End If
    
    sourcepath = GetFilePath(ActiveWorkbook.FullName)
    
    If left(sourcepath, Len(siteaddress)) = siteaddress Then
        NewSourcePath = "E:/Velox Financial Technology/Velox Shared Drive - Documents/" & Right(sourcepath, Len(sourcepath) - Len(siteaddress))
    Else
        NewSourcePath = sourcepath
    End If
        
    
    ubuntuhome = "\\wsl.localhost\Ubuntu\home\burtnolej\sambashare\veloxmon\excelvba"
    ubuntubookpath = ubuntuhome & "\" & bookname & "\"
    
    CreateDir ubuntubookpath
    ExportModules ActiveWorkbook, ubuntubookpath, _
        ""
        
    FileCopy bookname, NewSourcePath, ubuntubookpath

    ExportAllModules = ubuntubookpath
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
        path = sDirectory & vModulesNames(i) & sSuffix & ".bas"
        
        Call VBComp.Export(path)
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
                FuncLogIt sFuncName, "skipping " & sModuleName & " as exists and bOverwrite = False", C_MODULE_NAME, LogMsgType.Info
            ElseIf ModuleExists(xlwb, sModuleName) = True And bOverwrite = True Then
                FuncLogIt sFuncName, "deleting [" & sModuleName & "] as exists but overwrite=True", C_MODULE_NAME, LogMsgType.Info
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
                FuncLogIt sFuncName, "importing [" & sModuleName & "]", C_MODULE_NAME, LogMsgType.Info
                On Error GoTo 0
            End If
        Else
            FuncLogIt sFuncName, "skipping [" & sFile & "] as in ignore list", C_MODULE_NAME, LogMsgType.Info
        End If
    Next sFile
    ImportModules = iCount
    
End Function

