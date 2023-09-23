Attribute VB_Name = "OS_Utils"
Const C_MODULE_NAME = "OS_Utils"
Public Function ShellRun(ParamArray X()) As String
Dim s As String, sLine As String, sCmd As String, sFuncName As String, sCwd As String, sFileName As String
Dim oShell As Object, oExec As Object, oOutput As Object, oError As Object
Dim iX As Integer

    sFuncName = "ShellRun"
    sFileName = "C:\Users\burtnolej\vba_source_new.bat"
    sCwd = GetHomePath
    
    ChDir sCwd & "/runtime"

    FuncLogIt sFuncName, "running from directory [" & sCwd & "]", C_MODULE_NAME, LogMsgType.INFO

    On Error Resume Next 'in case running for first time and nothing to delete
    Call DeleteFile(sFileName)
    On Error GoTo 0
    Call TouchFile(sFileName)
    
    Call AppendFile(sFileName, "nircmd.exe win hide ititle " & DOUBLEQUOTE & "cmd.exe" & DOUBLEQUOTE & vbCrLf)
    Call AppendFile(sFileName, Join(X(0), " "))
    
    sCmd = sFileName
    Set oShell = CreateObject("WScript.Shell")
    oShell.Run sCmd, 0, True


End Function

Public Function ShellRun_Foreground(ParamArray X()) As String
Dim s As String, sLine As String, sCmd As String, sFuncName As String, sCwd As String, sFileName As String
Dim oShell As Object, oExec As Object, oOutput As Object, oError As Object
Dim iX As Integer

    sFuncName = "ShellRun"
    sFileName = "C:\Users\burtnolej\vba_source_new.bat"
    sCwd = GetHomePath
    
    ChDir sCwd & "/runtime"

    FuncLogIt sFuncName, "running from directory [" & sCwd & "]", C_MODULE_NAME, LogMsgType.INFO
    sCmd = "cmd /c " & Join(X(0), " ")

    'Run a shell command, returning the output as a string'
    Set oShell = CreateObject("WScript.Shell")

    Set oExec = oShell.exec(sCmd)
 
    Set oOutput = oExec.StdOut
    Set oError = oExec.StdErr

    'handle the results as they are written to and read from the StdOut object'
    While Not oOutput.AtEndOfStream
        sLine = oOutput.ReadLine
        If sLine <> "" Then s = s & sLine & vbCrLf
    Wend
   
    If oExec.ExitCode = 0 Then
        FuncLogIt sFuncName, "executed shell cmd [" & sCmd & "]", C_MODULE_NAME, LogMsgType.OK
    Else
        FuncLogIt sFuncName, "Error executing shell cmd [" & sCmd & "]", C_MODULE_NAME, LogMsgType.Failure
        ShellRun_Foreground = "-1"
        Exit Function
    End If
    
    ShellRun_Foreground = s

End Function

Public Function GetHomePath() As String
Dim oShell As Object
Dim iLen As Integer
Dim sSuffix As String, sNewSuffix As String

    sSuffix = "\OneDrive\Desktop"

    
    sNewSuffix = "Documents"
    iLen = Len(sSuffix)
    
    Set oShell = CreateObject("WScript.Shell")
    GetHomePath = oShell.SpecialFolders("Desktop")
    
    If Right(GetHomePath, iLen) = sSuffix Then
        GetHomePath = Left(GetHomePath, Len(GetHomePath) - iLen) & "\" & sNewSuffix
        Exit Function
    End If
    
    sSuffix = "\Desktop"
    iLen = Len(sSuffix)

    If Right(GetHomePath, iLen) = sSuffix Then
        GetHomePath = Left(GetHomePath, Len(GetHomePath) - iLen) & "\" & sNewSuffix
    End If
    
    
End Function

Public Function PopUpWindow(strText As String, strTitle As String, iType As Integer, _
            Optional iWait As Integer = 3000) As Integer
Dim oShell As Object

'https://ss64.com/vb/popup.html

    Set oShell = CreateObject("WScript.Shell")
    PopUpWindow = oShell.Popup(strText, iWait, strTitle, iType)

End Function

Public Sub PopUpInput()

   Debug.Print InputBox("Give me some input", "Hi", 1)

End Sub
