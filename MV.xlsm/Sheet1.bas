VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    EVENTSON
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
Dim sFolder As String, inputType As String, initFolder As String, selectedFolder As String
Dim inputTypeArray As Variant
Dim fso As Scripting.FileSystemObject
Dim myFolder As Folder
Dim myParentFolder As Folder
    
    Set fso = CreateObject("Scripting.FileSystemObject")

    inputTypeArray = Array("Monday Gdrive Path", "Monday Folder Path", "Output Report Folder")
    
    If Target.Rows.Count > 1 Then
        GoTo endsub
    End If
    
    If Target.Columns.Count > 1 Then
        GoTo endsub
    End If
        
    If ActiveSheet.Range("ONOFF").value = "OFF" Then
        GoTo endsub
    End If
    
    On Error Resume Next
    inputType = Target.offset(, -1).value
    On Error GoTo 0
    If Not Intersect(ActiveSheet.Range("INPUT"), Target) Is Nothing Then
        If Not IsInArray(inputType, inputTypeArray) Then
            GoTo endsub
        End If
        
        If Target.value <> "" Then
            Set myFolder = fso.GetFolder(Target.value)
            initFolder = myFolder.Path
            selectedFolder = myFolder.Name
        Else
            initFolder = "c:\users"
        End If
        Target.value = GetFolderSelection(initFolder, selectedFolder)
        ActiveSheet.Cells(1, 1).Select

    End If
    
endsub:
    Set fso = Nothing
    
End Sub

