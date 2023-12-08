VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public AllowSave As Boolean

Private Sub Workbook_Open()

    'ResizeWindow "vbautils.xlsm", 1920, 200
    'ResizeVBAUtilsWindow
    'MoveWindow "vbautils.xlsm", 0, 0
    
    'Application.DisplayFormulaBar = False

End Sub


Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    
   ' If Not AllowSave Then
   '     MsgBox "You can't save this workbook!"
   '     Cancel = True
   ' End If
End Sub


Public Sub CustomSave(Optional newFileName As String = "")
Dim workbookPath As String

    ThisWorkbook.AllowSave = True

    workbookPath = ActiveWorkbook.path
    
    ' TODO: Add the custom save logic
    If newFileName <> "" Then
        newFileName = workbookPath & "\" & newFileName
        ThisWorkbook.SaveAs newFileName
    Else
        newFileName = workbookPath & "\" & "copy_" & ActiveWorkbook.Name
        ThisWorkbook.SaveCopyAs newFileName
    End If
    
    
    
    MsgBox "A copy of this workbook has been saved successfully as : " & newFileName, , "Workbook saved"
    ThisWorkbook.AllowSave = False
End Sub
