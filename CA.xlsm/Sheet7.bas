VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    RefreshCapsuleData
End Sub

Private Sub CommandButton2_Click()
    TestWriteToRESTAPIFromSheet
End Sub

Private Sub CommandButton3_Click()
    TestDeleteEntity
End Sub

Private Sub CommandButton4_Click()
    TestGetCapsuleRecord
End Sub

Private Sub CommandButton5_Click()
    TestUpdateCapsuleRecordField
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    If ActiveSheet.Range("DEBUG") = "ON" Then
        GoTo exitsub
    End If
    
    If (Target.Rows.Count > 1 Or Target.Columns.Count > 1) And Target.MergeCells = False Then
        GoTo exitsub
    End If
    
    
    If Intersect(Target, ActiveSheet.Range("RECORD_TYPE")) Is Nothing Then
    Else
        SetupRecordDefaults
        GoTo exitsub
    End If
exitsub:

End Sub

