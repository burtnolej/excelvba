VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Sub Workbook_Open()
    AddReferences
    ActiveWorkbook.Sheets("Config").Activate
    ActiveSheet.Cells(1, 1).Select
End Sub

