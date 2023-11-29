VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
Dim oXl As Excel.Application
    
    'CommandBars("Document Recovery").Visible = False
    
    'ResizeWindow "vbautils.xlsm", 1920, 200
    ResizeVBAUtilsWindow
    'MoveWindow "vbautils.xlsm", 0, 0
    
    Application.DisplayFormulaBar = False



End Sub
