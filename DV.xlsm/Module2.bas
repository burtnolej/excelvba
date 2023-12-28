Attribute VB_Name = "Module2"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Sheets("PERSON").Select
    With ActiveWorkbook.Sheets("PERSON").Tab
        .Color = 4006690
        .TintAndShade = 0
    End With
    Sheets("EQUITIES_CSUITE_SENIOR_MIDLEVE").Select
End Sub
