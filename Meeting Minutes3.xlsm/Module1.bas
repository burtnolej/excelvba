Attribute VB_Name = "Module1"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    ActiveWorkbook.Worksheets("CLIENT").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("CLIENT").Sort.SortFields.Add2 Key:=Range( _
        "A2:A12423"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("CLIENT").Sort
        .SetRange Range("A2:L12423")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
