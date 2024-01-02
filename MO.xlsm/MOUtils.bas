Attribute VB_Name = "MOUtils"
Sub CloseBook(wb As Workbook)

    Application.DisplayAlerts = False
    wb.Close
    Application.DisplayAlerts = True

End Sub
