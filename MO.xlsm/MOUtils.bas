Attribute VB_Name = "MOUtils"
Sub CloseBook(wb As Workbook)

    Application.DisplayAlerts = False
    wb.Close
    Application.DisplayAlerts = True

End Sub

Sub EmailItemUpdateExec()
Dim toStr As String, ccStr As String, subjectStr As String
    toStr = ActiveWorkbook.Sheets("AddNewItems").Range("SEARCH_PULSE_ITEM_EMAIL").value
    ccStr = ""
    subjectStr = ""
    content = ActiveWorkbook.Sheets("AddNewItems").Range("UPDATE_CONTENT").value

    Application.Run "vbautils.xlsm!createEmail", toStr, ccStr, subjectStr, content, _
        "", True


End Sub
            
