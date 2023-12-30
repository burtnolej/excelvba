Attribute VB_Name = "Module1"

Sub TestDeleteButtons()
    DeleteButtons ActiveWorkbook.Sheets("Sheet1")
    
End Sub

Sub CheckInChangesExec(Optional param As String = "")
    Application.Run "vbautils.xlsm!CheckInChanges", ActiveWorkbook.Name
End Sub

Public Sub DeleteButtons(targetSheet As Worksheet)
Dim Buttons As Object, Button As Object, myShape As Shape

Dim RowNumber As Integer

    For Each myShape In targetSheet.Shapes
        If myShape.Type = msoOLEControlObject Or myShape.Type = msoAutoShape Or myShape.Type = msoFormControl Then
            'On Error Resume Next
            myShape.Delete
            On Error GoTo 0
        End If
    Next myShape

End Sub
Public Sub RefreshCapsuleDataExec(ByRef resultsDict As Dictionary)
Dim outputRange As Range
Dim RV As RibbonVariables
Dim url As String, domain As String

    Set RV = New RibbonVariables
    domain = CallByName(RV, "config__dataurl", VbGet)
    Set RV = Nothing
    
    Application.Run "vbautils.xlsm!SetEventsOff"
    
    On Error Resume Next
    
    
    ' entries_meetings
    url = domain + "entries_meetings.csv"
    Application.StatusBar = "loading " & url
    Set outputRange = Application.Run("vbautils.xlsm!HTTPDownloadFile", url, _
                Workbooks("CA.xlsm"), _
                "", "", 0, "start-of-day", "ENTRIES_MEETINGS", False, 0)
                
    Application.Run "vbautils.xlsm!SortRange", outputRange.Worksheet, outputRange, 7
    
    resultsDict.Add "ENTRIES_MEETINGS", outputRange.Rows.Count
    
    ' person
    url = domain + "person.csv"
    Application.StatusBar = "loading " & url
    Set outputRange = Application.Run("vbautils.xlsm!HTTPDownloadFile", url, _
                Workbooks("CA.xlsm"), _
                "", "", 0, "start-of-day", "PERSON", False, 0)
                
    
    colArray = Array(3, 6, 7, 4)
    
    Application.Run "vbautils.xlsm!CreateCalcNamedRange", outputRange.Worksheet, outputRange, "fullNameId", colArray
    
    Set outputRange = outputRange.Resize(, outputRange.Columns.Count + 1)
    
    Application.Run "vbautils.xlsm!AddNamedRange", outputRange.Worksheet, outputRange, 4, "PERSON_ID"
    Application.Run "vbautils.xlsm!AddNamedRange", outputRange.Worksheet, outputRange, 21, "PERSON_FULLNAME"

                
    Application.Run "vbautils.xlsm!SortRange", outputRange.Worksheet, outputRange, 13
    resultsDict.Add "PERSON", outputRange.Rows.Count
    
    
    ' opportunities
    url = domain + "opportunities.csv"
    Application.StatusBar = "loading " & url
    Set outputRange = Application.Run("vbautils.xlsm!HTTPDownloadFile", url, _
                Workbooks("CA.xlsm"), _
                "", "", 0, "start-of-day", "OPPORTUNITY", False, 0)
                
    colArray = Array(1, 8, 15)
    Application.Run "vbautils.xlsm!CreateCalcNamedRange", outputRange.Worksheet, outputRange, "opportunityMetaId", colArray
    
    Set outputRange = outputRange.Resize(, outputRange.Columns.Count + 1)
    
    Application.Run "vbautils.xlsm!AddNamedRange", outputRange.Worksheet, outputRange, 10, "OPPORTUNITY_ID"
    Application.Run "vbautils.xlsm!AddNamedRange", outputRange.Worksheet, outputRange, 21, "OPPORTUNITY_FULLNAME"
    
    
    Application.Run "vbautils.xlsm!SortRange", outputRange.Worksheet, outputRange, 15
    
    resultsDict.Add "OPPORTUNITY", outputRange.Rows.Count
    
    
    ' organisation
    url = domain + "organisation.csv"
    Application.StatusBar = "loading " & url
    Set outputRange = Application.Run("vbautils.xlsm!HTTPDownloadFile", url, _
                Workbooks("CA.xlsm"), _
                "", "", 0, "start-of-day", "CLIENT", False, 0)
    

    colArray = Array(2, 7, 6)
    Application.Run "vbautils.xlsm!CreateCalcNamedRange", outputRange.Worksheet, outputRange, "clientMetaId", colArray
    
    Set outputRange = outputRange.Resize(, outputRange.Columns.Count + 1)
    
    Application.Run "vbautils.xlsm!AddNamedRange", outputRange.Worksheet, outputRange, 6, "CLIENT_ID"
    Application.Run "vbautils.xlsm!AddNamedRange", outputRange.Worksheet, outputRange, 13, "CLIENT_FULLNAME"
    
    
    Application.Run "vbautils.xlsm!SortRange", outputRange.Worksheet, outputRange, 7
    'Application.Run "vbautils.xlsm!SortRange", outputRange.Worksheet, outputRange, 8
    On Error GoTo 0
    resultsDict.Add "CLIENT", outputRange.Rows.Count
    
    Application.Run "vbautils.xlsm!SetEventsOn"
    
End Sub

