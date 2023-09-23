Attribute VB_Name = "Date_Utils"
Private mlngStart As Long
'Private Declare Function GetTickCount Lib "kernel32" () As Long
Public Function GetDateString(sDate As String, Optional sFormat As String) As String
    If sFormat = "" Then
        sFormat = "ddmmyyHHMMSS"
    End If
    
    GetDateString = Format(sDate, sFormat)
End Function

Sub test()
Dim d1 As Date
Dim lResult As Double

    Debug.Print GetNow
End Sub
Public Function GetTicks() As Double
    GetTicks = GetTickCount
End Function




