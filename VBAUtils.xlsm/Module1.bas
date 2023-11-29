Attribute VB_Name = "Module1"

Private Declare PtrSafe Function OpenClipboard Lib "USER32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function CloseClipboard Lib "USER32" () As Long
Private Declare PtrSafe Function GetLastError Lib "USER32" () As String
Private Declare PtrSafe Function GetClipboardData Lib "USER32" (ByVal wFormat As Long) As Long
Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "USER32" (ByVal wFormat As Long) As Long
Private Declare PtrSafe Function RegisterClipboardFormat Lib "USER32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
' Memory functions:
'Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
'Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
'Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)



'Declare Function OpenClipboard Lib "User32" (ByVal hwnd As Long) As Long
'Declare Function CloseClipboard Lib "User32" () As Long
'Declare Function GetClipboardData Lib "User32" (ByVal wFormat As    Long) As Long
Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags&, ByVal dwBytes As Long) As Long
Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
 
Public Const GHND = &H42
Public Const CF_TEXT = 1
Public Const MAXSIZE = 4096


Sub ResizeVBAUtilsWindow(Optional param As Variant)
    ResizeWindow "vbautils.xlsm", 1920, 200
    MoveWindow "vbautils.xlsm", 0, 0
    'HideVBAUtilsTools "vbautils.xlsm"
End Sub
    
Sub HideVBAUtilsTools(param As String)
    HideDisplayables param
    HideMenuBar param
    HideSheets param
    HideFormulaBar param
End Sub
            
Sub TestClipBoard_GetData()

    Debug.Print ClipBoard_GetData
End Sub
Function ClipBoard_GetData() As String
   Dim hClipMemory As Long
   Dim lpClipMemory As Long
   Dim MyString As String
   Dim RetVal As Long
 
   If OpenClipboard(0&) = 0 Then
      MsgBox "Cannot open Clipboard. Another app. may have it open"
      Exit Function
   End If
          
   ' Obtain the handle to the global memory
   ' block that is referencing the text.
   hClipMemory = GetClipboardData(CF_TEXT)
   If IsNull(hClipMemory) Then
      MsgBox "Could not allocate memory"
      GoTo OutOfHere
   End If
 
   ' Lock Clipboard memory so we can reference
   ' the actual data string.
   lpClipMemory = GlobalLock(hClipMemory)
 
   If Not IsNull(lpClipMemory) Then
      MyString = Space$(MAXSIZE)
      Debug.Print MyString
      RetVal = lstrcpy(MyString, lpClipMemory)
      RetVal = GlobalUnlock(hClipMemory)
       
      ' Peel off the null terminating character.
      CloseClipboard
      'MyString = Mid(MyString, 1, InStr(1, MyString, Chr$(0), 0) - 1)
   Else
      MsgBox "Could not lock memory to copy string from."
   End If
 
OutOfHere:
 
   RetVal = CloseClipboard()
   ClipBoard_GetData = MyString
 
End Function


Public Function GetClipboardIDForCustomFormat(ByVal sName As String) As Long
    Dim wFormat As Long
    wFormat = RegisterClipboardFormat(sName & Chr$(0))
    If (wFormat > &HC000&) Then
        GetClipboardIDForCustomFormat = wFormat
    End If
End Function

Public Function GetClipboardDataAsString(ByVal hWndOwner As Long, ByVal lFormatID As Long) As String
    Dim bData() As Byte
    Dim hMem   As Long
    Dim lSize  As Long
    Dim lPtr   As Long

    ' Open the clipboard for access:
    If (OpenClipboard(hWndOwner)) Then
        ' Check if this data format is available:
        If (IsClipboardFormatAvailable(lFormatID) <> 0) Then
            ' Get the memory handle to the data:
            hMem = GetClipboardData(lFormatID)
            If (hMem <> 0) Then
                ' Get the size of this memory block:
                lSize = GlobalSize(hMem)
                If (lSize > 0) Then
                    ' Get a pointer to the memory:
                    lPtr = GlobalLock(hMem)
                    If (lPtr <> 0) Then
                        ' Resize the byte array to hold the data:
                        ReDim bData(0 To lSize - 1) As Byte
                        ' Copy from the pointer into the array:
                        CopyMemory bData(0), ByVal lPtr, lSize
                        ' Unlock the memory block:
                        GlobalUnlock hMem

                        ' Now return the data as a string:
                        GetClipboardDataAsString = StrConv(bData, vbUnicode)

                    End If
                End If
                Debug.Print GetLastError
            End If
        End If
        CloseClipboard
    End If

End Function
