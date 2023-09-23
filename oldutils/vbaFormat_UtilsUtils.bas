Attribute VB_Name = "Format_Utils"

Sub FormatParagraph(rTargetRange As Range)
    With rTargetRange
        .Columns.AutoFit
        .Rows.AutoFit
    End With
End Sub
Sub SetWrapOn(rTargetRange As Range)
    With rTargetRange
        .Select
        .WrapText = True
    End With
End Sub
Sub SetTopLeftAlignOn(rTargetRange As Range)
    With rTargetRange
        .Select
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
    End With
End Sub
Sub SetCenterAlignOn(rTargetRange As Range, Optional wbTmp As Workbook, Optional sSheetName As String)
    On Error GoTo err
    With rTargetRange
        .Select
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
err:
    If IsSet(wbTmp) And sSheetName <> "" Then
        wbTmp.Sheets(sSheetName).Range(rTargetRange.Address).HorizontalAlignment = xlCenter
        wbTmp.Sheets(sSheetName).Range(rTargetRange.Address).VerticalAlignment = xlCenter
    End If
        
End Sub
Sub SetColumnWidth(iColIdx As Integer, iColWidth As Integer, sSheetName As String)
Dim rColumn As Range
    With GetSheet(Application.ActiveWorkbook, sSheetName)
        Set rColumn = .Columns(iColIdx)
        rColumn.ColumnWidth = iColWidth
    End With
End Sub
Sub SetAllRowsHeight(iRowHeight As Integer, sSheetName As String, Optional rSource As Range)
Dim rColumn As Range
    With GetSheet(Application.ActiveWorkbook, sSheetName)
        .Activate
        If rSource Is Nothing Then
            Cells.Select
        Else
            rSource.Select
        End If
        Selection.RowHeight = iRowHeight
    End With
End Sub
Sub CopyFormat(wbSourceBook As Workbook, wbTargetbook As Workbook, sSourceSheetName As String, _
              sTargetSheetName As String, sSourceAddress As String, sTargetAddress As String)

    On Error GoTo 0
    
    With wbSourceBook.Sheets(sSourceSheetName)
        .Range(sSourceAddress).Copy
    End With
    
    With wbTargetbook.Sheets(sTargetSheetName)
        .Range(sTargetAddress).PasteSpecial Paste:=xlPasteFormats, operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    End With

End Sub

Sub CopyFormatRange(rSourceRange As Range, rTargetRange As Range)
    'rSourceRange.Select
    'Selection.Copy
    rSourceRange.Copy
    
    'rTargetRange.Worksheet.Activate
    'rTargetRange.Select
    'Application.CutCopyMode = False
    ' HERE
    'Selection.PasteSpecial Paste:=xlPasteFormats, operation:=xlNone, SkipBlanks:=False, Transpose:=False
    rTargetRange.PasteSpecial Paste:=xlPasteFormats, operation:=xlNone, SkipBlanks:=False, Transpose:=False
End Sub
Sub SetBgColor(sSheetName As String, rTarget As Range, iRed As Integer, iGreen As Integer, _
                    iBlue As Integer, Optional wbTmp As Workbook)
    
    If IsSet(wbTmp) = False Then
        Set wbTmp = ActiveWorkbook
    End If
    
    With GetSheet(wbTmp, sSheetName)
        With rTarget.Interior
            .Color = RGB(iRed, iGreen, iBlue)
        End With
    End With
End Sub
Sub SetBgColorByRGB(sSheetName As String, rTarget As Range, rgbClr As rgbColor, Optional wbTmp As Workbook)
    If IsSet(wbTmp) = False Then
        Set wbTmp = ActiveWorkbook
    End If

    With GetSheet(wbTmp, sSheetName)
        With rTarget.Interior
            .Color = RGB(rgbClr.Red, rgbClr.Green, rgbClr.Blue)
        End With
    End With
    
End Sub
Sub SetFgColor(sSheetName As String, sRange As String, iRed As Integer, iGreen As Integer, _
            iBlue As Integer, Optional wbTmp As Workbook)
    If IsSet(wbTmp) = False Then
        Set wbTmp = ActiveWorkbook
    End If
            
    With GetSheet(wbTmp, sSheetName)
        With .Range(sRange).Font
            .Color = RGB(iRed, iGreen, iBlue)
        End With
    End With
End Sub
Sub SetFgColorByRGB(sSheetName As String, sRange As String, rgbClr As rgbColor, Optional wbTmp As Workbook)
    If IsSet(wbTmp) = False Then
        Set wbTmp = ActiveWorkbook
    End If
            
    With GetSheet(wbTmp, sSheetName)
        With .Range(sRange).Font
            .Color = RGB(rgbClr.Red, rgbClr.Green, rgbClr.Blue)
        End With
    End With
End Sub
Sub SetFont(sSheetName As String, sRange As String, sName As String, iSize As Integer, sStyle As String, _
    Optional wbTmp As Workbook)
    
    If IsSet(wbTmp) = False Then
        Set wbTmp = ActiveWorkbook
    End If
    
    With GetSheet(wbTmp, sSheetName)
        With .Range(sRange).Font
            .Name = sName
            .FontStyle = sStyle
            .Size = iSize
            '.Underline = xlUnderlineStyleNone
        End With
    End With
End Sub
Sub SetBgColorFromString(sSheetName As String, rTarget As Range, sRGB As String, _
            Optional wbTmp As Workbook)
Dim aRGBElements() As String

    If IsSet(wbTmp) = False Then
        Set wbTmp = ActiveWorkbook
    End If
    
    aRGBElements = Split(sRGB, ",")
    SetBgColor sSheetName, rTarget, CInt(aRGBElements(0)), CInt(aRGBElements(1)), CInt(aRGBElements(2)), _
                wbTmp:=wbTmp

End Sub
Function GetBgColor(sSheetName As String, rTarget As Range) As rgbColor
Dim vColorVal As Variant
Dim cRGB As New rgbColor

    vColorVal = rTarget.Interior.Color
    
    cRGB.Red = (vColorVal Mod 256)
    cRGB.Green = ((vColorVal \ 256) Mod 256)
    cRGB.Blue = (vColorVal \ 65536)
    
    Set GetBgColor = cRGB

End Function

Function GetRGBFromExcelColor(vColorVal As Variant) As rgbColor
Dim cRGB As New rgbColor
    
    cRGB.Red = (vColorVal Mod 256)
    cRGB.Green = ((vColorVal \ 256) Mod 256)
    cRGB.Blue = (vColorVal \ 65536)
    
    Set GetRGBFromExcelColor = cRGB
End Function
