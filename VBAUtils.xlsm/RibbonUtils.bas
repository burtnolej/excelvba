Attribute VB_Name = "RibbonUtils"
Sub LoadCustRibbon()

Dim hFile As Long
Dim path As String, fileName As String, ribbonXML As String, user As String

hFile = FreeFile
user = Environ("Username")
path = "C:\Users\" & user & "\AppData\Local\Microsoft\Office\"
fileName = "Excel.officeUI"


ribbonXML = "<mso:customUI      xmlns:mso='http://schemas.microsoft.com/office/2009/07/customui'>" & vbNewLine
ribbonXML = ribbonXML + "  <mso:ribbon>" & vbNewLine
ribbonXML = ribbonXML + "    <mso:qat/>" & vbNewLine
ribbonXML = ribbonXML + "    <mso:tabs>" & vbNewLine
ribbonXML = ribbonXML + "      <mso:tab id='reportTab' label='Velox Tools' insertBeforeQ='mso:TabFormat'>" & vbNewLine
ribbonXML = ribbonXML + "        <mso:group id='veloxBooksGroup' label='Velox Books' autoScale='true'>" & vbNewLine
ribbonXML = ribbonXML + "          <mso:button id='openDesktop' label='Desktop.xlsm' " & vbNewLine
ribbonXML = ribbonXML + "imageMso='AppointmentColor3'      onAction='E:\Velox Financial Technology\Velox Shared Drive - Documents\General\Tools\VBAUtils!OpenDesktop'/>" & vbNewLine
ribbonXML = ribbonXML + "        </mso:group>" & vbNewLine
ribbonXML = ribbonXML + "        <mso:group id='veloxScriptsGroup' label='Velox Scripts' autoScale='true'>" & vbNewLine
ribbonXML = ribbonXML + "          <mso:button id='exportCode' label='Export Code' " & vbNewLine
ribbonXML = ribbonXML + "imageMso='AppointmentColor3'      onAction='E:\Velox Financial Technology\Velox Shared Drive - Documents\General\Tools\VBAUtils!ExportAllModules'/>" & vbNewLine
ribbonXML = ribbonXML + "        </mso:group>" & vbNewLine
ribbonXML = ribbonXML + "      </mso:tab>" & vbNewLine
ribbonXML = ribbonXML + "    </mso:tabs>" & vbNewLine
ribbonXML = ribbonXML + "  </mso:ribbon>" & vbNewLine
ribbonXML = ribbonXML + "</mso:customUI>"

ribbonXML = Replace(ribbonXML, """", "")

Open path & fileName For Output Access Write As hFile
Print #hFile, ribbonXML
Close hFile

End Sub

Sub ClearCustRibbon()

Dim hFile As Long
Dim path As String, fileName As String, ribbonXML As String, user As String

hFile = FreeFile
user = Environ("Username")
path = "C:\Users\" & user & "\AppData\Local\Microsoft\Office\"
fileName = "Excel.officeUI"

ribbonXML = "<mso:customUI           xmlns:mso=""http://schemas.microsoft.com/office/2009/07/customui"">" & _
"<mso:ribbon></mso:ribbon></mso:customUI>"

Open path & fileName For Output Access Write As hFile
Print #hFile, ribbonXML
Close hFile

End Sub

