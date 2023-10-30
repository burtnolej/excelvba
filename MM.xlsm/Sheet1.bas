VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Const DOCS_SHARED_RANGENAME = "DOCUMENTS_SHARED"
Const ARTEFACTS_RANGENAME = "ARTEFACTS"
Const DEMOS_USED_RANGENAME = "DEMO_USED"
Const DOCUMENTS_USED_RANGENAME = "DOCUMENTS_USED"

Const VELOX_ONEDRIVE_NAME = "Velox Financial Technology"
Const DEMO_USED_FOLDER_NAME = "\Velox Shared Drive - Documents\General\Sales Cycle\Demos & Screenshots"
Const DEMO_ARTEFACTS_FOLDER_NAME = "\Velox Shared Drive - Documents\General\Sales Cycle\In Sales Process"
Const DOCUMENTS_USED_FOLDER_NAME = "\Velox Shared Drive - Documents\General\Sales Cycle\In Sales Process"
Const DOCUMENTS_SHARED_FOLDER_NAME = "\Velox Shared Drive - Documents\General\Sales Cycle\In Sales Process"
Private Sub Worksheet_Change(ByVal Target As Range)
Dim rangeName As String

    If ActiveSheet.Range("DEBUG") = "ON" Then
        GoTo exitsub
    End If
    
    If Target.Rows.count > 1 Or Target.Columns.count > 1 Then
        GoTo exitsub
    End If
    
    'If Intersect(Target, ActiveSheet.Range("INPUT_FOCUS")) Is Nothing Then
    'Else
    '    rangeName = Right(Target.Value, Len(Target.Value) - 3)
    '    ChangeInputSheetFocus rangeName
    'End If

exitsub:
    
End Sub
Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    If ActiveSheet.Range("DEBUG") = "ON" Then
        GoTo exitsub
    End If
    
    If (Target.Rows.count > 1 Or Target.Columns.count > 1) And Target.MergeCells = False Then
        GoTo exitsub
    End If
    
    
    If Intersect(Target, ActiveSheet.Range("INPUT_ARTEFACTS_FOLDER1")) Is Nothing Then
    Else
        WriteFolderHyperlink DEMO_ARTEFACTS_FOLDER_NAME, Target
        GoTo exitsub
    End If

    If Intersect(Target, ActiveSheet.Range("INPUT_FILE_4")) Is Nothing Then
    Else
        WriteFileHyperlink DEMO_ARTEFACTS_FOLDER_NAME, Target
        GoTo exitsub
    End If


    If Intersect(Target, ActiveSheet.Range("INPUT_HIGHLIGHT_TIME1")) Is Nothing Then
    Else
        PopupTextBox Target.Rows(1).Address, ActiveSheet.Name
        GoTo exitsub
    End If

    If Intersect(Target, ActiveSheet.Range("INPUT_HIGHLIGHT_QUESTION_3")) Is Nothing Then
    Else
        PopupTextBox Target.Rows(1).Address, ActiveSheet.Name
        GoTo exitsub
    End If

    If Intersect(Target, ActiveSheet.Range("INPUT_HIGHLIGHT_ANSWER_4")) Is Nothing Then
    Else
        PopupTextBox Target.Rows(1).Address, ActiveSheet.Name
        GoTo exitsub
    End If
    
    'If Intersect(Target, ActiveSheet.Range("INPUT_CLIENTINFO1")) Is Nothing Then
    'Else
    '    PopupTextBox Target.Rows(1).Address, ActiveSheet.Name
    '    GoTo exitsub
    'End If

    'If Intersect(Target, ActiveSheet.Range("INPUT_CLIENTINFO3")) Is Nothing Then
    'Else
    '    PopupTextBox Target.Rows(1).Address, ActiveSheet.Name
    '    GoTo exitsub
    'End If
    


    If Intersect(Target, ActiveSheet.Range("INPUT_CLIENT_NAME")) Is Nothing Then
    Else
        PopupListBox Target.Rows(1).Address, ActiveSheet.Name, "CLIENT_NAME", "CLIENT"
        ProcessSelection Target.Value
        GoTo exitsub
    End If
    

    If Intersect(Target, ActiveSheet.Range("INPUT_OPPORTUNITY_NAME")) Is Nothing Then
    Else
        PopupListBox Target.Rows(1).Address, ActiveSheet.Name, "LOOKUPS_OPPORTUNITY_NAME", "LOOKUPS"
        'ProcessSelection Target.Value
        GoTo exitsub
    End If
    
    If Intersect(Target, ActiveSheet.Range("INPUT_ATTENDEES1")) Is Nothing Then
    Else
        PopupListBox Target.Rows(1).Address, ActiveSheet.Name, "LOOKUPS_PERSON_FULL_NAME", "LOOKUPS"
        'ProcessSelection Target.Value
        GoTo exitsub
    End If
    
    If Intersect(Target, ActiveSheet.Range("INPUT_OUTCOME_DESCRIPTION")) Is Nothing Then
    Else
        PopupTextBox Target.Rows(1).Address, ActiveSheet.Name
        'PopupTextBox Target.Address, ActiveSheet.Name
        'ProcessSelection Target.Value
        GoTo exitsub
    End If
    
    If Intersect(Target, ActiveSheet.Range("INPUT_PURPOSE")) Is Nothing Then
    Else
        PopupTextBox Target.Rows(1).Address, ActiveSheet.Name
        'ProcessSelection Target.Value
        GoTo exitsub
    End If
    
    
    If Intersect(Target, ActiveSheet.Range("INPUT_OPPO_CONCERNS")) Is Nothing Then
    Else
        PopupTextBox Target.Rows(1).Address, ActiveSheet.Name
        'ProcessSelection Target.Value
        GoTo exitsub
    End If
    
    If Intersect(Target, ActiveSheet.Range("INPUT_NEXT_STEPS")) Is Nothing Then
    Else
        PopupTextBox Target.Rows(1).Address, ActiveSheet.Name
        GoTo exitsub
    End If
    
    If Intersect(Target, ActiveSheet.Range("INPUT_MONDAY_NAME1")) Is Nothing Then
    Else
        PopupListBox Target.Rows(1).Address, ActiveSheet.Name, "MONDAY_FULLNAME", "MONDAY_META", wideFlag:=True
        'ProcessSelection Target.Value
        GoTo exitsub
    End If
    
    If Intersect(Target, ActiveSheet.Range("INPUT_LAST_MINUTES1")) Is Nothing Then
    Else
        PopupListBox Target.Rows(1).Address, ActiveSheet.Name, "LOOKUPS_MEETING_DISPLAY_NAME", "MEETING_MINUTES"
        'ProcessSelection Target.Value
        GoTo exitsub
    End If
    
    
    
exitsub:
    Set FSO = Nothing

End Sub




