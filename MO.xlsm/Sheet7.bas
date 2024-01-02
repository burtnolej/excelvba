VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal target As Range)
Dim boardRangeString As String: boardRangeString = "ADDITEM_BOARD_NAME"
Dim groupRangeString As String: groupRangeString = "ADD_ITEM_GROUP_NAMES"
Dim itemRangeString As String: itemRangeString = "ADD_ITEM_ITEM_NAMES"
Dim subItemRangeString As String: subItemRangeString = "ADD_ITEM_SUBITEM_NAMES"


Dim boardRange As Range, dropDownTarget As Range, groupRange As Range, itemRange As Range, subItemRange As Range

    If target.Rows.Count > 1 Then
        Exit Sub
    End If

    If target.Columns.Count > 1 Then
        Exit Sub
    End If
    
    If IsError(target.value) = True Then
        Exit Sub
    End If
    
    If target.value = "" Then
        Exit Sub
    End If
    
    'SetEventsOff
    
    Set boardRange = ActiveSheet.Range(boardRangeString)
    Set groupRange = ActiveSheet.Range(groupRangeString)
    Set itemRange = ActiveSheet.Range(itemRangeString)
    Set subItemRange = ActiveSheet.Range(subItemRangeString)
    
    If Not Intersect(boardRange, target) Is Nothing Then
        
        ActiveSheet.Range("SELECT_BOARD").value = target.value
        'Set dropDownTarget = target.Offset(, 1)
        Set dropDownTarget = target.Offset(1)
        CreateGroupNameDropdown dropDownTarget, "SELECT_GROUP_NAMES", "AddNewItems"
        
        
    ElseIf Not Intersect(groupRange, target) Is Nothing Then
        ActiveSheet.Range("SELECT_GROUP").value = target.value
        'Set dropDownTarget = target.Offset(, 3)
        Set dropDownTarget = target.Offset(3)
        CreateGroupNameDropdown dropDownTarget, "SELECT_ITEM_NAMES", "AddNewItems"
        
        'set new item name to blank
        ActiveSheet.Range("NEWITEM_NEWITEM_NAME").Rows(target.Row - 3).value = ""
        
        
    ElseIf Not Intersect(itemRange, target) Is Nothing Then
        ActiveSheet.Range("SELECT_ITEMS").value = target.value
        'Set dropDownTarget = target.Offset(, 3)
        Set dropDownTarget = target.Offset(3)
        CreateGroupNameDropdown dropDownTarget, "SELECT_SUBITEM_NAMES", "AddNewItems"

        'set new item name to N/A as an existing item has been added
        ActiveSheet.Range("NEWITEM_NEWITEM_NAME").Rows(target.Row - 3).value = "N/A"
        
        'set new sub item name to blank
        ActiveSheet.Range("NEWSUBITEM_NEWSUBITEM_NAME").Rows(target.Row - 3).value = ""
        
        'set new item is to blank
        ActiveSheet.Range("NEWITEM_ADDEDITEMID").Rows(target.Row - 3).value = ""
        
    ElseIf Not Intersect(subItemRange, target) Is Nothing Then
        
        'set new sub item name to blank
        ActiveSheet.Range("NEWSUBITEM_NEWSUBITEM_NAME").Rows(target.Row - 3).value = "N/A"
    End If
    
    'SetEventsOn
End Sub

