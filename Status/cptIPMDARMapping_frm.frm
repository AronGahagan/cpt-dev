VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptIPMDARMapping_frm 
   Caption         =   "IPMDAR Mapping"
   ClientHeight    =   6045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10125
   OleObjectBlob   =   "cptIPMDARMapping_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptIPMDARMapping_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v0.0.1</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Private Sub cboEnum_Change()
'objects
Dim aDefaults As Object
Dim aEnums As Object
Dim aFields As Object
'strings
Dim strCustomFieldName As String
'longs
Dim lngCol As Long
Dim lngType As Long
Dim lngCount As Long
Dim lngItem As Long
'integers
'doubles
'booleans
'variants
Dim vItem As Variant
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'set up array of field types and counts
  Set aFields = CreateObject("System.Collections.ArrayList")
  aFields.Add Array("Flag", 20)
  aFields.Add Array("Text", 30)
  'todo: other field types needed?
  'todo: add enterprise fields
  
  'define enums values for each enum
  Set aEnums = CreateObject("System.Collections.ArrayList")
  aEnums.Add Array("TaskTypeID*", Array("ACTIVITY", "MILESTONE", "SUMMARY", "HAMMOCK"))
  aEnums.Add Array("TaskSubTypeID", Array("RISK_MITIGATION_TASK", "SCHEDULE_VISIBILITY_TASK", "SCHEDULE_MARGIN", "CONTRACTUAL_MILESTONE"))
  aEnums.Add Array("TaskPlanningLevelID*", Array("SUMMARY_LEVEL_PLANNING_PACKAGE", "CONTROL_ACCOUNT", "PLANNING_PACKAGE", "WORK_PACKAGE", "ACTIVITY"))
  aEnums.Add Array("EarnedValueTechniqueID*", Array("APPORTIONED_EFFORT", "LEVEL_OF_EFFORT", "MILESTONE", "FIXED_0_100", "FIXED_100_0", "FIXED_X_Y", "PERCENT_COMPLETE", "STANDARDS", "UNITS", "OTHER_DISCRETE"))
  aEnums.Add Array("ElementOfCostID*", Array("LABOR", "MATERIAL", "OTHER_DIRECT_COSTS", "SUBCONTRACT"))
  
  'todo: alphabetize the custom field lists
  
  lngType = pjTask
  'reset form defaults
  Me.lblFieldType.Caption = "(Task Fields)"
  Me.cboWhereField.Clear
  Me.cboMapTo.ListWidth = "0 pt"
  Me.lblCOBRA.Visible = False
  Me.lblMPM.Visible = False

  Select Case Me.cboEnum
    Case "TaskTypeID*"
      
      'add task-specific items
      For Each vItem In Array("Summary", "Milestone", "Duration")
        With Me.cboWhereField
          .AddItem
          .List(.ListCount - 1, 0) = FieldNameToFieldConstant(vItem)
          .List(.ListCount - 1, 1) = vItem
        End With
      Next vItem
      
      'setup defaults
      Set aDefaults = CreateObject("System.Collections.ArrayList")
      aDefaults.Add Array("TaskTypeID", FieldNameToFieldConstant("Summary"), "Summary", "equals", "Yes", "SUMMARY")
      aDefaults.Add Array("TaskTypeID", FieldNameToFieldConstant("Milestone"), "Milestone", "equals", "Yes", "MILESTONE")
      aDefaults.Add Array("TaskTypeID", FieldNameToFieldConstant("Duration"), "Duration", "is greater than", 0, "ACTIVITY")
      'HAMMOCK not used in MSP
      
      'add defaults
      With Me.lboTaskTypeMap
        
        For lngItem = 0 To aDefaults.Count - 1
          .AddItem
          For lngCol = 0 To .ColumnCount - 1
            .List(.ListCount - 1, lngCol) = aDefaults(lngItem)(lngCol)
          Next lngCol
        Next lngItem
                      
      End With
    Case "TaskSubTypeID"
      
      'add task-specific items
      For Each vItem In Array("Name")
        With Me.cboWhereField
          .AddItem
          .List(.ListCount - 1, 0) = FieldNameToFieldConstant(vItem, lngType)
          .List(.ListCount - 1, 1) = vItem
        End With
      Next vItem
    
      Me.cboMapTo.ListWidth = 120
    Case "TaskPlanningLevelID*"
      Me.cboMapTo.ListWidth = 160
    Case "EarnedValueTechniqueID*"
      Me.cboMapTo.ListWidth = 120
      Me.lblCOBRA.Visible = True
      Me.lblMPM.Visible = True
    Case "ElementOfCostID*"
      
      Me.cboMapTo.ListWidth = 100
      lngType = pjResource
      Me.lblFieldType.Caption = "(Resource Fields)"
      
      'add resource-specific items
      For Each vItem In Array("Name", "Group", "Type") 'todo: other native resource fields?
        With Me.cboWhereField
          .AddItem
          .List(.ListCount - 1, 0) = FieldNameToFieldConstant(vItem, lngType)
          .List(.ListCount - 1, 1) = vItem
        End With
      Next vItem

  End Select
  
  'todo: add search? change cbo to find?
  
  'add common custom fields to cboWhereField
  With Me.cboWhereField
    For lngItem = 0 To aFields.Count - 1
      For lngCount = 1 To aFields(lngItem)(1)
        .AddItem
        .List(.ListCount - 1, 0) = FieldNameToFieldConstant(aFields(lngItem)(0) & lngCount, lngType)
        strCustomFieldName = CustomFieldGetName(FieldNameToFieldConstant(aFields(lngItem)(0) & lngCount, lngType))
        If Len(strCustomFieldName) > 0 Then
          .List(.ListCount - 1, 1) = strCustomFieldName & " (" & aFields(lngItem)(0) & lngCount & ")"
        Else
          .List(.ListCount - 1, 1) = aFields(lngItem)(0) & lngCount
        End If
      Next lngCount
    Next lngItem
  End With

  'update cboMapTo
  With Me.cboMapTo
    .Clear
    For lngItem = 0 To aEnums.Count - 1
      If aEnums(lngItem)(0) = Me.cboEnum.Value Then
        For lngCount = 0 To UBound(aEnums(lngItem)(1))
          .AddItem
          .List(.ListCount - 1, 0) = aEnums(lngItem)(1)(lngCount)
        Next
      End If
    Next lngItem
  End With
  
  Me.cboWhereField.ListIndex = 0
  
exit_here:
  On Error Resume Next
  Set aDefaults = Nothing
  Set aEnums = Nothing
  Set aFields = Nothing

  Exit Sub
err_here:
  'Call HandleErr("cptIPMDARMapping_frm", "cboFieldToMap_Change", Err)
  MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Error"
  Resume exit_here
End Sub

Sub cboOperator_Change()
'objects
Dim aCriteria As Object
Dim oCollection As Object
Dim oItem As Object
'strings
Dim strValue As String
'longs
Dim lngItem As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'set oItem and oCollection
  If Me.lblFieldType = "(Task Fields)" Then
    Set oCollection = ActiveProject.Tasks
  ElseIf Me.lblFieldType = "(Resource Fields)" Then
    Set oCollection = ActiveProject.Resources
  End If
  'todo: set type (task or resource)
  'todo: get lngField
  
  Select Case Me.cboOperator
    Case "equals"
      Me.txtCriteria.Visible = False
      Me.cboCriteria.Visible = True
      Me.cboCriteria.Clear
      If Not IsNull(Me.cboWhereField) Then
        'set up array to store values
        Set aCriteria = CreateObject("System.Collections.SortedList")
        'todo: get actual existing values
        For Each oItem In oCollection
          strValue = oItem.GetField(Me.cboWhereField.Value)
          If Not aCriteria.Contains(strValue) Then
            aCriteria.Add strValue, strValue
          End If
        Next oItem
        With Me.cboCriteria
          For lngItem = 0 To aCriteria.Count - 1
            .AddItem aCriteria.getByIndex(lngItem)
          Next lngItem
        End With
      End If
      Me.cboCriteria.SetFocus
      'Me.cboCriteria.DropDown
    Case "contains"
      Me.cboCriteria.Visible = False
      Me.txtCriteria.Visible = True
  End Select
    
exit_here:
  On Error Resume Next
  Set aCriteria = Nothing
  Set oCollection = Nothing
  Set oItem = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptIPMDARMapping_frm", "cboOperator_Change", Err, Erl)
  Resume exit_here
End Sub

Private Sub cboWhereField_Change()
  If IsNull(Me.cboOperator.Value) Then Me.cboOperator.Value = "equals"
  Me.cboOperator_Change
End Sub

Private Sub cdRemove_Click()
Dim lngItem As Long
  
  With Me.lboTaskTypeMap
    For lngItem = .ListCount - 1 To 0 Step -1
      If .Selected(lngItem) Then .RemoveItem (lngItem)
    Next lngItem
  End With
  
End Sub

Private Sub cmdAddMap_Click()
'objects
'strings
'longs
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  With Me.lboTaskTypeMap
    .AddItem
    'todo: confirm required
    .List(.ListCount - 1, 0) = Me.cboEnum.Value
    .List(.ListCount - 1, 1) = Me.cboWhereField.Value
    .List(.ListCount - 1, 2) = Me.cboWhereField.List(Me.cboWhereField.ListIndex, 1)
    .List(.ListCount - 1, 3) = Me.cboOperator.Value
    .List(.ListCount - 1, 4) = Me.txtCriteria.Value
    .List(.ListCount - 1, 5) = Me.cboMapTo.Value
    'todo: avoid conflicts
  End With

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptIPMDARMapping_frm", "cmdAddMap_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub lblCOBRA_Click()
  Call cptAddEVTDefaults("COBRA")
End Sub

Private Sub lblMPM_Click()
  Call cptAddEVTDefaults("MPM")
End Sub
