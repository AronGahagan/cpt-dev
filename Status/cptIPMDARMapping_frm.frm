VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptTaskTypeMapping_frm 
   Caption         =   "TaskTypeID Mapping"
   ClientHeight    =   6045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10125
   OleObjectBlob   =   "cptTaskTypeMapping_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptTaskTypeMapping_frm"
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
Dim aEnums As Object
Dim aFields As Object
'strings
Dim strCustomFieldName As String
'longs
Dim lngType As Long
Dim lngCount As Long
Dim lngItem As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Set aFields = CreateObject("System.Collections.ArrayList")
  aFields.Add Array("Flag", 20)
  aFields.Add Array("Text", 30)
  
  'todo: others needed?
  'todo: add enterprise fields?
  
  'define enums values for each enum
  Set aEnums = CreateObject("System.Collections.ArrayList")
  aEnums.Add Array("TaskTypeID*", Array("ACTIVITY", "MILESTONE", "SUMMARY", "HAMMOCK"))
  aEnums.Add Array("TaskSubTypeID", Array("RISK_MITIGATION_TASK", "SCHEDULE_VISIBILITY_TASK", "SCHEDULE_MARGIN", "CONTRACTUAL_MILESTONE"))
  aEnums.Add Array("TaskPlanningLevelID*", Array("SUMMARY_LEVEL_PLANNING_PACKAGE", "CONTROL_ACCOUNT", "PLANNING_PACKAGE", "WORK_PACKAGE", "ACTIVITY"))
  aEnums.Add Array("EarnedValueTechniqueID*", Array("APPORTIONED_EFFORT", "LEVEL_OF_EFFORT", "MILESTONE", "FIXED_0_100", "FIXED_100_0", "FIXED_X_Y", "PERCENT_COMPLETE", "STANDARDS", "UNITS", "OTHER_DISCRETE"))
  aEnums.Add Array("ElementOfCostID*", Array("LABOR", "MATERIAL", "OTHER_DIRECT_COSTS", "SUBCONTRACT"))
  
  'todo: alphabetize the custom field lists
  
  lngType = pjTask
  Me.lblFieldType.Caption = "(Task Fields)"
  
  Select Case Me.cboEnum
    Case "TaskTypeID*"
      'todo: setup defaults
      
    Case "TaskSubTypeID"
    Case "TaskPlanningLevelID*"
    Case "EarnedValueTechniqueID*"
    Case "ElementOfCostID*"
      lngType = pjResource
      Me.lblFieldType.Caption = "(Resource Fields)"
      'update cboWhereField
      With Me.cboWhereField
        .Clear
        .AddItem
        .List(.ListCount - 1, 0) = pjResourceType
        .List(.ListCount - 1, 1) = "[Resource]Type"
      End With
    
  End Select

  'update cboWhereField
  For lngItem = 0 To aFields.Count - 1
    For lngCount = 1 To aFields(lngItem)(1)
      .AddItem
      .List(.ListCount - 1, 0) = FieldNameToFieldConstant(aFields(lngItem)(0) & lngCount, pjResource)
      strCustomFieldName = CustomFieldGetName(FieldNameToFieldConstant(aFields(lngItem)(0) & lngCount, pjResource))
      If Len(strCustomFieldName) > 0 Then
        .List(.ListCount - 1, 1) = strCustomFieldName & " ([Resource]" & aFields(lngItem)(0) & lngCount & ")"
      Else
        .List(.ListCount - 1, 1) = "[Resource]" & aFields(lngItem)(0) & lngCount
      End If
    Next lngCount
  Next lngItem
  

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
  
exit_here:
  On Error Resume Next
  Set aEnums = Nothing
  Set aFields = Nothing

  Exit Sub
err_here:
  'Call HandleErr("cptTaskTypeMapping_frm", "cboFieldToMap_Change", Err)
  MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Error"
  Resume exit_here
End Sub

Private Sub cboOperator_Change()
'objects
'strings
'longs
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'todo: set oItem and oCollection
  'todo: set type (task or resource)
  'todo: get lngField
  
  Select Case Me.cboOperator
    Case "equals"
      'todo: get actual existing values
      
    Case "contains"
      'hide cbo show txt

  End Select
  
  'can't believe I just lost all of this.
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptTaskTypeMapping_frm", "cboOperator_Change", Err, Erl)
  MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Error"
  Resume exit_here
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

  With Me.lboPreview
    .AddItem
    'todo: confirm required
    .List(.ListCount - 1, 0) = Me.cboEnum.Value
    .List(.ListCount - 1, 1) = Me.cboWhereField.Value
    .List(.ListCount - 1, 2) = Me.cboWhereField.List(.ListIndex, 1)
    .List(.ListCount - 1, 3) = Me.cboOperator.Value
    .List(.ListCount - 1, 4) = Me.txtCriteria.Value
    .List(.ListCount - 1, 5) = Me.cboMapTo.Value
    'todo: avoid conflicts
  End With

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptTaskTypeMapping_frm", "cmdAddMap_Click", Err, Erl)
  Resume exit_here
End Sub
