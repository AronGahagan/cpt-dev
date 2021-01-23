Attribute VB_Name = "cptDynamicFilter_bas"
'<cpt_version>v1.2.4</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub cptShowDynamicFilter_frm()
'objects
'strings
'longs
'integers
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  '===
  'Validate users selected view type
  If ActiveProject.Application.ActiveWindow.ActivePane.View.Type <> pjTaskItem Then
    MsgBox "Please select a View with a Task Table.", vbInformation + vbOKOnly, "Dynamic Filter"
    GoTo exit_here
  End If
  'Validate users selected window pane - select the task table if not active
  If ActiveProject.Application.ActiveWindow.ActivePane.Index <> 1 Then
    ActiveProject.Application.ActiveWindow.TopPane.Activate
  End If
  '===
  
  With cptDynamicFilter_frm
    .txtFilter = ""
    With .cboField
      .Clear
      .AddItem "Task Name"
      '.AddItem "Work Package"
      '.AddItem "CAM"
      '.AddItem "WPM"
    End With
    With .cboOperator
      .Clear
      .AddItem "equals"
      .AddItem "does not equal"
      .AddItem "contains"
      .AddItem "does not contain"
    End With
    .cboField = "Task Name"
    .cboOperator = GetSetting("ClearPlanToolbar", "DynamicFilter", "Operator")
    .chkKeepSelected = GetSetting("ClearPlanToolbar", "DynamicFilter", "KeepSelected") = "1"
    .chkHideSummaries = GetSetting("ClearPlanToolbar", "DynamicFilter", "IncludeSummaries") = "1"
    .chkShowRelatedSummaries = GetSetting("ClearPlanToolbar", "DynamicFilter", "RelatedSummaries") = "1"
    .chkHighlight = GetSetting("ClearPlanToolbar", "DynamicFilter", "Highlight") = "1"
    .Show False
    .txtFilter.SetFocus
  End With
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptDynamicFilter_bas", "cptShowDynamicFilter_frm", Err, Erl)
  Resume exit_here
End Sub
