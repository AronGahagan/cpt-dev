Attribute VB_Name = "cptDynamicFilter_bas"
'<cpt_version>v1.0</cpt_version>
Sub ShowcptDynamicFilter_frm()
  ActiveWindow.TopPane.Activate
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
    .cboOperator = "contains"
    .chkKeepSelected = False
    .chkHideSummaries = False
    .chkShowRelatedSummaries = False
    .chkHighlight = False
    .Show False
    .txtFilter.SetFocus
  End With
End Sub
