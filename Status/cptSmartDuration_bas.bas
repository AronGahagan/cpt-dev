Attribute VB_Name = "cptSmartDuration_bas"
'<cpt_version>v2.0.2</cpt_version>

Sub cptShowSmartDuration_frm()
  'objects
  'strings
  Dim strKeepOpen As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Call cptUpdateSmartDurationForm
  With cptSmartDuration_frm
    .Caption = "Smart Duration (" & cptGetVersion("cptSmartDuration_frm") & ")"
    strKeepOpen = cptGetSetting("SmartDuration", "chkKeepOpen")
    If Len(strKeepOpen) = 0 Then
      .chkKeepOpen = False 'default to false
    Else
      .chkKeepOpen = CBool(strKeepOpen)
    End If
    .Show False
    If .txtTargetFinish.Enabled Then .txtTargetFinish.SetFocus
  End With
  
  cptCore_bas.cptStartEvents

exit_here:
  On Error Resume Next

  Exit Sub

err_here:
  Call cptHandleErr("cptSmartDuration_bas", "cptShowSmartDuration_frm", err, Erl)
  Resume exit_here
End Sub

Sub cptUpdateSmartDurationForm()
  'objects
  Dim oTasks As MSProject.Tasks
  'strings
  'longs
  'integers
  'doubles
  'booleans
  Dim blnValid As Boolean
  'variants
  'dates
  
  On Error Resume Next
  Set oTasks = Nothing
  Set oTasks = ActiveSelection.Tasks
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  blnValid = True
  If oTasks Is Nothing Then
    blnValid = False
  ElseIf oTasks.Count = 0 Then 'Group By Summary
    cptSmartDuration_frm.txtTargetFinish = "-"
    cptSmartDuration_frm.lblWeekday.Caption = "< invalid >"
    cptSmartDuration_frm.lblWeekday.ControlTipText = "Cannot adjust Group By Summary tasks."
    blnValid = False
  ElseIf oTasks(1) Is Nothing Then 'newly inserted task
    cptSmartDuration_frm.txtTargetFinish = "-"
    cptSmartDuration_frm.lblWeekday.Caption = "< invalid >"
    cptSmartDuration_frm.lblWeekday.ControlTipText = "Cannot adjust newly inserted tasks."
    blnValid = False
  ElseIf oTasks.Count > 1 Then 'too many
    cptSmartDuration_frm.txtTargetFinish = ""
    cptSmartDuration_frm.lblWeekday.Caption = "< focus >"
    cptSmartDuration_frm.lblWeekday.ControlTipText = "Please select a single task."
    blnValid = False
  ElseIf oTasks(1).Summary Then
    cptSmartDuration_frm.txtTargetFinish = ""
    cptSmartDuration_frm.lblWeekday.Caption = "< summary >"
    cptSmartDuration_frm.lblWeekday.ControlTipText = "Please select a Non-summary task."
    blnValid = False
  ElseIf IsDate(oTasks(1).ActualFinish) Then
    cptSmartDuration_frm.txtTargetFinish = ""
    cptSmartDuration_frm.lblWeekday.Caption = "< complete >"
    cptSmartDuration_frm.lblWeekday.ControlTipText = "Please select an incomplete task."
    blnValid = False
  End If
  
  With cptSmartDuration_frm
    If blnValid Then
      .lngUID = oTasks(1).UniqueID
      .StartDate = oTasks(1).Start
      .txtTargetFinish = FormatDateTime(oTasks(1).Finish, vbShortDate)
      .lblWeekday.Caption = Format(.txtTargetFinish.Text, "dddd")
      .lblWeekday.ControlTipText = ""
      .txtTargetFinish.Enabled = True
      '.txtTargetFinish.SetFocus 'this steals focus when user may not want it to
      .cmdApply.Enabled = True
    Else
      .txtTargetFinish.Enabled = False
      .cmdApply.Enabled = False
    End If
  End With

exit_here:
  On Error Resume Next
  Set oTasks = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptSmartDuration_bas", "cptUpdateSmartDurationForm", err, Erl)
  Resume exit_here
End Sub

