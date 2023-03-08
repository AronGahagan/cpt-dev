Attribute VB_Name = "cptSmartDuration_bas"
'<cpt_version>v2.0.0</cpt_version>

Sub cptShowSmartDuration_frm()
  'objects
  Dim oTasks As MSProject.Tasks
  'strings
  Dim strKeepOpen As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates

  On Error Resume Next
  Set oTasks = ActiveSelection.Tasks
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oTasks Is Nothing Then GoTo exit_here
  
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
    .txtTargetFinish.SetFocus
  End With
  
  cptCore_bas.cptStartEvents

exit_here:
  On Error Resume Next
  Set oTasks = Nothing
  Exit Sub

err_here:
  Call cptHandleErr("cptSmartDuration_bas", "cptShowSmartDuration_frm", Err, Erl)
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
  'variants
  'dates
  
  On Error Resume Next
  Set oTasks = ActiveSelection.Tasks
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  If oTasks Is Nothing Then GoTo exit_here
  If oTasks.Count > 1 Then
    cptSmartDuration_frm.txtTargetFinish = ""
    cptSmartDuration_frm.lblWeekday.Caption = "< focus >"
    cptSmartDuration_frm.lblWeekday.ControlTipText = "Please select a single task."
    GoTo exit_here
  ElseIf oTasks.Count = 0 Then
    cptSmartDuration_frm.txtTargetFinish = "-"
    GoTo exit_here
  End If
  
  If oTasks(1).Summary Then
    cptSmartDuration_frm.txtTargetFinish = ""
    cptSmartDuration_frm.lblWeekday.Caption = "< summary >"
    cptSmartDuration_frm.lblWeekday.ControlTipText = "Please select a Non-summary task."
    GoTo exit_here
  End If
  
  If IsDate(oTasks(1).ActualFinish) Then
    cptSmartDuration_frm.txtTargetFinish = ""
    cptSmartDuration_frm.lblWeekday.Caption = "< complete >"
    cptSmartDuration_frm.lblWeekday.ControlTipText = "Please select an incomplete task."
    GoTo exit_here
  End If
  
  If oTasks(1).Milestone Or oTasks(1).Duration = 0 Then
    If MsgBox("Proceed with editing a zero-duration milestone?", vbQuestion + vbYesNo, "Please confirm") = vbNo Then
      GoTo exit_here
    End If
  End If
  
  With cptSmartDuration_frm
    .lngUID = oTasks(1).UniqueID
    .StartDate = oTasks(1).Start
    .txtTargetFinish = FormatDateTime(oTasks(1).Finish, vbShortDate)
    .lblWeekday.Caption = Format(.txtTargetFinish.Text, "dddd")
    .lblWeekday.ControlTipText = ""
  End With

exit_here:
  On Error Resume Next
  Set oTasks = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptSmartDuration_bas", "cptUpdateSmartDurationForm", Err, Erl)
  Resume exit_here
End Sub

