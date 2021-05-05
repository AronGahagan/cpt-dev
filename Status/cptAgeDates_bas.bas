Attribute VB_Name = "cptAgeDates_bas"
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub cptShowAgeDates_frm()
  'objects
  'strings
  'longs
  Dim lngControl As Long
  Dim lngWeek As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'todo: create and apply table
  'todo: create and apply filter?
  'todo: create and apply view
  'todo: update table dynamically

  With cptAgeDates_frm
    .cboWeeks.Clear
    For lngWeek = 1 To 10
      .cboWeeks.AddItem lngWeek & IIf(lngWeek = 1, " week", " weeks")
      '.Controls("cboWeek" & lngWeek).Clear
      For lngControl = 1 To 10
        .Controls("cboWeek" & lngControl).AddItem "Start" & lngWeek & "/Finish" & lngWeek
      Next lngControl
    Next lngWeek
    
    .cboWeeks = cptGetSetting("AgeDates", "cboWeeks")
    For lngControl = 1 To 10
      .Controls("cboWeek" & lngControl).Value = cptGetSetting("AgeDates", "cboWeek" & lngControl)
    Next lngControl
    .CheckBox1 = CBool(cptGetSetting("AgeDates", "chkIncludeDurations"))
    .CheckBox2 = CBool(cptGetSetting("AgeDates", "chkUpdateCustomFieldNames"))
    
    .Show False
  End With
  
  
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptAgeDates_bas", "cptShowAgeDates_frm", Err, Erl)
  Resume exit_here
End Sub

Sub cptAgeDates()
'run this immediately prior to a status meeting
  'objects
  Dim oTask As Task
  'strings
  Dim strStatus As String
  'longs
  Dim lngControl As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  Dim dtStatus As Date
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Application.Calculation = pjManual
  Application.OpenUndoTransaction "Age Dates"
  dtStatus = Format(ActiveProject.StatusDate, "mm/dd/yy")
  strStatus = cptGetSetting("AgeDates", "LastCapture")
  If Len(strStatus) > 0 Then
    If dtStatus = CDate(strStatus) Then
      If MsgBox("Already captured status for " & Format(dtStatus, "mm/dd/yyyy") & ": proceed anyway?", vbExclamation + vbYesNo, "Confirm Overwrite") = vbYes Then
        'todo: prompt to only overwrite the single period?
      Else
        GoTo exit_here
      End If
    End If
  End If
  
  'note: pjSaveBaselineFrom enumeration <> pjSaveBaselineTo enumeration
  BaselineSave True, 3, 3 'Start2/Finish2 >> Start3/Finish3
  BaselineSave True, 2, 2 'Start1/Finish1 >> Start2/Finish2
  BaselineSave True, 0, 1 'Start/Finish >> Start1/Finish1
  
  For Each oTask In ActiveProject.Tasks
    oTask.SetField pjTaskDuration3, oTask.GetField(pjTaskDuration2)
    oTask.SetField pjTaskDuration2, oTask.GetField(pjTaskDuration1)
    oTask.SetField pjTaskDuration1, oTask.DurationText
  Next
  
  CustomFieldRename pjCustomTaskStart1, "Start (" & CStr(dtStatus) & ")"
  CustomFieldRename pjCustomTaskDuration1, "Duration (" & CStr(dtStatus) & ")"
  CustomFieldRename pjCustomTaskFinish1, "Finish (" & CStr(dtStatus) & ")"
  
  CustomFieldRename pjCustomTaskStart2, "Start (" & CStr(DateAdd("d", -7, dtStatus)) & ")"
  CustomFieldRename pjCustomTaskDuration2, "Duration (" & CStr(DateAdd("d", -7, dtStatus)) & ")"
  CustomFieldRename pjCustomTaskFinish2, "Finish (" & CStr(DateAdd("d", -7, dtStatus)) & ")"
  
  CustomFieldRename pjCustomTaskStart3, "Start (" & CStr(DateAdd("d", -14, dtStatus)) & ")"
  CustomFieldRename pjCustomTaskDuration3, "Duration (" & CStr(DateAdd("d", -14, dtStatus)) & ")"
  CustomFieldRename pjCustomTaskFinish3, "Finish (" & CStr(DateAdd("d", -14, dtStatus)) & ")"
  
  'save settings
  With cptAgeDates_frm
    cptSaveSetting "AgeDates", "LastCapture", Format(dtStatus, "mm/dd/yyyy")
    cptSaveSetting "AgeDates", "cboWeeks", .cboWeeks.Value
    For lngControl = 1 To 10
      cptSaveSetting "AgeDates", "cboWeek" & lngControl, .Controls("cboWeek" & lngControl).Value
    Next lngControl
    cptSaveSetting "AgeDates", "chkIncludeDurations", .CheckBox1.Value
    cptSaveSetting "AgeDates", "chkUpdateCustomFieldNames", .CheckBox2.Value
  End With
  
exit_here:
  On Error Resume Next
  Application.CloseUndoTransaction
  Application.Calculation = pjAutomatic
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptAgeDates_bas", "cptAgeDates", Err, Erl)
  Resume exit_here
End Sub

