Attribute VB_Name = "cptAgeDates_bas"
'<cpt_version>v1.0.0</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub cptShowAgeDates_frm()
  'objects
  'strings
  Dim strSetting As String
  'longs
  Dim lngFF As Long
  Dim lngFS As Long
  Dim lngAF As Long
  Dim lngAS As Long
  Dim lngControl As Long
  Dim lngWeek As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "Status Date required.", vbExclamation + vbOKOnly, "Age Dates"
    Application.ChangeStatusDate
  End If
  
  'todo: create and apply table
  'todo: create and apply filter?
  'todo: create and apply view
  'todo: update table dynamically
  'todo: multiple settings for multiple projects >> convert to ADODB
  
  'todo: avoid cpt conflicts with status sheet import
  strSetting = cptGetSetting("StatusSheetImport", "cboAS")
  If Len(strSetting) > 0 Then lngAS = CLng(strSetting) Else lngAS = 0
  strSetting = cptGetSetting("StatusSheetImport", "cboAF")
  If Len(strSetting) > 0 Then lngAF = CLng(strSetting) Else lngAF = 0
  strSetting = cptGetSetting("StatusSheetImport", "cboFS")
  If Len(strSetting) > 0 Then lngFS = CLng(strSetting) Else lngFS = 0
  strSetting = cptGetSetting("StatusSheetImport", "cboFF")
  If Len(strSetting) > 0 Then lngFF = CLng(strSetting) Else lngFF = 0
  
  With cptAgeDates_frm
    .lblStatus = "(" & FormatDateTime(ActiveProject.StatusDate, vbShortDate) & ")"
    .cboWeeks.Clear
    For lngWeek = 1 To 10
      .cboWeeks.AddItem lngWeek & IIf(lngWeek = 1, " week", " weeks")
      For lngControl = 1 To 10
        With .Controls("cboWeek" & lngControl)
          .AddItem
          .List(.ListCount - 1, 0) = lngWeek + 1
          .List(.ListCount - 1, 1) = lngWeek
          .List(.ListCount - 1, 2) = "Start" & lngWeek & "/Finish" & lngWeek
        End With
      Next lngControl
    Next lngWeek
    
    strSetting = cptGetSetting("AgeDates", "cboWeeks")
    If Len(strSetting) > 0 Then
      .cboWeeks.Value = strSetting
    Else
      .cboWeeks.Value = "3 weeks"
    End If
    For lngControl = 1 To 10
      strSetting = cptGetSetting("AgeDates", "cboWeek" & lngControl)
      If Len(strSetting) > 0 Then
        .Controls("cboWeek" & lngControl).Value = cptGetSetting("AgeDates", "cboWeek" & lngControl)
      End If
    Next lngControl
    strSetting = cptGetSetting("AgeDates", "chkIncludeDurations")
    If Len(strSetting) > 0 Then .chkIncludeDurations = CBool(strSetting)
    strSetting = cptGetSetting("AgeDates", "chkUpdateCustomFieldNames")
    If Len(strSetting) > 0 Then .chkUpdateCustomFieldNames = CBool(strSetting)
    
    .Show False 'False
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
  Dim strCustom As String
  Dim strStatus As String
  'longs
  Dim lngTest As Long
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
  
  On Error Resume Next
  lngTest = FieldNameToFieldConstant("Start (" & FormatDateTime(ActiveProject.StatusDate, vbShortDate) & ")")
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If lngTest > 0 Then
    MsgBox "Dates already aged for status date " & Format(dtStatus, "mm/dd/yyyy") & ".", vbExclamation + vbOKOnly, "Age Dates"
    GoTo exit_here
  End If

  With cptAgeDates_frm
    
    For lngControl = 10 To 1 Step -1
      If .Controls("cboWeek" & lngControl).Enabled Then
        If lngControl = 1 Then
          lngFrom = 0
        Else
          lngFrom = .Controls("cboWeek" & lngControl - 1).List(.Controls("cboWeek" & lngControl - 1).ListIndex, 0)
        End If
        lngTo = .Controls("cboWeek" & lngControl).List(.Controls("cboWeek" & lngControl).ListIndex, 1)
        BaselineSave True, lngFrom, lngTo
        'update custom field names
        If .chkUpdateCustomFieldNames Then
          If lngControl = 1 Then
            strCustom = "Start (" & FormatDateTime(dtStatus, vbShortDate) & ")"
            CustomFieldRename FieldNameToFieldConstant("Start" & lngControl), strCustom
            strCustom = "Finish (" & FormatDateTime(dtStatus, vbShortDate) & ")"
            CustomFieldRename FieldNameToFieldConstant("Finish" & lngControl), strCustom
            If .chkIncludeDurations Then
              strCustom = "Duration (" & FormatDateTime(dtStatus, vbShortDate) & ")"
              CustomFieldRename FieldNameToFieldConstant("Duration" & lngControl), strCustom
            End If
          Else
            strCustom = CustomFieldGetName(FieldNameToFieldConstant("Start" & lngControl - 1, pjTask))
            CustomFieldRename FieldNameToFieldConstant("Start" & lngControl - 1), ""
            CustomFieldRename FieldNameToFieldConstant("Start" & lngControl), strCustom
            strCustom = CustomFieldGetName(FieldNameToFieldConstant("Finish" & lngControl - 1, pjTask))
            CustomFieldRename FieldNameToFieldConstant("Finish" & lngControl - 1), ""
            CustomFieldRename FieldNameToFieldConstant("Finish" & lngControl), strCustom
            If .chkIncludeDurations Then
              strCustom = CustomFieldGetName(FieldNameToFieldConstant("Duration" & lngControl - 1, pjTask))
              CustomFieldRename FieldNameToFieldConstant("Duration" & lngControl - 1), ""
              CustomFieldRename FieldNameToFieldConstant("Duration" & lngControl), strCustom
            End If
          End If
        End If
      End If
    Next lngControl
    
    If .chkIncludeDurations Then
      For Each oTask In ActiveProject.Tasks
        For lngControl = 10 To 1 Step -1
          If .Controls("cboWeek" & lngControl).Enabled Then
            lngTo = cptRegEx(.Controls("cboWeek" & lngControl).List(.Controls("cboWeek" & lngControl).ListIndex, 1), "[0-9]")
            If lngControl = 1 Then
              oTask.SetField FieldNameToFieldConstant("Duration" & lngTo), oTask.DurationText
            Else
              lngFrom = cptRegEx(.Controls("cboWeek" & lngControl - 1).List(.Controls("cboWeek" & lngControl - 1).ListIndex, 2), "[0-9]")
              oTask.SetField FieldNameToFieldConstant("Duration" & lngTo), oTask.GetField(FieldNameToFieldConstant("Duration" & lngFrom))
            End If
          End If
        Next lngControl
      Next oTask
    End If
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
