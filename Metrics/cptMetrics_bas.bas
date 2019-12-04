Attribute VB_Name = "cptMetrics_bas"
'<cpt_version>v1.0.7</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

'add disclaimer: unburdened hours - not meant to be precise - generally within +/- 1%

Sub cptExportMetricsExcel()
'objects
'strings
'longs
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  MsgBox "Stay tuned...", vbInformation + vbOKOnly, "Under Construction..."

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptMetrics_bas", "cptExportMetricsExcel", Err, Erl)
  Resume exit_here
End Sub

Sub cptGetBAC()
  MsgBox Format(cptGetMetric("bac"), "#,##0.00h"), vbInformation + vbOKOnly, "Budget at Complete (BAC) - hours"
End Sub

Sub cptGetETC()
  MsgBox Format(cptGetMetric("etc"), "#,##0.00h"), vbInformation + vbOKOnly, "Estimate to Complete (ETC) - hours"
End Sub

Sub cptGetBCWS()

  'confirm status date
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "This project has no status date. Please update and try again.", vbExclamation + vbOKOnly, "Metrics"
    Exit Sub
  Else
    MsgBox Format(cptGetMetric("bcws"), "#,##0.00"), vbInformation + vbOKOnly, "Budgeted Cost of Work Scheduled (BCWS) - hours"
  End If

End Sub

Sub cptGetBCWP()
  
  If Not cptMetricsSettingsExist Then
    Call cptShowMetricsSettings_frm(True)
    If Not cptMetricsSettingsExist Then
      MsgBox "No settings saved. Cannot proceed.", vbExclamation + vbOKOnly, "Settings required."
      Exit Sub
    End If
  End If
  
  'confirm status date
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "This project has no status date. Please update and try again.", vbExclamation + vbOKOnly, "Metrics"
    Exit Sub
  Else
    MsgBox Format(cptGetMetric("bcwp"), "#,##0.00"), vbInformation + vbOKOnly, "Budgeted Cost of Work Performed (BCWP) - hours"
  End If
  
End Sub

Sub cptGetSPI()
  
  If Not cptMetricsSettingsExist Then
    Call cptShowMetricsSettings_frm(True)
    If Not cptMetricsSettingsExist Then
      MsgBox "No settings saved. Cannot proceed.", vbExclamation + vbOKOnly, "Settings required."
      Exit Sub
    End If
  End If
  
  'confirm status date
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "This project has no status date. Please update and try again.", vbExclamation + vbOKOnly, "Metrics"
    Exit Sub
  Else
    Call cptGET("SPI")
  End If
  
End Sub

Sub cptGetBEI()

  'confirm status date
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "This project has no status date. Please update and try again.", vbExclamation + vbOKOnly, "Metrics"
    Exit Sub
  Else
    Call cptGET("BEI")
  End If
  
End Sub

Sub cptGetCEI()
  
  'confirm status date
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "This project has no status date. Please update and try again.", vbExclamation + vbOKOnly, "Metrics"
    Exit Sub
  Else
    Call cptGET("CEI")
  End If
  
End Sub

Sub cptGetSV()
  
  If Not cptMetricsSettingsExist Then
    Call cptShowMetricsSettings_frm(True)
    If Not cptMetricsSettingsExist Then
      MsgBox "No settings saved. Cannot proceed.", vbExclamation + vbOKOnly, "Settings required."
      Exit Sub
    End If
  End If
  
  'confirm status date
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "This project has no status date. Please update and try again.", vbExclamation + vbOKOnly, "Metrics"
    Exit Sub
  Else
    Call cptGET("SV")
  End If

End Sub

Sub cptGetCPLI()
'objects
Dim oTasks As Tasks
Dim oPred As Task
Dim oTask As Task
'strings
Dim strMsg As String
Dim strTitle As String
'longs
Dim lngConstraintType As Long
Dim lngTS As Long
Dim lngMargin As Long
Dim lngCPL As Long
'integers
'doubles
'booleans
'variants
'dates
Dim dtStart As Date, dtFinish As Date
Dim dtConstraintDate As Date

  strTitle = "Critical Path Length Index (CPLI)"

  On Error Resume Next
  Set oTasks = ActiveSelection.Tasks
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If oTasks Is Nothing Then
    MsgBox "No Target Task selected.", vbExclamation + vbOKOnly, "Oops"
    GoTo exit_here
  End If

  'confirm a single, target oTask is selected
  If oTasks.Count <> 1 Then
    MsgBox "Please select a single, active, and non-summary target oTask.", vbExclamation + vbOKOnly, strTitle
    GoTo exit_here
  End If
  
  Set oTask = oTasks(1)
  
  'use MFO or MSO constraint
  If oTask.ConstraintType <> pjMFO And oTask.ConstraintType <> pjMSO Then
    strMsg = "No MSO/MFO constraint found; temporarily using Deadline..." & vbCrLf
    'if no MFO then use deadline as MFO
    If IsDate(oTask.Deadline) Then
      If IsDate(oTask.ConstraintDate) Then dtConstraintDate = oTask.ConstraintDate
      lngConstraintType = oTask.ConstraintType
      oTask.ConstraintDate = oTask.Deadline
      oTask.ConstraintType = pjMFO
      lngTS = oTask.TotalSlack
      dtFinish = oTask.Finish
      If CLng(dtConstraintDate) > 0 Then oTask.ConstraintDate = dtConstraintDate
      oTask.ConstraintType = lngConstraintType
    Else
      strMsg = strMsg & "No Deadline found; temporarily using Baseline Finish..." & vbCrLf
      If Not IsDate(oTask.BaselineFinish) Then
        strMsg = strMsg & "No Baseline Finish found." & vbCrLf & vbCrLf
        strMsg = strMsg & "In order to calculate the CPLI, the target Task should be (at least temporarily) constrained with a MFO or Deadline." & vbCrLf & vbCrLf
        strMsg = strMsg & "Please constrain the Target Task and try again."
        MsgBox strMsg, vbExclamation + vbOKOnly, strTitle
        GoTo exit_here
      Else
        If IsDate(oTask.ConstraintDate) Then dtConstraintDate = oTask.ConstraintDate
        lngConstraintType = oTask.ConstraintType
        oTask.ConstraintDate = oTask.BaselineFinish
        oTask.ConstraintType = pjMFO
        lngTS = oTask.TotalSlack
        dtFinish = oTask.Finish
        If CLng(dtConstraintDate) > 0 Then oTask.ConstraintDate = dtConstraintDate
        oTask.ConstraintType = lngConstraintType
      End If
    End If
  Else
    lngTS = oTask.TotalSlack
    dtFinish = oTask.Finish
  End If
      
  'use status date if exists
  If IsDate(ActiveProject.StatusDate) Then
    dtStart = ActiveProject.StatusDate
  Else
    dtStart = FormatDateTime(Now(), vbShortDate) & " 08:00 AM"
  End If
  
  'use earliest start date
  'NOTE: cannot account for schedule margin due to possibility
  'of dual paths, one with and one without, a particular SM Task
  
  If oTask Is Nothing Then GoTo exit_here
  If oTask.Summary Then GoTo exit_here
  If Not oTask.Active Then GoTo exit_here
  HighlightDrivingPredecessors Set:=True
  For Each oPred In ActiveProject.Tasks
    If oPred.PathDrivingPredecessor Then
      If IsDate(oPred.ActualStart) Then
        If oPred.Stop < dtStart Then dtStart = oPred.Stop
      Else
        If oPred.Start < dtStart Then dtStart = oPred.Start
      End If
    End If
  Next oPred
  'calculate the CPL
  lngCPL = Application.DateDifference(dtStart, dtFinish)
  'convert values to days
  lngCPL = lngCPL / 480
  lngTS = lngTS / 480
  'notify user
  strMsg = strMsg & vbCrLf & "CPL = Critical Path Length" & vbCrLf
'  strMsg = strMsg & "CPL = Target Finish - Timenow (or CP start)" & vbCrLf
'  strMsg = strMsg & "CPL = " & FormatDateTime(dtFinish, vbShortDate) & " - " & FormatDateTime(dtStart, vbShortDate) & vbCrLf
'  strMsg = strMsg & "CPL = " & lngCPL & vbCrLf
  strMsg = strMsg & "TS = Total Slack" & vbCrLf & vbCrLf
  strMsg = strMsg & "CPLI = ( CPL + TS ) / CPL" & vbCrLf
  strMsg = strMsg & "CPLI = ( " & lngCPL & " + " & lngTS & " ) / " & lngCPL & vbCrLf & vbCrLf
  strMsg = strMsg & "CPLI = " & Round((lngCPL + lngTS) / lngCPL, 3) & vbCrLf & vbCrLf
  strMsg = strMsg & "Note: Schedule Margin Tasks are not considered."
  
  MsgBox strMsg, vbInformation + vbOKOnly, "Critical Path Length Index (CPLI)"
    
exit_here:
  On Error Resume Next
  Set oTasks = Nothing
  Set oPred = Nothing
  Application.CloseUndoTransaction
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptMetrics_bas", "cptGetCPLI", Err, Erl)
  Resume exit_here
End Sub

Sub cptGET(strWhat As String)
'todo: need to store weekly bcwp, etc data somewhere
'objects
'strings
Dim strMsg As String
'longs
Dim lngBEI_AF As Long
Dim lngBEI_BF As Long
'integers
'doubles
Dim dblBCWS As Double
Dim dblBCWP As Double
Dim dblResult As Double
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'validate tasks exist
  If ActiveProject.Tasks.Count = 0 Then
    MsgBox "This project has no tasks.", vbExclamation + vbOKOnly, "No Tasks"
    GoTo exit_here
  End If

  Select Case strWhat
    Case "BEI"
      lngBEI_BF = CLng(cptGetMetric("bei_bf"))
      If lngBEI_BF = 0 Then
        MsgBox "No baseline finishes found.", vbExclamation + vbOKOnly, "No BEI"
        GoTo exit_here
      End If
      lngBEI_AF = CLng(cptGetMetric("bei_af"))
      strMsg = "BEI = # Actual Finishes / # Planned Finishes" & vbCrLf
      strMsg = strMsg & "BEI = " & Format(lngBEI_AF, "#,##0") & " / " & Format(lngBEI_BF, "#,##0") & vbCrLf & vbCrLf
      strMsg = strMsg & "BEI = " & Format((lngBEI_AF / lngBEI_BF), "#0.#0")
      MsgBox strMsg, vbInformation + vbOKOnly, "Baseline Execution Index (BEI)"
      
    Case "CEI"
      'todo: need to track previous week's plan
      
    Case "SPI"
      dblBCWP = cptGetMetric("bcwp")
      dblBCWS = cptGetMetric("bcws")
      If dblBCWS = 0 Then
        MsgBox "No BCWS found.", vbExclamation + vbOKOnly, "Schedule Performance Index (SPI) - Hours"
        GoTo exit_here
      End If
      strMsg = "SPI = BCWP / BCWS" & vbCrLf
      strMsg = strMsg & "SPI = " & Format(dblBCWP, "#,##0h") & " / " & Format(dblBCWS, "#,##0h") & vbCrLf & vbCrLf
      strMsg = strMsg & "SPI = ~" & Round(dblBCWP / dblBCWS, 2) '& vbCrLf & vbCrLf
      'strMsg = strMsg & "(Assumes EV% in Physical % Complete.)"
      MsgBox strMsg, vbInformation + vbOKOnly, "Schedule Performance Index (SPI) - Hours"
      
    Case "SV"
      dblBCWP = cptGetMetric("bcwp")
      dblBCWS = cptGetMetric("bcws")
      If dblBCWS = 0 Then
        MsgBox "No BCWS found.", vbExclamation + vbOKOnly, "Schedule Variance (SV) - Hours"
        GoTo exit_here
      End If
      strMsg = strMsg & "Schedule Variance (SV)" & vbCrLf
      strMsg = strMsg & "SV = BCWP - BCWS" & vbCrLf
      strMsg = strMsg & "SV = " & Format(dblBCWP, "#,##0h") & " - " & Format(dblBCWS, "#,##0h") & vbCrLf
      strMsg = strMsg & "SV = ~" & Format(dblBCWP - dblBCWS, "#,##0.0h") & vbCrLf & vbCrLf
      strMsg = strMsg & "Schedule Variance % (SV%)" & vbCrLf
      strMsg = strMsg & "SV% = ( SV / BCWS ) * 100" & vbCrLf
      strMsg = strMsg & "SV% = ( " & Format((dblBCWP - dblBCWS), "#,##0.0h") & " / " & Format(dblBCWS, "#,##0.0h") & " ) * 100" & vbCrLf
      strMsg = strMsg & "SV% = " & Format(((dblBCWP - dblBCWS) / dblBCWS), "0.00%") '& vbCrLf & vbCrLf
      'strMsg = strMsg & "(Assumes EV% in Physical % Complete.)"
      
      MsgBox strMsg, vbInformation + vbOKOnly, "Schedule Variance (SV) - Hours"
      
    Case "es" 'earned schedule
          'todo: earned schedule
    
  End Select
  
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptMerics_Bas", "cptGet", Err, Erl)
  Resume exit_here
End Sub

Sub cptGetHitTask()
'objects
Dim oTask As Task
'strings
Dim strMsg As String
'longs
Dim lngAF As Long
Dim lngBLF As Long
'integers
'doubles
'booleans
'variants
'dates
Dim dtStatus As Date

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'confirm status date
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "This project has no status date. Please update and try again.", vbExclamation + vbOKOnly, "Metrics"
    GoTo exit_here
  Else
    dtStatus = ActiveProject.StatusDate
  End If
  
  'find it
  For Each oTask In ActiveProject.Tasks
    If oTask Is Nothing Then GoTo next_task
    If oTask.Summary Then GoTo next_task
    If Not oTask.Active Then GoTo next_task
    If IsDate(oTask.BaselineFinish) Then
      'was task baselined to finish before status date?
      If oTask.BaselineFinish <= dtStatus Then
        lngBLF = lngBLF + 1
        'did it?
        If IsDate(oTask.ActualFinish) Then
          If oTask.ActualFinish <= oTask.BaselineFinish Then
            lngAF = lngAF + 1
          End If
        End If
      End If
    End If
next_task:
  Next oTask

  strMsg = "BF = # Tasks Baselined to Finish ON or before Status Date" & vbCrLf
  strMsg = strMsg & "AF = # BF that Actually Finished ON or before Baseline Finish" & vbCrLf & vbCrLf
  strMsg = strMsg & "Hit Task % = (AF / BF) / 100" & vbCrLf
  strMsg = strMsg & "Hit Task % = (" & Format(lngAF, "#,##0") & " / " & Format(lngBLF, "#,##0") & ") / 100" & vbCrLf & vbCrLf
  strMsg = strMsg & "Hit Task % = " & Format((lngAF / lngBLF), "0%")
  MsgBox strMsg, vbInformation + vbOKOnly, "Hit Task %"

exit_here:
  On Error Resume Next
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptMetrics_bas", "cptGetHitTask", Err, Erl)
  Resume exit_here
End Sub

Function cptGetMetric(strGet As String) As Double
'todo: no screen changes!
'objects
Dim oAssignment As Assignment
Dim tsv As TimeScaleValue
Dim tsvs As TimeScaleValues
Dim oTasks As Tasks
Dim oTask As Task
'strings
Dim strLOE As String
'longs
Dim lngLOEField As Long
Dim lngEVP As Long
Dim lngYears As Long
'integers
'doubles
Dim dblResult As Double
'booleans
'variants
'dates
Dim dtStatus As Date

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  lngYears = Year(ActiveProject.ProjectFinish) - Year(ActiveProject.ProjectStart) + 1
  
  'confirm status date
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "This project has no status date. Please update and try again.", vbExclamation + vbOKOnly, "Metrics"
    GoTo exit_here
  Else
    dtStatus = ActiveProject.StatusDate
  End If
  
  cptSpeed True
  FilterClear
  GroupClear
  OptionsViewEx displaysummarytasks:=True, displaynameindent:=True
  On Error Resume Next
  If Not OutlineShowAllTasks Then
    Sort "ID", , , , , , False, True
    OutlineShowAllTasks
  End If
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  SelectAll
  Set oTasks = ActiveSelection.Tasks
  For Each oTask In oTasks
    If oTask Is Nothing Then GoTo next_task
    If oTask.ExternalTask Then GoTo next_task
    If oTask.Summary Then GoTo next_task
    If Not oTask.Active Then GoTo next_task
    If oTask.BaselineWork > 0 Then 'idea here was to limit tasks to PMB tasks only
                                  'but won't work for non-resource loaded schedules
      Select Case strGet
        Case "bac"
          For Each oAssignment In oTask.Assignments
            If oAssignment.ResourceType = pjResourceTypeWork Then
              dblResult = dblResult + (oAssignment.BaselineWork / 60)
            End If
          Next oAssignment
          
        Case "etc"
          For Each oAssignment In oTask.Assignments
            If oAssignment.ResourceType = pjResourceTypeWork Then
              dblResult = dblResult + (oAssignment.RemainingWork / 60)
            End If
          Next oAssignment

          
        Case "bcws"
          If oTask.BaselineStart < dtStatus Then
            For Each oAssignment In oTask.Assignments
              If oAssignment.ResourceType = pjResourceTypeWork Then
                Set tsvs = oAssignment.TimeScaleData(oTask.BaselineStart, dtStatus, pjAssignmentTimescaledBaselineWork, pjTimescaleWeeks)
                For Each tsv In tsvs
                  dblResult = dblResult + (IIf(tsv.Value = "", 0, tsv.Value) / 60)
                Next
              End If
            Next oAssignment
          End If
          
        Case "bcwp"
          lngEVP = CLng(cptGetSetting("Metrics", "cboEVP"))
          lngLOEField = CLng(cptGetSetting("Metrics", "cboLOEField"))
          strLOE = cptGetSetting("Metrics", "txtLOE")
          
          For Each oAssignment In oTask.Assignments
            If oAssignment.ResourceType = pjResourceTypeWork Then
              If oTask.GetField(lngLOEField) = strLOE Then
                If oTask.BaselineStart < dtStatus Then
                  Set tsvs = oAssignment.TimeScaleData(oTask.BaselineStart, dtStatus, pjAssignmentTimescaledBaselineWork, pjTimescaleWeeks, 1)
                  For Each tsv In tsvs
                    dblResult = dblResult + (IIf(tsv.Value = "", 0, tsv.Value) / 60)
                  Next
                End If
              Else
                dblResult = dblResult + ((oAssignment.BaselineWork / 60) * (CLng(cptRegEx(oTask.GetField(lngEVP), "[0-9]*")) / 100))
              End If
            End If
          Next oAssignment
      End Select
    End If 'bac>0
    Select Case strGet
    
      Case "bei_bf"
        dblResult = dblResult + IIf(oTask.BaselineFinish <= dtStatus, 1, 0)
          
      Case "bei_af"
        dblResult = dblResult + IIf(oTask.ActualFinish <= dtStatus, 1, 0)

    End Select
next_task:
    Application.StatusBar = "Calculating " & UCase(strGet) & "..."
  Next

  cptGetMetric = dblResult

exit_here:
  On Error Resume Next
  Set oAssignment = Nothing
  Application.StatusBar = ""
  cptSpeed False
  Set tsv = Nothing
  Set tsvs = Nothing
  Set oTasks = Nothing
  Set oTask = Nothing

  Exit Function
err_here:
  'Debug.Print Task.UniqueID & ": " & Task.Name
  Call cptHandleErr("cptMetrics_bas", "cptGetMetric", Err, Erl)
  Resume exit_here

End Function

Sub cptShowMetricsSettings_frm(Optional blnModal As Boolean = False)
  'objects
  'strings
  Dim strCustomName As String
  Dim strLOE As String
  Dim strLOEField As String
  Dim strEVP As String
  'longs
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  With cptMetricsSettings_frm
  
    .cboEVP.Clear
    .cboEVP.AddItem
    .cboEVP.List(.cboEVP.ListCount - 1, 0) = FieldNameToFieldConstant("Physical % Complete")
    .cboEVP.List(.cboEVP.ListCount - 1, 1) = "Physical % Complete"
    For lngItem = 1 To 20
      .cboEVP.AddItem
      .cboEVP.List(.cboEVP.ListCount - 1, 0) = FieldNameToFieldConstant("Number" & lngItem)
      .cboEVP.List(.cboEVP.ListCount - 1, 1) = "Number" & lngItem
      strCustomName = CustomFieldGetName(FieldNameToFieldConstant("Number" & lngItem))
      If Len(strCustomName) > 0 Then
        .cboEVP.List(.cboEVP.ListCount - 1, 1) = strCustomName & " (Number" & lngItem & ")"
      End If
    Next lngItem
    
    .cboLOEField.Clear
    For lngItem = 1 To 30
      .cboLOEField.AddItem
      .cboLOEField.List(.cboLOEField.ListCount - 1, 0) = FieldNameToFieldConstant("Text" & lngItem)
      .cboLOEField.List(.cboLOEField.ListCount - 1, 1) = "Text" & lngItem
      strCustomName = CustomFieldGetName(FieldNameToFieldConstant("Text" & lngItem))
      If Len(strCustomName) > 0 Then
        .cboLOEField.List(.cboLOEField.ListCount - 1, 1) = strCustomName & " (Text" & lngItem & ")"
      End If
    Next lngItem
    
    strEVP = cptGetSetting("Metrics", "cboEVP")
    If Len(strEVP) > 0 Then .cboEVP.Value = CLng(strEVP)
    strLOEField = cptGetSetting("Metrics", "cboLOEField")
    If Len(strLOEField) > 0 Then .cboLOEField.Value = CLng(strLOEField)
    strLOE = cptGetSetting("Metrics", "txtLOE")
    If Len(strLOE) > 0 Then .txtLOE = strLOE
    If blnModal Then
      .Show
    Else
      .Show False
    End If
  End With
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  'Call HandleErr("cptMetrics_bas", "cptShowMetricsSettings_frm", Err)
  MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Error"
  Resume exit_here
End Sub

Function cptMetricsSettingsExist() As Boolean
  'objects
  'strings
  Dim strLOE As String
  Dim strLOEField As String
  Dim strEVP As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
    
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  strEVP = cptGetSetting("Metrics", "cboEVP")
  strLOEField = cptGetSetting("Metrics", "cboLOEField")
  strLOE = cptGetSetting("Metrics", "txtLOE")
  
  If Len(strEVP) = 0 Or Len(strLOEField) = 0 Or Len(strLOE) = 0 Then
    cptMetricsSettingsExist = False
  Else
    cptMetricsSettingsExist = True
  End If

exit_here:
  On Error Resume Next

  Exit Function
err_here:
  Call cptHandleErr("cptMetrics_bas", "cptMetricsSettingsExist", Err, Erl)
  Resume exit_here
End Function
Sub cptEarnedSchedule()
'objects
Dim rng As Excel.Range 'Object
Dim Task As Task
Dim TSVS_ACTUAL As TimeScaleValues
Dim TSVS_WORK As TimeScaleValues
Dim TSV As TimeScaleValue
Dim TSVS As TimeScaleValues
Dim Worksheet As Excel.Worksheet 'object
Dim Workbook As Excel.Workbook 'object
Dim xlApp As Excel.Application 'Object
'strings
'longs
Dim lngBAC As Long
Dim lngLastRow As Long
Dim lngES As Long
Dim lngAD As Long
Dim lngEVP As Long
Dim lngCurrentRow As Long
Dim lngEVT As Long
'integers
'doubles
Dim dblBCWP As Double
'booleans
'variants
'dates
Dim dtETC As Date
Dim dtStatus As Date
Dim dtEnd As Date
Dim dtStart As Date

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'validate has tasks, baseline, status date
  If ActiveProject.Tasks.Count = 0 Then
    MsgBox "This project has no tasks.", vbExclamation + vbOKOnly, "Earned Schedule"
    GoTo exit_here
  End If
  If Not IsDate(ActiveProject.ProjectSummaryTask.BaselineStart) Then
    MsgBox "This project has not been properly baselined.", vbExclamation + vbOKOnly, "Earned Schedule"
    GoTo exit_here
  End If
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "This project has no status date.", vbExclamation + vbOKOnly, "Earned Schedule"
    GoTo exit_here
  End If
  
  cptSpeed True
  
  'set up the workbook
  Set xlApp = CreateObject("Excel.Application")
  Set Workbook = xlApp.Workbooks.Add
  xlApp.ScreenUpdating = False
  xlApp.Calculation = xlCalculationManual
  Set Worksheet = Workbook.Sheets(1)
  xlApp.ActiveWindow.Zoom = 85
  Worksheet.Name = "Earned Schedule"
  Worksheet.[A1:D1] = Array("WEEK", "BCWS", "BCWP", "ETC")
  Worksheet.[A2].Select
  xlApp.ActiveWindow.FreezePanes = True
  
  lngEVT = FieldNameToFieldConstant("EVT")
  
  'get week-over-week BCWS
  dtStart = xlApp.WorksheetFunction.Min(ActiveProject.ProjectSummaryTask.Start, ActiveProject.ProjectSummaryTask.BaselineStart)
  dtEnd = xlApp.WorksheetFunction.Max(ActiveProject.ProjectSummaryTask.Finish, ActiveProject.ProjectSummaryTask.BaselineFinish)
  
  Set TSVS = ActiveProject.ProjectSummaryTask.TimeScaleData(dtStart, dtEnd, pjTaskTimescaledBaselineWork, pjTimescaleWeeks, 1)
  For Each TSV In TSVS
    Worksheet.Cells(TSV.Index + 1, 1) = DateAdd("d", 5, TSV.StartDate) 'ensure Friday
    Worksheet.Cells(TSV.Index + 1, 2) = Val(TSV.Value) / 60
    Worksheet.Columns("A:B").AutoFit
    'get work
    Set TSVS_WORK = ActiveProject.ProjectSummaryTask.TimeScaleData(TSV.StartDate, TSV.EndDate, pjTaskTimescaledWork, pjTimescaleWeeks, 1)
    'get actual work (per msp)
    Set TSVS_ACTUAL = ActiveProject.ProjectSummaryTask.TimeScaleData(TSV.StartDate, TSV.EndDate, pjTaskTimescaledActualWork, pjTimescaleWeeks, 1)
    'return remaining work (= work - actual work)
    Worksheet.Cells(TSV.Index + 1, 4) = (Val(TSVS_WORK(1)) - Val(TSVS_ACTUAL(1))) / 60
  Next TSV
  
  'convert BCWS to cumulative
  Worksheet.[E2].FormulaR1C1 = "=RC[-3]"
  Worksheet.[E3].FormulaR1C1 = "=R[-1]C+RC[-3]"
  Worksheet.Range(Worksheet.[E3], Worksheet.[A1048576].End(xlUp).Offset(0, 4)).FillDown
  
  'convert ETC to cumulative?
  Worksheet.[F2].FormulaR1C1 = "=RC[-2]"
  Worksheet.[F3].FormulaR1C1 = "=R[-1]C+RC[-2]"
  Worksheet.Range(Worksheet.[F3], Worksheet.[A1048576].End(xlUp).Offset(0, 5)).FillDown
  
  'calculate once
  Worksheet.Calculate
  
  'paste bcws values
  Worksheet.Range(Worksheet.[E2], Worksheet.[E2].End(xlDown)).Copy
  Worksheet.[B2].PasteSpecial xlValues
  Worksheet.Range(Worksheet.[E2], Worksheet.[E2].End(xlDown)).Clear
  
  'paste etc values
  Worksheet.Range(Worksheet.[F2], Worksheet.[F2].End(xlDown)).Copy
  Worksheet.[D2].PasteSpecial xlValues
  Worksheet.Range(Worksheet.[F2], Worksheet.[F2].End(xlDown)).Clear
  
  'format the ranges
  Worksheet.Range(Worksheet.[D2], Worksheet.[A1048576].End(xlUp).Offset(0, 1)).NumberFormat = "#,##0.00"
  
  'add borders
  Set rng = Worksheet.Range(Worksheet.[A1], Worksheet.[A1].End(xlDown).Offset(0, 3))
  rng.Borders(xlDiagonalDown).LineStyle = xlNone
  rng.Borders(xlDiagonalUp).LineStyle = xlNone
  With rng.Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .Weight = xlThin
  End With
  With rng.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .Weight = xlThin
  End With
  With rng.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .Weight = xlThin
  End With
  With rng.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .Weight = xlThin
  End With
  With rng.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .ThemeColor = 1
    .TintAndShade = -0.249946594869248
    .Weight = xlThin
  End With
  With rng.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .ThemeColor = 1
    .TintAndShade = -0.249946594869248
    .Weight = xlThin
  End With
  Set rng = Worksheet.Range("A1:D1")
  rng.HorizontalAlignment = xlCenter
  rng.Font.Bold = True
  With rng
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
  End With
  With rng.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .ThemeColor = xlThemeColorDark1
    .TintAndShade = -0.149998478032896
    .PatternTintAndShade = 0
  End With
  
  'get current BCWP
  lngEVP = FieldNameToFieldConstant("Physical % Complete")
  For Each Task In ActiveProject.Tasks
    If Task Is Nothing Then GoTo next_task
    If Task.ExternalTask Then GoTo next_task
    If Task.Summary Then GoTo next_task
    If Not Task.Active Then GoTo next_task
    If Task.BaselineWork > 0 Then
      dblBCWP = dblBCWP + ((Task.BaselineWork * Val(Task.GetField(lngEVP)) / 100) / 60)
    End If
next_task:
  Next Task
  dtStatus = FormatDateTime(ActiveProject.StatusDate, vbShortDate)
  If Weekday(ActiveProject.StatusDate) <> 6 Then
    'todo: adjust status date to Friday
  End If
  lngCurrentRow = Worksheet.Columns(1).Find(dtStatus, lookat:=xlWhole).Row
  Worksheet.Cells(lngCurrentRow, 3).Value = dblBCWP
  
  'highlight BCWP
  With Worksheet.Cells(lngCurrentRow, 3).Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 65535
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  
  'get actual duration (in weeks)
  lngAD = lngCurrentRow - 1
  
  'get earned schedule (in weeks)
  If dblBCWP < xlApp.WorksheetFunction.Min(Worksheet.Range(Worksheet.[B2], Worksheet.[B2].End(xlDown))) Then
    lngES = 1
  Else
    lngES = xlApp.WorksheetFunction.Match(dblBCWP, Worksheet.Range(Worksheet.[B2], Worksheet.[B2].End(xlDown)), 1)
  End If
  xlApp.ActiveWindow.ScrollRow = xlApp.WorksheetFunction.Min(lngES, lngAD)
  'highlight it
  With Worksheet.Cells(lngES + 1, 2).Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 65535
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  
  'todo: get rd using same method as above
  'todo: how to account for schedule margin?
  
  'add ES
  Worksheet.Cells(lngES + 1, 6).Value = "Earned Schedule:"
  Worksheet.Cells(lngES + 1, 6).HorizontalAlignment = xlRight
  Worksheet.Cells(lngES + 1, 7).FormulaR1C1 = "=COUNTA(R2C1:R" & lngES + 1 & "C1)"
  Workbook.Names.Add "ES", Worksheet.Cells(lngES + 1, 7)
  Worksheet.Cells(lngES + 1, 8).Value = "wks"
    
  'add AD
  Worksheet.Cells(lngAD + 1, 6).Value = "Actual Duration:"
  Worksheet.Cells(lngAD + 1, 6).HorizontalAlignment = xlRight
  Worksheet.Cells(lngAD + 1, 7).FormulaR1C1 = "=COUNTA(R2C1:R" & lngAD + 1 & "C1)"
  Workbook.Names.Add "AD", Worksheet.Cells(lngAD + 1, 7)
  Worksheet.Cells(lngAD + 1, 8).Value = "wks"
  
  'add SPI(t)
  Worksheet.Cells(lngAD + 3, 6).Value = "SPI(t):"
  Worksheet.Cells(lngAD + 3, 6).HorizontalAlignment = xlRight
  Worksheet.Cells(lngAD + 3, 7).FormulaR1C1 = "=ES/AD"
  Workbook.Names.Add "SPIt", Worksheet.Cells(lngAD + 3, 7)
  Worksheet.Cells(lngAD + 3, 7).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* "" - ""??_);_(@_)"
  
  'add PDWR
  'todo: do we remove schedule margin from this calc?
  'todo: if yes then base it on earliest date of BAC; if not base on baseline finish (assumes that schedule margin is included in roll-up)
  lngLastRow = Worksheet.[F1048546].End(xlUp).Row + 2
  Worksheet.Cells(lngLastRow, 6).Value = "Planned Duration of Work Remaining:"
  Worksheet.Cells(lngLastRow, 6).HorizontalAlignment = xlRight
  'ignoring schedule margin
  lngBAC = xlApp.WorksheetFunction.Match(Worksheet.[B1048576].End(xlUp).Value, Worksheet.Range(Worksheet.[B2], Worksheet.[B2].End(xlDown)), 0) + 1
  Worksheet.Cells(lngLastRow, 7).FormulaR1C1 = "=COUNTA(R" & lngES + 2 & "C2:R" & lngBAC & "C2)"
  Workbook.Names.Add "PDWR", Worksheet.Cells(lngLastRow, 7)
  Worksheet.Cells(lngLastRow, 8).Value = "wks"
  
  'add RD
  dtETC = Worksheet.[A1048576].End(xlUp)
  lngLastRow = Worksheet.[F1048546].End(xlUp).Row + 2
  Worksheet.Rows.Replace "-", "/"
  Worksheet.Cells(lngLastRow, 6).Value = "Remaining Duration:"
  Worksheet.Cells(lngLastRow, 6).HorizontalAlignment = xlRight
  lngETC = xlApp.WorksheetFunction.Match(Worksheet.[D1048576].End(xlUp).Value, Worksheet.Range(Worksheet.[D2], Worksheet.[D2].End(xlDown)), 0) + 1
  lngLastRow = Worksheet.[F1048546].End(xlUp).Row + 2
  'todo: do we include schedule margin? if yes, then use latest week; if no then use earliest week of BAC
  Worksheet.Cells(lngLastRow, 6) = "Estimated Completion Date:"
  Worksheet.Cells(lngLastRow, 6).HorizontalAlignment = xlRight
  'ignoring schedule margin:
  Worksheet.Cells(lngLastRow, 7).FormulaR1C1 = "=" & Worksheet.Cells(lngETC, 1).Address(True, True, xlR1C1) 'dtETC
  'considering schedule margin:
  'Worksheet.Cells(lngLastRow, 7).FormulaR1C1 = "=" & Worksheet.[A1048576].end(xlup).Address(True, True, xlR1C1) 'dtETC
  Workbook.Names.Add "ECD", Worksheet.Cells(lngLastRow, 7)
  Worksheet.Cells(lngLastRow, 7).NumberFormat = "mm/dd/yyyy"
  lngLastRow = Worksheet.[F1048546].End(xlUp).Row + 2
  Workbook.Names.Add "SD", Worksheet.Cells(lngAD + 1, 1)
  Worksheet.Cells(lngLastRow - 4, 7).FormulaR1C1 = "=NETWORKDAYS(SD+1,ECD)/5"
  Workbook.Names.Add "RD", Worksheet.Cells(lngLastRow - 4, 7)
  Worksheet.Cells(lngLastRow - 4, 8).Value = "wks"
  
  'add TSPI(ed)
  lngLastRow = Worksheet.[F1048546].End(xlUp).Row + 2
  Worksheet.Cells(lngLastRow, 6).Value = "TSPI(ed):"
  Worksheet.Cells(lngLastRow, 6).HorizontalAlignment = xlRight
  Worksheet.Cells(lngLastRow, 7).FormulaR1C1 = "=PDWR/RD"
  Workbook.Names.Add "TSPIed", Worksheet.Cells(lngLastRow, 7)
  Worksheet.Cells(lngLastRow, 7).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
  
  'add threshold
  lngLastRow = Worksheet.[F1048546].End(xlUp).Row + 2
  Worksheet.Cells(lngLastRow, 6).Value = "'|SPI(t)-TSPI(ed)|:"
  Worksheet.Cells(lngLastRow, 6).HorizontalAlignment = xlRight
  Worksheet.Cells(lngLastRow, 7).FormulaR1C1 = "=ABS(SPIt-TSPIed)"
  Worksheet.Cells(lngLastRow, 7).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
  'add interpretation
  Worksheet.Cells(lngLastRow, 8).FormulaR1C1 = "=IF(RC[-1]>0.1,""OUT OF RANGE"",""IN RANGE"")"
  
  'add threshold
  lngLastRow = Worksheet.[F1048546].End(xlUp).Row + 2
  Worksheet.Cells(lngLastRow, 6).Value = "SPI(t)-TSPI(ed):"
  Worksheet.Cells(lngLastRow, 6).HorizontalAlignment = xlRight
  Worksheet.Cells(lngLastRow, 7).FormulaR1C1 = "=SPIt-TSPIed"
  Worksheet.Cells(lngLastRow, 7).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
  'add interpretation
  Worksheet.Cells(lngLastRow, 8).FormulaR1C1 = "=IF(RC[-1]<0.1,""OPTIMISTIC"",IF(RC[-1]>0.1,""PESSIMISTIC"",""REASONABLE""))"
  
  'add IECD(es)
  lngLastRow = Worksheet.[F1048546].End(xlUp).Row + 2
  Worksheet.Cells(lngLastRow, 6).Value = "IECD(es):"
  Worksheet.Cells(lngLastRow, 6).HorizontalAlignment = xlRight
  Worksheet.Cells(lngLastRow, 7).FormulaR1C1 = "=SD+((PDWR*5)/SPIt)"
  Workbook.Names.Add "IECDes", Worksheet.Cells(lngLastRow, 7)
  Worksheet.Cells(lngLastRow, 7).NumberFormat = "m/d/yyyy"
  lngLastRow = Worksheet.[F1048546].End(xlUp).Row + 2
  Worksheet.Cells(lngLastRow, 6).Value = "DELTA:"
  Worksheet.Cells(lngLastRow, 6).HorizontalAlignment = xlRight
  Worksheet.Cells(lngLastRow, 7).FormulaR1C1 = "=NETWORKDAYS(IECDes,ECD)"
  Worksheet.Cells(lngLastRow, 7).HorizontalAlignment = xlCenter
  Worksheet.Cells(lngLastRow, 7).NumberFormat = "#,##0_);[Red](#,##0)"
  Worksheet.Cells(lngLastRow, 8).Value = "work days"
  
  'add indicator
'  Worksheet.Cells(lngLastRow + 1, 8).Value = "Predicted Growth:"
'  Worksheet.Cells(lngLastRow + 1, 8).HorizontalAlignment = xlRight
'  Worksheet.Cells(lngLastRow + 1, 9).FormulaR1C1 = "=ABS(R[-1]C)/((RD*5)+ABS(R[-1]C))"
'  Worksheet.Cells(lngLastRow + 1, 9).NumberFormat = "0%"
'  Worksheet.Cells(lngLastRow + 1, 9).HorizontalAlignment = xlCenter
  lngLastRow = Worksheet.[F1048546].End(xlUp).Row + 2
  Worksheet.Cells(lngLastRow, 6).Value = "IMS IS:"
  Worksheet.Cells(lngLastRow, 6).HorizontalAlignment = xlRight
  'todo: fix this formula
  Worksheet.Cells(lngLastRow, 7).FormulaR1C1 = "=IF(ABS(R[-2]C)<=5,""REASONABLE"",IF(R[-2]C<5,""OPTIMISTIC"",""PESSIMISTIC""))"
  Worksheet.Cells(lngLastRow, 7).HorizontalAlignment = xlCenter
  
  'adjust columns
  Worksheet.Columns(5).ColumnWidth = 2
  Worksheet.Columns.AutoFit
  
  'select the IECD(es)
  Worksheet.Cells(lngLastRow, 7).Select
  
  xlApp.Visible = True
  
exit_here:
  On Error Resume Next
  Set rng = Nothing
  Set Task = Nothing
  Set TSVS_ACTUAL = Nothing
  Set TSVS_WORK = Nothing
  Set TSV = Nothing
  Set TSVS = Nothing
  Set Worksheet = Nothing
  xlApp.ScreenUpdating = True
  xlApp.Calculation = xlCalculationAutomatic
  cptSpeed False
  Set Workbook = Nothing
  Set xlApp = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptMetrics_bas", "cptEarnedSchedule", Err, Erl)
  Resume exit_here
End Sub
