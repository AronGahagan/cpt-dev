Attribute VB_Name = "cptMetrics_bas"
'<cpt_version>v1.0.6</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
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
  
  'confirm status date
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "This project has no status date. Please update and try again.", vbExclamation + vbOKOnly, "Metrics"
    Exit Sub
  Else
    MsgBox Format(cptGetMetric("bcwp"), "#,##0.00") & vbCrLf & vbCrLf & "(Assumes EV% in Physical % Complete.)", vbInformation + vbOKOnly, "Budgeted Cost of Work Performed (BCWP) - hours"
  End If
  
End Sub

Sub cptGetSPI()
  
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
      strMsg = strMsg & "SPI = ~" & Round(dblBCWP / dblBCWS, 2) & vbCrLf & vbCrLf
      strMsg = strMsg & "(Assumes EV% in Physical % Complete.)"
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
      strMsg = strMsg & "SV% = " & Format(((dblBCWP - dblBCWS) / dblBCWS), "0.00%") & vbCrLf & vbCrLf
      strMsg = strMsg & "(Assumes EV% in Physical % Complete.)"
      
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
Dim TSV As TimeScaleValue
Dim TSVS As TimeScaleValues
Dim oTasks As Tasks
Dim oTask As Task
'strings
'longs
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
                Set TSVS = oAssignment.TimeScaleData(oTask.BaselineStart, dtStatus, pjAssignmentTimescaledBaselineWork, pjTimescaleWeeks)
                For Each TSV In TSVS
                  dblResult = dblResult + IIf(TSV.Value = "", 0, TSV.Value) / 60
                Next
              End If
            Next oAssignment
          End If
          
        Case "bcwp"
          For Each oAssignment In oTask.Assignments
            If oAssignment.ResourceType = pjResourceTypeWork Then
              dblResult = dblResult + ((oAssignment.BaselineWork / 60) * (oTask.PhysicalPercentComplete / 100))
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
  Set TSV = Nothing
  Set TSVS = Nothing
  Set oTasks = Nothing
  Set oTask = Nothing

  Exit Function
err_here:
  'Debug.Print Task.UniqueID & ": " & Task.Name
  Call cptHandleErr("cptMetrics_bas", "cptGetMetric", Err, Erl)
  Resume exit_here

End Function
