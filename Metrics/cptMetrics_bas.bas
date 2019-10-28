Attribute VB_Name = "cptMetrics_bas"
'<cpt_version>v1.0.0</cpt_version>
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
  Call cptHandleErr("cptMetrics_bas", "cptExportMetricsExcel", err, Erl)
  Resume exit_here
End Sub

Sub cptGetBAC()
  MsgBox Format(cptGetMetric("bac"), "#,##0.00"), vbInformation + vbOKOnly, "Budget at Complete (BAC) - hours"
End Sub

Sub cptGetETC()
  MsgBox Format(cptGetMetric("etc"), "#,##0.00"), vbInformation + vbOKOnly, "Estimate to Complete (ETC) - hours"
End Sub

Sub cptGetBCWS()
  MsgBox Format(cptGetMetric("bcws"), "#,##0.00"), vbInformation + vbOKOnly, "Budgeted Cost of Work Scheduled (BCWS) - hours"
End Sub

Sub cptGetBCWP()
  MsgBox Format(cptGetMetric("bcwp"), "#,##0.00"), vbInformation + vbOKOnly, "Budgeted Cost of Work Performed (BCWP) - hours"
End Sub

Sub cptGetSPI()
  Call cptGET("SPI")
End Sub

Sub cptGetBEI()
  Call cptGET("BEI")
End Sub

Sub cptGetCEI()
  Call cptGET("CEI")
End Sub

Sub cptGetSV()
  Call cptGET("SV")
End Sub

Sub cptGetCPLI()
'objects
Dim Pred As Task
Dim Task As Task
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

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  strTitle = "Critical Path Length Index (CPLI)"

  'confirm a single, target task is selected
  If ActiveSelection.Tasks.Count <> 1 Then
    MsgBox "Please select a single, active, and non-summary target task.", vbExclamation + vbOKOnly, strTitle
    GoTo exit_here
  End If
  
  Set Task = ActiveSelection.Tasks(1)
  
  'use MFO or MSO constraint
  If Task.ConstraintType <> pjMFO And Task.ConstraintType <> pjMSO Then
    strMsg = "No MSO/MFO constraint found; temporarily using Deadline..." & vbCrLf
    'if no MFO then use deadline as MFO
    If IsDate(Task.Deadline) Then
      If IsDate(Task.ConstraintDate) Then dtConstraintDate = Task.ConstraintDate
      lngConstraintType = Task.ConstraintType
      Task.ConstraintDate = Task.Deadline
      Task.ConstraintType = pjMFO
      lngTS = Task.TotalSlack
      dtFinish = Task.Finish
      If CLng(dtConstraintDate) > 0 Then Task.ConstraintDate = dtConstraintDate
      Task.ConstraintType = lngConstraintType
    Else
      strMsg = strMsg & "No Deadline found; temporarily using Baseline Finish..." & vbCrLf
      If Not IsDate(Task.BaselineFinish) Then
        strMsg = strMsg & "No Baseline Finish found." & vbCrLf & vbCrLf
        strMsg = strMsg & "In order to calculate the CPLI, the target task should be (at least temporarily) constrained with a MFO or Deadline." & vbCrLf & vbCrLf
        strMsg = strMsg & "Please constrain the task and try again."
        MsgBox strMsg, vbExclamation + vbOKOnly, strTitle
        GoTo exit_here
      Else
        If IsDate(Task.ConstraintDate) Then dtConstraintDate = Task.ConstraintDate
        lngConstraintType = Task.ConstraintType
        Task.ConstraintDate = Task.BaselineFinish
        Task.ConstraintType = pjMFO
        lngTS = Task.TotalSlack
        dtFinish = Task.Finish
        If CLng(dtConstraintDate) > 0 Then Task.ConstraintDate = dtConstraintDate
        Task.ConstraintType = lngConstraintType
      End If
    End If
  Else
    lngTS = Task.TotalSlack
    dtFinish = Task.Finish
  End If
      
  'use status date if exists
  If IsDate(ActiveProject.StatusDate) Then
    dtStart = ActiveProject.StatusDate
  Else
    dtStart = FormatDateTime(Now(), vbShortDate) & " 08:00 AM"
  End If
  
  'use earliest start date
  'NOTE: cannot account for schedule margin due to possibility
  'of dual paths, one with and one without, a particular SM task
  
  If Task Is Nothing Then GoTo exit_here
  If Task.Summary Then GoTo exit_here
  If Not Task.Active Then GoTo exit_here
  HighlightDrivingPredecessors Set:=True
  For Each Pred In ActiveProject.Tasks
    If Pred.PathDrivingPredecessor Then
      If IsDate(Pred.ActualStart) Then
        If Pred.Stop < dtStart Then dtStart = Pred.Stop
      Else
        If Pred.Start < dtStart Then dtStart = Pred.Start
      End If
    End If
  Next Pred
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
  strMsg = strMsg & "Note: schedule margin tasks are not considered."
  
  MsgBox strMsg, vbInformation + vbOKOnly, "Critical Path Length Index (CPLI)"
    
exit_here:
  On Error Resume Next
  Set Pred = Nothing
  Application.CloseUndoTransaction
  Set Task = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptMetrics_bas", "cptGetCPLI", err, Erl)
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
      strMsg = "SPI = BCWP / BCWS" & vbCrLf
      strMsg = strMsg & "SPI = " & Format(dblBCWP, "#,##0h") & " / " & Format(dblBCWS, "#,##0h") & vbCrLf & vbCrLf
      strMsg = strMsg & "SPI = ~" & Round(dblBCWP / dblBCWS, 2)
      MsgBox strMsg, vbInformation + vbOKOnly, "Schedule Performance Index (SPI) - Hours"
      
    Case "SV"
      dblBCWP = cptGetMetric("bcwp")
      dblBCWS = cptGetMetric("bcws")
      strMsg = strMsg & "Schedule Variance (SV)" & vbCrLf
      strMsg = strMsg & "SV = BCWP - BCWS" & vbCrLf
      strMsg = strMsg & "SV = " & Format(dblBCWP, "#,##0h") & " - " & Format(dblBCWS, "#,##0h") & vbCrLf
      strMsg = strMsg & "SV = ~" & Format(dblBCWP - dblBCWS, "#,##0.0h") & vbCrLf & vbCrLf
      strMsg = strMsg & "Schedule Variance % (SV%)" & vbCrLf
      strMsg = strMsg & "SV% = ( SV / BCWS ) * 100" & vbCrLf
      strMsg = strMsg & "SV% = ( " & Format((dblBCWP - dblBCWS), "#,##0.0h") & " / " & Format(dblBCWS, "#,##0.0h") & " ) * 100" & vbCrLf
      strMsg = strMsg & "SV% = " & Format(((dblBCWP - dblBCWS) / dblBCWS), "0.00%")
      
      MsgBox strMsg, vbInformation + vbOKOnly, "Schedule Variance (SV) - Hours"
      
    Case "es" 'earned schedule
          'todo: earned schedule
    
  End Select
  
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptMerics_Bas", "cptGet", err, Erl)
  Resume exit_here
End Sub

Sub cptGetHitTask()
'objects
Dim Task As Task
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
  Else
    dtStatus = ActiveProject.StatusDate
  End If
  
  'find it
  For Each Task In ActiveProject.Tasks
    If Not Task Is Nothing Then
      If Task.Summary Then GoTo next_task
      If Not Task.Active Then GoTo next_task
      If IsDate(Task.BaselineFinish) Then
        'was task baselined to finish before status date?
        If Task.BaselineFinish <= dtStatus Then
          lngBLF = lngBLF + 1
          'did it?
          If IsDate(Task.ActualFinish) Then
            If Task.ActualFinish <= Task.BaselineFinish Then
              lngAF = lngAF + 1
            End If
          End If
        End If
      End If
    End If
next_task:
  Next

  strMsg = "BF = # Tasks Baselined to Finish ON or before Status Date" & vbCrLf
  strMsg = strMsg & "AF = # BF that Actually Finished ON or before Baseline Finish" & vbCrLf & vbCrLf
  strMsg = strMsg & "Hit Task % = (AF / BF) / 100" & vbCrLf
  strMsg = strMsg & "Hit Task % = (" & Format(lngAF, "#,##0") & " / " & Format(lngBLF, "#,##0") & ") / 100" & vbCrLf & vbCrLf
  strMsg = strMsg & "Hit Task % = " & Format((lngAF / lngBLF), "0%")
  MsgBox strMsg, vbInformation + vbOKOnly, "Hit Task %"

exit_here:
  On Error Resume Next
  Set Task = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptMetrics_bas", "cptGetHitTask", err, Erl)
  Resume exit_here
End Sub

Function cptGetMetric(strGet As String) As Double
'todo: no screen changes!
'objects
Dim TSV As Object 'TimeScaleValue
Dim TSVS As Object 'TimeScaleValues
Dim Tasks As Object 'Tasks
Dim Task As Object 'Task
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
  Else
    dtStatus = ActiveProject.StatusDate
  End If
  
  cptSpeed True
  Call cptResetAll
  SelectAll
  Set Tasks = ActiveSelection.Tasks
  For Each Task In Tasks
    If Not Task Is Nothing Then
      If Task.ExternalTask Then GoTo next_task
      If Task.Summary Then GoTo next_task
      If Not Task.Active Then GoTo next_task
      If Task.BaselineWork > 0 Then
        Select Case strGet
          Case "bac"
            dblResult = dblResult + (Task.BaselineWork / 60)
            
          Case "etc"
            dblResult = dblResult + (Task.RemainingWork / 60)
            
          Case "bcws"
            If Task.Start < dtStatus Then
              Set TSVS = Task.TimeScaleData(Task.Start, dtStatus, pjTaskTimescaledBaselineWork, pjTimescaleWeeks)
              For Each TSV In TSVS
                dblResult = dblResult + IIf(TSV.Value = "", 0, TSV.Value) / 60
              Next
            End If
            
          Case "bcwp"
            'todo: user has not identified where ev% is kept
            dblResult = dblResult + ((Task.BaselineWork / 60) * (Task.PhysicalPercentComplete / 100))
            
          Case "bei_bf"
            dblResult = dblResult + IIf(Task.BaselineFinish <= dtStatus, 1, 0)
            If Task.BaselineFinish <= dtStatus Then Task.Text23 = "BF"
            
          Case "bei_af"
            dblResult = dblResult + IIf(Task.ActualFinish <= dtStatus, 1, 0)
            If Task.ActualFinish <= dtStatus Then Task.Text24 = "AF"
            
        End Select
      End If 'bac>0
    End If 'not nothing
next_task:
    Application.StatusBar = "Calculating " & UCase(strGet) & "..."
  Next

  cptGetMetric = dblResult

exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  cptSpeed False
  Set TSV = Nothing
  Set TSVS = Nothing
  Set Tasks = Nothing
  Set Task = Nothing

  Exit Function
err_here:
  'Debug.Print Task.UniqueID & ": " & Task.Name
  Call cptHandleErr("cptMetrics_bas", "cptGetMetric", err, Erl)
  Resume exit_here

End Function
